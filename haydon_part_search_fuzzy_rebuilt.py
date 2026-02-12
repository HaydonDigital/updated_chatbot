import os
from pathlib import Path
import re
import pandas as pd
import streamlit as st

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(layout="wide", page_title="Haydon Cross-Reference Search")

BASE_DIR = Path(__file__).parent
CROSS_FILE = BASE_DIR / "Updated File - 7-10-25.xlsx"
IMAGES_FILE = BASE_DIR / "Image and Submittals.xlsx"
MATERIAL_DESC_FILE = BASE_DIR / "Material Description.xlsx"

# Optional PDF support (text-based PDFs)
try:
    from PyPDF2 import PdfReader
    HAS_PYPDF2 = True
except Exception:
    HAS_PYPDF2 = False


# =========================================================
# UTILITY FUNCTIONS
# =========================================================
def normalize(text):
    """Strip non-alphanumerics and lowercase for fuzzy-ish matching."""
    if pd.isna(text):
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(text)).lower()


# =========================================================
# LONG-DESCRIPTION PARSING (for Material Description.xlsx)
# =========================================================
UNICODE_FRACTIONS = {
    "¼": "1/4", "½": "1/2", "¾": "3/4",
    "⅛": "1/8", "⅜": "3/8", "⅝": "5/8", "⅞": "7/8",
    "⅓": "1/3", "⅔": "2/3",
}


def normalize_desc_text(s: str) -> str:
    s = (s or "").strip()
    for k, v in UNICODE_FRACTIONS.items():
        s = s.replace(k, v)
    # normalize curly quotes to inch mark
    s = s.replace("″", '"').replace("”", '"').replace("“", '"').replace("’", "'")
    s = re.sub(r"\bw/\b", "with", s, flags=re.I)
    s = re.sub(r"\s+", " ", s)
    return s


def frac_to_float(frac: str):
    """Supports 1-5/8, 1 5/8, 5/8, 1"""
    frac = (frac or "").strip()

    m = re.fullmatch(r"(\d+)[-\s]+(\d+)\s*/\s*(\d+)", frac)
    if m:
        whole, num, den = map(int, m.groups())
        return whole + (num / den)

    m = re.fullmatch(r"(\d+)\s*/\s*(\d+)", frac)
    if m:
        num, den = map(int, m.groups())
        return num / den

    m = re.fullmatch(r"\d+", frac)
    if m:
        return float(frac)

    return None


def float_to_inches_str(x: float) -> str:
    """Best-effort normalize numeric inches back to the common text in the file (e.g., 1.625 -> 1-5/8")."""
    if x is None:
        return ""
    # common eighths for strut dimensions
    eighths = round(x * 8)
    whole = eighths // 8
    rem = eighths % 8
    if rem == 0:
        return f"{whole}\""
    # reduce fraction rem/8
    from math import gcd
    g = gcd(rem, 8)
    num, den = rem // g, 8 // g
    if whole == 0:
        return f"{num}/{den}\""
    return f"{whole}-{num}/{den}\""


def extract_width_in(desc: str):
    s = normalize_desc_text(desc)
    m = re.search(r"(?P<w>\d+(?:[-\s]+\d+\s*/\s*\d+|\s*/\s*\d+)?)\s*(?:\"|in\b)", s, flags=re.I)
    if not m:
        return None
    return frac_to_float(m.group("w"))


def extract_gauge(desc: str):
    s = normalize_desc_text(desc)
    m = re.search(r"\b(\d{1,2})\s*(?:ga\.?|gauge)\b", s, flags=re.I)
    return int(m.group(1)) if m else None


def extract_length(desc: str):
    s = normalize_desc_text(desc)
    # feet: 10', 10 ft, 10ft
    m = re.search(r"\b(\d{1,2})\s*(?:'|ft\b)", s, flags=re.I)
    if m:
        return f"{int(m.group(1))}'"
    # inches: 120", 12"
    m = re.search(r"\b(\d{1,3})\s*\"\b", s)
    if m:
        return f"{int(m.group(1))}\""
    return None


def extract_finish(desc: str):
    s = normalize_desc_text(desc).upper()
    if "PRE" in s and "GALV" in s:
        return "Pre-Galvanized"
    if "HOT" in s and "DIP" in s:
        return "Hot Dipped Galvanized"
    if "HDG" in s:
        return "Hot Dipped Galvanized"
    if "PLAIN" in s and "STEEL" in s:
        return "Plain Steel"
    if "STAINLESS" in s and "316" in s:
        return "Stainless Steel 316"
    if "STAINLESS" in s and "304" in s:
        return "Stainless Steel 304"
    if "ALUMIN" in s:
        return "Aluminum"
    return None


def is_slotted(desc: str) -> bool:
    s = normalize_desc_text(desc).upper()
    return ("SLOT" in s) or ("OPEN SLOT" in s) or ("OS" in s and "HOLE" in s)


def looks_like_long_description(q: str) -> bool:
    """Heuristic: treat as long description if it has spaces and dimension/gauge-ish tokens."""
    if not q:
        return False
    s = str(q)
    if len(s) < 10:
        return False
    if " " not in s:
        return False
    s2 = s.lower()
    return any(t in s2 for t in ['gauge', 'ga', 'galv', 'slotted', 'channel', 'steel', 'stainless', 'aluminum', '"', "'"])

def split_parts_from_text(raw: str) -> list[str]:
    """Split pasted text into part tokens (newline/comma/space separated)."""
    if not raw:
        return []
    tokens = re.split(r"[\n,;\t ]+", raw.strip())
    return [t.strip() for t in tokens if t.strip()]


def extract_text_from_pdf(uploaded_file) -> str:
    """Extract text from a PDF (works for text-based PDFs, not scanned images)."""
    if not HAS_PYPDF2:
        return ""
    reader = PdfReader(uploaded_file)
    chunks = []
    for page in reader.pages:
        chunks.append(page.extract_text() or "")
    return "\n".join(chunks)


def extract_part_candidates_from_pdf_text(text: str) -> list[str]:
    """
    Pull likely part-number-like strings from PDF text.
    Tune regex as needed to reduce noise.
    """
    if not text:
        return []

    # General "part-like" token: letters/numbers with -, /, ., # allowed
    pattern = re.compile(r"\b[A-Z0-9][A-Z0-9\-\/\.\#]{3,39}\b", re.IGNORECASE)
    hits = pattern.findall(text)

    # Clean + dedupe preserving order
    seen = set()
    out = []
    for h in hits:
        h2 = h.strip().strip(".,;:()[]{}")
        if len(h2) < 4:
            continue
        key = h2.upper()
        if key not in seen:
            seen.add(key)
            out.append(h2)
    return out


def dedupe_preserve_order(items: list[str]) -> list[str]:
    seen = set()
    out = []
    for x in items:
        k = (x or "").strip().upper()
        if k and k not in seen:
            seen.add(k)
            out.append((x or "").strip())
    return out


def get_haydon_candidates(part):
    """
    Given a Haydon part like 'H-1234-XG', progressively shorten it
    so we can try to match in the image/submittal file.
    """
    part = str(part).upper()
    tokens = re.split(r"[ \-X()]+", part)
    for i in range(len(tokens), 0, -1):
        yield "-".join(tokens[:i])


# =========================================================
# DATA LOADERS (CACHED)
# =========================================================
@st.cache_data
def load_cross_reference():
    if not CROSS_FILE.exists():
        raise FileNotFoundError(f"Cross-reference file not found at: {CROSS_FILE}")
    df = pd.read_excel(CROSS_FILE, sheet_name="Export", engine="openpyxl")
    df["Normalized Haydon Part"] = df["Haydon Part Description"].apply(normalize)
    df["Normalized Vendor Part"] = df["Vendor Part #"].apply(normalize)
    return df


@st.cache_data
def load_images():
    if not IMAGES_FILE.exists():
        raise FileNotFoundError(f"Image/submittals file not found at: {IMAGES_FILE}")
    df = pd.read_excel(IMAGES_FILE, sheet_name="Sheet1")
    if "Name" in df.columns:
        df["Name_upper"] = df["Name"].astype(str).str.upper()
    return df


@st.cache_data
def load_material_descriptions():
    """Haydon material master export used for long-description matching."""
    if not MATERIAL_DESC_FILE.exists():
        raise FileNotFoundError(f"Material Description file not found at: {MATERIAL_DESC_FILE}")
    df = pd.read_excel(MATERIAL_DESC_FILE, engine="openpyxl")

    # Normalized helpers for fast filtering
    if "Material Description" in df.columns:
        df["_desc_norm"] = df["Material Description"].astype(str).map(normalize_desc_text).str.upper()
    else:
        df["_desc_norm"] = ""

    for c in ["Width", "Height", "Length", "Gauge", "Material Group 1 Name", "Material Group Name", "Material Group 3 Name"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    return df


# =========================================================
# SEARCH
# =========================================================
def search_parts_contains(cross_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Contains search (your current behavior)."""
    norm_query = normalize(query)
    return cross_df[
        cross_df["Normalized Haydon Part"].str.contains(norm_query, na=False)
        | cross_df["Normalized Vendor Part"].str.contains(norm_query, na=False)
    ]


def search_parts_exact(cross_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Exact match on normalized values (recommended for bulk)."""
    norm_query = normalize(query)
    if not norm_query:
        return cross_df.iloc[0:0]
    return cross_df[
        (cross_df["Normalized Haydon Part"] == norm_query)
        | (cross_df["Normalized Vendor Part"] == norm_query)
    ]


def match_long_description(material_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Match a competitor long description to Haydon materials using structured attributes + scoring."""
    q = normalize_desc_text(query)
    q_up = q.upper()

    width_in = extract_width_in(q)
    gauge = extract_gauge(q)
    finish = extract_finish(q)
    length = extract_length(q)
    slotted = is_slotted(q)

    # Start with strut-ish only if we can
    df = material_df
    if "Material Group Name" in df.columns:
        df = df[df["Material Group Name"].astype(str).str.contains("Strut", case=False, na=False)]

    # Fast filters where available
    if width_in is not None and "Width" in df.columns:
        width_text = float_to_inches_str(width_in)
        # file stores width like 1-5/8"
        df = df[df["Width"].astype(str).str.replace(" ", "").str.contains(width_text.replace(" ", ""), na=False)]

    if gauge is not None and "Gauge" in df.columns:
        df = df[df["Gauge"].astype(str).str.contains(str(gauge), na=False)]

    if finish and "Material Group 1 Name" in df.columns:
        df = df[df["Material Group 1 Name"].astype(str).str.contains(finish, case=False, na=False)]

    if length and "Length" in df.columns:
        # only filter if explicitly present
        df = df[df["Length"].astype(str).str.strip().eq(length)]

    if slotted and "Material Group 3 Name" in df.columns:
        df = df[df["Material Group 3 Name"].astype(str).str.contains("SLOT", case=False, na=False)]

    if df.empty:
        return df

    # Score candidates (helps when length is missing or multiple variants exist)
    def score_row(r) -> int:
        s = 0
        if width_in is not None and "Width" in r and str(r["Width"]).strip() != "nan":
            if float_to_inches_str(width_in).replace(" ", "") in str(r["Width"]).replace(" ", ""):
                s += 3
        if gauge is not None and "Gauge" in r and str(gauge) in str(r["Gauge"]):
            s += 2
        if finish and "Material Group 1 Name" in r and finish.lower() in str(r["Material Group 1 Name"]).lower():
            s += 2
        if slotted and "Material Group 3 Name" in r and "SLOT" in str(r["Material Group 3 Name"]).upper():
            s += 2
        if "CHANNEL" in q_up and "Material Description" in r and "H-132" in str(r["Material Description"]).upper():
            # shallow channel hint
            s += 1
        return s

    scored = df.copy()
    scored["_score"] = scored.apply(score_row, axis=1)
    scored = scored.sort_values(by=["_score"], ascending=False)
    return scored


def bulk_search(cross_df: pd.DataFrame, material_df: pd.DataFrame, parts: list[str], allow_contains_fallback: bool) -> pd.DataFrame:
    """
    Flattened output: one row per match (or one row per input if not found).
    """
    rows = []
    for raw in parts:
        raw = (raw or "").strip()
        if not raw:
            continue

        found = search_parts_exact(cross_df, raw)
        if found.empty and allow_contains_fallback:
            found = search_parts_contains(cross_df, raw)

        if found.empty:
            # If it looks like a long description, try matching against Material Description.xlsx
            if material_df is not None and looks_like_long_description(raw):
                m = match_long_description(material_df, raw)
                if not m.empty:
                    top = m.head(10)
                    for _, r in top.iterrows():
                        rows.append(
                            {
                                "Input": raw,
                                "Status": "Found (Material)",
                                "Vendor Part #": None,
                                "Vendor": None,
                                "Category": r.get("Material Group Name"),
                                "Haydon Part Description": r.get("Material Description"),
                                "Haydon Material": r.get("Material"),
                                "Weight Each": r.get("Weight Each"),
                                "Match Count (for Input)": len(top),
                            }
                        )
                    continue

            rows.append(
                {
                    "Input": raw,
                    "Status": "Not Found",
                    "Vendor Part #": None,
                    "Vendor": None,
                    "Category": None,
                    "Haydon Part Description": None,
                    "Haydon Material": None,
                    "Weight Each": None,
                    "Match Count (for Input)": 0,
                }
            )
        else:
            for _, r in found.iterrows():
                rows.append(
                    {
                        "Input": raw,
                        "Status": "Found",
                        "Vendor Part #": r.get("Vendor Part #"),
                        "Vendor": r.get("Vendor"),
                        "Category": r.get("Category"),
                        "Haydon Part Description": r.get("Haydon Part Description"),
                        "Haydon Material": None,
                        "Weight Each": None,
                        "Match Count (for Input)": len(found),
                    }
                )
    return pd.DataFrame(rows)


# =========================================================
# APP UI
# =========================================================
st.title("Haydon Cross-Reference Search")

try:
    cross_df = load_cross_reference()
    image_df = load_images()
    material_df = load_material_descriptions()
except FileNotFoundError as e:
    st.error(str(e))
    st.stop()

tab_single, tab_bulk = st.tabs(["Single", "Bulk"])

# =========================================================
# SINGLE
# =========================================================
with tab_single:
    query = st.text_input("Enter part number (Haydon or Vendor):")

    if query:
        used_material = looks_like_long_description(query)
        if used_material:
            m = match_long_description(material_df, query)
            if not m.empty:
                st.subheader(f"Material matches ({min(len(m), 25)} shown)")
                show_cols = ["Material", "Material Description", "Material Group 1 Name", "Weight Each", "Gauge", "Height", "Width", "Length"]
                show = m[[c for c in show_cols if c in m.columns]].head(25)
                st.dataframe(show, use_container_width=True)
            else:
                st.error("No material match found for that description.")
            results = None
        else:
            results = search_parts_contains(cross_df, query)

        if results is not None and not results.empty:
            display_cols = ["Vendor Part #", "Vendor", "Category", "Haydon Part Description"]
            display_df = results[[c for c in display_cols if c in results.columns]]

            st.subheader(f"Found {len(display_df)} matching entries")
            st.dataframe(display_df, use_container_width=True)

            # Sidebar: product preview / submittals for first result
            first_row = results.iloc[0]
            haydon_part = first_row["Haydon Part Description"]

            with st.sidebar:
                st.markdown("### Haydon Product Preview")
                match_found = False

                candidates = [haydon_part] + list(get_haydon_candidates(haydon_part))
                for candidate in candidates:
                    matched = image_df[image_df["Name_upper"] == str(candidate).upper()]
                    if not matched.empty:
                        row = matched.iloc[0]

                        if "Cover Image" in row and pd.notna(row["Cover Image"]):
                            st.image(
                                row["Cover Image"],
                                caption=row.get("Name", candidate),
                                use_container_width=True,
                            )

                        if "Files" in row and pd.notna(row["Files"]):
                            st.markdown(f"[View Submittal]({row['Files']})", unsafe_allow_html=True)

                        match_found = True
                        break

                if not match_found:
                    st.warning("No product preview or submittal found.")
        elif results is not None:
            st.error(
                "No match found. Send the Haydon part number and the customer/competitor part number to "
                "[marketing@haydoncorp.com](mailto:marketing@haydoncorp.com)."
            )
    else:
        st.write("Enter a part number above to begin.")

# =========================================================
# BULK
# =========================================================
with tab_bulk:
    st.markdown("Paste a list of parts and/or upload a PDF to cross-reference in bulk.")

    col1, col2 = st.columns(2)

    with col1:
        pasted = st.text_area(
            "Paste part numbers (one per line, or comma/space separated):",
            height=220,
            placeholder="Example:\nPS-1100-AS-4-EG\nTSN-802\nP1000",
        )

    with col2:
        pdf = st.file_uploader("Upload a PDF (text-based PDFs):", type=["pdf"])
        allow_contains = st.checkbox(
            "If exact match fails, try contains fallback (may increase false matches)",
            value=True,
        )
        if pdf is not None and not HAS_PYPDF2:
            st.warning("PyPDF2 is not installed in this environment. Add it to requirements.txt to enable PDF parsing.")

    parts = []
    if pasted and pasted.strip():
        parts.extend(split_parts_from_text(pasted))

    if pdf is not None and HAS_PYPDF2:
        pdf_text = extract_text_from_pdf(pdf)
        if not pdf_text.strip():
            st.warning("Could not extract text from this PDF (it may be scanned or protected).")
        else:
            candidates = extract_part_candidates_from_pdf_text(pdf_text)
            st.info(f"Extracted {len(candidates)} candidate tokens from the PDF.")
            with st.expander("View extracted candidates"):
                st.write(candidates)
            parts.extend(candidates)

    parts_unique = dedupe_preserve_order(parts)

    st.write(f"Inputs ready: {len(parts_unique)}")

    run = st.button("Run Bulk Cross-Reference", type="primary", disabled=(len(parts_unique) == 0))

    if run:
        out_df = bulk_search(cross_df, material_df, parts_unique, allow_contains_fallback=allow_contains)

        st.subheader(f"Bulk Results ({len(out_df)} rows)")
        st.dataframe(out_df, use_container_width=True)

        found_rows = int((out_df["Status"] == "Found").sum())
        not_found_rows = int((out_df["Status"] == "Not Found").sum())
        st.write(f"Found rows: {found_rows} | Not found rows: {not_found_rows}")

        csv_bytes = out_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download results as CSV",
            data=csv_bytes,
            file_name="haydon_cross_reference_bulk_results.csv",
            mime="text/csv",
        )
