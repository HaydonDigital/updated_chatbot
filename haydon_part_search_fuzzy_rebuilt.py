import os
from pathlib import Path
import re
from math import gcd
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

# =========================================================
# NAMING-CONVENTION AWARE NORMALIZATION
#   - key_full: keeps finish/slot tokens
#   - key_base: strips finish-ish tokens so variants still match
# =========================================================
FINISH_TOKENS = {
    "HDG", "H.D.G", "HOTDIP", "HOT-DIP", "HOTDIPPED", "HOTDIPGALV", "HOT-DIPGALV",
    "EG", "E.G", "ELECTRO", "ELECTROGALV", "ELECTRO-GALV", "ZINC", "ZN",
    "PG", "PREGALV", "PRE-GALV", "PREGALVANIZED", "PRE-GALVANIZED",
    "SS", "SST", "STAINLESS", "AL", "ALUM", "ALUMINUM", "PLAIN", "STEEL"
}

def normalize_common(s: str) -> str:
    """Normalize punctuation/spacing so tokens are consistent for keying/matching."""
    if pd.isna(s) or s is None:
        return ""
    s = str(s).strip().upper()

    # normalize curly quotes
    s = s.replace("″", '"').replace("”", '"').replace("“", '"').replace("’", "'")

    # normalize separators
    s = s.replace("\\", "/")
    s = re.sub(r"\s+", " ", s)

    # normalize X separators: " x " -> " X "
    s = re.sub(r"\s*[xX]\s*", " X ", s)

    # normalize feet marks: 10 FT, 10FT -> 10'
    s = re.sub(r"\b(\d{1,2})\s*(?:FT|FEET)\b", r"\1'", s)

    # normalize common vendor shorthand
    s = re.sub(r"\bw/\b", "WITH", s, flags=re.I)

    return s

def key_full(s: str) -> str:
    """Full key: keep important tokens, strip punctuation/spaces."""
    s = normalize_common(s)
    return re.sub(r"[^A-Z0-9]", "", s)

def key_base(s: str) -> str:
    """
    Base key: remove finish-ish tokens so variants match:
      H-112 X 10' <-> H-112-HDG X 10
      P1000 SL 10 <-> P1000 SL 10PG
    """
    s = normalize_common(s)
    tokens = re.split(r"[ \-_/()]+", s)
    tokens = [t for t in tokens if t]

    cleaned = []
    for t in tokens:
        t2 = t.replace(".", "")
        if t2 in FINISH_TOKENS:
            continue
        cleaned.append(t)

    return re.sub(r"[^A-Z0-9]", "", " ".join(cleaned))

# =========================================================
# ORDER-INSENSITIVE TOKEN MATCHING
#   Fixes: "1/4 flat washer" vs "FLAT WASHER 1/4"
#   AND:   tokens spread across Category + Vendor + Haydon
# =========================================================
def tokenize_for_search(s: str) -> list[str]:
    """
    Tokenize text for order-insensitive matching.
    Keeps fractions like 1/4 as '1/4', and also adds '14' as a fallback token.
    """
    s = normalize_common(s)
    tokens = re.findall(r"[A-Z0-9]+(?:/[0-9]+)?", s)

    out = []
    for t in tokens:
        out.append(t)
        if "/" in t:
            out.append(t.replace("/", ""))  # 1/4 -> 14
    out = [t for t in out if len(t) >= 2]
    return out

def contains_all_tokens(field_value: str, tokens: list[str]) -> bool:
    """
    Order-insensitive "contains" check on a single combined field.
    """
    hay = normalize_common(field_value)
    hay_compact = re.sub(r"[^A-Z0-9/]", "", hay)

    for tok in tokens:
        tok_compact = re.sub(r"[^A-Z0-9/]", "", str(tok).upper())
        if tok_compact and tok_compact not in hay_compact:
            return False
    return True

# =========================================================
# LONG-DESCRIPTION PARSING (Material Description.xlsx)
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
    """Normalize numeric inches back to common text (e.g., 1.625 -> 1-5/8")."""
    if x is None:
        return ""
    eighths = round(x * 8)
    whole = eighths // 8
    rem = eighths % 8
    if rem == 0:
        return f"{whole}\""
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

    m = re.search(r"\b(\d{1,2})\s*(?:'|ft\b)", s, flags=re.I)
    if m:
        return f"{int(m.group(1))}'"

    m = re.search(r"\b(\d{1,3})\s*\"\b", s)
    if m:
        return f"{int(m.group(1))}\""

    return None

def extract_finish(desc: str):
    s = normalize_desc_text(desc).upper()

    if ("PRE" in s and "GALV" in s) or ("PREGALV" in s) or ("PRE-GALV" in s) or re.search(r"\bPG\b", s):
        return "Pre-Galvanized"
    if ("HOT" in s and "DIP" in s) or re.search(r"\bHDG\b", s):
        return "Hot Dipped Galvanized"
    if "ELECTRO" in s or re.search(r"\bEG\b", s):
        return "Electro Galvanized"
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
    return ("SLOT" in s) or ("OPEN SLOT" in s) or ("SLOTTED" in s) or ("SLOTTED HOLES" in s)

def looks_like_long_description(q: str) -> bool:
    """
    Route to Material Description matcher only when it really looks like
    a strut/channel long description (avoid routing washer/category phrases).
    """
    if not q:
        return False
    s = str(q).strip()
    if len(s) < 14:
        return False
    if " " not in s:
        return False
    s2 = s.lower()
    return any(t in s2 for t in [
        'gauge', 'ga', 'galv', 'pre-galv', 'pregalv', 'hdg', 'hot dip',
        'slotted', 'channel', 'strut', 'steel', 'stainless', 'aluminum'
    ]) and (('"') in s or ("'") in s or (" in " in s2) or (" x " in s2))

# =========================================================
# BULK INPUT HELPERS
# =========================================================
def split_parts_from_text(raw: str) -> list[str]:
    """
    Split pasted text into inputs.
    IMPORTANT: keep spaces inside long descriptions, so only split on newline/comma/semicolon/tab.
    """
    if not raw:
        return []
    tokens = re.split(r"[\n,;\t]+", raw.strip())
    return [t.strip() for t in tokens if t and t.strip()]

def dedupe_preserve_order(items: list[str]) -> list[str]:
    seen = set()
    out = []
    for x in items:
        k = (x or "").strip().upper()
        if k and k not in seen:
            seen.add(k)
            out.append((x or "").strip())
    return out

# =========================================================
# IMAGE/SUBMITTAL CANDIDATES (naming convention tolerant)
# =========================================================
def get_haydon_candidates(part: str):
    """Generate likely lookup variants for image/submittal Name matching."""
    p = normalize_common(part)
    variants = set()

    variants.add(p)
    variants.add(p.replace(" X ", "X"))
    variants.add(p.replace("-", " - "))
    variants.add(p.replace(" - ", "-"))

    # base variant without finish tokens
    tokens = re.split(r"[ \-_/()]+", p)
    tokens = [t for t in tokens if t and t.replace(".", "") not in FINISH_TOKENS]
    variants.add(" ".join(tokens))

    for v in sorted(variants, key=len, reverse=True):
        yield v

# =========================================================
# DATA LOADERS (CACHED)
# =========================================================
@st.cache_data
def load_cross_reference():
    if not CROSS_FILE.exists():
        raise FileNotFoundError(f"Cross-reference file not found at: {CROSS_FILE}")

    df = pd.read_excel(CROSS_FILE, sheet_name="Export", engine="openpyxl")

    # Robust keys for matching (Haydon, Vendor, and Category)
    df["HaydonKeyFull"] = df["Haydon Part Description"].apply(key_full)
    df["HaydonKeyBase"] = df["Haydon Part Description"].apply(key_base)

    df["VendorKeyFull"] = df["Vendor Part #"].apply(key_full)
    df["VendorKeyBase"] = df["Vendor Part #"].apply(key_base)

    df["CategoryKeyFull"] = df["Category"].apply(key_full)
    df["CategoryKeyBase"] = df["Category"].apply(key_base)

    # Normalized raw strings used for token matching
    df["HaydonNorm"] = df["Haydon Part Description"].apply(normalize_common)
    df["VendorNorm"] = df["Vendor Part #"].apply(normalize_common)
    df["CategoryNorm"] = df["Category"].apply(normalize_common)

    # ✅ Combined field so tokens can match across Category + Vendor + Haydon
    df["CombinedNorm"] = (
        df["CategoryNorm"].fillna("") + " " +
        df["VendorNorm"].fillna("") + " " +
        df["HaydonNorm"].fillna("")
    ).str.strip()

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

    if "Material Description" in df.columns:
        df["_desc_norm"] = df["Material Description"].astype(str).map(normalize_desc_text).str.upper()
    else:
        df["_desc_norm"] = ""

    for c in [
        "Width", "Height", "Length", "Gauge",
        "Material Group 1 Name", "Material Group Name", "Material Group 3 Name",
        "Weight Each"
    ]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    return df

# =========================================================
# SEARCH (cross file)
# =========================================================
def search_parts_exact(cross_df: pd.DataFrame, query: str) -> pd.DataFrame:
    qf = key_full(query)
    qb = key_base(query)
    if not qf and not qb:
        return cross_df.iloc[0:0]

    return cross_df[
        (cross_df["HaydonKeyFull"].eq(qf)) |
        (cross_df["VendorKeyFull"].eq(qf)) |
        (cross_df["CategoryKeyFull"].eq(qf)) |
        (cross_df["HaydonKeyBase"].eq(qb)) |
        (cross_df["VendorKeyBase"].eq(qb)) |
        (cross_df["CategoryKeyBase"].eq(qb))
    ]

def search_parts_contains(cross_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """
    Contains search that supports:
      - key-based contains (fast)
      - AND token-based contains (order-insensitive, across combined field)
    """
    q = (query or "").strip()
    if not q:
        return cross_df.iloc[0:0]

    qf = key_full(q)
    qb = key_base(q)

    # 1) Fast key-based contains
    fast = cross_df[
        cross_df["HaydonKeyFull"].str.contains(qf, na=False) |
        cross_df["VendorKeyFull"].str.contains(qf, na=False) |
        cross_df["CategoryKeyFull"].str.contains(qf, na=False) |
        cross_df["HaydonKeyBase"].str.contains(qb, na=False) |
        cross_df["VendorKeyBase"].str.contains(qb, na=False) |
        cross_df["CategoryKeyBase"].str.contains(qb, na=False)
    ]
    if not fast.empty:
        return fast

    # 2) Token-based order-insensitive contains across Category+Vendor+Haydon
    tokens = tokenize_for_search(q)
    if not tokens:
        return cross_df.iloc[0:0]

    mask = cross_df["CombinedNorm"].apply(lambda v: contains_all_tokens(v, tokens))
    return cross_df[mask]

# =========================================================
# SEARCH (material file / long description)
# =========================================================
def match_long_description(material_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """
    Match a competitor long description to Haydon materials using
    structured attributes + internal ranking (no score column returned).
    """
    q = normalize_desc_text(query)
    q_up = q.upper()

    width_in = extract_width_in(q)
    gauge = extract_gauge(q)
    finish = extract_finish(q)
    length = extract_length(q)
    slotted = is_slotted(q)

    df = material_df

    if "Material Group Name" in df.columns:
        df = df[df["Material Group Name"].astype(str).str.contains("STRUT|CHANNEL|FRAM", case=False, na=False)]

    if width_in is not None and "Width" in df.columns:
        width_text = float_to_inches_str(width_in).replace(" ", "")
        df = df[df["Width"].astype(str).str.replace(" ", "").str.contains(width_text, na=False)]

    if gauge is not None and "Gauge" in df.columns:
        df = df[df["Gauge"].astype(str).str.contains(str(gauge), na=False)]

    if finish and "Material Group 1 Name" in df.columns:
        df = df[df["Material Group 1 Name"].astype(str).str.contains(finish, case=False, na=False)]

    if length and "Length" in df.columns:
        df = df[df["Length"].astype(str).str.strip().eq(length)]

    if slotted and "Material Group 3 Name" in df.columns:
        df = df[df["Material Group 3 Name"].astype(str).str.contains("SLOT", case=False, na=False)]

    if df.empty:
        return df

    def score_row(r) -> int:
        s = 0

        if width_in is not None and "Width" in r and str(r.get("Width", "")).strip().lower() != "nan":
            if float_to_inches_str(width_in).replace(" ", "") in str(r["Width"]).replace(" ", ""):
                s += 3

        if gauge is not None and "Gauge" in r and str(gauge) in str(r.get("Gauge", "")):
            s += 2

        if finish and "Material Group 1 Name" in r and finish.lower() in str(r.get("Material Group 1 Name", "")).lower():
            s += 2

        if slotted and "Material Group 3 Name" in r and "SLOT" in str(r.get("Material Group 3 Name", "")).upper():
            s += 2

        if ("CHANNEL" in q_up or "W CHANNEL" in q_up) and "Material Description" in r:
            md = str(r.get("Material Description", "")).upper()
            if "H-132" in md:
                s += 1

        return s

    ranked = df.copy()
    ranked["_tmp_rank"] = ranked.apply(score_row, axis=1)
    ranked = ranked.sort_values(by=["_tmp_rank"], ascending=False).drop(columns=["_tmp_rank"])
    return ranked

# =========================================================
# BULK SEARCH
# =========================================================
def bulk_search(cross_df: pd.DataFrame, material_df: pd.DataFrame, parts: list[str], allow_contains_fallback: bool) -> pd.DataFrame:
    """
    Flattened output: one row per match (or one row per input if not found).
    Includes Weight Each for Material matches (from Material Description.xlsx).
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
    query = st.text_input("Enter part number (Haydon/Vendor), Category phrase, OR paste a long description:")

    if query:
        used_material = looks_like_long_description(query)

        if used_material:
            m = match_long_description(material_df, query)
            if not m.empty:
                st.subheader(f"Material matches ({min(len(m), 25)} shown)")
                show_cols = [
                    "Material",
                    "Material Description",
                    "Material Group Name",
                    "Material Group 1 Name",
                    "Material Group 3 Name",
                    "Weight Each",
                    "Gauge",
                    "Height",
                    "Width",
                    "Length",
                ]
                show = m[[c for c in show_cols if c in m.columns]].head(25)
                st.dataframe(show, use_container_width=True)
            else:
                st.error("No material match found for that description.")
        else:
            # Exact first, then contains (contains includes token-based order-insensitive)
            results = search_parts_exact(cross_df, query)
            if results.empty:
                results = search_parts_contains(cross_df, query)

            if not results.empty:
                display_cols = ["Vendor Part #", "Vendor", "Category", "Haydon Part Description"]
                display_df = results[[c for c in display_cols if c in results.columns]]

                st.subheader(f"Found {len(display_df)} matching entries")
                st.dataframe(display_df, use_container_width=True)

                first_row = results.iloc[0]
                haydon_part = first_row.get("Haydon Part Description", "")

                with st.sidebar:
                    st.markdown("### Haydon Product Preview")
                    match_found = False

                    for candidate in get_haydon_candidates(haydon_part):
                        matched = image_df[image_df.get("Name_upper", "").astype(str) == str(candidate).upper()]
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
            else:
                st.error(
                    "No match found. Send the Haydon part number and the customer/competitor part number to "
                    "[marketing@haydoncorp.com](mailto:marketing@haydoncorp.com)."
                )
    else:
        st.write("Enter a part number, category phrase, or long description above to begin.")

# =========================================================
# BULK (no PDF upload)
# =========================================================
with tab_bulk:
    st.markdown(
        "Paste a list of inputs (one per line). "
        "Tip: keep long descriptions on a single line."
    )

    col1, col2 = st.columns(2)

    with col1:
        pasted = st.text_area(
            "Paste inputs (one per line):",
            height=220,
            placeholder=(
                "Example:\n"
                "P1000 SL 10PG\n"
                "H-112 X 10' HDG\n"
                "SQUARE WASHER - 1 5/8 in x 1 5/8 in\n"
                "1/4\" SPRING NUTS\n"
                "1/4\" flat washer\n"
                "1-5/8\" W Channel w/ Slotted Holes - Steel Pre-Galvanized 12 Gauge"
            ),
        )

    with col2:
        allow_contains = st.checkbox(
            "If exact match fails, try contains fallback (may increase false matches)",
            value=True,
        )

    parts = []
    if pasted and pasted.strip():
        parts.extend(split_parts_from_text(pasted))

    parts_unique = dedupe_preserve_order(parts)

    st.write(f"Inputs ready: {len(parts_unique)}")

    run = st.button("Run Bulk Cross-Reference", type="primary", disabled=(len(parts_unique) == 0))

    if run:
        out_df = bulk_search(cross_df, material_df, parts_unique, allow_contains_fallback=allow_contains)

        st.subheader(f"Bulk Results ({len(out_df)} rows)")
        st.dataframe(out_df, use_container_width=True)

        found_rows = int((out_df["Status"] == "Found").sum())
        found_material_rows = int((out_df["Status"] == "Found (Material)").sum())
        not_found_rows = int((out_df["Status"] == "Not Found").sum())
        st.write(f"Found rows: {found_rows} | Found (Material): {found_material_rows} | Not found rows: {not_found_rows}")

        csv_bytes = out_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download results as CSV",
            data=csv_bytes,
            file_name="haydon_cross_reference_bulk_results.csv",
            mime="text/csv",
        )
