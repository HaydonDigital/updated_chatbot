import re
from pathlib import Path

import pandas as pd
import streamlit as st

# =========================================================
# PDF SUPPORT (text + OCR fallback)
# =========================================================
# Text-based PDFs
try:
    from PyPDF2 import PdfReader  # pip install PyPDF2
    HAS_PYPDF2 = True
except Exception:
    HAS_PYPDF2 = False

# Scanned/image PDFs (OCR)
try:
    from pdf2image import convert_from_bytes  # pip install pdf2image
    import pytesseract  # pip install pytesseract
    HAS_OCR = True
except Exception:
    HAS_OCR = False


# =========================================================
# CONFIG
# =========================================================
st.set_page_config(layout="wide", page_title="Haydon Cross-Reference Search")

BASE_DIR = Path(__file__).parent
CROSS_FILE = BASE_DIR / "Updated File - 7-10-25.xlsx"
IMAGES_FILE = BASE_DIR / "Image and Submittals.xlsx"


# =========================================================
# NORMALIZATION / PARSING
# =========================================================
def normalize(text) -> str:
    """Strip non-alphanumerics and lowercase for matching."""
    if pd.isna(text):
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(text)).lower()


def split_parts_from_text(raw: str) -> list[str]:
    """Split pasted input into part tokens (newline/comma/space separated)."""
    if not raw:
        return []
    tokens = re.split(r"[\n,;\t ]+", raw.strip())
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


def get_haydon_candidates(part: str):
    """
    Progressively shorten a Haydon part so we can try to match the image/submittal file.
    """
    part = str(part).upper()
    tokens = re.split(r"[ \-X()]+", part)
    for i in range(len(tokens), 0, -1):
        yield "-".join(tokens[:i])


# =========================================================
# PDF EXTRACTION
# =========================================================
def extract_text_from_pdf_text_layer(uploaded_file) -> str:
    """
    Extract embedded text from a text-based PDF. Returns "" if not possible.
    """
    if not HAS_PYPDF2:
        return ""
    try:
        reader = PdfReader(uploaded_file)
        pages_text = []
        for page in reader.pages:
            pages_text.append(page.extract_text() or "")
        return "\n".join(pages_text)
    except Exception:
        return ""


def extract_text_from_pdf_ocr(uploaded_file, dpi: int = 250, max_pages: int = 5) -> str:
    """
    OCR scanned PDFs (image-based).
    Requires:
      - pdf2image + pytesseract
      - System packages: poppler-utils, tesseract-ocr
    """
    if not HAS_OCR:
        return ""
    try:
        pdf_bytes = uploaded_file.getvalue()
        images = convert_from_bytes(
            pdf_bytes,
            dpi=dpi,
            first_page=1,
            last_page=max_pages,
        )
        chunks = []
        for img in images:
            chunks.append(pytesseract.image_to_string(img))
        return "\n".join(chunks)
    except Exception:
        return ""


def extract_text_smart(uploaded_file, ocr_enabled: bool, ocr_dpi: int, ocr_pages: int) -> str:
    """
    Try text-layer extraction first. If empty and OCR is enabled, run OCR.
    """
    # Important: PdfReader consumes the file-like object in some contexts.
    # Reset pointer when possible.
    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    text = extract_text_from_pdf_text_layer(uploaded_file)
    if text and len(text.strip()) > 50:
        return text

    if not ocr_enabled:
        return text or ""

    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    ocr_text = extract_text_from_pdf_ocr(uploaded_file, dpi=ocr_dpi, max_pages=ocr_pages)
    return ocr_text or ""


def extract_part_candidates_from_text(text: str) -> list[str]:
    """
    Pull likely part-number-like tokens from extracted text.
    Tune the regex to your common formats if you want tighter/looser behavior.
    """
    if not text:
        return []

    # General "part-like" token: letters/numbers with -, /, ., # allowed (4-40 chars)
    pattern = re.compile(r"\b[A-Z0-9][A-Z0-9\-\/\.\#]{3,39}\b", re.IGNORECASE)
    hits = pattern.findall(text)

    cleaned = []
    seen = set()
    for h in hits:
        h2 = h.strip().strip(".,;:()[]{}")
        if len(h2) < 4:
            continue
        key = h2.upper()
        if key not in seen:
            seen.add(key)
            cleaned.append(h2)
    return cleaned


# =========================================================
# DATA LOADERS
# =========================================================
@st.cache_data
def load_cross_reference():
    if not CROSS_FILE.exists():
        raise FileNotFoundError(f"Cross-reference file not found at: {CROSS_FILE}")
    df = pd.read_excel(CROSS_FILE, sheet_name="Export", engine="openpyxl")

    if "Haydon Part Description" not in df.columns or "Vendor Part #" not in df.columns:
        raise ValueError("Cross-reference sheet must include 'Haydon Part Description' and 'Vendor Part #' columns.")

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
    else:
        df["Name_upper"] = ""
    return df


# =========================================================
# SEARCH
# =========================================================
def search_parts_contains(cross_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Contains search (matches partial strings)."""
    norm_query = normalize(query)
    if not norm_query:
        return cross_df.iloc[0:0]
    return cross_df[
        cross_df["Normalized Haydon Part"].str.contains(norm_query, na=False)
        | cross_df["Normalized Vendor Part"].str.contains(norm_query, na=False)
    ]


def search_parts_exact(cross_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Exact normalized match (recommended for bulk)."""
    norm_query = normalize(query)
    if not norm_query:
        return cross_df.iloc[0:0]
    return cross_df[
        (cross_df["Normalized Haydon Part"] == norm_query)
        | (cross_df["Normalized Vendor Part"] == norm_query)
    ]


def bulk_search(cross_df: pd.DataFrame, parts: list[str], allow_contains_fallback: bool) -> pd.DataFrame:
    """
    Output: one row per match (or one row per input if not found).
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
            rows.append(
                {
                    "Input": raw,
                    "Status": "Not Found",
                    "Vendor Part #": None,
                    "Vendor": None,
                    "Category": None,
                    "Haydon Part Description": None,
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
                        "Match Count (for Input)": len(found),
                    }
                )
    return pd.DataFrame(rows)


# =========================================================
# UI
# =========================================================
st.title("Haydon Cross-Reference Search")

# Load data once
try:
    cross_df = load_cross_reference()
    image_df = load_images()
except (FileNotFoundError, ValueError) as e:
    st.error(str(e))
    st.stop()

tab_single, tab_bulk = st.tabs(["Single", "Bulk"])


# =========================================================
# SINGLE TAB
# =========================================================
with tab_single:
    query = st.text_input("Enter part number (Haydon or Vendor):")

    if query:
        results = search_parts_contains(cross_df, query)

        if not results.empty:
            display_cols = ["Vendor Part #", "Vendor", "Category", "Haydon Part Description"]
            display_df = results[[c for c in display_cols if c in results.columns]]

            st.subheader(f"Found {len(display_df)} matching entries")
            st.dataframe(display_df, use_container_width=True)

            # Sidebar preview for first result
            first_row = results.iloc[0]
            haydon_part = first_row.get("Haydon Part Description", "")

            with st.sidebar:
                st.markdown("### Haydon Product Preview")
                match_found = False

                candidates = [haydon_part] + list(get_haydon_candidates(haydon_part))
                for candidate in candidates:
                    matched = image_df[image_df["Name_upper"] == str(candidate).upper()]
                    if not matched.empty:
                        row = matched.iloc[0]

                        if "Cover Image" in row and pd.notna(row.get("Cover Image")):
                            st.image(
                                row["Cover Image"],
                                caption=row.get("Name", candidate),
                                use_container_width=True,
                            )

                        if "Files" in row and pd.notna(row.get("Files")):
                            st.markdown(
                                f"[View Submittal]({row['Files']})",
                                unsafe_allow_html=True,
                            )

                        match_found = True
                        break

                if not match_found:
                    st.warning("No product preview or submittal found.")
        else:
            st.error(
                "No match found. Send the Haydon part number and the customer/competitor part number to "
                "marketing@haydoncorp.com."
            )
    else:
        st.write("Enter a part number above to begin.")


# =========================================================
# BULK TAB
# =========================================================
with tab_bulk:
    st.markdown("Paste part numbers and/or upload a PDF to cross-reference in bulk.")

    col1, col2 = st.columns(2)

    with col1:
        pasted = st.text_area(
            "Paste part numbers (one per line, or comma/space separated):",
            height=220,
            placeholder="Example:\nPS-1100-AS-4-EG\nG582-A-OS-1\nEG0815-10",
        )

    with col2:
        pdf = st.file_uploader("Upload a PDF:", type=["pdf"])

        allow_contains = st.checkbox(
            "If exact match fails, try contains fallback (may increase false matches)",
            value=True,
        )

        st.markdown("PDF extraction options")
        ocr_enabled = st.checkbox(
            "Enable OCR fallback for scanned PDFs (recommended)",
            value=True,
            disabled=not HAS_OCR,
        )
        ocr_pages = st.number_input("OCR pages (first N pages)", min_value=1, max_value=50, value=5, step=1)
        ocr_dpi = st.number_input("OCR DPI (higher = better, slower)", min_value=150, max_value=400, value=250, step=10)

        if pdf is not None:
            if not HAS_PYPDF2:
                st.info("PyPDF2 not available. Add PyPDF2 to requirements.txt for text-based PDF extraction.")
            if not HAS_OCR:
                st.info(
                    "OCR not available. Add pdf2image and pytesseract to requirements.txt, and install "
                    "system packages poppler-utils and tesseract-ocr."
                )

    parts = []
    if pasted and pasted.strip():
        parts.extend(split_parts_from_text(pasted))

    extracted_candidates = []
    if pdf is not None:
        extracted_text = extract_text_smart(pdf, ocr_enabled=ocr_enabled, ocr_dpi=int(ocr_dpi), ocr_pages=int(ocr_pages))

        if not extracted_text or len(extracted_text.strip()) < 20:
            st.warning("Could not extract text from this PDF (it may be scanned, protected, or OCR is not set up).")
        else:
            extracted_candidates = extract_part_candidates_from_text(extracted_text)
            st.info(f"Extracted {len(extracted_candidates)} candidate tokens from the PDF.")
            with st.expander("View extracted candidates"):
                st.write(extracted_candidates)

            parts.extend(extracted_candidates)

    parts_unique = dedupe_preserve_order(parts)
    st.write(f"Inputs ready: {len(parts_unique)}")

    run = st.button("Run Bulk Cross-Reference", type="primary", disabled=(len(parts_unique) == 0))

    if run:
        out_df = bulk_search(cross_df, parts_unique, allow_contains_fallback=allow_contains)

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

st.sidebar.markdown("---")
st.sidebar.markdown("Support: marketing@haydoncorp.com")
