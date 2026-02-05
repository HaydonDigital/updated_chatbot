import re
from pathlib import Path

import pandas as pd
import streamlit as st

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(layout="wide", page_title="Haydon Cross-Reference Search")

BASE_DIR = Path(__file__).parent
CROSS_FILE = BASE_DIR / "Updated File - 7-10-25.xlsx"
IMAGES_FILE = BASE_DIR / "Image and Submittals.xlsx"


# =========================================================
# UTILITY FUNCTIONS
# =========================================================
def normalize(text):
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
    return df


# =========================================================
# SEARCH
# =========================================================
def search_parts(cross_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Contains search (kept as your original behavior)."""
    norm_query = normalize(query)
    if not norm_query:
        return cross_df.iloc[0:0]
    return cross_df[
        cross_df["Normalized Haydon Part"].str.contains(norm_query, na=False)
        | cross_df["Normalized Vendor Part"].str.contains(norm_query, na=False)
    ]


def search_parts_exact(cross_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Exact match on normalized values (used first in bulk)."""
    norm_query = normalize(query)
    if not norm_query:
        return cross_df.iloc[0:0]
    return cross_df[
        (cross_df["Normalized Haydon Part"] == norm_query)
        | (cross_df["Normalized Vendor Part"] == norm_query)
    ]


def bulk_search(cross_df: pd.DataFrame, parts: list[str], allow_contains_fallback: bool) -> pd.DataFrame:
    rows = []
    for raw in parts:
        raw = (raw or "").strip()
        if not raw:
            continue

        found = search_parts_exact(cross_df, raw)
        if found.empty and allow_contains_fallback:
            found = search_parts(cross_df, raw)

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
# APP UI
# =========================================================
st.title("Haydon Cross-Reference Search")

try:
    cross_df = load_cross_reference()
    image_df = load_images()
except (FileNotFoundError, ValueError) as e:
    st.error(str(e))
    st.stop()

tab_single, tab_bulk = st.tabs(["Single", "Bulk"])


# =========================================================
# SINGLE TAB (SIDEBAR UNCHANGED)
# =========================================================
with tab_single:
    query = st.text_input("Enter part number (Haydon or Vendor):")

    if query:
        results = search_parts(cross_df, query)

        if not results.empty:
            display_cols = ["Vendor Part #", "Vendor", "Category", "Haydon Part Description"]
            display_df = results[[c for c in display_cols if c in results.columns]]

            st.subheader(f"Found {len(display_df)} matching entries")
            st.dataframe(display_df, use_container_width=True)

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
# BULK TAB (PASTE LIST ONLY)
# =========================================================
with tab_bulk:
    st.markdown("Paste a list of part numbers to cross-reference in bulk.")

    pasted = st.text_area(
        "Paste part numbers (one per line, or comma/space separated):",
        height=260,
        placeholder="Example:\nPS-1100-AS-4-EG\nG582-A-OS-1\nEG0815-10",
    )

    allow_contains = st.checkbox(
        "If exact match fails, try contains fallback (may increase false matches)",
        value=True,
    )

    parts = dedupe_preserve_order(split_parts_from_text(pasted))
    st.write(f"Inputs ready: {len(parts)}")

    run = st.button("Run Bulk Cross-Reference", type="primary", disabled=(len(parts) == 0))

    if run:
        out_df = bulk_search(cross_df, parts, allow_contains_fallback=allow_contains)

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
