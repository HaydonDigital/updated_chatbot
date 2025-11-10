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
PRICING_FILE = BASE_DIR / "Standard Pricing - June 2025.xlsx"
IMAGES_FILE = BASE_DIR / "Image and Submittals.xlsx"


# =========================================================
# UTILITY FUNCTIONS
# =========================================================
def normalize(text):
    """Strip non-alphanumerics and lowercase for fuzzy-ish matching."""
    if pd.isna(text):
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(text)).lower()


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
def load_pricing():
    if not PRICING_FILE.exists():
        raise FileNotFoundError(f"Pricing file not found at: {PRICING_FILE}")
    df = pd.read_excel(PRICING_FILE)
    # standardize join key
    df["Cross Reference Name"] = (
        df["Cross Reference Name"].astype(str).str.strip().str.upper()
    )
    # drop dupes so merge is cleaner
    df = df.drop_duplicates(subset=["Cross Reference Name"])
    return df


@st.cache_data
def load_images():
    if not IMAGES_FILE.exists():
        raise FileNotFoundError(f"Image/submittals file not found at: {IMAGES_FILE}")
    df = pd.read_excel(IMAGES_FILE, sheet_name="Sheet1")
    # normalize name column for easier matching later
    if "Name" in df.columns:
        df["Name_upper"] = df["Name"].astype(str).str.upper()
    return df


# =========================================================
# SEARCH
# =========================================================
def search_parts(cross_df: pd.DataFrame, query: str) -> pd.DataFrame:
    norm_query = normalize(query)
    return cross_df[
        cross_df["Normalized Haydon Part"].str.contains(norm_query, na=False)
        | cross_df["Normalized Vendor Part"].str.contains(norm_query, na=False)
    ]


# =========================================================
# APP UI
# =========================================================
st.title("Haydon Cross-Reference Search")

query = st.text_input("Enter part number (Haydon or Vendor):")

if query:
    try:
        cross_df = load_cross_reference()
        pricing_df = load_pricing()
        image_df = load_images()
    except FileNotFoundError as e:
        st.error(str(e))
        st.stop()

    results = search_parts(cross_df, query)

    if not results.empty:
        # enrich with pricing
        enriched = pd.merge(
            results,
            pricing_df[
                [
                    "Cross Reference Name",
                    "Macola\nItem",
                    "SAP Item Nbr",
                    "LESS THAN TRUCKLOAD PRICE",
                    "TRUCKLOAD PRICE",
                ]
            ],
            left_on="Haydon Part Description",
            right_on="Cross Reference Name",
            how="left",
        )

        # final display columns
        display_cols = [
            "Vendor Part #",
            "Vendor",
            "Category",
            "Haydon Part Description",
            "Macola\nItem",
            "SAP Item Nbr",
            "LESS THAN TRUCKLOAD PRICE",
            "TRUCKLOAD PRICE",
        ]
        display_df = enriched[[c for c in display_cols if c in enriched.columns]]

        st.subheader(f"Found {len(display_df)} matching entries")
        st.dataframe(display_df, use_container_width=True)

        # -------------------------------------------------
        # Sidebar: product preview / submittals
        # -------------------------------------------------
        first_row = enriched.iloc[0]
        haydon_part = first_row["Haydon Part Description"]

        with st.sidebar:
            st.markdown("### Haydon Product Preview")
            match_found = False
            # start from the exact part, then try shorter versions
            candidates = [haydon_part] + list(get_haydon_candidates(haydon_part))
            for candidate in candidates:
                matched = image_df[image_df["Name_upper"] == str(candidate).upper()]
                if not matched.empty:
                    row = matched.iloc[0]
                    # show image if present
                    if "Cover Image" in row and pd.notna(row["Cover Image"]):
                        st.image(
                            row["Cover Image"],
                            caption=row.get("Name", candidate),
                            use_container_width=True,
                        )
                    # show submittal link if present
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
    st.write("Enter a part number above to begin.")
