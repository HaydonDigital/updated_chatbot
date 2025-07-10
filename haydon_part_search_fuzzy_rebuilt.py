
import pandas as pd
import re
import streamlit as st
import os

# Normalize special characters and finish codes
def normalize_part_for_pricing(part):
    part = str(part).strip().upper()
    part = re.sub(r"(X\s*\d+)(?!')", lambda m: m.group(1) + "'", part)

    finish_map = {
        "GR": "GREEN",
        "HDG": "HDG",
        "PG": "PRE GALV",
        "PL": "PLAIN",
    }

    for code, desc in finish_map.items():
        part = re.sub(rf"\b{code}\b", desc, part)

    part = re.sub(r"\s{2,}", " ", part).strip()
    return part

def normalize(part):
    if pd.isna(part):
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(part)).lower()

def load_cross_reference():
    path = os.path.join(os.path.dirname(__file__), "Updated File - 7-10-25.xlsx")
    df = pd.read_excel(path, sheet_name="Export", engine="openpyxl")
    df["Normalized Haydon Part"] = df["Haydon Part Description"].apply(normalize)
    df["Normalized Vendor Part"] = df["Vendor Part #"].apply(normalize)
    return df

def load_haydon_reference():
    path = os.path.join(os.path.dirname(__file__), "Image and Submittals.xlsx")
    return pd.read_excel(path, sheet_name="Sheet1")

def load_pricing_reference():
    path = os.path.join(os.path.dirname(__file__), "Standard Pricing - June 2025.xlsx")
    return pd.read_excel(path)

def get_haydon_candidates(part):
    part = str(part).upper()
    tokens = re.split(r"[ \-X()]+", part)
    for i in range(len(tokens), 0, -1):
        yield "-".join(tokens[:i])

def search_parts(df, query):
    norm_query = normalize(query)
    return df[
        df["Normalized Haydon Part"].str.contains(norm_query, na=False) |
        df["Normalized Vendor Part"].str.contains(norm_query, na=False)
    ]

# Streamlit UI
st.set_page_config(layout="wide")
st.title("Haydon Cross-Reference Search")

query = st.text_input("Enter part number (Haydon or Vendor):")

if query:
    cross_ref_df = load_cross_reference()
    image_ref_df = load_haydon_reference()
    pricing_df = load_pricing_reference()
    results = search_parts(cross_ref_df, query)

    if not results.empty:
        # Normalize pricing sheet for matching
        pricing_df["Cross Reference Name"] = pricing_df["Cross Reference Name"].astype(str).str.strip().str.upper()
        updated_ref = pd.read_excel(os.path.join(os.path.dirname(__file__), "Updated File - 7-10-25.xlsx"), sheet_name="Export")
        updated_ref["Haydon Part Description"] = updated_ref["Haydon Part Description"].astype(str).str.strip().str.upper()

        # Merge pricing info by matching Cross Reference Name <-> Haydon Part Description
        enriched = pd.merge(
            results,
            updated_ref[["Haydon Part Description", "Haydon Part Description"]].drop_duplicates(),
            on="Haydon Part Description",
            how="left"
        )

        enriched = pd.merge(
            enriched,
            pricing_df[
    ["Cross Reference Name", "Macola Number", "SAP Number", "LESS THAN TRUCKLOAD PRICE", "TRUCKLOAD PRICE"]
].drop_duplicates(subset=["Cross Reference Name"]),
            left_on="Haydon Part Description",
            right_on="Cross Reference Name",
            how="left"
        )

        # Final column order
        display_df = enriched[[
            "Vendor Part #", "Vendor", "Category", "Haydon Part Description",
            "Macola Number", "SAP Number", "LESS THAN TRUCKLOAD PRICE", "TRUCKLOAD PRICE"
        ]]

        st.subheader(f"Found {len(display_df)} matching entries")
        st.dataframe(display_df)

        # Sidebar preview (unchanged)
        first_row = enriched.iloc[0]
        haydon_part = first_row["Haydon Part Description"]

        with st.sidebar:
            st.markdown("### Haydon Product Preview")
            match_found = False
            candidates = [haydon_part] + list(get_haydon_candidates(haydon_part))
            for candidate in candidates:
                matched_ref = image_ref_df[image_ref_df["Name"].str.upper() == candidate]
                if not matched_ref.empty:
                    ref_row = matched_ref.iloc[0]
                    image_url = ref_row["Cover Image"]
                    submittal_url = ref_row["Files"]
                    if pd.notna(image_url):
                        st.image(image_url, caption=ref_row["Name"], use_container_width=True)
                    if pd.notna(submittal_url):
                        st.markdown(f"[📄 View Submittal]({submittal_url})", unsafe_allow_html=True)
                    match_found = True
                    break
            if not match_found:
                st.warning("No product preview or submittal found for this Haydon part.")
    else:
        st.warning("No matches found.")
