
import pandas as pd
import re
import streamlit as st
import os

# Normalize special characters and finish codes
def normalize_part_for_pricing(part):
    part = str(part).strip().upper()

    # Normalize 'X 10' to "X 10'" to match pricing sheet format
    part = re.sub(r"(X\s*\d+)(?!')", lambda m: m.group(1) + "'", part)

    # Map suffix codes to full descriptions
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
    path = os.path.join(os.path.dirname(__file__), "Updated File - 3-24.xlsx")
    df = pd.read_excel(path, sheet_name="Export", engine="openpyxl")
    df["Normalized Haydon Part"] = df["Haydon Part #"].apply(normalize)
    df["Normalized Vendor Part"] = df["Vendor Part #"].apply(normalize)
    return df

def load_haydon_reference():
    path = os.path.join(os.path.dirname(__file__), "Image.xlsx")
    return pd.read_excel(path, sheet_name="Sheet1")

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
    results = search_parts(cross_ref_df, query)

    if not results.empty:
        pricing_path = os.path.join(os.path.dirname(__file__), "Standard Pricing - June 2025.xlsx")
        pricing_df = pd.read_excel(pricing_path)
        pricing_df["Name"] = pricing_df["Name"].astype(str).str.strip()
        pricing_df["Current Item Nbr"] = pricing_df["Current Item Nbr"].astype(str).str.strip()

        # Create normalized key in both DataFrames
        results["Pricing Key"] = results["Haydon Part #"].apply(normalize_part_for_pricing)
        pricing_df["Pricing Key"] = pricing_df["Name"].apply(normalize_part_for_pricing)

        # Merge on normalized key
        merged = pd.merge(results, pricing_df[["Pricing Key", "LESS THAN TRUCKLOAD PRICE"]],
                          on="Pricing Key", how="left")

        # Fallback by Current Item Nbr
        fallback_merge = pd.merge(results, pricing_df[["Current Item Nbr", "LESS THAN TRUCKLOAD PRICE"]],
                                  left_on="Haydon Part #", right_on="Current Item Nbr", how="left")

        merged["Price"] = merged["LESS THAN TRUCKLOAD PRICE"]
        fallback_price = fallback_merge["LESS THAN TRUCKLOAD PRICE"]
        merged["Price"] = merged["Price"].fillna(fallback_price)

        merged.drop(columns=["LESS THAN TRUCKLOAD PRICE", "Pricing Key"], inplace=True, errors="ignore")

        # DEBUG TABLE
        st.write("🔧 DEBUG PRICE MATCHING")
        st.dataframe(merged[["Haydon Part #", "Price"]])

        st.subheader(f"Found {len(merged)} matching entries")
        st.dataframe(merged.drop(columns=["Normalized Haydon Part", "Normalized Vendor Part"], errors="ignore"))

        first_row = merged.iloc[0]
        haydon_part = first_row["Haydon Part #"]

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
