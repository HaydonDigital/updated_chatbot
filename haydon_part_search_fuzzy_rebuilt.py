
import pandas as pd
import re
import streamlit as st
import os

# Build the pricing match key in exact format, e.g., "H-132-OS X 10' GREEN"
def build_pricing_key(haydon_part):
    if pd.isna(haydon_part):
        return ""

    part = str(haydon_part).strip().upper()

    # Extract base, length, and finish
    base_match = re.match(r"^(H-[0-9A-Z]+(?:-OS|OSA)?)", part)
    length_match = re.search(r"X\s*(\d+)", part)
    finish_match = re.search(r"\b(GR|PG|HDG|PL)\b", part)

    base = base_match.group(1) if base_match else ""
    length = length_match.group(1) + "'" if length_match else ""
    finish_map = {
        "GR": "GREEN",
        "PG": "PRE GALV",
        "HDG": "HDG",
        "PL": "PLAIN"
    }
    finish = finish_map.get(finish_match.group(1)) if finish_match else ""

    # Assemble formatted name like in pricing sheet
    components = [base]
    if length:
        components.append(f"X {length}")
    if finish:
        components.append(finish)

    return " ".join(components).strip()

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

        # Build match key using exact reconstruction
        results["Pricing Key"] = results["Haydon Part #"].apply(build_pricing_key)

        # Merge using the exact name
        merged = pd.merge(results, pricing_df[["Name", "LESS THAN TRUCKLOAD PRICE"]],
                          left_on="Pricing Key", right_on="Name", how="left")

        # Fallback using Current Item Nbr
        pricing_df["Current Item Nbr"] = pricing_df["Current Item Nbr"].astype(str).str.strip()
        fallback_merge = pd.merge(results, pricing_df[["Current Item Nbr", "LESS THAN TRUCKLOAD PRICE"]],
                                  left_on="Haydon Part #", right_on="Current Item Nbr", how="left")

        # Combine both sources of price
        merged["Price"] = merged["LESS THAN TRUCKLOAD PRICE"]
        fallback_price = fallback_merge["LESS THAN TRUCKLOAD PRICE"]
        merged["Price"] = merged["Price"].fillna(fallback_price)

        merged.drop(columns=["LESS THAN TRUCKLOAD PRICE", "Name", "Pricing Key"], inplace=True, errors="ignore")

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
