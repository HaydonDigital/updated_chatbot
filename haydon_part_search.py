
import pandas as pd
import re
import streamlit as st
import os

# Load the Excel data
def load_cross_reference():
    file_path = os.path.join(os.path.dirname(__file__), "Updated File - 3-24.xlsx")
    df = pd.read_excel(file_path, sheet_name="Export", engine="openpyxl")
    df["Normalized Haydon Part"] = df["Haydon Part #"].apply(normalize)
    df["Normalized Vendor Part"] = df["Vendor Part #"].apply(normalize)
    return df

def load_haydon_reference():
    path = os.path.join(os.path.dirname(__file__), "Image.xlsx")
    df = pd.read_excel(path, sheet_name="Sheet1")
    return df

def normalize(part):
    if pd.isna(part):
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(part)).lower()

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
        st.subheader(f"Found {len(results)} matching entries")
        st.dataframe(results.drop(columns=["Normalized Haydon Part", "Normalized Vendor Part"]))

        first_row = results.iloc[0]
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
