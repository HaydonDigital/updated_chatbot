import pandas as pd
import re
import streamlit as st
import os

# --- Utility Functions ---

def normalize(text):
    if pd.isna(text):
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(text)).lower()

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

def get_haydon_candidates(part):
    part = str(part).upper()
    tokens = re.split(r"[ \-X()]+", part)
    for i in range(len(tokens), 0, -1):
        yield "-".join(tokens[:i])

# --- Load Data Functions ---

def load_cross_reference():
    path = os.path.join(os.path.dirname(__file__), "Updated File - 7-10-25.xlsx")
    df = pd.read_excel(path, sheet_name="Export", engine="openpyxl")
    df["Normalized Haydon Part"] = df["Haydon Part Description"].apply(normalize)
    df["Normalized Vendor Part"] = df["Vendor Part #"].apply(normalize)
    return df

def load_pricing():
    path = os.path.join(os.path.dirname(__file__), "Standard Pricing - June 2025.xlsx")
    df = pd.read_excel(path)
    df["Cross Reference Name"] = df["Cross Reference Name"].astype(str).str.strip().str.upper()
    return df.drop_duplicates(subset=["Cross Reference Name"])

def load_images():
    path = os.path.join(os.path.dirname(__file__), "Image and Submittals.xlsx")
    return pd.read_excel(path, sheet_name="Sheet1")

def search_parts(df, query):
    norm_query = normalize(query)
    return df[
        df["Normalized Haydon Part"].str.contains(norm_query, na=False) |
        df["Normalized Vendor Part"].str.contains(norm_query, na=False)
    ]

# --- Streamlit App ---

st.set_page_config(layout="wide")
st.title("Haydon Cross-Reference Search")

query = st.text_input("Enter part number (Haydon or Vendor):")

if query:
    cross_df = load_cross_reference()
    pricing_df = load_pricing()
    image_df = load_images()

    results = search_parts(cross_df, query)

    if not results.empty:
        enriched = pd.merge(
            results,
            pricing_df[["Cross Reference Name", "Macola\nItem", "SAP Item Nbr", "LESS THAN TRUCKLOAD PRICE", "TRUCKLOAD PRICE"]],
            left_on="Haydon Part Description",
            right_on="Cross Reference Name",
            how="left"
        )

        display_df = enriched[[
            "Vendor Part #", "Vendor", "Category", "Haydon Part Description",
            "Macola\nItem", "SAP Item Nbr", "LESS THAN TRUCKLOAD PRICE", "TRUCKLOAD PRICE"
        ]]

        st.subheader(f"Found {len(display_df)} matching entries")
        st.dataframe(display_df)

        first_row = enriched.iloc[0]
        haydon_part = first_row["Haydon Part Description"]

        with st.sidebar:
            st.markdown("### Haydon Product Preview")
            match_found = False
            candidates = [haydon_part] + list(get_haydon_candidates(haydon_part))
            for candidate in candidates:
                matched = image_df[image_df["Name"].str.upper() == candidate]
                if not matched.empty:
                    row = matched.iloc[0]
                    if pd.notna(row["Cover Image"]):
                        st.image(row["Cover Image"], caption=row["Name"], use_container_width=True)
                    if pd.notna(row["Files"]):
                        st.markdown(f"[📄 View Submittal]({row['Files']})", unsafe_allow_html=True)
                    match_found = True
                    break
            if not match_found:
                st.warning("No product preview or submittal found.")
    else:
        st.error(
            "No match found? Send the Haydon part number and the customer/competitor part number to "
            "[marketing@haydoncorp.com](mailto:marketing@haydoncorp.com). "
            "Prefer Teams? [Start a chat with Brock]"
            "(https://teams.microsoft.com/l/chat/0/0?users=b.bernholtz@haydoncorp.com) "
            "and include any details you have."
        )
