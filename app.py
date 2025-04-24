import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("📊 Excel Viewer – Full Edit, Filter & Download")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

def filter_data(df, filters, logic):
    if not filters:
        return df

    masks = []
    for col, values in filters.items():
        if values:
            masks.append(df[col].isin(values))

    if not masks:
        return df

    if logic == "AND":
        combined_mask = masks[0]
        for m in masks[1:]:
            combined_mask &= m
    else:  # OR
        combined_mask = masks[0]
        for m in masks[1:]:
            combined_mask |= m

    return df[combined_mask]

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='FilteredData')
    output.seek(0)
    return output

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    
    st.subheader("✏️ Editable Full Data")
    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")

    # Global Search
    st.subheader("🔎 Global Search (like Ctrl+F)")
    search_term = st.text_input("Search across all columns", placeholder="Type to filter...")

    searched_df = edited_df.copy()
    if search_term:
        mask = searched_df.astype(str).apply(lambda row: row.str.contains(search_term, case=False, na=False)).any(axis=1)
        searched_df = searched_df[mask]

    st.subheader("🔍 Filter Options")
    columns_to_filter = st.multiselect("Select up to 3 columns to filter", searched_df.columns, max_selections=3)
    filter_logic = st.radio("Filter logic", ["AND", "OR"], horizontal=True)

    filters = {}
    for col in columns_to_filter:
        options = searched_df[col].dropna().unique()
        selected = st.multiselect(f"Values for {col}", options)
        if selected:
            filters[col] = selected

    filtered_df = filter_data(searched_df, filters, filter_logic)

    st.subheader("📄 Filtered & Edited Data")
    st.write(f"{len(filtered_df)} rows found")
    st.dataframe(filtered_df, use_container_width=True)

    if not filtered_df.empty:
        excel_file = convert_df_to_excel(filtered_df)
        st.download_button(
            label="📥 Download filtered + edited data",
            data=excel_file,
            file_name="filtered_edited_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
