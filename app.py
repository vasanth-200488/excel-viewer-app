import streamlit as st
import pandas as pd
from io import BytesIO
import os

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("📊 Excel Viewer – Full Edit, Filter & Download")

# Path to the uploaded or preloaded Excel file
EXCEL_FILE_PATH = "January Supplier Timesheet 25.xlsx"

# Load Excel data if file exists
if os.path.exists(EXCEL_FILE_PATH):
    df = pd.read_excel(EXCEL_FILE_PATH)
else:
    st.error("Excel file not found.")
    st.stop()

# Function to apply multiple filters (AND/OR logic)
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
    else:  # OR logic
        combined_mask = masks[0]
        for m in masks[1:]:
            combined_mask |= m

    return df[combined_mask]

# Filter options
st.sidebar.subheader("Filter Data")
filter_columns = df.columns.tolist()

filters = {}
for col in filter_columns:
    unique_values = df[col].dropna().unique().tolist()
    filters[col] = st.sidebar.multiselect(f"Filter by {col}", unique_values)

logic = st.sidebar.radio("Filter Logic", ["AND", "OR"])

# Apply filters
filtered_df = filter_data(df, filters, logic)

# Editable Data Table
st.subheader("✏️ Editable Full Data")
edited_df = st.data_editor(filtered_df, use_container_width=True, num_rows="dynamic")

# Display the edited DataFrame
st.dataframe(edited_df)

# Function to convert DataFrame to Excel and provide download
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    output.seek(0)
    return output

# Download edited data
st.subheader("Download Updated Data")
excel_data = convert_df_to_excel(edited_df)
st.download_button(label="Download Excel File", data=excel_data, file_name="edited_data.xlsx", mime="application/vnd.ms-excel")
