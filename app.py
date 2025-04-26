import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("📊 Excel Viewer – Edit, Filter & Download")

# Path of Excel file in your repo
EXCEL_FILE_PATH = "January Supplier Timesheet 25.xlsx"

# Load existing Excel file if present
if os.path.exists(EXCEL_FILE_PATH):
    df = pd.read_excel(EXCEL_FILE_PATH)
else:
    df = pd.DataFrame()

# File upload (optional for user)
uploaded_file = st.file_uploader("Upload a new Excel file (optional)", type=["xlsx", "xls"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)

# Allow user to select filters
if not df.empty:
    st.subheader("🔎 Filter Data")

    filter_columns = st.multiselect("Select columns to filter by", df.columns)
    filters = {}
    for col in filter_columns:
        options = df[col].dropna().unique().tolist()
        selected = st.multiselect(f"Filter values for {col}", options)
        if selected:
            filters[col] = selected

    logic = st.radio("Filter Logic", ["AND", "OR"], horizontal=True)

    # Apply filters
    def filter_data(df, filters, logic):
        if not filters:
            return df
        masks = []
        for col, values in filters.items():
            masks.append(df[col].isin(values))
        if logic == "AND":
            combined = masks[0]
            for m in masks[1:]:
                combined &= m
        else:
            combined = masks[0]
            for m in masks[1:]:
                combined |= m
        return df[combined]

    filtered_df = filter_data(df, filters, logic)

    # Display filtered data and editable
    st.subheader("✏️ Editable Data Table")
    edited_df = st.data_editor(filtered_df, use_container_width=True, num_rows="dynamic")

    # Month column adjustment
    if "Month" in edited_df.columns:
        try:
            edited_df["Month"] = edited_df["Month"].apply(lambda x: f"{x} 2025" if isinstance(x, str) and not str(x).endswith("2025") else x)
        except Exception as e:
            st.warning(f"Couldn't update 'Month' column automatically: {e}")

    # Download buttons
    st.subheader("⬇️ Download Updated Data")

    def convert_df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
        processed_data = output.getvalue()
        return processed_data

    excel_data = convert_df_to_excel(edited_df)

    st.download_button(
        label="Download Excel File",
        data=excel_data,
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("📂 Upload or load an Excel file to get started.")

