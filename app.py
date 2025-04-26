import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("📊 Viewer App")

# Define the path to the shared Excel file
EXCEL_FILE_PATH = 'C:/Users/chand/excel-viewer-app/January Supplier Timesheet 25.xlsx'  # Modify this to your shared folder location

df = pd.read_excel(EXCEL_FILE_PATH)

# (Optional) Allow user to upload another file if they want
uploaded_file = st.file_uploader("Or upload another Excel file", type=["xlsx", "xls"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.info("Loaded uploaded Excel file.")

st.dataframe(df)

# Allow users to upload a new file if they want
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.info("Loaded data from the uploaded file.")
    except Exception as e:
        st.error(f"Error loading uploaded file: {e}")

# Display the dataframe
if 'df' in locals():
    st.dataframe(df)
else:
    st.warning("No data available to display.")

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

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # 🗓 Format 'Month' column (adjust if your month column has a different name)
        if 'Month' in df.columns:
            df['Month'] = pd.to_datetime(df['Month'], errors='coerce').dt.strftime('%B %Y')

        # 🔧 Fix for data_editor serialization issue
        df = df.astype(str)

        st.subheader("Filters")
        filter_cols = st.multiselect("Select columns to filter", options=df.columns)

        filters = {}
        for col in filter_cols:
            options = df[col].unique().tolist()
            selected = st.multiselect(f"Filter by {col}", options=options, key=col)
            if selected:
                filters[col] = selected

        logic = st.radio("Apply filter logic", ["AND", "OR"], horizontal=True)

        filtered_df = filter_data(df, filters, logic)

        st.subheader("📄 View Full Data")
        st.dataframe(df, use_container_width=True)  # View-only data table

        st.subheader("✏️ Editable Filtered Data")
        edited_df = st.data_editor(filtered_df, use_container_width=True, num_rows="dynamic")

        # Download edited data as Excel
        output = BytesIO()
        edited_df.to_excel(output, index=False, engine='openpyxl')
        st.download_button(
            label="Download as Excel",
            data=output.getvalue(),
            file_name="filtered_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error reading the file: {e}")
else:
    st.info("Please upload an Excel file.")
