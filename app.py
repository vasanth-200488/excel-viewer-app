import streamlit as st
import pandas as pd
from io import BytesIO
import os

# Set page config
st.set_page_config(page_title="Excel Viewer", layout="wide")
st.title("📊 Excel Viewer – Full Edit, Filter & Download")

st.write("Files in the current directory:", os.listdir())



# Default hardcoded file
DEFAULT_FILE = "Data.xlsx"

# Upload file widget
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

# User selects source
st.subheader("📂 Select Data Source")
source_option = st.selectbox(
    "Choose data source:",
    ("Uploaded File", "Default File", "Dummy Data")
)

# Filter function
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

# Create dummy data
def create_dummy_data():
    dummy_data = {
        "Name": ["Alice", "Bob", "Charlie", "David"],
        "Grade": ["A", "B", "C", "B"],
        "Month": ["2025-04-01", "2025-04-01", "2025-04-01", "2025-04-01"]
    }
    return pd.DataFrame(dummy_data)

# Load Data
if source_option == "Uploaded File":
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.success("✅ Uploaded file loaded successfully.")
    else:
        st.warning("⚠️ Please upload a file to proceed.")
        st.stop()

elif source_option == "Default File":
    if os.path.exists(DEFAULT_FILE):
        df = pd.read_excel(DEFAULT_FILE)
        st.success(f"✅ Loaded default file: {DEFAULT_FILE}")
    else:
        st.error(f"❌ Default file '{DEFAULT_FILE}' not found. Please check the file.")
        st.stop()

else:  # Dummy Data
    df = create_dummy_data()
    st.success("✅ Dummy data generated.")

# Format 'Month' column if exists
if 'Month' in df.columns:
    df['Month'] = pd.to_datetime(df['Month'], errors='coerce').dt.strftime('%B %Y')

# Fix datatypes for editing
df = df.astype(str)

# Filters
st.subheader("🔎 Filters")
filter_cols = st.multiselect("Select columns to filter", options=df.columns)

filters = {}
for col in filter_cols:
    options = df[col].unique().tolist()
# **NEW**: Create a search input box for each column
    search_term = st.text_input(f"Search in {col}", "")

# **NEW**: Filter options based on search term
    filtered_options = [option for option in options if search_term.lower() in str(option).lower()]

# **NEW**: Create a multiselect dropdown for filtered options
    selected = st.multiselect(f"Filter by {col}", options=filtered_options, key=col)
    if selected:
        filters[col] = selected

logic = st.radio("Apply filter logic", ["AND", "OR"], horizontal=True)

filtered_df = filter_data(df, filters, logic)

# View full data
st.subheader("📄 View Full Data")
st.dataframe(df, use_container_width=True)

# Editable filtered data
st.subheader("✏️ Editable Filtered Data")
edited_df = st.data_editor(filtered_df, use_container_width=True, num_rows="dynamic")

# Download button
output = BytesIO()
edited_df.to_excel(output, index=False, engine='openpyxl')
st.download_button(
    label="💾 Download as Excel",
    data=output.getvalue(),
    file_name="filtered_data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
