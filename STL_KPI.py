import streamlit as st
import pandas as pd
import os

# Define the required columns
required_columns = [
    "Generic ID", "RIO", "STL_SC", "Site ID", "STL Site Code", "Site Class Category",
    "On Air Date From GP end", "GP Circle", "RFAI Date", "Start Date", "Last Date",
    "Total Day", "Hour", "Total Hour", "Total Unavailable Hr", "Site wise KPI", "Remarks"
]

# Define SLA values for each month
sla_values = {
    "January": 99.65,
    "February": 99.65,
    "March": 99.65,
    "April": 99.4,
    "May": 99.4,
    "June": 99.4,
    "July": 99.55,
    "August": 99.55,
    "September": 99.55,
    "October": 99.65,
    "November": 99.65,
    "December": 99.65
}

# Function to process a single file
def process_file(file, month):
    try:
        # Read the Excel file
        df = pd.read_excel(file, sheet_name="Total Hour Calculation", header=2)  # Header starts from row 3
    except Exception as e:
        st.error(f"Error reading {month} file: {e}")
        return None

    # Check if all required columns are present
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        st.error(f"Unable to find column(s) {', '.join(missing_columns)} in {month} file.")
        return None

    # Add a new column for the month
    df["Month"] = month

    # Add SLA column
    df["SLA"] = sla_values[month]

    # Convert "Site wise KPI" to percentage format
    if "Site wise KPI" in df.columns:
        df["Site wise KPI"] = df["Site wise KPI"].apply(lambda x: f"{x}%")

    return df

# Streamlit app
st.title("Monthly Data Consolidation")

# Upload 12 files for each month
months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

uploaded_files = {}
for month in months:
    uploaded_files[month] = st.file_uploader(f"Upload {month} file", type=["xlsx", "xls"])

# Process files when all are uploaded
if all(uploaded_files.values()):
    combined_df = pd.DataFrame()

    for month, file in uploaded_files.items():
        df = process_file(file, month)
        if df is not None:
            combined_df = pd.concat([combined_df, df], ignore_index=True)

    if not combined_df.empty:
        st.success("All files processed successfully!")
        st.write(combined_df)

        # Save the combined DataFrame to a new Excel file
        output_file = "combined_data.xlsx"
        combined_df.to_excel(output_file, index=False)
        st.download_button(
            label="Download Combined File",
            data=open(output_file, "rb").read(),
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Please upload all 12 files to proceed.")
