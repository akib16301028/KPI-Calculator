import streamlit as st
import pandas as pd

# Define expected columns
EXPECTED_COLUMNS = [
    "Generic ID", "RIO", "BL Circle", "BL RO", "BL site ID", "TOWERCO Site ID", "RFAI Date", 
    "Start Date", "Last Date", "Total Day", "Hour", "Total Hour", "Site Wise total Downtime", 
    "Site Wise KPI", "SLA"
]

# Dictionary to map file index to month
MONTHS = {
    1: "January", 2: "February", 3: "March", 4: "April", 5: "May", 6: "June",
    7: "July", 8: "August", 9: "September", 10: "October", 11: "November", 12: "December"
}

st.title("Excel File Merger for 12 Months")

uploaded_files = st.file_uploader("Upload 12 Excel files (one for each month)", type=["xls", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    if len(uploaded_files) != 12:
        st.error("Please upload exactly 12 files, one for each month.")
    else:
        all_data = []
        
        for i, file in enumerate(uploaded_files, start=1):
            try:
                df = pd.read_excel(file, sheet_name="Total Hour Calculation", skiprows=2)  # Read specific sheet and skip first two rows
                missing_columns = [col for col in EXPECTED_COLUMNS if col not in df.columns]
                if missing_columns:
                    st.error(f"Unable to find column name {missing_columns} in month {MONTHS[i]}")
                    continue
                df = df[EXPECTED_COLUMNS]  # Select only required columns
                df["Month"] = MONTHS[i]  # Assign month name based on order of upload
                all_data.append(df)
            except Exception as e:
                st.error(f"Error processing {file.name}: {e}")
                
        if all_data:
            merged_df = pd.concat(all_data, ignore_index=True)
            output_file = "merged_data.xlsx"
            merged_df.to_excel(output_file, index=False)
            
            st.success("Files merged successfully!")
            st.download_button(
                label="Download Merged Excel File",
                data=open(output_file, "rb").read(),
                file_name="Merged_Excel_File.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
