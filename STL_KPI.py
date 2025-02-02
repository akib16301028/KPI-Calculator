import streamlit as st
import pandas as pd

# Define expected columns
EXPECTED_COLUMNS = [
    "Generic ID", "RIO", "BL Circle", "BL RO", "BL site ID", "TOWERCO Site ID", "RFAI Date", 
    "Start Date", "Last Date", "Total Day", "Hour", "Total Hour", "Site Wise total Downtime", 
    "Site Wise KPI", "SLA"
]

# Dictionary to map month names
MONTHS = {
    "January": None, "February": None, "March": None, "April": None, "May": None, "June": None,
    "July": None, "August": None, "September": None, "October": None, "November": None, "December": None
}

st.title("Excel File Merger for 12 Months")

# File uploaders for each month
for month in MONTHS.keys():
    MONTHS[month] = st.file_uploader(f"Upload Excel file for {month}", type=["xls", "xlsx", "xlsb"], key=month)

if all(MONTHS.values()):
    all_data = []
    
    for month, file in MONTHS.items():
        try:
            df = pd.read_excel(file, sheet_name="Total Hour Calculation", skiprows=2, engine="pyxlsb" if file.name.endswith(".xlsb") else None)  # Read specific sheet and skip first two rows
            missing_columns = [col for col in EXPECTED_COLUMNS if col not in df.columns]
            if missing_columns:
                st.error(f"Unable to find column name {missing_columns} in month {month}")
                continue
            df = df[EXPECTED_COLUMNS]  # Select only required columns
            df["Month"] = month  # Assign month name
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
else:
    st.warning("Please upload all 12 files before proceeding.")
