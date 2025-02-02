import streamlit as st
import pandas as pd

# Define expected columns for GP (excluding SLA since we'll add it manually)
EXPECTED_COLUMNS = [
    "Generic ID", "RIO", "STL_SC", "Site ID", "STL Site Code", "Site Class Category", 
    "On Air Date From GP end", "GP Circle", "RFAI Date", "Start Date", "Last Date", "Total Day", 
    "Hour", "Total Hour", "Total Unavailable Hr", "Site wise KPI", "Remarks"
]

# SLA values for each month
SLA_VALUES = {
    "January": 99.65, "February": 99.65, "March": 99.65, "April": 99.4, "May": 99.4, "June": 99.4,
    "July": 99.55, "August": 99.55, "September": 99.55, "October": 99.65, "November": 99.65, "December": 99.65
}

# Dictionary to map month names
MONTHS = {
    "January": None, "February": None, "March": None, "April": None, "May": None, "June": None,
    "July": None, "August": None, "September": None, "October": None, "November": None, "December": None
}

st.title("GP Excel File Merger for 12 Months")

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
            df["SLA"] = SLA_VALUES[month]  # Assign SLA value
            
            # Convert 'Site wise KPI' to percentage format
            df["Site wise KPI"] = df["Site wise KPI"] / 100
            
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
