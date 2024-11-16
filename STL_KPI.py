import streamlit as st
import pandas as pd

# Function to process files and thresholds
def process_files(month_data, thresholds):
    results = {}
    for month, data in month_data.items():
        if data is None:  # Skip if the file is not uploaded
            continue
        try:
            # Read the required sheet starting from the correct header row
            sheet_data = pd.read_excel(data, sheet_name="Total Hour Calculation", header=2)
            # Extract necessary columns
            site_kpi = sheet_data[["Site ID", "Site wise KPI"]]
            # Compare KPI values with the threshold and add a Pass/Fail column
            site_kpi["Pass/Fail"] = site_kpi["Site wise KPI"].apply(
                lambda x: "Pass" if x >= thresholds[month] else "Fail"
            )
            results[month] = site_kpi
        except KeyError as e:
            st.error(f"Error processing {month}: Missing required columns. {e}")
        except Exception as e:
            st.error(f"Unexpected error with {month}: {e}")
    return results

# Title of the Streamlit app
st.title("Monthly KPI Comparison Tool")

# Initialize month names
months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Step 1: File upload and threshold input
month_data = {}
thresholds = {}

st.sidebar.header("Upload Files and Set Thresholds")
for month in months:
    st.sidebar.subheader(f"{month}")
    uploaded_file = st.sidebar.file_uploader(f"Upload {month} File", type=["xlsx"], key=month)
    if uploaded_file is not None:
        month_data[month] = uploaded_file  # Store the uploaded file object
        thresholds[month] = st.sidebar.number_input(f"{month} KPI Threshold", min_value=0.0, value=0.0)
    else:
        st.sidebar.write(f"Ignoring {month}.")

# Step 2: Process data when "Process" is clicked
if st.button("Process Files"):
    # Ensure at least one file is uploaded
    if all(data is None for data in month_data.values()):
        st.warning("Please upload at least one file!")
    else:
        results = process_files(month_data, thresholds)
        if results:
            # Combine results into a single Excel workbook
            with pd.ExcelWriter("KPI_Results.xlsx", engine="openpyxl") as writer:
                for month, df in results.items():
                    df.to_excel(writer, sheet_name=month, index=False)
            st.success("Processing Complete!")
            
            # Provide download button
            with open("KPI_Results.xlsx", "rb") as f:
                st.download_button("Download Results", data=f, file_name="KPI_Results.xlsx")
        else:
            st.warning("No files were processed.")
