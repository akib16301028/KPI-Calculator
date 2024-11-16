import streamlit as st
import pandas as pd

# Default KPI thresholds for GP and BL
GP_THRESHOLDS = {
    "January": 99.65, "February": 99.65, "March": 99.65, "April": 99.40,
    "May": 99.40, "June": 99.40, "July": 99.55, "August": 99.55,
    "September": 99.55, "October": 99.65, "November": 99.65, "December": 99.65
}
BL_THRESHOLDS = {
    "January": 99.76, "February": 99.6, "March": 99.49, "April": 99.95,
    "May": 99.05, "June": 99.55, "July": 99.57, "August": 99.65,
    "September": 99.66, "October": 99.7, "November": 99.78, "December": 99.77
}

# Function to process files and thresholds
def process_files(client, month_data, thresholds):
    results = {}
    fail_summary = pd.DataFrame()

    for month, data in month_data.items():
        if data is None:  # Skip if the file is not uploaded
            continue
        try:
            # Read the required sheet starting from the correct header row
            if client == "BL":
                # For BL, use pyxlsb to read binary Excel files
                sheet_data = pd.read_excel(data, sheet_name="Total Hour Calculation", engine="pyxlsb", header=2)
                site_kpi = sheet_data[["Generic ID", "Site Wise KPI", "RIO"]]
                site_kpi.rename(columns={"Generic ID": "Site ID", "Site Wise KPI": "Site wise KPI"}, inplace=True)
            else:  # GP
                sheet_data = pd.read_excel(data, sheet_name="Total Hour Calculation", header=2)
                site_kpi = sheet_data[["Site ID", "Site wise KPI", "RIO", "STL_SC"]]

            # Ignore rows with KPI == 0
            site_kpi = site_kpi[site_kpi["Site wise KPI"] > 0]

            # Add Threshold and Pass/Fail columns
            site_kpi["Threshold"] = thresholds[month]
            site_kpi["Pass/Fail"] = site_kpi["Site wise KPI"].apply(
                lambda x: "Pass" if x >= thresholds[month] else "Fail"
            )

            # Add Fail results to the summary
            site_kpi["Month"] = month
            fail_summary = pd.concat(
                [fail_summary, site_kpi[site_kpi["Pass/Fail"] == "Fail"]],
                ignore_index=True
            )

            # Sort data by KPI in descending order
            site_kpi = site_kpi.sort_values(by="Site wise KPI", ascending=False)
            results[month] = site_kpi
        except KeyError as e:
            st.error(f"Error processing {month}: Missing required columns. {e}")
        except Exception as e:
            st.error(f"Unexpected error with {month}: {e}")

    return results, fail_summary

# Function to analyze fails
def analyze_fails(client, fail_summary):
    fail_summary["Month Order"] = fail_summary["Month"].apply(lambda m: months.index(m))
    fail_summary = fail_summary.sort_values(["Site ID", "Month Order"])

    # Calculate total fails for each site
    group_columns = ["Site ID", "RIO"] if client == "BL" else ["Site ID", "RIO", "STL_SC"]
    total_fails = (
        fail_summary.groupby(group_columns)
        .size()
        .reset_index(name="Total_Fails")
        .query("Total_Fails >= 5")
        .sort_values(by="Total_Fails", ascending=False)
    )

    # Identify consecutive streaks
    fail_summary["Consecutive Group"] = (
        fail_summary.groupby("Site ID")["Month Order"].diff().fillna(1).ne(1).cumsum()
    )
    streaks = (
        fail_summary.groupby(["Site ID", "Consecutive Group"])
        .agg(
            Fail_Streak=("Month", "count"),
            Months=("Month", lambda x: ", ".join(x)),
            RIO=("RIO", "first"),
            STL_SC=("STL_SC", "first") if client == "GP" else None
        )
        .reset_index()
    )
    consecutive_fails = streaks[streaks["Fail_Streak"] >= 3].drop(columns=["Consecutive Group"])
    consecutive_fails = consecutive_fails.sort_values(by="Fail_Streak", ascending=False)

    return total_fails, consecutive_fails

# Streamlit App
st.title("KPI Comparison Tool with Fail Analysis")

# Client selection
st.sidebar.header("Client Selection")
client = st.sidebar.selectbox("Select Client", ["GP", "BL"], index=0)

# Assign thresholds based on client
thresholds = GP_THRESHOLDS if client == "GP" else BL_THRESHOLDS

# Initialize month names
months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Step 1: File upload
month_data = {}
file_type = "xlsb" if client == "BL" else "xlsx"
st.sidebar.header("Upload Files")
for month in months:
    st.sidebar.subheader(f"{month} (Threshold: {thresholds[month]})")
    uploaded_file = st.sidebar.file_uploader(f"Upload {month} File ({file_type})", type=[file_type], key=month)
    month_data[month] = uploaded_file

# Step 2: Process data when "Process" is clicked
if st.button("Process Files"):
    if all(data is None for data in month_data.values()):
        st.warning("Please upload at least one file!")
    else:
        results, fail_summary = process_files(client, month_data, thresholds)
        if results:
            # Analyze fails
            total_fails, consecutive_fails = analyze_fails(client, fail_summary)

            # Display tables
            st.subheader("Sites with Total KPI Failures (5 or More)")
            if not total_fails.empty:
                st.write(total_fails)
            else:
                st.write("No sites with 5 or more total failures.")

            st.subheader("Sites with 3 or More Consecutive Month Failures")
            if not consecutive_fails.empty:
                st.write(consecutive_fails)
            else:
                st.write("No sites with 3 or more consecutive month failures.")

            # Export to Excel
            with pd.ExcelWriter("KPI_Results_with_Analysis.xlsx", engine="openpyxl") as writer:
                for month, df in results.items():
                    df.to_excel(writer, sheet_name=month, index=False)
                total_fails.to_excel(writer, sheet_name="Total_Failures", index=False)
                consecutive_fails.to_excel(writer, sheet_name="Consecutive_Fails", index=False)

            with open("KPI_Results_with_Analysis.xlsx", "rb") as f:
                st.download_button("Download Results", data=f, file_name="KPI_Results_with_Analysis.xlsx")
        else:
            st.warning("No files were processed.")
