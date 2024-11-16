import streamlit as st
import pandas as pd

# Function to process files and thresholds
def process_files(month_data, thresholds):
    results = {}
    fail_summary = pd.DataFrame()

    for month, data in month_data.items():
        if data is None:  # Skip if the file is not uploaded
            continue
        try:
            # Read the required sheet starting from the correct header row
            sheet_data = pd.read_excel(data, sheet_name="Total Hour Calculation", header=2)
            # Extract necessary columns
            site_kpi = sheet_data[["Site ID", "Site wise KPI", "RIO", "STL_SC"]]
            # Fill missing KPI values with a placeholder (e.g., 0 or leave NaN)
            site_kpi["Site wise KPI"] = site_kpi["Site wise KPI"].fillna(0)

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

            results[month] = site_kpi
        except KeyError as e:
            st.error(f"Error processing {month}: Missing required columns. {e}")
        except Exception as e:
            st.error(f"Unexpected error with {month}: {e}")

    return results, fail_summary

# Function to analyze fails
def analyze_fails(fail_summary):
    # Map months to their order
    fail_summary["Month Order"] = fail_summary["Month"].apply(lambda m: months.index(m))
    fail_summary = fail_summary.sort_values(["Site ID", "Month Order"])

    # Calculate total fails for each site
    total_fails = (
        fail_summary.groupby("Site ID")
        .size()
        .reset_index(name="Total Fails")
        .query("`Total Fails` >= 5")
    )

    # Calculate the gap between consecutive failures
    fail_summary["Consecutive Group"] = (
        fail_summary.groupby("Site ID")["Month Order"].diff().fillna(1).ne(1).cumsum()
    )

    # Group by Site ID and Consecutive Group to identify consecutive streaks
    streaks = (
        fail_summary.groupby(["Site ID", "Consecutive Group"])
        .agg(
            Fail_Streak=("Month", "count"),
            Months=("Month", lambda x: ", ".join(x)),
            RIO=("RIO", "first"),
            STL_SC=("STL_SC", "first"),
        )
        .reset_index()
    )

    # Filter streaks for 3 or more consecutive months
    consecutive_fails = streaks[streaks["Fail_Streak"] >= 3].drop(columns=["Consecutive Group"])

    return total_fails, consecutive_fails

# Streamlit App
st.title("Monthly KPI Comparison Tool with Fail Analysis")

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
        month_data[month] = uploaded_file
        thresholds[month] = st.sidebar.number_input(f"{month} KPI Threshold", min_value=0.0, value=0.0)
    else:
        st.sidebar.write(f"Ignoring {month}.")

# Step 2: Process data when "Process" is clicked
if st.button("Process Files"):
    # Ensure at least one file is uploaded
    if all(data is None for data in month_data.values()):
        st.warning("Please upload at least one file!")
    else:
        results, fail_summary = process_files(month_data, thresholds)
        if results:
            # Analyze fails
            total_fails, consecutive_fails = analyze_fails(fail_summary)

            # Display the table for sites with 5 or more total fails
            st.subheader("Sites with Total KPI Failures (5 or More)")
            if not total_fails.empty:
                st.write(total_fails)
            else:
                st.write("No sites with 5 or more total failures.")

            # Display the table for 3 or more consecutive month failures
            st.subheader("Sites with 3 or More Consecutive Month Failures")
            if not consecutive_fails.empty:
                st.write(consecutive_fails)
            else:
                st.write("No sites with 3 or more consecutive month failures.")

            # Combine results into a single Excel workbook
            with pd.ExcelWriter("KPI_Results_with_Analysis.xlsx", engine="openpyxl") as writer:
                for month, df in results.items():
                    df.to_excel(writer, sheet_name=month, index=False)

                # Add summary sheets for streaks
                total_fails.to_excel(writer, sheet_name="Total_Failures", index=False)
                consecutive_fails.to_excel(writer, sheet_name="Consecutive_Fails", index=False)

            # Provide download button
            with open("KPI_Results_with_Analysis.xlsx", "rb") as f:
                st.download_button("Download Results", data=f, file_name="KPI_Results_with_Analysis.xlsx")
        else:
            st.warning("No files were processed.")
