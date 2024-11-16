import streamlit as st
import pandas as pd

# Function to process files and thresholds
def process_files(client, month_data, thresholds):
    results = {}
    fail_summary = pd.DataFrame()

    for month, data in month_data.items():
        if data is None:  # Skip if the file is not uploaded
            continue
        try:
            # Read the file and handle BL-specific binary format
            if client == "BL":
                sheet_data = pd.read_excel(
                    data,
                    sheet_name="Total Hour Calculation",
                    skiprows=2,  # Skip the first two rows for BL files
                    engine="pyxlsb"
                )
            else:
                sheet_data = pd.read_excel(
                    data,
                    sheet_name="Total Hour Calculation",
                    header=2  # Start reading from the third row for GP files
                )

            # Extract necessary columns based on client
            if client == "GP":
                site_kpi = sheet_data[["Site ID", "Site wise KPI", "RIO", "STL_SC"]]
            else:  # For BL
                site_kpi = sheet_data[["Generic ID", "Site Wise KPI", "RIO"]]
                site_kpi.rename(columns={"Generic ID": "Site ID", "Site Wise KPI": "Site wise KPI"}, inplace=True)

            # Skip rows where "Site wise KPI" is 0
            site_kpi = site_kpi[site_kpi["Site wise KPI"] > 0]

            # Add Threshold and Pass/Fail columns
            site_kpi["Threshold"] = thresholds[month]
            site_kpi["Pass/Fail"] = site_kpi["Site wise KPI"].apply(
                lambda x: "Pass" if x >= thresholds[month] else "Fail"
            )

            # Sort data by KPI in descending order
            site_kpi = site_kpi.sort_values(by="Site wise KPI", ascending=False)

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
    # Add a numerical order for months to sort
    fail_summary["Month Order"] = fail_summary["Month"].apply(lambda m: months.index(m))
    fail_summary = fail_summary.sort_values(["Site ID", "Month Order"])

    # Aggregation for total failures by site
    if "STL_SC" in fail_summary.columns:
        total_fails = (
            fail_summary.groupby("Site ID")
            .agg(
                Total_Fails=("Pass/Fail", "size"),
                RIO=("RIO", "first"),
                STL_SC=("STL_SC", "first")  # Include only if STL_SC exists
            )
            .reset_index()
            .query("Total_Fails >= 5")
            .sort_values("Total_Fails", ascending=False)
        )
    else:  # Skip STL_SC for BL
        total_fails = (
            fail_summary.groupby("Site ID")
            .agg(
                Total_Fails=("Pass/Fail", "size"),
                RIO=("RIO", "first")
            )
            .reset_index()
            .query("Total_Fails >= 5")
            .sort_values("Total_Fails", ascending=False)
        )

    # Calculate consecutive fail streaks
    fail_summary["Consecutive Group"] = (
        fail_summary.groupby("Site ID")["Month Order"].diff().fillna(1).ne(1).cumsum()
    )
    streaks = (
        fail_summary.groupby(["Site ID", "Consecutive Group"])
        .agg(
            Fail_Streak=("Month", "count"),
            Months=("Month", lambda x: ", ".join(x)),
            RIO=("RIO", "first")
        )
        .reset_index()
        .sort_values("Fail_Streak", ascending=False)
    )

    # Filter streaks with 3 or more consecutive fails
    consecutive_fails = streaks[streaks["Fail_Streak"] >= 3].drop(columns=["Consecutive Group"])

    return total_fails, consecutive_fails


# Streamlit App
st.title("Monthly KPI Comparison Tool with Fail Analysis")

# Client selection
client = st.selectbox("Select Client", ["GP", "BL"])

# Initialize months and thresholds
months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]
thresholds = {
    "GP": {
        "January": 99.65, "February": 99.65, "March": 99.65, "April": 99.40,
        "May": 99.40, "June": 99.40, "July": 99.55, "August": 99.55,
        "September": 99.55, "October": 99.65, "November": 99.65, "December": 99.65,
    },
    "BL": {
        "January": 0.9976, "February": 0.9960, "March": 0.9949, "April": 0.9995,
        "May": 0.9905, "June": 0.9955, "July": 0.9957, "August": 0.9965,
        "September": 0.9966, "October": 0.9970, "November": 0.9978, "December": 0.9977,
    }
}[client]

# Step 1: File upload
month_data = {}
st.sidebar.header("Upload Files")
for month in months:
    file_type = "Excel Binary" if client == "BL" else "Excel"
    uploaded_file = st.sidebar.file_uploader(f"Upload {month} File ({file_type})", type=["xlsb", "xlsx"], key=month)
    month_data[month] = uploaded_file

# Step 2: Process data
if st.button("Process Files"):
    if all(data is None for data in month_data.values()):
        st.warning("Please upload at least one file!")
    else:
        results, fail_summary = process_files(client, month_data, thresholds)
        if results:
            total_fails, consecutive_fails = analyze_fails(fail_summary)

            # Display results
            st.subheader("Sites with Total KPI Failures (5 or More)")
            st.write(total_fails)

            st.subheader("Sites with 3 or More Consecutive Month Failures")
            st.write(consecutive_fails)

            # Combine into a single workbook
            file_name = f"{client}_KPI_Results_with_Analysis.xlsx"
            with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
                for month, df in results.items():
                    df.to_excel(writer, sheet_name=month, index=False)
                total_fails.to_excel(writer, sheet_name="Total_Failures", index=False)
                consecutive_fails.to_excel(writer, sheet_name="Consecutive_Fails", index=False)

            with open(file_name, "rb") as f:
                st.download_button("Download Results", data=f, file_name=file_name)
        else:
            st.warning("No files were processed.")
