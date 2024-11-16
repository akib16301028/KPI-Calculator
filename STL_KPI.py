import streamlit as st
import pandas as pd
import os

# KPI thresholds
THRESHOLDS = {
    "GP": {
        "January": 99.65, "February": 99.65, "March": 99.65,
        "April": 99.40, "May": 99.40, "June": 99.40,
        "July": 99.55, "August": 99.55, "September": 99.55,
        "October": 99.65, "November": 99.65, "December": 99.65
    },
    "BL": {
        "January": 0.9976, "February": 0.996, "March": 0.9949,
        "April": 0.9995, "May": 0.9905, "June": 0.9955,
        "July": 0.9957, "August": 0.9965, "September": 0.9966,
        "October": 0.997, "November": 0.9978, "December": 0.9977
    }
}

# Initialize month names
MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Process files and calculate KPI failures
def process_files(client, month_data):
    thresholds = THRESHOLDS[client]
    results = {}
    fail_summary = pd.DataFrame()

    for month, data in month_data.items():
        if data is None:  # Skip if the file is not uploaded
            continue

        try:
            # Load file based on client type
            if client == "GP":
                sheet_data = pd.read_excel(data, sheet_name="Total Hour Calculation", header=2)
                site_col, kpi_col = "Site ID", "Site wise KPI"
            elif client == "BL":
                sheet_data = pd.read_excel(data, sheet_name="Total Hour Calculation", header=2, engine="pyxlsb")
                site_col, kpi_col = "Generic ID", "Site Wise KPI"

            # Extract relevant data
            site_kpi = sheet_data[[site_col, kpi_col, "RIO"]].rename(
                columns={site_col: "Site ID", kpi_col: "Site wise KPI"}
            )
            if client == "GP":
                site_kpi["STL_SC"] = sheet_data["STL_SC"]

            # Remove rows with Site wise KPI = 0
            site_kpi = site_kpi[site_kpi["Site wise KPI"] != 0]

            # Add threshold and pass/fail information
            site_kpi["Threshold"] = thresholds[month]
            site_kpi["Pass/Fail"] = site_kpi["Site wise KPI"].apply(
                lambda x: "Pass" if x >= thresholds[month] else "Fail"
            )
            site_kpi["Month"] = month

            # Add failing sites to the summary
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

# Analyze failures
def analyze_fails(client, fail_summary):
    # Map months to order for consecutive calculations
    fail_summary["Month Order"] = fail_summary["Month"].apply(lambda m: MONTHS.index(m) + 1)

    # Consecutive streaks
    fail_summary["Consecutive Group"] = (
        fail_summary.groupby("Site ID")["Month Order"].diff().fillna(1).ne(1).cumsum()
    )
    aggregation = {
        "Fail_Streak": ("Month", "count"),
        "Months": ("Month", lambda x: ", ".join(x)),
        "RIO": ("RIO", "first"),
    }

    if client == "GP":
        aggregation["STL_SC"] = ("STL_SC", "first")

    streaks = (
        fail_summary.groupby(["Site ID", "Consecutive Group"])
        .agg(**aggregation)
        .reset_index()
    )

    # Filter streaks of 3 or more consecutive fails
    consecutive_fails = streaks[streaks["Fail_Streak"] >= 3].drop(columns=["Consecutive Group"])
    consecutive_fails = consecutive_fails.sort_values(by="Fail_Streak", ascending=False)

    # Total fails
    fail_count = fail_summary.groupby("Site ID").size().reset_index(name="Fail Count")
    fail_count = fail_count[fail_count["Fail Count"] >= 5]
    fail_count = fail_count.merge(fail_summary[["Site ID", "RIO"]].drop_duplicates(), on="Site ID", how="left")

    if client == "GP":
        fail_count = fail_count.merge(fail_summary[["Site ID", "STL_SC"]].drop_duplicates(), on="Site ID", how="left")

    return fail_count.sort_values(by="Fail Count", ascending=False), consecutive_fails

# Streamlit App
st.title("KPI Analysis Tool")

# Choose client type
client = st.selectbox("Select Client", options=["GP", "BL"])

# Display thresholds for selected client
st.sidebar.header(f"KPI Thresholds for {client}")
thresholds = THRESHOLDS[client]
for month in MONTHS:
    st.sidebar.write(f"{month}: {thresholds[month]}")

# File uploads
st.sidebar.header(f"Upload Files for {client}")
month_data = {}
for month in MONTHS:
    month_data[month] = st.sidebar.file_uploader(f"Upload {month} File", type=["xlsx", "xlsb"], key=f"{client}_{month}")

# Process files on button click
if st.button("Process Files"):
    if all(file is None for file in month_data.values()):
        st.warning("Please upload at least one file!")
    else:
        results, fail_summary = process_files(client, month_data)
        if results:
            # Analyze fails
            total_fails, consecutive_fails = analyze_fails(client, fail_summary)

            # Save results to an Excel file
            output_file = f"{client}_KPI_Analysis_Result.xlsx"
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for month, df in results.items():
                    df.sort_values(by="Site wise KPI", ascending=False).to_excel(writer, sheet_name=month, index=False)
                total_fails.to_excel(writer, sheet_name="Total Failures", index=False)
                consecutive_fails.to_excel(writer, sheet_name="Consecutive Failures", index=False)

            # Display results
            st.subheader(f"Sites with Total KPI Failures (5 or More) for {client}")
            st.write(total_fails)

            st.subheader(f"Sites with 3 or More Consecutive Month Failures for {client}")
            st.write(consecutive_fails)

            # Download button
            with open(output_file, "rb") as f:
                st.download_button(
                    label="Download Results", data=f, file_name=output_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
