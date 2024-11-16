import streamlit as st
import pandas as pd

# Define thresholds for GP and BL
THRESHOLDS_GP = {
    "January": 99.65, "February": 99.65, "March": 99.65,
    "April": 99.40, "May": 99.40, "June": 99.40,
    "July": 99.55, "August": 99.55, "September": 99.55,
    "October": 99.65, "November": 99.65, "December": 99.65
}

THRESHOLDS_BL = {
    "January": 99.76, "February": 99.6, "March": 99.49,
    "April": 99.95, "May": 99.05, "June": 99.55,
    "July": 99.57, "August": 99.65, "September": 99.66,
    "October": 99.7, "November": 99.78, "December": 99.77
}

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Process GP and BL files and calculate KPI failures
def process_file(month_data, client_type):
    fail_summary = pd.DataFrame()
    thresholds = THRESHOLDS_GP if client_type == "GP" else THRESHOLDS_BL

    for month, data in month_data.items():
        if data is None:
            continue

        try:
            # Load the file (for BL use pyxlsb)
            if client_type == "BL":
                sheet_data = pd.read_excel(data, sheet_name="Total Hour Calculation", header=2, engine="pyxlsb")
                site_col, kpi_col = "Generic ID", "Site Wise KPI"
            else:
                sheet_data = pd.read_excel(data, sheet_name="Total Hour Calculation", header=2)
                site_col, kpi_col = "Site ID", "Site wise KPI"

            # Extract and clean data
            site_kpi = sheet_data[[site_col, kpi_col]].rename(
                columns={site_col: "Site ID", kpi_col: "Site wise KPI"}
            )

            # Remove rows where KPI is 0
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

        except KeyError as e:
            st.error(f"Error processing {month}: Missing required columns. {e}")
        except Exception as e:
            st.error(f"Unexpected error with {month}: {e}")

    return fail_summary

# Analyze failures: Sites with 5+ total failures and consecutive month failures
def analyze_fails(fail_summary):
    fail_summary["Month Order"] = fail_summary["Month"].apply(lambda m: MONTHS.index(m) + 1)

    fail_summary["Consecutive Group"] = (
        fail_summary.groupby("Site ID")["Month Order"].diff().fillna(1).ne(1).cumsum()
    )
    
    aggregation = {
        "Fail_Streak": ("Month", "count"),
        "Months": ("Month", lambda x: ", ".join(x)),
        "Site wise KPI": ("Site wise KPI", "first"),
    }

    streaks = (
        fail_summary.groupby(["Site ID", "Consecutive Group"])
        .agg(**aggregation)
        .reset_index()
    )

    consecutive_fails = streaks[streaks["Fail_Streak"] >= 3].drop(columns=["Consecutive Group"])
    consecutive_fails = consecutive_fails.sort_values(by="Fail_Streak", ascending=False)

    fail_count = fail_summary.groupby("Site ID").size().reset_index(name="Fail Count")
    fail_count = fail_count[fail_count["Fail Count"] >= 5]
    fail_count = fail_count.merge(fail_summary[["Site ID"]].drop_duplicates(), on="Site ID", how="left")

    return fail_count.sort_values(by="Fail Count", ascending=False), consecutive_fails

# Streamlit App for GP and BL Clients
st.title("KPI Analysis Tool")

# Ask user for client type
client_type = st.selectbox("Select Client", ["GP", "BL"])

# File uploads
st.sidebar.header(f"Upload Files for {client_type}")
month_data = {}
for month in MONTHS:
    month_data[month] = st.sidebar.file_uploader(f"Upload {month} File", type=["xlsb"], key=f"{client_type}_{month}")

if st.button("Process Files"):
    if all(file is None for file in month_data.values()):
        st.warning("Please upload at least one file!")
    else:
        fail_summary = process_file(month_data, client_type)
        if not fail_summary.empty:
            total_fails, consecutive_fails = analyze_fails(fail_summary)

            output_file = f"{client_type}_KPI_Analysis_Results.xlsx"
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                fail_summary.to_excel(writer, sheet_name="Fail Summary", index=False)
                total_fails.to_excel(writer, sheet_name="Total Failures", index=False)
                consecutive_fails.to_excel(writer, sheet_name="Consecutive Failures", index=False)

            st.subheader("Sites with Total KPI Failures (5 or More)")
            st.write(total_fails)

            st.subheader("Sites with 3 or More Consecutive Month Failures")
            st.write(consecutive_fails)

            with open(output_file, "rb") as f:
                st.download_button(
                    label="Download Results", data=f, file_name=output_file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("No failure data found for the uploaded files.")
