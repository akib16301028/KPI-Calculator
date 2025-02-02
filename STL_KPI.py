import streamlit as st
import pandas as pd

# Streamlit app
st.title("Site Wise KPI Converter")

# Upload file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)

        # Check if "Site Wise KPI" column exists
        if "Site Wise KPI" in df.columns:
            # Convert percentage values to numeric format
            df["Site Wise KPI"] = df["Site wise KPI"].astype(str).str.rstrip("%").astype(float)

            # Display the updated DataFrame
            st.success("File processed successfully!")
            st.write(df)

            # Download the updated file
            output_file = "updated_kpi_data.xlsx"
            df.to_excel(output_file, index=False)
            st.download_button(
                label="Download Updated File",
                data=open(output_file, "rb").read(),
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("The file does not contain a 'Site Wise KPI' column.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload an Excel file to proceed.")
