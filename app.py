import streamlit as st
import pandas as pd
from io import BytesIO

# --------- Main Processing Function ---------
def format_data(df):
    def format_change_details(row):
        try:
            start = pd.to_datetime(row['PlannedStart']).strftime('%-dth %B')
            end = pd.to_datetime(row['PlannedEnd']).strftime('%-dth %B %Y')
        except:
            start = end = "Date Missing"
        location = row['Location'] if pd.notnull(row['Location']) else ""
        return f"{start} - {end}\n{location}\n{row['Title']}"

    def format_change_number(row):
        change_id = row['ChangeId'] if pd.notnull(row['ChangeId']) else "Unknown"
        f4f = row['F4F'] if pd.notnull(row['F4F']) else "Unknown"
        risk = row['RiskLevel'].split('_')[-1].capitalize() if pd.notnull(row['RiskLevel']) else "Unknown"
        return f"{change_id}/{f4f}\n{risk}"

    def format_other_details(row):
        platform = f"Platform: {row['Location']}" if pd.notnull(row['Location']) else ""
        trading_scope = "Trading assets in scope:Yes" if 'Trading' in str(row['Title']) else "Trading assets in scope:No"
        other_bc_apps = "Other BC Apps: Yes" if pd.notnull(row['BC']) and "(0 BC)" not in str(row['BC']) else "Other BC Apps: No"
        return f"{platform}\n{trading_scope}\n\n{other_bc_apps}"

    return pd.DataFrame({
        'Change_Details': df.apply(format_change_details, axis=1),
        'Change_number': df.apply(format_change_number, axis=1),
        'Other Details': df.apply(format_other_details, axis=1)
    })


# --------- Streamlit UI ---------
st.set_page_config(page_title="Change Record Formatter", layout="centered")
st.title("📋 Excel Formatter for Change Records")
st.write("Upload an Excel file, and download a neatly formatted version for reporting.")

uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.success("✅ File uploaded and read successfully!")

        formatted_df = format_data(df)

        # Prepare file to download
        buffer = BytesIO()
        formatted_df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            label="📥 Download Formatted Excel",
            data=buffer,
            file_name="Formatted_Changes_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.dataframe(formatted_df.head())  # Preview first few rows

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
