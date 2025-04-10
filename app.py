import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from io import BytesIO


def generate_formatted_excel(df):
    # Extract metadata from first row
    start_date = df['PlannedStart'].iloc[0]
    end_date = df['PlannedEnd'].iloc[0]
    location = df['Location'].iloc[0]
    outage_type = df['OnLine/Outage'].iloc[0]

    # Extract Business Affected
    business_affected = df['BusinessAffected'].dropna().unique()

    # Filter and parse BC apps with RelationType = Direct
    bc_direct_raw = df['BC'].dropna().tolist()
    direct_apps = []
    for item in bc_direct_raw:
        if "(RelationType = Direct)" in item:
            app_name = item.split("(RelationType = Direct)")[0].strip().rstrip(',')
            direct_apps.append(app_name)

    trading_apps = sorted([app for app in direct_apps if app.startswith("ST")])
    other_apps = sorted([app for app in direct_apps if not app.startswith("ST")])

    # Count CIs and BCs
    total_cis = df['CI'].nunique()
    bc_count = len(direct_apps)
    non_bc_count = total_cis - bc_count

    # Create Excel workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formatted Output"
    bold = Font(bold=True)

    # Row 1 & 2: Planned Dates
    ws["A1"] = "Planned Start Date:"
    ws["A1"].font = bold
    ws["B1"] = str(start_date)

    ws["A2"] = "Planned End Date:"
    ws["A2"].font = bold
    ws["B2"] = str(end_date)

    # Row 3: Blank line
    ws.append([])

    # Row 4: Summary line
    summary = f"{location}, {outage_type}, {total_cis} CIs, {bc_count} BC, {non_bc_count} Non BC"
    ws.append([summary])

    # Row 5: Blank line
    ws.append([])

    # Row 6+: Business Affected
    for biz in business_affected:
        ws.append([biz])

    # Row after Business Affected: Empty line + Column C header
    ws.append([])
    ws.append(["", "", "Platform: FCI"])
    ws.append(["", "", "Trading assets in scope: Yes"])
    ws.append([])
    ws.append(["", "", "Trading BC Apps:"])

    for app in trading_apps:
        ws.append(["", "", app])

    ws.append(["", "", "Other BC Apps:"])
    for app in other_apps:
        ws.append(["", "", app])

    # Autofit columns
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# Streamlit UI
st.set_page_config(page_title="Change Formatter App")
st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("File uploaded successfully.")
        st.subheader("Preview of Uploaded Data")
        st.dataframe(df.head())

        excel_data = generate_formatted_excel(df)

        st.download_button(
            label="📥 Download Formatted Excel",
            data=excel_data,
            file_name="formatted_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
