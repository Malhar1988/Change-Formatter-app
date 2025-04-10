import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from io import BytesIO

def generate_formatted_excel(df):
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formatted Output"

    bold = Font(bold=True)

    row_idx = 1

    for _, row in df.iterrows():
        # Planned start and end dates
        ws[f"A{row_idx}"] = "Planned Start Date:"
        ws[f"A{row_idx}"].font = bold
        ws[f"B{row_idx}"] = str(row['PlannedStart'])
        row_idx += 1

        ws[f"A{row_idx}"] = "Planned End Date:"
        ws[f"A{row_idx}"].font = bold
        ws[f"B{row_idx}"] = str(row['PlannedEnd'])
        row_idx += 1

        # Spacer
        row_idx += 1

        # Summary line
        location = row['Location']
        outage_type = row['OnLine/Outage']
        total_cis = len(str(row['CI']).split(",")) if pd.notna(row['CI']) else 0

        # Extract BC Direct apps from BC column
        bc_apps_raw = str(row['BC']).split(",") if pd.notna(row['BC']) else []
        bc_direct_apps = [app.split(" (")[0].strip() for app in bc_apps_raw if "(RelationType = Direct)" in app]
        trading_bc_apps = [app for app in bc_direct_apps if app.startswith("ST")]
        other_bc_apps = [app for app in bc_direct_apps if not app.startswith("ST")]

        bc_count = len(bc_direct_apps)
        non_bc_count = total_cis - bc_count

        ws[f"A{row_idx}"] = f"{location}, {outage_type}, {total_cis} CIs, {bc_count} BC, {non_bc_count} Non BC"
        row_idx += 1

        # Spacer
        row_idx += 1

        # Business affected
        if pd.notna(row['BusinessGroups']):
            for line in str(row['BusinessGroups']).split(","):
                ws[f"A{row_idx}"] = line.strip()
                row_idx += 1

        # Column B - Change Request
        ws[f"B{row_idx - 1}"] = str(row['Change Request'])

        # Column C content
        ws[f"C{row_idx - 2}"] = "Platform: FCI"
        row_idx += 1
        ws[f"C{row_idx - 1}"] = "Trading assets in scope: Yes"
        row_idx += 1
        ws[f"C{row_idx - 1}"] = ""
        row_idx += 1
        ws[f"C{row_idx - 1}"] = "Trading BC Apps: " + ", ".join(trading_bc_apps)
        row_idx += 1
        ws[f"C{row_idx - 1}"] = "Other BC Apps: " + ", ".join(other_bc_apps)

        # Spacer between records
        row_idx += 2

    # Auto-fit columns
    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    wb.save(output)
    output.seek(0)
    return output

# ----------------------------
# Streamlit App
# ----------------------------

st.title("Change Formatter App")
uploaded_file = st.file_uploader("Upload your Changes.xlsx file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Preview of Uploaded Data")
    st.dataframe(df.head())

    try:
        excel_data = generate_formatted_excel(df)
        st.download_button(
            label="📥 Download Formatted Excel",
            data=excel_data,
            file_name="formatted_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
