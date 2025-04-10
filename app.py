import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from io import BytesIO

def generate_formatted_excel(df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formatted Output"
    bold = Font(bold=True)

    # Write each change entry
    row_pointer = 1  # Track where we are writing in Excel
    for idx, row in df.iterrows():
        # Extract fields
        start_date = row.get('PlannedStart', '')
        end_date = row.get('PlannedEnd', '')
        location = row.get('Location', '')
        outage_type = row.get('OnLine/Outage', '')
        business_affected = row.get('BusinessGroups', '')
        ci_count = 1 if pd.notna(row.get('CI')) else 0

        # Extract BC apps with (RelationType = Direct)
        bc_raw = row.get('BC', '')
        direct_apps = []
        if isinstance(bc_raw, str):
            entries = [entry.strip() for entry in bc_raw.split(',')]
            for entry in entries:
                if '(RelationType = Direct)' in entry:
                    app_name = entry.replace('(RelationType = Direct)', '').strip()
                    direct_apps.append(app_name)

        # Separate Trading BC Apps and Other BC Apps
        trading_apps = [app for app in direct_apps if app.startswith('ST')]
        other_apps = [app for app in direct_apps if not app.startswith('ST')]

        # Planned dates
        ws[f"A{row_pointer}"] = "Planned Start Date:"
        ws[f"A{row_pointer}"].font = bold
        ws[f"B{row_pointer}"] = str(start_date)

        ws[f"A{row_pointer+1}"] = "Planned End Date:"
        ws[f"A{row_pointer+1}"].font = bold
        ws[f"B{row_pointer+1}"] = str(end_date)

        # Spacer
        row_pointer += 2
        ws.append([])
        row_pointer += 1

        # Summary
        summary = f"{location}, {outage_type}, {ci_count} CIs, {len(trading_apps)} BC, {ci_count - len(trading_apps)} Non BC"
        ws[f"A{row_pointer}"] = summary

        row_pointer += 1
        ws.append([])
        row_pointer += 1

        # Business Affected
        ws[f"A{row_pointer}"] = "Business Affected"
        row_pointer += 1
        ws[f"A{row_pointer}"] = business_affected

        # Spacer before trading section
        row_pointer += 1
        ws.append([])
        row_pointer += 1

        # Trading Apps Section
        ws[f"C{row_pointer}"] = "Trading assets in scope: Yes"
        row_pointer += 1
        ws.append([])
        row_pointer += 1

        ws[f"C{row_pointer}"] = "Trading BC Apps:"
        for app in trading_apps:
            row_pointer += 1
            ws[f"C{row_pointer}"] = app

        # Spacer before other BC apps
        row_pointer += 1
        ws.append([])
        row_pointer += 1

        ws[f"C{row_pointer}"] = "Other BC Apps:"
        for app in other_apps:
            row_pointer += 1
            ws[f"C{row_pointer}"] = app

        # Add a spacer after each change
        row_pointer += 2
        ws.append([])
        row_pointer += 1

    # Auto-fit column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 5

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -----------------------
# Streamlit app interface
# -----------------------
st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Preview of Uploaded Data")
    st.dataframe(df.head())

    excel_data = generate_formatted_excel(df)

    st.download_button(
        label="📥 Download Formatted Excel",
        data=excel_data,
        file_name="formatted_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
