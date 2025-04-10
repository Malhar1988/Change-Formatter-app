import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from io import BytesIO
import re

def generate_formatted_excel(df):
    # Extract planned dates and metadata from first row
    start_date = df['PlannedStart'].iloc[0]
    end_date = df['PlannedEnd'].iloc[0]
    location = df['Location'].iloc[0]
    outage_type = df['OnLine/Outage'].iloc[0]

    # Extract app names from BC column where value contains '(RelationType = Direct)'
    bc_apps_raw = df['BC'].dropna().tolist()
    trading_bc_apps = []
    other_bc_apps = []

    for item in bc_apps_raw:
        if '(RelationType = Direct)' in item:
            app_name = re.sub(r'\s*\(RelationType\s*=\s*Direct\)', '', item).strip()
            if app_name.startswith("ST"):
                trading_bc_apps.append(app_name)
            else:
                other_bc_apps.append(app_name)

    # Count metrics
    total_cis = df["CI"].nunique()
    bc_count = len(trading_bc_apps) + len(other_bc_apps)
    non_bc_count = total_cis - bc_count

    # Create workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formatted Output"
    bold = Font(bold=True)

    # Column A - Planned Dates & Summary
    ws["A1"] = "Planned Start Date:"
    ws["A1"].font = bold
    ws["B1"] = str(start_date)

    ws["A2"] = "Planned End Date:"
    ws["A2"].font = bold
    ws["B2"] = str(end_date)

    ws.append([])  # spacer

    summary = f"{location}, {outage_type}, {total_cis} CIs, {bc_count} BC [Direct], {non_bc_count} NON BC [Direct]"
    ws.append([summary])

    ws.append([])  # spacer
    ws.append(["Business Affected"])

    # Column C - Trading & Other BC Apps
    ws["C1"] = "Platform: FCI"
    ws["C2"] = "Trading assets in scope: Yes"
    ws["C3"] = ""  # spacer
    ws["C4"] = "Trading BC Apps:"
    for i, app in enumerate(trading_bc_apps, start=5):
        ws[f"C{i}"] = app

    # Leave one row then start Other BC Apps
    other_start = 5 + len(trading_bc_apps) + 1
    ws[f"C{other_start - 1}"] = "Other BC Apps:"
    for i, app in enumerate(other_bc_apps, start=other_start):
        ws[f"C{i}"] = app

    # Auto-fit column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    # Save to memory
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ----------------------------
# Streamlit UI
# ----------------------------

st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

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
