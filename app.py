import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from io import BytesIO

def generate_formatted_excel(df):
    # Extract basic info
    start_date = df['PlannedStart'].iloc[0]
    end_date = df['PlannedEnd'].iloc[0]
    location = df['Location'].iloc[0]
    outage_type = df['OnLine/Outage'].iloc[0]

    # BC App extraction
    direct_apps = df['BC'].dropna().astype(str)
    bc_direct_apps = direct_apps[direct_apps.str.contains(r"\(RelationType\s*=\s*Direct\)", case=False)]

    # Extract app names without (RelationType = Direct)
    clean_app_names = bc_direct_apps.str.extract(r"^(.*?)\s*\(RelationType\s*=\s*Direct\)")[0].dropna()

    # Split into ST and other apps
    trading_bc_apps = sorted([app for app in clean_app_names if app.startswith("ST-")])
    other_bc_apps = sorted([app for app in clean_app_names if not app.startswith("ST-")])

    # Count metrics
    total_cis = df["CI"].nunique()
    bc_count = len(clean_app_names)
    non_bc_count = total_cis - bc_count

    # Business groups
    business_groups = df["BusinessGroups"].dropna().unique()

    # Create Excel file
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

    # Row 3: Spacer
    ws.append([])

    # Row 4: Summary
    summary = f"{location}, {outage_type}, {total_cis} CIs, {bc_count} BC [Direct], {non_bc_count} NON BC [Direct]"
    ws.append([summary])

    # Row 5: Spacer
    ws.append([])

    # Row 6: Business Affected
    ws.append(["Business Affected"])
    for bg in business_groups:
        ws.append([bg])

    # Row: Spacer
    ws.append([])

    # Column C content
    ws["C1"] = "Platform: FCI"
    ws["C2"] = "Trading assets in scope: Yes"
    ws["C4"] = "Trading BC Apps:"
    row_trading = 5
    for app in trading_bc_apps:
        ws[f"C{row_trading}"] = app
        row_trading += 1

    row_other_start = row_trading + 1
    ws[f"C{row_other_start}"] = "Other BC Apps:"
    row_other = row_other_start + 1
    for app in other_bc_apps:
        ws[f"C{row_other}"] = app
        row_other += 1

    # Auto-fit columns
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    # Save to memory
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit App UI
st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.subheader("Preview of Uploaded Data")
    st.dataframe(df.head())

    formatted_excel = generate_formatted_excel(df)

    st.download_button(
        label="📥 Download Formatted Excel",
        data=formatted_excel,
        file_name="formatted_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

