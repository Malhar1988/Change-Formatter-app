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

    bold_font = Font(bold=True)
    row_num = 1

    for _, row in df.iterrows():
        # Planned Dates
        ws[f"A{row_num}"] = "Planned Start Date:"
        ws[f"A{row_num}"].font = bold_font
        ws[f"B{row_num}"] = str(row['PlannedStart'])
        row_num += 1

        ws[f"A{row_num}"] = "Planned End Date:"
        ws[f"A{row_num}"].font = bold_font
        ws[f"B{row_num}"] = str(row['PlannedEnd'])
        row_num += 1

        # Spacer
        row_num += 1

        # Summary line
        location = row['Location']
        outage_type = row['OnLine/Outage']
        total_cis = len(str(row['CI']).split(",")) if pd.notna(row['CI']) else 0
        bc_apps = []
        other_apps = []

        if pd.notna(row['BC']):
            for item in str(row['BC']).split(","):
                item = item.strip()
                if "(RelationType = Direct)" in item:
                    app_name = item.replace("(RelationType = Direct)", "").strip()
                    if app_name.startswith("ST"):
                        bc_apps.append(app_name)
                    else:
                        other_apps.append(app_name)

        bc_count = len(bc_apps)
        non_bc_count = total_cis - bc_count

        summary = f"{location}, {outage_type}, {total_cis} CIs, {bc_count} BC, {non_bc_count} Non BC"
        ws[f"A{row_num}"] = summary
        row_num += 1

        # Spacer
        row_num += 1

        # Business Groups
        if pd.notna(row['BusinessGroups']):
            for line in str(row['BusinessGroups']).split(","):
                ws[f"A{row_num}"] = line.strip()
                row_num += 1

        # Spacer
        row_num += 1

        # Change ID and Criticality in Column B
        ws[f"B{row_num}"] = str(row['ChangeId'])
        row_num += 1

        # Spacer
        row_num += 1

        # Third column with BC details
        ws[f"C{row_num}"] = "Platform: FCI"
        row_num += 1

        ws[f"C{row_num}"] = "Trading assets in scope: Yes"
        row_num += 1

        # Spacer
        row_num += 1

        ws[f"C{row_num}"] = "Trading BC Apps:"
        row_num += 1

        for app in bc_apps:
            ws[f"C{row_num}"] = app
            row_num += 1

        # Spacer
        row_num += 1

        ws[f"C{row_num}"] = "Other BC Apps:"
        row_num += 1

        for app in other_apps:
            ws[f"C{row_num}"] = app
            row_num += 1

        # Large Spacer between rows
        row_num += 3

    wb.save(output)
    output.seek(0)
    return output

# Streamlit App UI
st.title("Change Formatter App")
uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.subheader("Preview of Uploaded Data")
    st.dataframe(df.head())

    try:
        formatted_excel = generate_formatted_excel(df)
        st.download_button(
            label="📥 Download Formatted Excel",
            data=formatted_excel,
            file_name="formatted_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
