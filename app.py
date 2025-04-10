import pandas as pd
import openpyxl
from openpyxl.styles import Font
from io import BytesIO
import streamlit as st

def generate_formatted_excel(df):
    # Create workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formatted Output"
    bold = Font(bold=True)

    for index, row in df.iterrows():
        # Planned dates
        start_date = str(row['PlannedStart'])
        end_date = str(row['PlannedEnd'])
        ws.append([f"Planned Start Date:"])
        ws["A" + str(ws.max_row)].font = bold
        ws.append([start_date])
        ws.append([f"Planned End Date:"])
        ws["A" + str(ws.max_row)].font = bold
        ws.append([end_date])

        ws.append([])  # Spacer

        # Summary
        location = row['Location']
        outage_type = row['OnLine/Outage']
        total_cis = len(str(row['CI']).split(","))
        bc_list = []
        if pd.notna(row['BC']):
            bc_items = str(row['BC']).split(",")
            for item in bc_items:
                if "(RelationType = Direct)" in item:
                    app_name = item.split(" (")[0].strip()
                    bc_list.append(app_name)
        bc_count = sum(1 for app in bc_list if app)
        non_bc_count = total_cis - bc_count
        summary = f"{location}, {outage_type}, {total_cis} CIs, {bc_count} BC, {non_bc_count} Non BC"
        ws.append([summary])

        ws.append([])  # Spacer

        # Business affected
        business_affected = row.get("BusinessGroups", "")
        ws.append([business_affected])

        # Column C: Trading and Other BC Apps
        ws.append([])  # Spacer
        ws.append([None, None, "Platform: FCI"])
        ws.append([None, None, "Trading assets in scope: Yes"])
        ws.append([])

        trading_apps = [app for app in bc_list if app.startswith("ST")]
        other_apps = [app for app in bc_list if not app.startswith("ST")]

        ws.append([None, None, "Trading BC Apps: " + ", ".join(trading_apps)])
        ws.append([None, None, "Other BC Apps: " + ", ".join(other_apps)])
        ws.append([])  # Spacer between records

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
# Streamlit UI begins here
# ----------------------------
st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

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
