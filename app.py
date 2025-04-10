import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

def generate_formatted_excel(df):
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formatted Output"
    
    # Define common styles
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style="thin"),
                         right=Side(style="thin"),
                         top=Side(style="thin"),
                         bottom=Side(style="thin"))
    
    # Start row counter
    row_num = 1

    # Process each record from the DataFrame into a block of rows
    for idx, row in df.iterrows():
        # ===============================
        # Block Start: Planned Dates Section
        # ===============================
        # Row for Planned Start Date
        ws.cell(row=row_num, column=1, value="Planned Start Date:").font = bold_font
        planned_start = row['PlannedStart'] if pd.notna(row['PlannedStart']) else ""
        ws.cell(row=row_num, column=2, value=str(planned_start))
        row_num += 1

        # Row for Planned End Date
        ws.cell(row=row_num, column=1, value="Planned End Date:").font = bold_font
        planned_end = row['PlannedEnd'] if pd.notna(row['PlannedEnd']) else ""
        ws.cell(row=row_num, column=2, value=str(planned_end))
        row_num += 1

        # Spacer row
        row_num += 1
        
        # ===============================
        # Summary Line Section (merged across several columns)
        # ===============================
        location = row['Location'] if pd.notna(row['Location']) else ""
        outage_type = row['OnLine/Outage'] if pd.notna(row['OnLine/Outage']) else ""
        # Calculate total CIs from the CI column (comma-separated list)
        total_cis = len(str(row['CI']).split(",")) if pd.notna(row.get('CI')) else 0
        
        # Process BC details to separate BC apps and other apps
        bc_apps = []
        other_apps = []
        if pd.notna(row.get('BC')):
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
        ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=3)
        summary_cell = ws.cell(row=row_num, column=1, value=summary)
        summary_cell.alignment = center_align
        row_num += 1
        
        # Spacer row
        row_num += 1
        
        # ===============================
        # Business Groups Section
        # ===============================
        if pd.notna(row.get('BusinessGroups')):
            business_groups = str(row['BusinessGroups']).split(",")
            for bg in business_groups:
                ws.cell(row=row_num, column=1, value=bg.strip())
                row_num += 1
        
        # Spacer row
        row_num += 1
        
        # ===============================
        # Change ID Section (placed in Column B)
        # ===============================
        change_id = str(row['ChangeId']) if pd.notna(row.get('ChangeId')) else ""
        ws.cell(row=row_num, column=2, value=change_id)
        row_num += 1
        
        # Spacer row
        row_num += 1
        
        # ===============================
        # Additional Details in Column C
        # ===============================
        # Platform & Trading Assets lines
        ws.cell(row=row_num, column=3, value="Platform: FCI")
        row_num += 1
        ws.cell(row=row_num, column=3, value="Trading assets in scope: Yes")
        row_num += 1
        
        # Spacer row
        row_num += 1
        
        # Trading BC Apps
        ws.cell(row=row_num, column=3, value="Trading BC Apps:")
        ws.cell(row=row_num, column=3).font = bold_font
        row_num += 1
        if bc_apps:
            for app in bc_apps:
                ws.cell(row=row_num, column=3, value=app)
                row_num += 1
        else:
            ws.cell(row=row_num, column=3, value="None")
            row_num += 1
        
        # Spacer row
        row_num += 1
        
        # Other BC Apps
        ws.cell(row=row_num, column=3, value="Other BC Apps:")
        ws.cell(row=row_num, column=3).font = bold_font
        row_num += 1
        if other_apps:
            for app in other_apps:
                ws.cell(row=row_num, column=3, value=app)
                row_num += 1
        else:
            ws.cell(row=row_num, column=3, value="None")
            row_num += 1
        
        # Large Spacer between records
        row_num += 3

    # Optionally, adjust column widths to improve appearance
    for col in range(1, 4):
        max_length = 0
        col_letter = get_column_letter(col)
        for cell in ws[col_letter]:
            try:
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(output)
    output.seek(0)
    return output

# -------------------------
# Streamlit App UI
# -------------------------
st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Read the uploaded Excel file
        df = pd.read_excel(uploaded_file)
        st.subheader("Preview of Uploaded Data")
        st.dataframe(df.head())

        # Generate the formatted Excel file
        formatted_excel = generate_formatted_excel(df)
        st.download_button(
            label="📥 Download Formatted Excel",
            data=formatted_excel,
            file_name="formatted_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
