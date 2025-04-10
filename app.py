import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO

def generate_formatted_excel(df):
    # Create a new workbook and worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Output Final"
    
    # Define some common styles
    bold = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    
    # Start at row 1
    current_row = 1
    
    # Process each record in the DataFrame
    for idx, row in df.iterrows():
        # ------------------------------------------------------
        # Cell-by-cell assignment for this record's block
        # ------------------------------------------------------
        
        # 1. Planned Start Date (e.g., cell A1 and B1)
        ws.cell(row=current_row, column=1, value="Planned Start Date:").font = bold
        # Convert the date (or value) to string if necessary
        planned_start = str(row['PlannedStart']) if pd.notna(row['PlannedStart']) else ""
        ws.cell(row=current_row, column=2, value=planned_start)
        current_row += 1  # move to next row
        
        # 2. Planned End Date (e.g., cell A2 and B2)
        ws.cell(row=current_row, column=1, value="Planned End Date:").font = bold
        planned_end = str(row['PlannedEnd']) if pd.notna(row['PlannedEnd']) else ""
        ws.cell(row=current_row, column=2, value=planned_end)
        current_row += 1
        
        # 3. Spacer row
        current_row += 1
        
        # 4. Summary line (merged across columns A to C)
        # Build summary text. (Custom calculation as per your data)
        location = str(row['Location']) if pd.notna(row['Location']) else ""
        outage_type = str(row['OnLine/Outage']) if pd.notna(row['OnLine/Outage']) else ""
        
        # Count total CIs (assuming comma-separated list in "CI")
        total_cis = len(str(row['CI']).split(",")) if pd.notna(row['CI']) else 0
        
        # Separate BC apps (if value in column "BC") and others
        bc_apps = []
        other_apps = []
        if pd.notna(row['BC']):
            for item in str(row['BC']).split(","):
                item = item.strip()
                if "(RelationType = Direct)" in item:
                    # Remove indicator text and strip spaces
                    app_name = item.replace("(RelationType = Direct)", "").strip()
                    # If the app name starts with "ST" assume it belongs to BC apps
                    if app_name.startswith("ST"):
                        bc_apps.append(app_name)
                    else:
                        other_apps.append(app_name)
        bc_count = len(bc_apps)
        non_bc_count = total_cis - bc_count
        
        summary_text = f"{location}, {outage_type}, {total_cis} CIs, {bc_count} BC, {non_bc_count} Non BC"
        # Merge cells from A to C in current row to host the summary text
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        summary_cell = ws.cell(row=current_row, column=1, value=summary_text)
        summary_cell.alignment = center_align
        current_row += 1
        
        # 5. Spacer row
        current_row += 1
        
        # 6. Business Groups (placed in column A, one per row)
        if pd.notna(row['BusinessGroups']):
            groups = str(row['BusinessGroups']).split(",")
            for group in groups:
                ws.cell(row=current_row, column=1, value=group.strip())
                current_row += 1
        
        # 7. Spacer row
        current_row += 1
        
        # 8. Change ID placed in Column B
        change_id = str(row['ChangeId']) if pd.notna(row['ChangeId']) else ""
        ws.cell(row=current_row, column=2, value=change_id)
        current_row += 1
        
        # 9. Spacer row
        current_row += 1
        
        # 10. Additional Details in Column C
        # a. Platform (cell in column C)
        ws.cell(row=current_row, column=3, value="Platform: FCI")
        current_row += 1
        
        # b. Trading assets in scope
        ws.cell(row=current_row, column=3, value="Trading assets in scope: Yes")
        current_row += 1
        
        # c. Spacer row
        current_row += 1
        
        # d. Trading BC Apps header
        bc_header = ws.cell(row=current_row, column=3, value="Trading BC Apps:")
        bc_header.font = bold
        current_row += 1
        
        # e. List each BC app cell by cell; if none, write "None"
        if bc_apps:
            for app in bc_apps:
                ws.cell(row=current_row, column=3, value=app)
                current_row += 1
        else:
            ws.cell(row=current_row, column=3, value="None")
            current_row += 1
        
        # f. Spacer row
        current_row += 1
        
        # g. Other BC Apps header
        other_bc_header = ws.cell(row=current_row, column=3, value="Other BC Apps:")
        other_bc_header.font = bold
        current_row += 1
        
        # h. List each Other BC app; if none, write "None"
        if other_apps:
            for app in other_apps:
                ws.cell(row=current_row, column=3, value=app)
                current_row += 1
        else:
            ws.cell(row=current_row, column=3, value="None")
            current_row += 1
        
        # i. Final Spacer before next record
        current_row += 3
        
    # Save the workbook to a BytesIO stream so it can be downloaded
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit user interface
st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Read the uploaded Excel file as a DataFrame
        df = pd.read_excel(uploaded_file)
        st.subheader("Preview of Uploaded Data")
        st.dataframe(df.head())
        
        # Create the formatted Excel file by calling the function cell by cell
        formatted_excel = generate_formatted_excel(df)
        
        # Provide a download button for the file
        st.download_button(
            label="📥 Download Formatted Excel",
            data=formatted_excel,
            file_name="formatted_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
