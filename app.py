import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO

def generate_formatted_excel(df):
    # Create a new workbook and select the active worksheet.
    wb = Workbook()
    ws = wb.active
    ws.title = "Output Final"
    
    # Define common styles.
    bold = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    
    # For illustration, we use only the first record from the dataframe.
    # (You can loop over all records with appropriate row offsets.)
    record = df.iloc[0]
    
    # --- Row 1: Planned Start Date details ---
    # Cell A1: Label
    ws['A1'] = "Planned Start Date:"
    ws['A1'].font = bold
    # Cell B1: Planned Start Date value
    planned_start = str(record['PlannedStart']) if pd.notna(record['PlannedStart']) else ""
    ws['B1'] = planned_start
    # (Optional: You could assign cell C1 if needed, e.g., a note or formatting detail.)
    
    # --- Row 2: Planned End Date details ---
    # Cell A2: Label for Planned End Date.
    ws['A2'] = "Planned End Date:"
    ws['A2'].font = bold
    # Cell B2: Planned End Date value.
    planned_end = str(record['PlannedEnd']) if pd.notna(record['PlannedEnd']) else ""
    ws['B2'] = planned_end
    
    # Cell C2: Additional details for row 2.
    # For example, we can put a summary of location and outage type.
    location = str(record['Location']) if pd.notna(record['Location']) else ""
    outage_type = str(record['OnLine/Outage']) if pd.notna(record['OnLine/Outage']) else ""
    summary = f"{location}, {outage_type}"
    ws['C2'] = summary
    ws['C2'].alignment = center_align

    # --- (Optional) Additional cell-by-cell assignments ---
    # For example, assign a summary line in row 3:
    # Cell A3 could have a merged cell for overall summary.
    total_cis = len(str(record['CI']).split(",")) if pd.notna(record.get('CI')) else 0
    # Process BC details (this is an example extraction; adjust per your logic)
    bc_apps = []
    if pd.notna(record.get('BC')):
        for item in str(record['BC']).split(","):
            item = item.strip()
            if "(RelationType = Direct)" in item:
                app_name = item.replace("(RelationType = Direct)", "").strip()
                bc_apps.append(app_name)
    bc_count = len(bc_apps)
    non_bc_count = total_cis - bc_count
    overall_summary = f"{total_cis} CIs, {bc_count} BC, {non_bc_count} Non BC"
    
    # Merge cells A3 to C3 and center the summary.
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=3)
    summary_cell = ws.cell(row=3, column=1, value=overall_summary)
    summary_cell.alignment = center_align

    # Save the workbook to a BytesIO stream.
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ----------------------
# Streamlit App UI
# ----------------------
st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Read the uploaded Excel file as a DataFrame.
        df = pd.read_excel(uploaded_file)
        st.subheader("Preview of Uploaded Data")
        st.dataframe(df.head())
        
        # Generate the formatted Excel file (cell-by-cell assignment).
        formatted_excel = generate_formatted_excel(df)
        
        # Provide a download button for the file.
        st.download_button(
            label="📥 Download Formatted Excel",
            data=formatted_excel,
            file_name="formatted_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
