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

    # Loop through each row in the DataFrame
    for index, row in df.iterrows():
        # --- Planned Dates Section ---
        # Planned Start Date
        ws[f"A{row_num}"] = "Planned Start Date:"
        ws[f"A{row_num}"].font = bold_font
        planned_start = row['PlannedStart'] if pd.notna(row['PlannedStart']) else ""
        ws[f"B{row_num}"] = str(planned_start)
        row_num += 1

        # Planned End Date
        ws[f"A{row_num}"] = "Planned End Date:"
        ws[f"A{row_num}"].font = bold_font
        planned_end = row['PlannedEnd'] if pd.notna(row['PlannedEnd']) else ""
        ws[f"B{row_num}"] = str(planned_end)
        row_num += 1

        # Spacer
        row_num += 1

        # --- Summary Line Section ---
        # Retrieve required fields with safety in mind
        location = row['Location'] if pd.notna(row['Location']) else ""
        outage_type = row['OnLine/Outage'] if pd.notna(row['OnLine/Outage']) else ""
        
        # Count total CIs
        if pd.notna(row.get('CI')):
            total_cis = len(str(row['CI']).split(",")) 
        else:
            total_cis = 0

        # Process BC column: separate BC apps and other apps
        bc_apps = []
        other_apps = []
        if pd.notna(row.get('BC')):
            for item in str(row['BC']).split(","):
                item = item.strip()
                if "(RelationType = Direct)" in item:
                    app_name = item.replace("(RelationType = Direct)", "").strip()
                    # Distinguish BC apps by name (e.g., starting with "ST") – adjust as needed
                    if app_name.startswith("ST"):
                        bc_apps.append(app_name)
                    else:
                        other_apps.append(app_name)

        bc_count = len(bc_apps)
        non_bc_count = total_cis - bc_count

        # Create summary string
        summary = f"{location}, {outage_type}, {total_cis} CIs, {bc_count} BC, {non_bc_count} Non BC"
        ws[f"A{row_num}"] = summary
        row_num += 1

        # Spacer
        row_num += 1

        # --- Business Groups Section ---
        if pd.notna(row.get('BusinessGroups')):
            business_groups = str(row['BusinessGroups']).split(",")
            for bg in business_groups:
                ws[f"A{row_num}"] = bg.strip()
                row_num += 1

        # Spacer
        row_num += 1

        # --- Change ID Section (placed in column B) ---
        ws[f"B{row_num}"] = str(row['ChangeId']) if pd.notna(row.get('ChangeId')) else ""
        row_num += 1

        # Spacer
        row_num += 1

        # --- Column C Details Section ---
        # Platform Information
        ws[f"C{row_num}"] = "Platform: FCI"
        row_num += 1
        
        # Trading Assets in Scope
        ws[f"C{row_num}"] = "Trading assets in scope: Yes"
        row_num += 1

        # Spacer
        row_num += 1

        # Trading BC Apps and details
        ws[f"C{row_num}"] = "Trading BC Apps:"
        row_num += 1
        if bc_apps:
            for app in bc_apps:
                ws[f"C{row_num}"] = app
                row_num += 1
        else:
            # In case there are no BC apps
            ws[f"C{row_num}"] = "None"
            row_num += 1

        # Spacer
        row_num += 1

        ws[f"C{row_num}"] = "Other BC Apps:"
        row_num += 1
        if other_apps:
            for app in other_apps:
                ws[f"C{row_num}"] = app
                row_num += 1
        else:
            # In case there are no other BC apps
            ws[f"C{row_num}"] = "None"
            row_num += 1

        # Large spacer before processing the next record
        row_num += 3

    wb.save(output)
    output.seek(0)
    return output

# --- Streamlit App UI ---
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
