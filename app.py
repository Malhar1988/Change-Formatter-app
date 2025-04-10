import streamlit as st
import pandas as pd
from openpyxl import Workbook
from io import BytesIO

def generate_formatted_excel(df):
    """
    For each record in the input dataframe (one row),
    build a three-column output where:
      - Column 1 combines PlannedStart, PlannedEnd, Title, a summary line, and BusinessGroups.
      - Column 2 combines ChangeId (with /F4F appended) and a processed RiskLevel.
      - Column 3 checks BC column for trading assets (apps starting with 'ST') and lists any other BC apps.
    All content for each record is built as multi-line text.
    """
    
    # Create a new workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Output Final"
    
    # Process each record (each row in df)
    for idx, row in df.iterrows():
        # --- Column 1: Record Details ---
        # Line 1: PlannedStart - PlannedEnd
        planned_start = str(row['PlannedStart']) if pd.notna(row['PlannedStart']) else ""
        planned_end   = str(row['PlannedEnd'])   if pd.notna(row['PlannedEnd'])   else ""
        line1 = f"{planned_start} - {planned_end}" if planned_start or planned_end else ""
        
        # Line 2: Title
        line2 = str(row['Title']) if pd.notna(row['Title']) else ""
        
        # Line 3: Summary with Location, OnLine/Outage, and counts from CI, BC and NONBC
        location = str(row['Location']) if pd.notna(row['Location']) else ""
        online_outage = str(row['OnLine/Outage']) if pd.notna(row['OnLine/Outage']) else ""
        
        # For counts, split the string by comma if not empty; else count 0
        ci_count = len(str(row['CI']).split(",")) if pd.notna(row['CI']) else 0
        bc_count_total = len(str(row['BC']).split(",")) if pd.notna(row['BC']) else 0
        nonbc_count = len(str(row['NONBC']).split(",")) if pd.notna(row['NONBC']) else 0
        line3 = f"{location}, {online_outage}, CI ({ci_count} CIs), BC ({bc_count_total} BC), NONBC ({nonbc_count} NONBC)"
        
        # Line 4: BusinessGroups
        line4 = str(row['BusinessGroups']) if pd.notna(row['BusinessGroups']) else ""
        
        col1_text = "\n".join([line for line in [line1, line2, line3, line4] if line])
        
        # --- Column 2: Change and Risk ---
        # First line: ChangeId with /F4F appended
        change_id = str(row['ChangeId']) if pd.notna(row['ChangeId']) else ""
        change_text = f"{change_id}/F4F" if change_id else ""
        
        # Second line: Process the RiskLevel value.
        risk_value = str(row['RiskLevel']) if pd.notna(row['RiskLevel']) else ""
        # Remove a possible SHELL_ prefix (case-insensitive) and capitalize the result.
        if risk_value.upper().startswith("SHELL_"):
            risk_value = risk_value[6:]
        risk_value = risk_value.capitalize()
        col2_text = "\n".join([change_text, risk_value]).strip()
        
        # --- Column 3: Trading Assets & Other BC Apps ---
        trading_apps = []
        other_apps = []
        if pd.notna(row['BC']):
            bc_items = [item.strip() for item in str(row['BC']).split(",")]
            for item in bc_items:
                if "(RelationType = Direct)" in item:
                    # Remove the relation tag and any extra spaces
                    app_name = item.replace("(RelationType = Direct)", "").strip()
                    if app_name.upper().startswith("ST"):
                        trading_apps.append(app_name)
                    else:
                        other_apps.append(app_name)
                else:
                    # If the relation tag is not found, consider it as an "other" app
                    other_apps.append(item)
                    
        trading_scope = "Trading assets in scope: Yes" if trading_apps else "Trading assets in scope: No"
        other_apps_line = "Other BC Apps: " + (", ".join(other_apps) if other_apps else "None")
        col3_text = "\n".join([trading_scope, other_apps_line])
        
        # --- Write to the output worksheet ---
        # Each record occupies one row; set columns A, B, and C.
        output_row = idx + 1  # Excel rows start at 1
        ws.cell(row=output_row, column=1, value=col1_text)
        ws.cell(row=output_row, column=2, value=col2_text)
        ws.cell(row=output_row, column=3, value=col3_text)
    
    # Save the workbook into a BytesIO stream so it can be downloaded
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --------------------
# Streamlit App UI
# --------------------
st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read the input file into a DataFrame
        df = pd.read_excel(uploaded_file)
        st.subheader("Preview of Input Data")
        st.dataframe(df.head())
        
        # Generate the formatted output Excel file
        formatted_excel = generate_formatted_excel(df)
        
        # Download button to save the output file as output_final.xlsx
        st.download_button(
            label="📥 Download Formatted Output",
            data=formatted_excel,
            file_name="output_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
