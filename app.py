import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

def ordinal(n):
    """Return the ordinal string for a number (e.g., 9 -> '9th')."""
    if 11 <= (n % 100) <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return str(n) + suffix

def format_date(val):
    """
    Convert a date string like '09-04-2025  05:00:00' to '9th April 2025'.
    If parsing fails, returns the original value as a string.
    """
    if pd.isna(val):
        return ""
    if isinstance(val, datetime):
        dt = val
    else:
        try:
            dt = datetime.strptime(str(val).strip(), "%d-%m-%Y %H:%M:%S")
        except Exception:
            return str(val)
    return f"{ordinal(dt.day)} {dt.strftime('%B')} {dt.year}"

def generate_formatted_excel(df):
    """
    Generate an Excel file using XlsxWriter with three columns:
    
      Column 1 (Record Details):
        - Line 1 (bold): Formatted date line: PlannedStart - PlannedEnd.
        - Line 2: Title.
        - Line 3: Summary (Location, OnLine/Outage, CI/BC/NONBC counts).
        - Line 4: BusinessGroups.
        Each line is separated by extra newlines.
      
      Column 2 (Change & Risk):
        - Line 1: F4F column content (if present after stripping headers)
                  or ChangeId with "/F4F" appended.
        - Line 2: Processed RiskLevel with "SHELL_" stripped.
      
      Column 3 (Trading Assets & BC Apps):
        - Line 1 (bold heading): "Trading assets in scope:" followed by Yes/No.
        - Line 2 (bold heading): "Trading BC Apps:" followed by a comma-separated list of trading apps.
        - Line 3 (bold heading): "Other BC Apps:" followed by a comma-separated list
          of other BC apps (or "No"/"None" as specified).
    """
    output = BytesIO()
    
    # Use a context manager so the workbook is properly saved
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Output Final")
        
        # Create formats with white background and black font.
        bold_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'bg_color': 'white', 'font_color': 'black'
        })
        normal_format = workbook.add_format({
            'text_wrap': True, 'bg_color': 'white', 'font_color': 'black'
        })
        
        # Set column widths.
        worksheet.set_column(0, 0, 50)  # Column A
        worksheet.set_column(1, 1, 30)  # Column B
        worksheet.set_column(2, 2, 50)  # Column C
        
        # Iterate over the DataFrame records.
        for idx, row in df.iterrows():
            row_num = idx  # XlsxWriter rows are zero-indexed.
            
            # -------- Column 1: Record Details --------
            planned_start = format_date(row['PlannedStart']) if pd.notna(row['PlannedStart']) else ""
            planned_end   = format_date(row['PlannedEnd'])   if pd.notna(row['PlannedEnd'])   else ""
            date_line = f"{planned_start} - {planned_end}".strip()
            title_line = str(row['Title']) if pd.notna(row['Title']) else ""
            location = str(row['Location']) if pd.notna(row['Location']) else ""
            online_outage = str(row['OnLine/Outage']) if pd.notna(row['OnLine/Outage']) else ""
            ci_count = len(str(row['CI']).split(",")) if pd.notna(row['CI']) else 0
            bc_count = len(str(row['BC']).split(",")) if pd.notna(row['BC']) else 0
            nonbc_count = len(str(row['NONBC']).split(",")) if pd.notna(row['NONBC']) else 0
            summary_line = f"{location}, {online_outage}, CI ({ci_count} CIs), BC ({bc_count} BC), NONBC ({nonbc_count} NONBC)".strip()
            business_groups_line = str(row['BusinessGroups']) if pd.notna(row['BusinessGroups']) else ""
            
            # Prepend an empty string so the first segment can be formatted.
            col1_parts = [
                "", bold_format, date_line,
                normal_format, "\n\n",
                normal_format, title_line,
                normal_format, "\n\n",
                normal_format, summary_line,
                normal_format, "\n\n",
                normal_format, business_groups_line
            ]
            
            # -------- Column 2: Change & Risk --------
            # Check for "F4F" column (after stripping header spaces).
            if 'F4F' in df.columns:
                f4f_val = str(row['F4F']) if pd.notna(row['F4F']) else ""
            else:
                change_id = str(row['ChangeId']) if pd.notna(row['ChangeId']) else ""
                f4f_val = f"{change_id}/F4F" if change_id else ""
            risk = str(row['RiskLevel']) if pd.notna(row['RiskLevel']) else ""
            if risk.upper().startswith("SHELL_"):
                risk = risk[6:]
            risk = risk.capitalize()
            
            col2_parts = [
                "", normal_format, f4f_val,
                normal_format, "\n\n",
                normal_format, risk
            ]
            
            # -------- Column 3: Trading Assets & BC Apps --------
            trading_apps = []
            other_apps = []
            if pd.notna(row['BC']):
                for item in str(row['BC']).split(","):
                    item = item.strip()
                    if "(RelationType = Direct)" in item:
                        app_name = item.replace("(RelationType = Direct)", "").strip()
                        if app_name.upper().startswith("ST"):
                            trading_apps.append(app_name)
                        else:
                            other_apps.append(app_name)
            trading_scope = "Yes" if trading_apps else "No"
            trading_bc_apps_content = ", ".join(trading_apps) if trading_apps else "None"
            if not trading_apps:
                other_bc_apps_content = "No"
            else:
                other_bc_apps_content = ", ".join(other_apps) if other_apps else "None"
            
            col3_parts = [
                "", bold_format, "Trading assets in scope: ",
                normal_format, trading_scope,
                normal_format, "\n\n",
                bold_format, "Trading BC Apps: ",
                normal_format, trading_bc_apps_content,
                normal_format, "\n\n",
                bold_format, "Other BC Apps: ",
                normal_format, other_bc_apps_content
            ]
            
            worksheet.write_rich_string(row_num, 0, *col1_parts)
            worksheet.write_rich_string(row_num, 1, *col2_parts)
            worksheet.write_rich_string(row_num, 2, *col3_parts)
        
        # The context manager ensures writer.save() is called.
    output.seek(0)
    return output

# ----------------------- Streamlit App UI -----------------------
st.title("Change Formatter App")
uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()  # Clean header names.
        st.subheader("Preview of Input Data")
        st.dataframe(df.head())
        formatted_excel = generate_formatted_excel(df)
        st.download_button(
            label="📥 Download Formatted Output",
            data=formatted_excel,
            file_name="output_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
