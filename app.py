import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# --- Helper Functions ---
def ordinal(n):
    """Return the ordinal string for a number (e.g., 9 -> '9th')."""
    if 11 <= (n % 100) <= 13:
        suffix = "th"
    else:
        suffix = {1:"st", 2:"nd", 3:"rd"}.get(n % 10, "th")
    return str(n) + suffix

def format_date(val):
    """
    Convert a date string like '09-04-2025  05:00:00' to '9th April 2025'.
    If conversion fails, returns the original value as a string.
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

# --- Main Function to Generate the Final Formatted Excel File ---
def generate_formatted_excel(df):
    output = BytesIO()
    
    # Use ExcelWriter with XlsxWriter in a context manager.
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Output Final")
        
        # Create formats with explicit white background and black text.
        bold_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'bg_color': 'white',
            'font_color': 'black',
            'font_size': 12
        })
        normal_format = workbook.add_format({
            'text_wrap': True,
            'bg_color': 'white',
            'font_color': 'black',
            'font_size': 12
        })
        
        # Set column widths.
        worksheet.set_column(0, 0, 50)  # Column A: Record Details
        worksheet.set_column(1, 1, 30)  # Column B: Change & Risk
        worksheet.set_column(2, 2, 50)  # Column C: Trading Assets & BC Apps
        
        st.write("Processing", len(df), "records...")
        
        # (Optional) Write a header row; you can remove or comment this block later.
        # Here, we write a test header row in Excel row 1.
        worksheet.write_rich_string(0, 0,
                                    "", bold_format, "Test Header Row",
                                    normal_format, " (Row 1)")
        st.write("Test header row written at Excel row 1.")
        
        # Process each record from the input file and write one output row per record.
        output_row = 1  # We start writing processed records from Excel row 2.
        for idx, row in df.iterrows():
            # ---------- Column 1: Record Details ----------
            planned_start = format_date(row.get('PlannedStart', ''))
            planned_end = format_date(row.get('PlannedEnd', ''))
            date_line = f"{planned_start} - {planned_end}".strip()
            
            title_line = str(row.get('Title', ''))
            
            location = str(row.get('Location', ''))
            online_outage = str(row.get('OnLine/Outage', ''))
            ci_val = row.get('CI', '')
            bc_val = row.get('BC', '')
            nonbc_val = row.get('NONBC', '')
            ci_count = len(str(ci_val).split(",")) if pd.notna(ci_val) and str(ci_val).strip() != "" else 0
            bc_count = len(str(bc_val).split(",")) if pd.notna(bc_val) and str(bc_val).strip() != "" else 0
            nonbc_count = len(str(nonbc_val).split(",")) if pd.notna(nonbc_val) and str(nonbc_val).strip() != "" else 0
            summary_line = f"{location}, {online_outage}, CI ({ci_count} CIs), BC ({bc_count} BC), NONBC ({nonbc_count} NONBC)".strip()
            
            business_groups_line = str(row.get('BusinessGroups', ''))
            
            col1_parts = [
                "", bold_format, date_line,
                normal_format, "\n\n",
                normal_format, title_line,
                normal_format, "\n\n",
                normal_format, summary_line,
                normal_format, "\n\n",
                normal_format, business_groups_line
            ]
            
            # ---------- Column 2: Change & Risk ----------
            # Check for the column "F4F" (which contains the change numbers).
            if 'F4F' in df.columns:
                change_val = str(row.get('F4F', ''))
            else:
                change_id = str(row.get('ChangeId', ''))
                change_val = f"{change_id}/F4F" if change_id != "" else ""
            col2_line1 = f"Change: {change_val}"
            
            risk = str(row.get('RiskLevel', ''))
            if risk.upper().startswith("SHELL_"):
                risk = risk[6:]
            risk = risk.capitalize()
            col2_line2 = f"Risk: {risk}"
            
            col2_parts = [
                "", normal_format, col2_line1,
                normal_format, "\n\n",
                normal_format, col2_line2
            ]
            
            # ---------- Column 3: Trading Assets & BC Apps ----------
            trading_apps = []
            other_apps = []
            if pd.notna(bc_val) and str(bc_val).strip() != "":
                for item in str(bc_val).split(","):
                    item = item.strip()
                    if "(RelationType = Direct)" in item:
                        app_name = item.replace("(RelationType = Direct)", "").strip()
                        if app_name.upper().startswith("ST"):
                            trading_apps.append(app_name)
                        else:
                            other_apps.append(app_name)
            trading_scope = "Yes" if trading_apps else "No"
            trading_bc_apps_text = ", ".join(trading_apps) if trading_apps else "None"
            # If no trading apps, then Other BC Apps should show "No"; otherwise, list other apps if present.
            if not trading_apps:
                other_bc_apps_text = "No"
            else:
                other_bc_apps_text = ", ".join(other_apps) if other_apps else "None"
            
            col3_parts = [
                "", bold_format, "Trading assets in scope: ",
                normal_format, trading_scope,
                normal_format, "\n\n",
                bold_format, "Trading BC Apps: ",
                normal_format, trading_bc_apps_text,
                normal_format, "\n\n",
                bold_format, "Other BC Apps: ",
                normal_format, other_bc_apps_text
            ]
            
            # Log diagnostics in Streamlit.
            st.write(f"Record {idx} -> Excel Row {output_row+1}:")
            st.write("  Col1:", date_line, "|", title_line)
            st.write("  Col2:", col2_line1, "|", col2_line2)
            st.write("  Col3:", trading_scope, "| Trading Apps:", trading_bc_apps_text, "| Other Apps:", other_bc_apps_text)
            
            # Write the rich strings to the worksheet.
            worksheet.write_rich_string(output_row, 0, *col1_parts)
            worksheet.write_rich_string(output_row, 1, *col2_parts)
            worksheet.write_rich_string(output_row, 2, *col3_parts)
            
            output_row += 1
        
        # (When the with-block exits, the file is finalized.)
    output.seek(0)
    return output

# ----------------------- Streamlit App UI -----------------------
st.title("Change Formatter App")
uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        # Clean header names (remove extra spaces)
        df.columns = df.columns.str.strip()
        st.write("DataFrame shape:", df.shape)
        st.write("Columns:", df.columns.tolist())
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
