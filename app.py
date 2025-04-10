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
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
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

# --- Function to Generate the Excel File ---
def generate_formatted_excel(df):
    output = BytesIO()
    # Use a context manager for ExcelWriter. No need to call save() explicitly.
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Output Final")
        
        # Create formats with explicit white background and black font.
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
        
        st.write("Generating output for", len(df), "rows...")
        
        # -------- Write a hard-coded test row at row 0 --------
        worksheet.write_rich_string(0, 0, "", bold_format, "Test Bold", normal_format, " Test Normal")
        st.write("Test row written at row 0.")
        
        # -------- Write processed rows starting at row 1 --------
        output_row = 1
        for idx, row in df.iterrows():
            # ---- Column 1: Record Details ----
            planned_start = format_date(row.get('PlannedStart', ''))
            planned_end   = format_date(row.get('PlannedEnd', ''))
            date_line = f"{planned_start} - {planned_end}".strip()
            
            title_line = str(row.get('Title', ''))
            location = str(row.get('Location', ''))
            online_outage = str(row.get('OnLine/Outage', ''))
            ci_count = len(str(row.get('CI', '')).split(",")) if pd.notna(row.get('CI', '')) else 0
            bc_count = len(str(row.get('BC', '')).split(",")) if pd.notna(row.get('BC', '')) else 0
            nonbc_count = len(str(row.get('NONBC', '')).split(",")) if pd.notna(row.get('NONBC', '')) else 0
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
            
            # ---- Column 2: Change & Risk ----
            if 'F4F' in df.columns:
                f4f_val = str(row.get('F4F', ''))
            else:
                change_id = str(row.get('ChangeId', ''))
                f4f_val = f"{change_id}/F4F" if change_id else ""
            risk = str(row.get('RiskLevel', ''))
            if risk.upper().startswith("SHELL_"):
                risk = risk[6:]
            risk = risk.capitalize()
            
            col2_parts = [
                "", normal_format, f4f_val,
                normal_format, "\n\n",
                normal_format, risk
            ]
            
            # ---- Column 3: Trading Assets & BC Apps ----
            trading_apps = []
            other_apps = []
            bc_val = row.get('BC', '')
            if pd.notna(bc_val):
                for item in str(bc_val).split(","):
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
            
            # Log diagnostics in Streamlit.
            st.write(f"Row {idx} written to output row {output_row}:")
            st.write("  Col1 ->", date_line, "|", title_line)
            st.write("  Col2 ->", f4f_val, "|", risk)
            st.write("  Col3 ->", trading_scope, "|", trading_bc_apps_content, "|", other_bc_apps_content)
            
            worksheet.write_rich_string(output_row, 0, *col1_parts)
            worksheet.write_rich_string(output_row, 1, *col2_parts)
            worksheet.write_rich_string(output_row, 2, *col3_parts)
            
            output_row += 1
        
        # The context manager closes and finalizes the file automatically.
    
    # Move to the start of the BytesIO stream.
    output.seek(0)
    return output

# ----------------------- Streamlit App UI -----------------------
st.title("Change Formatter App")
uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()  # Clean header names.
        st.write("DataFrame shape:", df.shape)
        st.write("DataFrame columns:", df.columns.tolist())
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
