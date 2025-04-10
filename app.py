import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re

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
    if pd.isna(val) or val == " ":
        return ""
    if isinstance(val, datetime):
        dt = val
    else:
        try:
            dt = datetime.strptime(str(val).strip(), "%d-%m-%Y %H:%M:%S")
        except Exception:
            return str(val)
    return f"{ordinal(dt.day)} {dt.strftime('%B')} {dt.year}"

def build_summary(location, online_outage, ci_val, bc_val, nonbc_val):
    """
    Build a summary string with:
      - For CI: if count > 0: output e.g. "3 CIs"
      - For BC: if count > 0: output e.g. "2 BC"
      - For NONBC: if count > 0: output e.g. "1 NONBC"
    Other parts (location, online_outage) are joined with commas.
    If a field is empty, it is omitted.
    """
    parts = []
    if location.strip():
        parts.append(location.strip())
    if online_outage.strip():
        parts.append(online_outage.strip())
    # For CI.
    ci_count = len(str(ci_val).split(",")) if str(ci_val).strip() != "" else 0
    if ci_count > 0:
        parts.append(f"{ci_count} CIs")
    # For BC.
    bc_count = len(str(bc_val).split(",")) if str(bc_val).strip() != "" else 0
    if bc_count > 0:
        parts.append(f"{bc_count} BC")
    # For NONBC.
    nonbc_count = len(str(nonbc_val).split(",")) if str(nonbc_val).strip() != "" else 0
    if nonbc_count > 0:
        parts.append(f"{nonbc_count} NONBC")
    return ", ".join(parts)

# --- Main Function to Generate the Final Formatted Excel File ---
def generate_formatted_excel(df):
    output = BytesIO()
    
    # Replace blank values in individual cells with a single space.
    df.fillna(" ", inplace=True)
    
    # Use ExcelWriter with XlsxWriter in a context manager.
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Output Final")
        
        # Create formats with white background and black text.
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
        
        # Write a test header row (optional).
        worksheet.write_rich_string(0, 0,
                                    " ", bold_format, "Test Header Row",
                                    normal_format, " (Row 1)")
        st.write("Test header row written at Excel row 1.")
        
        # Process each record and write one output row per record, starting at Excel row 2.
        output_row = 1
        for idx, row in df.iterrows():
            # ---------------- Column 1: Record Details ----------------
            planned_start = format_date(row.get('PlannedStart', " "))
            planned_end   = format_date(row.get('PlannedEnd', " "))
            date_line = f"{planned_start} - {planned_end}".strip()
            title_line = str(row.get('Title', " ")).strip()
            
            location = str(row.get('Location', " "))
            online_outage = str(row.get('OnLine/Outage', " "))
            ci_val = row.get('CI', " ")
            bc_val = row.get('BC', " ")
            nonbc_val = row.get('NONBC', " ")
            
            summary_line = build_summary(location, online_outage, ci_val, bc_val, nonbc_val)
            business_groups_line = str(row.get('BusinessGroups', " ")).strip()
            
            col1_parts = [
                " ", bold_format, date_line,
                normal_format, "\n\n",
                normal_format, title_line,
                normal_format, "\n\n",
                normal_format, summary_line,
                normal_format, "\n\n",
                normal_format, business_groups_line
            ]
            
            # ---------------- Column 2: Change & Risk ----------------
            # We need both ChangeId and F4F.
            change_id = str(row.get('ChangeId', " ")).strip()
            f4f_val = str(row.get('F4F', " ")).strip()
            # Concatenate with a slash between them.
            if change_id != " " and f4f_val != " ":
                change_text = f"{change_id}/{f4f_val}"
            elif change_id != " ":
                change_text = change_id
            elif f4f_val != " ":
                change_text = f4f_val
            else:
                change_text = " "
            risk = str(row.get('RiskLevel', " ")).strip()
            if risk.upper().startswith("SHELL_"):
                risk = risk[6:]
            risk = risk.capitalize().strip()
            # Use a single newline between the two lines.
            col2_parts = [
                " ", normal_format, change_text,
                normal_format, "\n",
                normal_format, risk
            ]
            
            # ---------------- Column 3: Trading Assets & BC Apps ----------------
            trading_apps = []
            other_apps = []
            # Determine the separator: if bc_val contains newline, split on newline; otherwise on comma.
            bc_content = str(bc_val).strip()
            if "\n" in bc_content:
                items = bc_content.split("\n")
            else:
                items = bc_content.split(",")
            for item in items:
                item = item.strip()
                # Ignore if the item starts with "(" (e.g., "(12 BC)").
                if item.startswith("("):
                    continue
                # Only consider items with "(RelationType = Direct)".
                if "(RelationType = Direct)" in item:
                    app_name = item.replace("(RelationType = Direct)", "").strip()
                    if app_name.upper().startswith("ST"):
                        trading_apps.append(app_name)
                    else:
                        other_apps.append(app_name)
            trading_scope = "Yes" if trading_apps else "No"
            trading_bc_apps_text = ", ".join(trading_apps) if trading_apps else "None"
            # For other apps: if there are no trading apps, output "No". 
            # Otherwise, if there are trading apps and other apps exist then list them; if no other apps, output "No".
            if not trading_apps:
                other_bc_apps_text = "No"
            else:
                other_bc_apps_text = ", ".join(other_apps) if other_apps else "No"
            
            col3_parts = [
                " ", bold_format, "Trading assets in scope: ",
                normal_format, trading_scope,
                normal_format, "\n\n",
                bold_format, "Trading BC Apps: ",
                normal_format, trading_bc_apps_text,
                normal_format, "\n\n",
                bold_format, "Other BC Apps: ",
                normal_format, other_bc_apps_text
            ]
            
            # Log diagnostics.
            st.write(f"Record {idx} -> Excel Row {output_row+1}:")
            st.write("  Col1:", date_line, "|", title_line, "|", summary_line, "|", business_groups_line)
            st.write("  Col2:", change_text, "|", risk)
            st.write("  Col3:", trading_scope, "| Trading Apps:", trading_bc_apps_text, "| Other Apps:", other_bc_apps_text)
            
            worksheet.write_rich_string(output_row, 0, *col1_parts)
            worksheet.write_rich_string(output_row, 1, *col2_parts)
            worksheet.write_rich_string(output_row, 2, *col3_parts)
            
            output_row += 1
        
    output.seek(0)
    return output

# ----------------------- Streamlit App UI -----------------------
st.title("Change Formatter App")
uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        # Replace only blank cells with a space.
        df.fillna(" ", inplace=True)
        df.columns = df.columns.str.strip()  # Clean header names.
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
