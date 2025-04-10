import streamlit as st
import pandas as pd
import re
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
    If conversion fails, return the original value as a string.
    """
    if pd.isna(val) or str(val).strip() == "":
        return ""
    if isinstance(val, datetime):
        dt = val
    else:
        try:
            dt = datetime.strptime(str(val).strip(), "%d-%m-%Y %H:%M:%S")
        except Exception:
            return str(val)
    return f"{ordinal(dt.day)} {dt.strftime('%B')} {dt.year}"

def parse_first_int(text):
    """
    Return the first integer found in the text, or 0 if none.
    e.g. '21 under CI' returns 21.
    """
    m = re.search(r"\d+", text)
    return int(m.group(0)) if m else 0

def count_direct(text):
    """
    Count how many items (split by comma) contain "(RelationType = Direct)".
    """
    if not isinstance(text, str):
        return 0
    items = text.split(",")
    count = 0
    for item in items:
        if "(RelationType = Direct)" in item:
            count += 1
    return count

def build_summary(location, online_outage, ci_text, bc_text, nonbc_text):
    """
    Build the summary for Column 1.
    
    - For Location, if blank, it becomes an empty string.
    - For CI: extract the first integer (if any) and output "X CIs" if X > 0.
    - For BC: count only items with "(RelationType = Direct)"; if count > 0, output "X BC (Direct)".
    - For NONBC: count only direct items; if count > 0, output "X NON BC (Direct)".
    
    All parts are joined with ", ". For example, if Location is blank, online_outage is "Online",
    CI has "21 under CI", BC has 1 item but indirect only (so count 0), and NONBC has 2 items with one direct,
    the output will be: ", Online, 21 CIs, 1 NON BC (Direct)"
    """
    parts = []
    # For location: if the stripped value is empty, use empty string.
    loc = location.strip() if isinstance(location, str) else ""
    parts.append(loc)  # This may be empty
    parts.append(online_outage.strip())
    
    # CI count using parse_first_int
    ci_count = parse_first_int(ci_text)
    if ci_count > 0:
        parts.append(f"{ci_count} CIs")
        
    # BC count: count only direct items.
    bc_direct = count_direct(bc_text)
    if bc_direct > 0:
        parts.append(f"{bc_direct} BC (Direct)")
        
    # NONBC count: count only direct items.
    nonbc_direct = count_direct(nonbc_text)
    if nonbc_direct > 0:
        parts.append(f"{nonbc_direct} NON BC (Direct)")
    
    return ", ".join(parts)

def generate_formatted_excel(df):
    output = BytesIO()
    
    # Replace blank (NaN) values in individual cells with a space.
    df.fillna(" ", inplace=True)
    
    # Use ExcelWriter with XlsxWriter.
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Output Final")
        
        # Create formats.
        bold_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 
            'bg_color': 'white', 'font_color': 'black', 'font_size': 12
        })
        normal_format = workbook.add_format({
            'text_wrap': True, 
            'bg_color': 'white', 'font_color': 'black', 'font_size': 12
        })
        
        # Set column widths.
        worksheet.set_column(0, 0, 50)  # Column A
        worksheet.set_column(1, 1, 30)  # Column B
        worksheet.set_column(2, 2, 50)  # Column C
        
        # (Optional) If desired, write a header row. Here we choose to leave it blank.
        # Starting with processed rows at Excel row 1 (which appears as row 2 in Excel).
        output_row = 1
        
        for idx, row in df.iterrows():
            # -------------- Column 1: Record Details --------------
            planned_start = format_date(row.get('PlannedStart', " "))
            planned_end = format_date(row.get('PlannedEnd', " "))
            # If the two dates are identical, show just one date.
            if planned_start == planned_end:
                date_line = planned_start
            else:
                date_line = f"{planned_start} - {planned_end}".strip()
            title_line = str(row.get('Title', " ")).strip()
            
            location_val = str(row.get('Location', " "))
            online_val = str(row.get('OnLine/Outage', " "))
            ci_text = str(row.get('CI', " "))
            bc_text = str(row.get('BC', " "))
            nonbc_text = str(row.get('NONBC', " "))
            
            summary_line = build_summary(location_val, online_val, ci_text, bc_text, nonbc_text)
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
            
            # -------------- Column 2: Change & Risk --------------
            change_id = str(row.get('ChangeId', " ")).strip()
            f4f_val = str(row.get('F4F', " ")).strip()
            # Combine both with a slash—if both exist, otherwise use the one that exists.
            if change_id and f4f_val and change_id != " " and f4f_val != " ":
                change_text = f"{change_id}/{f4f_val}"
            elif change_id and change_id != " ":
                change_text = change_id
            elif f4f_val and f4f_val != " ":
                change_text = f4f_val
            else:
                change_text = " "
            risk_val = str(row.get('RiskLevel', " ")).strip()
            if risk_val.upper().startswith("SHELL_"):
                risk_val = risk_val[6:]
            risk_val = risk_val.capitalize().strip()
            
            col2_parts = [
                " ", normal_format, change_text,
                normal_format, "\n",
                normal_format, risk_val
            ]
            
            # -------------- Column 3: Trading Assets & BC Apps --------------
            trading_apps = []
            other_apps = []
            bc_content = str(bc_text).strip()
            # Split by newline if present; otherwise by comma.
            if "\n" in bc_content:
                items = bc_content.split("\n")
            else:
                items = bc_content.split(",")
            for item in items:
                item = item.strip()
                # Ignore items like "(12 BC)".
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
            # If no trading apps, then we omit the "Trading BC Apps:" line.
            if trading_scope == "No":
                other_bc_apps_text = ", ".join(other_apps) if other_apps else "No"
                col3_parts = [
                    " ", bold_format, "Trading assets in scope: ",
                    normal_format, trading_scope,
                    normal_format, "\n\n",
                    bold_format, "Other BC Apps: ",
                    normal_format, other_bc_apps_text
                ]
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
            
            # Write the cells.
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
        df.fillna(" ", inplace=True)
        df.columns = df.columns.str.strip()
        
        formatted_excel = generate_formatted_excel(df)
        st.download_button(
            label="📥 Download Formatted Output",
            data=formatted_excel,
            file_name="output_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
