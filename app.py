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
    Convert a date string like '09-04-2025 05:00:00' or '09 April 2025' 
    to a formatted string such as '9th April 2025'.
    If conversion fails, returns the original value as a string.
    """
    if pd.isna(val) or str(val).strip() == "":
        return ""
    # Try to parse using a couple of common formats.
    for fmt in ("%d %B %Y", "%d-%m-%Y %H:%M:%S"):
        try:
            dt = datetime.strptime(str(val).strip(), fmt)
            return f"{ordinal(dt.day)} {dt.strftime('%B')} {dt.year}"
        except Exception:
            continue
    return str(val)

def split_items(text):
    """
    Split the text into items.
    If a newline exists, split on newline; otherwise, split on commas.
    Returns a list of trimmed non-empty items.
    """
    t = str(text).strip()
    if t == "":
        return []
    if "\n" in t:
        items = t.split("\n")
    else:
        items = t.split(",")
    return [x.strip() for x in items if x.strip() != ""]

def count_items(text):
    """Return the count of items in text (split by newline or comma)."""
    return len(split_items(text))

def count_direct_items(text):
    """Return the count of items (from split_items) that contain '(RelationType = Direct)'."""
    items = split_items(text)
    return len([x for x in items if "(RelationType = Direct)" in x])

def build_summary(location, online_outage, ci_val, bc_val, nonbc_val):
    """
    Build the summary for Column 1:
      - Use trimmed location (if blank, becomes "").
      - Use trimmed OnLine/Outage.
      - For CI: count all items (as X CI or X CIs).
      - For BC and NON BC: count only items that contain '(RelationType = Direct)'.
    Join the parts with ", " (which may produce a leading comma if location is blank).
    """
    parts = []
    loc = location.strip() if isinstance(location, str) else ""
    onl = online_outage.strip() if isinstance(online_outage, str) else ""
    parts.append(loc)
    parts.append(onl)
    
    ci_count = count_items(ci_val)
    if ci_count > 0:
        # Use singular "CI" if count == 1.
        parts.append(f"{ci_count} CI" if ci_count == 1 else f"{ci_count} CIs")
    
    bc_direct = count_direct_items(bc_val)
    if bc_direct > 0:
        parts.append(f"{bc_direct} BC (Direct)")
    
    nonbc_direct = count_direct_items(nonbc_val)
    if nonbc_direct > 0:
        parts.append(f"{nonbc_direct} NON BC (Direct)")
    
    # Join parts with ", ". If the first part (location) is empty, the result starts with a comma.
    return ", ".join(parts)

# --- Main Function to Generate the Formatted Excel File ---
def generate_formatted_excel(df):
    output = BytesIO()
    
    # Replace blank cells with a single space.
    df.fillna(" ", inplace=True)
    df.columns = df.columns.str.strip()
    
    # Use ExcelWriter with XlsxWriter.
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
        
        # No header row is written (we want pure data rows).
        output_row = 0
        for idx, row in df.iterrows():
            # ----- Column 1: Record Details -----
            planned_start = format_date(row.get('PlannedStart', " "))
            planned_end = format_date(row.get('PlannedEnd', " "))
            if planned_start == planned_end:
                date_line = planned_start
            else:
                date_line = f"{planned_start} - {planned_end}".strip()
            title_line = str(row.get('Title', " ")).strip()
            
            location_val = str(row.get('Location', " "))
            online_val = str(row.get('OnLine/Outage', " "))
            ci_val = row.get('CI', " ")
            bc_val = row.get('BC', " ")
            nonbc_val = row.get('NONBC', " ")
            
            summary_line = build_summary(location_val, online_val, ci_val, bc_val, nonbc_val)
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
            
            # ----- Column 2: Change & Risk -----
            change_id = str(row.get('ChangeId', " ")).strip()
            f4f_val = str(row.get('F4F', " ")).strip()
            if change_id and f4f_val:
                change_text = f"{change_id}/{f4f_val}"
            elif change_id:
                change_text = change_id
            elif f4f_val:
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
            
            # ----- Column 3: Trading Assets & BC Apps -----
            trading_apps = []
            other_apps = []
            bc_text = str(bc_val).strip()
            if "\n" in bc_text:
                items = bc_text.split("\n")
            else:
                items = bc_text.split(",")
            for item in items:
                item = item.strip()
                if item.startswith("("):
                    continue
                if "(RelationType = Direct)" in item:
                    app_name = item.replace("(RelationType = Direct)", "").strip()
                    if app_name and app_name.upper().startswith("ST"):
                        trading_apps.append(app_name)
                    elif app_name:
                        other_apps.append(app_name)
            trading_scope = "Yes" if trading_apps else "No"
            trading_bc_apps_text = ", ".join(trading_apps) if trading_apps else "None"
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
            
            # Write rich strings to the worksheet.
            worksheet.write_rich_string(output_row, 0, *col1_parts)
            worksheet.write_rich_string(output_row, 1, *col2_parts)
            worksheet.write_rich_string(output_row, 2, *col3_parts)
            
            output_row += 1
        
    output.seek(0)
    return output

# ==============================
#       Streamlit UI
# ==============================

st.title("Change Formatter App")
uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        # Replace blank cells with a space.
        df.fillna(" ", inplace=True)
        df.columns = df.columns.str.strip()
        
        formatted_excel = generate_formatted_excel(df)
        st.download_button(
            label="Download Formatted Output",
            data=formatted_excel,
            file_name="output_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
