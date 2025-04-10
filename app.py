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
    if isinstance(val, datetime):
        dt = val
    else:
        for fmt in ("%d %B %Y", "%d-%m-%Y %H:%M:%S"):
            try:
                dt = datetime.strptime(str(val).strip(), fmt)
                break
            except Exception:
                continue
        else:
            return str(val)
    return f"{ordinal(dt.day)} {dt.strftime('%B')} {dt.year}"

def split_items(text):
    """
    Split text into items.
    If a newline exists, split on newline; otherwise, split on commas.
    Returns a list of trimmed, non-empty items.
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
    """Return the count of items in text that contain '(RelationType = Direct)'."""
    return len([x for x in split_items(text) if "(RelationType = Direct)" in x])

def preserve(val):
    """
    Preserve the value if it is exactly a single space; otherwise, trim it.
    This is used to ensure that if an input cell is blank (filled as " "),
    it remains " " instead of becoming empty.
    """
    s = str(val)
    return s if s == " " else s.strip()

def build_summary(location, online_outage, ci_val, bc_val, nonbc_val):
    """
    Build the summary string for Column 1 (Line 3):
      - Use the preserved Location (so if blank, it remains " " and yields a leading comma).
      - Use the trimmed OnLine/Outage.
      - Count CI as all comma-separated items.
      - Count BC and NONBC only for items containing "(RelationType = Direct)".
    Return the parts joined by ", ".
    """
    parts = []
    parts.append(preserve(location))
    parts.append(online_outage.strip())
    
    ci_count = count_items(ci_val)
    if ci_count > 0:
        parts.append("1 CI" if ci_count == 1 else f"{ci_count} CIs")
    
    bc_direct = count_direct_items(bc_val)
    if bc_direct > 0:
        parts.append(f"{bc_direct} BC (Direct)")
    
    nonbc_direct = count_direct_items(nonbc_val)
    if nonbc_direct > 0:
        parts.append(f"{nonbc_direct} NON BC (Direct)")
    
    return ", ".join(parts)

# --- Main Function to Generate the Formatted Excel File ---

def generate_formatted_excel(df):
    output = BytesIO()
    
    # Replace blank (NaN) cells with a single space.
    df.fillna(" ", inplace=True)
    df.columns = df.columns.str.strip()
    
    # Use ExcelWriter with XlsxWriter.
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Output Final")
        
        # Create formats.
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
        
        # No header row is written.
        output_row = 0
        
        for idx, row in df.iterrows():
            # ---------- Column 1: Record Details ----------
            planned_start = format_date(row.get('PlannedStart', " "))
            planned_end = format_date(row.get('PlannedEnd', " "))
            if planned_start == planned_end:
                date_line = planned_start
            else:
                date_line = f"{planned_start} - {planned_end}".strip()
            title_line = preserve(row.get('Title', " "))
            
            location_val = preserve(row.get('Location', " "))
            online_val = str(row.get('OnLine/Outage', " ")).strip()
            ci_val = row.get('CI', " ")
            bc_val = row.get('BC', " ")
            nonbc_val = row.get('NONBC', " ")
            
            summary_line = build_summary(location_val, online_val, ci_val, bc_val, nonbc_val)
            business_groups_line = preserve(row.get('BusinessGroups', " "))
            
            col1_parts = [
                " ", bold_format, date_line,
                normal_format, "\n\n",
                normal_format, title_line,
                normal_format, "\n\n",
                normal_format, summary_line,
                normal_format, "\n\n",
                normal_format, business_groups_line
            ]
            
            # ---------- Column 2: Change & Risk ----------
            change_id = preserve(row.get('ChangeId', " "))
            f4f_field = preserve(row.get('F4F', " "))
            if f4f_field == " ":
                change_text = change_id
            else:
                # If F4F is not blank, combine both.
                if change_id != " ":
                    change_text = f"{change_id}/{f4f_field}"
                else:
                    change_text = f4f_field
            risk_val = str(row.get('RiskLevel', " ")).strip()
            if risk_val.upper().startswith("SHELL_"):
                risk_val = risk_val[6:]
            risk_val = risk_val.capitalize().strip()
            
            col2_parts = [
                " ", normal_format, change_text,
                normal_format, "\n",
                normal_format, risk_val
            ]
            
            # ---------- Column 3: Trading Assets & BC Apps ----------
            trading_apps = []
            other_apps = []
            bc_content = str(bc_val).strip()
            if "\n" in bc_content:
                items = bc_content.split("\n")
            else:
                items = bc_content.split(",")
            for item in items:
                item = item.strip()
                if item.startswith("("):
                    continue  # ignore items like "(12 BC)"
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
            
            # Write the rich strings to the worksheet.
            worksheet.write_rich_string(output_row, 0, *col1_parts)
            worksheet.write_rich_string(output_row, 1, *col2_parts)
            worksheet.write_rich_string(output_row, 2, *col3_parts)
            
            output_row += 1
        
    output.seek(0)
    return output

# ==============================
#       Streamlit App UI
# ==============================

st.title("Change Formatter App")
uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
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
