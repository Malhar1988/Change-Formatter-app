import streamlit as st
import pandas as pd
from io import BytesIO
import re
from datetime import datetime

# ====================
#   HELPER FUNCTIONS
# ====================

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
    If parsing fails, return the original value as a string.
    """
    if pd.isna(val) or str(val).strip() == "":
        return ""
    if isinstance(val, datetime):
        dt = val
    else:
        try:
            dt = datetime.strptime(str(val).strip(), "%d %B %Y")
        except Exception:
            try:
                dt = datetime.strptime(str(val).strip(), "%d-%m-%Y %H:%M:%S")
            except Exception:
                return str(val)
    return f"{ordinal(dt.day)} {dt.strftime('%B')} {dt.year}"

def singular_or_plural_ci(count):
    """Return '1 CI' or 'X CIs' depending on the count."""
    if count == 1:
        return "1 CI"
    else:
        return f"{count} CIs"

def count_items(text):
    """Count the number of non-empty, comma-separated items in text."""
    t = str(text).strip()
    if t == "":
        return 0
    # Split by commas, ignoring blank items
    items = [x.strip() for x in t.split(",") if x.strip() != ""]
    return len(items)

def count_direct_items(text):
    """Count how many comma-separated items contain '(RelationType = Direct)'."""
    t = str(text).strip()
    if t == "":
        return 0
    count = 0
    for x in t.split(","):
        x = x.strip()
        if x and "(RelationType = Direct)" in x:
            count += 1
    return count

def build_summary(location, online_outage, ci_val, bc_val, nonbc_val):
    """
    Build Column 1's summary line (Line 3):
      - location (trimmed); if blank, becomes '' => leading comma
      - online_outage (trimmed)
      - X CI or X CIs => count items from ci_val
      - X BC (Direct) => only if direct_count>0
      - X NON BC (Direct) => only if direct_count>0
    Then join them with ", ".
    """
    parts = []
    
    # Trim location. If it's empty, we keep it as '' => leading comma.
    loc = location.strip() if isinstance(location, str) else ""
    onl = online_outage.strip() if isinstance(online_outage, str) else ""
    
    parts.append(loc)   # might be empty
    parts.append(onl)   # might be empty
    
    # For CI: count all comma-separated items.
    ci_count = count_items(ci_val)
    if ci_count > 0:
        parts.append(singular_or_plural_ci(ci_count))
    
    # For BC and NON BC: only count direct items
    bc_count = count_direct_items(bc_val)
    if bc_count > 0:
        # e.g. '4 BC (Direct)'
        parts.append(f"{bc_count} BC (Direct)")
    nonbc_count = count_direct_items(nonbc_val)
    if nonbc_count > 0:
        parts.append(f"{nonbc_count} NON BC (Direct)")
    
    # Join with comma and space. If loc is '', we get a leading comma => e.g. ', Outage, 3 CI, ...'
    return ", ".join([p for p in parts if p != ""])

# ====================
#   MAIN EXCEL CODE
# ====================

def generate_formatted_excel(df):
    output = BytesIO()
    
    # Fill blanks in each cell with a single space
    df.fillna(" ", inplace=True)
    df.columns = df.columns.str.strip()
    
    import xlsxwriter
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Output Final")
        
        # Create formats
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
        
        # Set col widths
        worksheet.set_column(0, 0, 50)  # A
        worksheet.set_column(1, 1, 30)  # B
        worksheet.set_column(2, 2, 50)  # C
        
        output_row = 0  # no header row
        for idx, row_data in df.iterrows():
            # ---- Column 1 (Record Details) ----
            
            # LINE 1: If planned_start == planned_end => single date else => date range
            planned_start = format_date(row_data.get('PlannedStart', " "))
            planned_end = format_date(row_data.get('PlannedEnd', " "))
            if planned_start == planned_end:
                date_line = planned_start
            else:
                date_line = (planned_start + " - " + planned_end).strip()
            
            # LINE 2: Title
            title_line = str(row_data.get('Title', " ")).strip()
            
            # LINE 3: summary => location, online, X CI or X CIs, X BC (Direct), X NON BC (Direct)
            location_val = str(row_data.get('Location', " "))
            online_val = str(row_data.get('OnLine/Outage', " "))
            ci_val = row_data.get('CI', " ")
            bc_val = row_data.get('BC', " ")
            nonbc_val = row_data.get('NONBC', " ")
            summary_line = build_summary(location_val, online_val, ci_val, bc_val, nonbc_val)
            
            # LINE 4: businessGroups
            business_groups_line = str(row_data.get('BusinessGroups', " ")).strip()
            
            col1_parts = [
                " ", bold_format, date_line,
                normal_format, "\n\n",
                normal_format, title_line,
                normal_format, "\n\n",
                normal_format, summary_line,
                normal_format, "\n\n",
                normal_format, business_groups_line
            ]
            
            # ---- Column 2 (Change & Risk) ----
            change_id = str(row_data.get('ChangeId', " ")).strip()
            f4f_val = str(row_data.get('F4F', " ")).strip()
            if change_id and f4f_val:
                change_text = f"{change_id}/{f4f_val}"
            elif change_id:
                change_text = change_id
            elif f4f_val:
                change_text = f4f_val
            else:
                change_text = " "
            
            risk_val = str(row_data.get('RiskLevel', " ")).strip()
            if risk_val.upper().startswith("SHELL_"):
                risk_val = risk_val[6:]
            risk_val = risk_val.capitalize()
            
            col2_parts = [
                " ", normal_format, change_text,
                normal_format, "\n",
                normal_format, risk_val
            ]
            
            # ---- Column 3 (Trading Assets & BC Apps) ----
            bc_text = str(row_data.get('BC', " ")).strip()
            trading_apps = []
            other_apps = []
            if "\n" in bc_text:
                items = bc_text.split("\n")
            else:
                items = bc_text.split(",")
            for it in items:
                it = it.strip()
                if it.startswith("("):
                    continue
                if "(RelationType = Direct)" in it:
                    app_name = it.replace("(RelationType = Direct)", "").strip()
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
            
            # Write columns
            worksheet.write_rich_string(output_row, 0, *col1_parts)
            worksheet.write_rich_string(output_row, 1, *col2_parts)
            worksheet.write_rich_string(output_row, 2, *col3_parts)
            
            output_row += 1
        
    output.seek(0)
    return output


# ==============================
#   Streamlit App UI
# ==============================
st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        # Replace blank cells with a single space
        df.fillna(" ", inplace=True)
        # Remove extra spaces from header names
        df.columns = df.columns.str.strip()
        
        # Build the output file
        formatted_excel = generate_formatted_excel(df)
        
        # Provide a download button
        st.download_button(
            label="Download Formatted Output",
            data=formatted_excel,
            file_name="output_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
