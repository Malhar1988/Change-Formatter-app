import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# --- Helper Functions ---
def ordinal(n):
    """Return the ordinal string for a number (e.g. 9 -> '9th')."""
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
    if pd.isna(val) or val.strip() == "":
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
    Find the first integer in the given text, e.g. '21 under CI' -> 21.
    If none is found, return 0.
    """
    if not isinstance(text, str):
        return 0
    m = re.search(r"\d+", text)
    return int(m.group(0)) if m else 0

def count_direct_items(text):
    """
    Count how many lines in 'text' contain '(RelationType = Direct)'.
    We split on commas or newlines, ignoring lines that don't have the direct tag.
    """
    if not isinstance(text, str):
        return 0
    lines = []
    if "\n" in text:
        lines = text.split("\n")
    else:
        lines = text.split(",")
    direct_count = 0
    for line in lines:
        line = line.strip()
        # If it has "(RelationType = Direct)" then it counts as direct
        if "(RelationType = Direct)" in line:
            direct_count += 1
    return direct_count

def build_summary(location, online_outage, ci_val, bc_val, nonbc_val):
    """
    Build the summary for Column 1:
      - If we find an integer in ci_val, we label it as e.g. "21 CIs" (if >0).
      - For bc_val and nonbc_val, we look only for how many direct items exist. 
        e.g. if bc_direct_count=1, we add "1 BC (Direct)".
        if bc_direct_count=0, skip it.
      - If location is blank, it remains an empty string, so the final might start 
        with ", Online, 21 CIs, 1 NON BC (Direct)".
    """
    parts = []
    
    # Possibly strip location / online
    loc = location.strip() if isinstance(location, str) else ""
    onl = online_outage.strip() if isinstance(online_outage, str) else ""
    parts.append(loc)     # can be empty
    parts.append(onl)     # can also be empty
    
    # 1) CI
    ci_total = parse_first_int(ci_val)
    if ci_total > 0:
        parts.append(f"{ci_total} CIs")
    
    # 2) BC => consider direct items only
    bc_direct_count = count_direct_items(bc_val)
    if bc_direct_count > 0:
        parts.append(f"{bc_direct_count} BC (Direct)")
    
    # 3) NONBC => consider direct items only
    nonbc_direct_count = count_direct_items(nonbc_val)
    if nonbc_direct_count > 0:
        parts.append(f"{nonbc_direct_count} NON BC (Direct)")
    
    return ", ".join(parts)

def generate_formatted_excel(df):
    output = BytesIO()
    df.fillna(" ", inplace=True)  # Replace blank cells with spaces
    
    import xlsxwriter
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Output Final")
        
        # Two basic formats
        bold_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'bg_color': 'white', 'font_color': 'black', 'font_size': 12
        })
        normal_format = workbook.add_format({
            'text_wrap': True, 'bg_color': 'white', 'font_color': 'black', 'font_size': 12
        })
        
        # Column widths
        worksheet.set_column(0, 0, 50)  # A
        worksheet.set_column(1, 1, 30)  # B
        worksheet.set_column(2, 2, 50)  # C
        
        # Optional test header row
        worksheet.write_rich_string(0, 0,
                                    " ", bold_format, "Test Header Row",
                                    normal_format, " (Row 1)")
        st.write("Test header row at Excel row 1.")
        
        # Process rows
        output_row = 1  # Excel row 2
        for idx, row_data in df.iterrows():
            # --- Column 1 ---
            planned_start = format_date(str(row_data.get('PlannedStart', " ")))
            planned_end = format_date(str(row_data.get('PlannedEnd', " ")))
            date_line = f"{planned_start} - {planned_end}".strip()
            title_line = str(row_data.get('Title', " ")).strip()
            
            location_val = str(row_data.get('Location', " "))
            online_val = str(row_data.get('OnLine/Outage', " "))
            ci_val = str(row_data.get('CI', " "))
            bc_val = str(row_data.get('BC', " "))
            nonbc_val = str(row_data.get('NONBC', " "))
            
            summary_line = build_summary(location_val, online_val, ci_val, bc_val, nonbc_val)
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
            
            # --- Column 2 ---
            # combine ChangeId and F4F with slash
            change_id = str(row_data.get('ChangeId', " ")).strip()
            f4f_val = str(row_data.get('F4F', " ")).strip()
            if change_id != "" and f4f_val != "":
                change_line = f"{change_id}/{f4f_val}"
            elif change_id != "":
                change_line = change_id
            elif f4f_val != "":
                change_line = f4f_val
            else:
                change_line = " "
            
            risk_val = str(row_data.get('RiskLevel', " ")).strip()
            if risk_val.upper().startswith("SHELL_"):
                risk_val = risk_val[6:]
            risk_val = risk_val.capitalize()
            
            col2_parts = [
                " ", normal_format, change_line,
                normal_format, "\n",
                normal_format, risk_val
            ]
            
            # --- Column 3 ---
            # unchanged from your logic
            trading_apps = []
            other_apps = []
            bc_text = str(bc_val).strip()
            if "\n" in bc_text:
                bc_lines = bc_text.split("\n")
            else:
                bc_lines = bc_text.split(",")
            for line in bc_lines:
                line = line.strip()
                if line.startswith("("):
                    continue
                if "(RelationType = Direct)" in line:
                    app_name = line.replace("(RelationType = Direct)", "").strip()
                    if app_name.upper().startswith("ST"):
                        trading_apps.append(app_name)
                    else:
                        other_apps.append(app_name)
            trading_scope = "Yes" if trading_apps else "No"
            trading_bc_apps_text = ", ".join(trading_apps) if trading_apps else "None"
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
            
            # Diagnostics
            st.write(f"Row {idx} -> Excel Row {output_row+1}")
            st.write("  Col1:", date_line, "| Title:", title_line, "| Summary:", summary_line)
            st.write("  Col2:", change_line, "|", risk_val)
            st.write("  Col3:", trading_scope, "| Trading Apps:", trading_bc_apps_text, "| Other Apps:", other_bc_apps_text)
            
            # Write cells
            worksheet.write_rich_string(output_row, 0, *col1_parts)
            worksheet.write_rich_string(output_row, 1, *col2_parts)
            worksheet.write_rich_string(output_row, 2, *col3_parts)
            
            output_row += 1
        
    output.seek(0)
    return output


# ----------------------------- Streamlit App ----------------------------- #
st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx","xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        # Replace blank cells with a space
        df.fillna(" ", inplace=True)
        df.columns = df.columns.str.strip()
        
        st.write("DataFrame shape:", df.shape)
        st.write("Columns:", df.columns.tolist())
        st.subheader("Preview of Input Data")
        st.dataframe(df.head())
        
        formatted_excel = generate_formatted_excel(df)
        
        st.download_button(
            label="Download Formatted Output",
            data=formatted_excel,
            file_name="output_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
