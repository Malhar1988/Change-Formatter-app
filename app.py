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
    Generate an Excel file using XlsxWriter. The output file will have three columns:
      - Column 1 (Record Details): Contains the formatted date line (PlannedStart - PlannedEnd),
        title, a summary line (Location, OnLine/Outage, CI/BC/NONBC counts), and BusinessGroups.
      - Column 2 (Change & Risk): Contains the F4F value (or ChangeId/F4F) and a processed RiskLevel.
      - Column 3 (Trading Assets & BC Apps): Contains bold headings for 
        'Trading assets in scope:', 'Trading BC Apps:' and 'Other BC Apps:' along with their values.
    Each element is separated by extra newlines for spacing.
    """
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet("Output Final")
    
    # Create formats
    bold_format = workbook.add_format({'bold': True, 'text_wrap': True})
    normal_format = workbook.add_format({'text_wrap': True})
    
    # Set column widths for a better appearance.
    worksheet.set_column(0, 0, 50)  # Column A: Record Details
    worksheet.set_column(1, 1, 30)  # Column B: Change & Risk
    worksheet.set_column(2, 2, 50)  # Column C: Trading Assets & BC Apps
    
    # Iterate over the DataFrame records.
    for idx, row in df.iterrows():
        row_num = idx  #_
