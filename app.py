import pandas as pd
import openpyxl
from openpyxl.styles import Font
from io import BytesIO

def generate_formatted_excel(df):
    # Extract planned dates and metadata from first row
    start_date = df['PlannedStart'].iloc[0]
    end_date = df['PlannedEnd'].iloc[0]
    location = df['Location'].iloc[0]
    outage_type = df['OnLine/Outage'].iloc[0]

    # Filter BC apps where RelationType is Direct
    bc_direct_apps = df[(df["RelationType"] == "Direct") & (df["BC"].notna())]

    # Count metrics
    total_cis = df["CI"].nunique()
    bc_count = bc_direct_apps.shape[0]
    non_bc_count = total_cis - bc_count

    # Create workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formatted Output"
    bold = Font(bold=True)

    # Row 1 & 2: Planned Dates
    ws["A1"] = "Planned Start Date:"
    ws["A1"].font = bold
    ws["B1"] = str(start_date)

    ws["A2"] = "Planned End Date:"
    ws["A2"].font = bold
    ws["B2"] = str(end_date)

    # Row 3: Spacer
    ws.append([])

    # Row 4: Summary Line
    summary = f"{location}, {outage_type}, {total_cis} CIs, {bc_count} BC, {non_bc_count} Non-BC"
    ws.append([summary])

    # Row 5: Spacer
    ws.append([])

    # Row 6: Business Affected
    ws.append(["Business Affected"])

    # Row 7: Column C heading
    ws["C7"] = "BC Apps (RelationType = Direct)"
    ws["C7"].font = bold

    # Row 8 onward: App names from BC column
    for i, app in enumerate(bc_direct_apps["BC"], start=8):
        ws[f"C{i}"] = app

    # Auto-fit column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    # Save to memory
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
    # Required imports
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
from io import BytesIO

# Your Excel formatting function
def generate_formatted_excel(df):
    # ... your function code ...
    return output  # ⬅️ THIS is the last line of the function

# ----------------------------
# Streamlit UI begins here
# ----------------------------

st.title("Change Formatter App")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Preview of Uploaded Data")
    st.dataframe(df.head())

    excel_data = generate_formatted_excel(df)

    st.download_button(
        label="📥 Download Formatted Excel",
        data=excel_data,
        file_name="formatted_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

