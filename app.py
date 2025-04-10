import pandas as pd
import openpyxl
from openpyxl.styles import Font
from io import BytesIO
import streamlit as st

def generate_formatted_excel(df):
    start_date = df['PlannedStart'].iloc[0]
    end_date = df['PlannedEnd'].iloc[0]
    location = df['Location'].iloc[0]
    outage_type = df['OnLine/Outage'].iloc[0]

    bc_direct_apps = df[(df["RelationType"] == "Direct") & (df["BC"].notna())]

    total_cis = df["CI"].nunique()
    bc_count = bc_direct_apps.shape[0]
    non_bc_count = total_cis - bc_count

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formatted Output"
    bold = Font(bold=True)

    ws["A1"] = "Planned Start Date:"
    ws["A1"].font = bold
    ws["B1"] = str(start_date)

    ws["A2"] = "Planned End Date:"
    ws["A2"].font = bold
    ws["B2"] = str(end_date)

    ws.append([])

    summary = f"{location}, {outage_type}, {total_cis} CIs, {bc_count} BC, {non_bc_count} Non-BC"
    ws.append([summary])
    ws.append([])

    ws.append(["Business Affected"])
    ws["C7"] = "BC Apps (RelationType = Direct)"
    ws["C7"].font = bold

    for i, app in enumerate(bc_direct_apps["BC"], start=8):
        ws[f"C{i}"] = app

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- Streamlit UI ---
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
