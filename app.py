import streamlit as st
import pandas as pd
from io import BytesIO
import xlsxwriter  # Import xlsxwriter directly

# Minimal helper: (you can add date formatting later)
def create_test_excel(df):
    output = BytesIO()
    # Create an in‑memory workbook.
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Output")
    
    # Create formats with white background and black font.
    bold_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'bg_color': 'white',
        'font_color': 'black'
    })
    normal_format = workbook.add_format({
        'text_wrap': True,
        'bg_color': 'white',
        'font_color': 'black'
    })
    
    # Set column widths (for clarity)
    worksheet.set_column(0, 0, 50)
    
    # Write a hard-coded test row at row 0.
    # Using write_rich_string: the first argument is an empty string so that formatting is applied.
    worksheet.write_rich_string(0, 0, "", bold_format, "Test Bold", normal_format, " Test Normal")
    
    # Check if the DataFrame is non-empty.
    if not df.empty:
        # For demonstration, take the first row and build dummy content.
        # We'll simply combine the "Title" and "PlannedStart" values.
        title = str(df.iloc[0].get("Title", "No Title"))
        planned_start = str(df.iloc[0].get("PlannedStart", "No Start"))
        content = f"{planned_start} - {title}"
        # Write this dummy content into row 1, column 0.
        worksheet.write(1, 0, content, normal_format)
    
    # Close the workbook so it finalizes the file.
    workbook.close()
    # Move to the beginning of the BytesIO object.
    output.seek(0)
    return output

st.title("Test XlsxWriter Output")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read the uploaded file into a DataFrame.
        df = pd.read_excel(uploaded_file)
        st.write("DataFrame preview:")
        st.dataframe(df.head())
        
        # Generate the test Excel file.
        output = create_test_excel(df)
        
        # Provide a download button for the generated Excel file.
        st.download_button(
            label="Download Test Excel File",
            data=output,
            file_name="test_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
