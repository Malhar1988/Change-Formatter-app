import streamlit as st
import pandas as pd
from io import BytesIO
import xlsxwriter  # We use XlsxWriter directly for this minimal example

def create_debug_excel(df):
    output = BytesIO()
    # Create an in-memory workbook.
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Output")
    
    # Create basic formats with white background and black text.
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
    
    # Set column width for clarity.
    worksheet.set_column(0, 0, 50)
    
    # Write a hard-coded test row in row 0.
    # Here, we start with a non-empty string "Test: " so the rich text is displayed.
    worksheet.write_rich_string(0, 0,
                                "Test: ", bold_format, "Test Bold", 
                                normal_format, " Test Normal")
    st.write("Hard-coded test row written at Excel row 1.")
    
    # Write dummy content from the first row of the input, if available (to row 1).
    if not df.empty:
        first_row = df.iloc[0]
        title = str(first_row.get("Title", "No Title"))
        planned_start = str(first_row.get("PlannedStart", "No Start"))
        dummy_content = f"{planned_start} - {title}"
        st.write("Dummy content for Excel row 2 is:", dummy_content)
        worksheet.write(1, 0, dummy_content, normal_format)
    else:
        st.write("DataFrame is empty!")
    
    # Write a manual test value in row 2 (Excel row 3).
    worksheet.write(2, 0, "Manual test value in row 3", normal_format)
    st.write("Manual test value written at Excel row 3.")
    
    workbook.close()
    output.seek(0)
    return output

st.title("Debug XlsxWriter Output")
uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()  # Clean header names.
        st.write("DataFrame shape:", df.shape)
        st.write("DataFrame columns:", df.columns.tolist())
        st.subheader("Preview of Input Data")
        st.dataframe(df.head())
        
        debug_excel = create_debug_excel(df)
        st.download_button(
            label="Download Debug Excel File",
            data=debug_excel,
            file_name="test_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
