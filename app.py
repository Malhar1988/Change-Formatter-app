import streamlit as st
import pandas as pd
from io import BytesIO

def create_debug_excel(df):
    output = BytesIO()
    import xlsxwriter
    # Create a workbook in memory.
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Output")
    
    # Create formats with a white background and black font.
    bold_format = workbook.add_format({
        'bold': True, 'text_wrap': True, 'bg_color': 'white', 'font_color': 'black'
    })
    normal_format = workbook.add_format({
        'text_wrap': True, 'bg_color': 'white', 'font_color': 'black'
    })
    
    # Set column width for clarity.
    worksheet.set_column(0, 0, 50)
    
    # Write a hard-coded test row in row 0.
    row_counter = 0
    worksheet.write_rich_string(row_counter, 0, "", bold_format, "Test Bold", normal_format, " Test Normal")
    st.write("Hard-coded test row written at Excel row 1.")
    row_counter += 1  # Move to the next row.
    
    # If the input file has data, generate dummy content from the first row.
    if not df.empty:
        first_row = df.iloc[0]
        title = str(first_row.get("Title", "No Title"))
        planned_start = str(first_row.get("PlannedStart", "No Start"))
        dummy_content = f"{planned_start} - {title}"
        st.write("Dummy content for Excel row 2 is:", dummy_content)
        worksheet.write(row_counter, 0, dummy_content, normal_format)
        row_counter += 1  # Move to the next row.
    else:
        st.write("DataFrame is empty!")
    
    # Write a manual test value in the next row.
    worksheet.write(row_counter, 0, "Manual test value in row 3", normal_format)
    st.write("Manual test value written at Excel row", row_counter + 1)
    
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
