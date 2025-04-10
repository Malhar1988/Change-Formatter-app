import streamlit as st
import pandas as pd
from io import BytesIO

# --- Debugging Helper Function ---
def create_debug_excel(df):
    output = BytesIO()
    # Create a workbook using XlsxWriter directly.
    import xlsxwriter
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Output")
    
    # Create a basic format.
    test_format = workbook.add_format({'bold': True, 'text_wrap': True, 'bg_color': 'white', 'font_color': 'black'})
    normal_format = workbook.add_format({'text_wrap': True, 'bg_color': 'white', 'font_color': 'black'})
    
    # Set column width for clarity.
    worksheet.set_column(0, 0, 50)
    
    # Write a hard-coded test row in row 0.
    worksheet.write_rich_string(0, 0, "", test_format, "Test Bold", normal_format, " Test Normal")
    st.write("Hard-coded test row written at row 0.")
    
    # If the input file is not empty, generate dummy content from the first row.
    if not df.empty:
        # Create dummy content from the first row.
        title = str(df.iloc[0].get("Title", "No Title"))
        planned_start = str(df.iloc[0].get("PlannedStart", "No Start"))
        content = f"{planned_start} - {title}"
        st.write("Dummy content for row 1 is:", content)
        worksheet.write(1, 0, content, normal_format)
    else:
        st.write("DataFrame is empty!")
    
    # Write a manual test value in row 2 to check further.
    worksheet.write(2, 0, "Manual test value in row 3", normal_format)
    
    workbook.close()
    output.seek(0)
    return output

# --- Streamlit App UI ---
st.title("Debug XlsxWriter Output")
uploaded_file = st.file_uploader("Upload your Changes Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read the file into a DataFrame.
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()  # Remove extra spaces from headers.
        st.write("DataFrame shape:", df.shape)
        st.write("DataFrame columns:", df.columns.tolist())
        st.subheader("Preview of Input Data")
        st.dataframe(df.head())
        
        # Create the debug Excel file.
        debug_excel = create_debug_excel(df)
        st.download_button(
            label="Download Debug Excel File",
            data=debug_excel,
            file_name="test_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
