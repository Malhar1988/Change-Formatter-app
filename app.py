import streamlit as st
import pandas as pd
from openpyxl import Workbook
from io import BytesIO

def generate_output_excel(df):
    """
    Combine all rows of the first column of the DataFrame into one string,
    preserving newlines, then write this text into cell A1 of a new Excel workbook.
    """
    # We assume the data is in the first column (adjust if needed)
    # Convert all non-null rows to strings and strip extra spaces
    lines = [str(cell).strip() if pd.notna(cell) else "" for cell in df.iloc[:, 0]]
    # Join the lines with newline characters. Blank lines in input are preserved.
    combined_text = "\n".join(lines)
    
    # Create a new workbook and set the active worksheet title
    wb = Workbook()
    ws = wb.active
    ws.title = "Output"
    
    # Write the combined text into cell A1
    ws["A1"] = combined_text
    
    # Optionally, you can adjust cell A1's alignment or wrap text if desired:
    ws["A1"].alignment = ws["A1"].alignment.copy(wrapText=True)
    
    # Save the workbook to a BytesIO stream for downloading
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

st.title("Single Cell Converter")

uploaded_file = st.file_uploader("Upload your input Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read the uploaded Excel file. We assume there's no header.
        df = pd.read_excel(uploaded_file, header=None)
        st.subheader("Input Data Preview:")
        st.dataframe(df)
        
        # Generate the output Excel file with the combined text in one cell.
        output_excel = generate_output_excel(df)
        
        # Create a download button for the formatted file.
        st.download_button(
            label="📥 Download Formatted Output",
            data=output_excel,
            file_name="output_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
