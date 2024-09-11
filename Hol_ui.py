import streamlit as st
import pdfplumber
import pandas as pd
import logging
import re
import io
# from docx2pdf import convert
import tempfile
import os
import openpyxl

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Functions from your original code
def extract_data_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        all_data = []
        first_table_skipped = False
        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            logger.info(f"Page {page_num}: Found {len(tables)} tables")
            for table_num, table in enumerate(tables, 1):
                if not first_table_skipped:
                    first_table_skipped = True
                    logger.info(f"Skipped first table on page {page_num}")
                    continue
                logger.info(f"Processing table {table_num} on page {page_num}")
                all_data.extend(table)
    return all_data

def clean_and_structure_data(raw_data):
    structured_data = []
    current_row = []
    headers = ["ISIN", "ISIN Description", "Account Description", "Quantity", "Total Balance"]
    
    for row in raw_data:
        if any(cell.strip() for cell in row if cell):  # Check if row is not empty
            filtered_row = []
            for i, cell in enumerate(row):
                if cell and cell.strip():
                    if i != 2:  # Skip the third column (index 2) which is 'ISIN Status'
                        filtered_row.append(cell.strip())
            
            # Skip the row if it contains header values or "Pledge"
            if any(header in filtered_row for header in headers) or "Pledge" in filtered_row:
                continue
            
            # Remove any cell that contains "Pledge" and the next cell (number)
            i = 0
            while i < len(filtered_row):
                if "Pledge" in filtered_row[i]:
                    filtered_row.pop(i)
                    if i < len(filtered_row):  # Remove the next cell if it exists
                        filtered_row.pop(i)
                else:
                    i += 1
            
            current_row.extend(filtered_row)
            
            while len(current_row) >= 5:
                structured_data.append(current_row[:5])
                current_row = current_row[5:]
    
    # Add any remaining data
    if current_row:
        logger.warning(f"Incomplete row detected: {current_row}")
    
    logger.info(f"Structured {len(structured_data)} rows of data")
    
    # columns = ['Column1', 'Column2', 'Column3', 'Column4', 'Column5']
    columns = ["ISIN", "ISIN Description", "Account Description", "Quantity", "Total Balance"]
    df = pd.DataFrame(structured_data, columns=columns)
    return df

# Streamlit app

def word_to_pdf(docx_file):
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_docx_path = os.path.join(temp_dir, "temp.docx")
        temp_pdf_path = os.path.join(temp_dir, "temp.pdf")
        
        with open(temp_docx_path, "wb") as f:
            f.write(docx_file.getvalue())
        
        convert(temp_docx_path, temp_pdf_path)
        
        with open(temp_pdf_path, "rb") as f:
            pdf_data = f.read()
    
    return pdf_data

# Streamlit app
def main():
    st.title("Document Converter By Tejas Gavale")

    option = st.radio("Choose conversion type:", 
                      ("Word to PDF", "PDF to Excel"))

    if option == "Word to PDF":
        st.subheader("Word to PDF Converter")
        uploaded_file = st.file_uploader("Choose a Word file", type="docx", key="word_uploader")

        if uploaded_file is not None:
            st.success("File successfully uploaded!")

            if st.button("Convert to PDF", key="word_convert_button"):
                try:
                    pdf_data = word_to_pdf(uploaded_file)
                    
                    st.success("Conversion completed successfully!")
                    
                    # Offer PDF file for download
                    st.download_button(
                        label="Download PDF file",
                        data=pdf_data,
                        file_name="converted_document.pdf",
                        mime="application/pdf",
                        key="pdf_download_button"
                    )
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    logger.error(f"An error occurred: {str(e)}")

    elif option == "PDF to Excel":
        st.subheader("PDF to Excel Converter")
        uploaded_file = st.file_uploader("Choose a PDF file", type="pdf", key="pdf_uploader")

        if uploaded_file is not None:
            st.success("File successfully uploaded!")

            if st.button("Convert to Excel", key="pdf_convert_button"):
                try:
                    # Extract data from PDF
                    raw_data = extract_data_from_pdf(uploaded_file)
                    logger.info(f"Extracted {len(raw_data)} rows of raw data")

                    if not raw_data:
                        st.error("No data extracted from PDF")
                        return

                    # Clean and structure data
                    structured_data = clean_and_structure_data(raw_data)

                    if structured_data.empty:
                        st.error("No structured data generated")
                        return

                    # Display the data on the screen
                    st.subheader("Converted Data:")
                    st.dataframe(structured_data)

                    # Convert DataFrame to Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        structured_data.to_excel(writer, index=False, sheet_name='Sheet1')
                    excel_data = output.getvalue()

                    # Sidebar for download options
                    st.sidebar.subheader("Download Options")
                    
                    # Ask user for custom file name in the sidebar
                    custom_file_name = st.sidebar.text_input("Enter a name for your Excel file:", "converted_data", key="file_name_input")
                    
                    # Ensure the file name ends with .xlsx
                    if not custom_file_name.endswith('.xlsx'):
                        custom_file_name += '.xlsx'

                    # Offer Excel file for download in the sidebar
                    st.sidebar.download_button(
                        label="Download Excel file",
                        data=excel_data,
                        file_name=custom_file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="excel_download_button"
                    )

                    st.success("Conversion completed successfully!")
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                    logger.error(f"An error occurred: {str(e)}")

if __name__ == '__main__':
    main()