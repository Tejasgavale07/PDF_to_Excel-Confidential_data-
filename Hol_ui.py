import streamlit as st
import pandas as pd
from io import BytesIO, StringIO

def add_empty_line(input_content, target_line):
    output = StringIO()
    for line in input_content.split('\n'):
        output.write(line + '\n')
        if line.strip() == target_line.strip():
            output.write('\n')
    return output.getvalue()

def add_line_breaker_to_content(content):
    sections = content.split('^PART-I - Details of Tax Deducted at Source^')
    
    if len(sections) < 2:
        raise ValueError("Expected header not found in the file")

    header_section = sections[0]
    data_section = sections[1]

    lines = data_section.strip().split('\n')
    modified_lines = []
    header_found = False

    for line in lines:
        if not header_found and "Sr. No." in line:
            modified_lines.append(line)
            modified_lines.append(' ' * 1)
            header_found = True
        else:
            modified_lines.append(line)

    modified_content = header_section + '^PART-I - Details of Tax Deducted at Source^' + '\n'.join(modified_lines)
    return modified_content

def read_data_from_content(content):
    sections = content.split('^PART-I - Details of Tax Deducted at Source^')[1].split('\n\n')

    all_data = []
    header = None

    for section in sections:
        lines = section.strip().split('\n')
        if not lines:
            continue

        deductor_info = lines[0].split('^')
        if len(deductor_info) < 3:
            continue

        deductor_number = deductor_info[0]
        deductor_name = deductor_info[1]
        deductor_tan = deductor_info[2]

        for line in lines[1:]:
            if line.strip() == '':
                continue
            
            cols = [col.strip() for col in line.split('^') if col.strip()]
            if not header and "Sr. No." in cols:
                header = cols
            elif header and cols and cols[0].isdigit() and len(cols) == len(header):
                all_data.append([deductor_number, deductor_name, deductor_tan] + cols)

    if not header:
        raise ValueError("Header not found in the file")

    full_header = ['Deductor Number', 'Name of Deductor', 'TAN of Deductor'] + header
    return full_header, all_data

def create_dataframe(header, data):
    df = pd.DataFrame(data, columns=header)

    numeric_columns = ['Amount Paid / Credited(Rs.)', 'Tax Deducted(Rs.)', 'TDS Deposited(Rs.)']

    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    return df

st.sidebar.title("File Input")
uploaded_file = st.sidebar.file_uploader("Upload a Text File", type=["txt"])

if st.sidebar.button("Submit") and uploaded_file is not None:
    try:
        content = uploaded_file.getvalue().decode("utf-8")

        target_line = "Sr. No.^Name of Deductor^TAN of Deductor^^^^^Total Amount Paid / Credited(Rs.)^Total Tax Deducted(Rs.)^Total TDS Deposited(Rs.)"
        content_with_empty_line = add_empty_line(content, target_line)
        modified_content = add_line_breaker_to_content(content_with_empty_line)
        header, data = read_data_from_content(modified_content)
        df = create_dataframe(header, data)

        df = df.drop(columns=['Deductor Number', 'Sr. No.'], errors='ignore')
        df.insert(0, 'Sr. No.', range(1, len(df) + 1))

        st.write("### Updated Extracted Data", df)

        # Group by 'Name of Deductor' and 'TAN of Deductor' to create the aggregated DataFrame
        aggregated_df = df.groupby(['Name of Deductor', 'TAN of Deductor']).agg({
            'Amount Paid / Credited(Rs.)': 'sum',
            'Tax Deducted(Rs.)': 'sum',
            'TDS Deposited(Rs.)': 'sum'
        }).reset_index()

        st.write("### Aggregated Totals by Deductor", aggregated_df)

        @st.cache_data
        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()
        def convert_df_to_excel1(aggregated_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        excel_data = convert_df_to_excel(df)
        excel_data_agg = convert_df_to_excel1(df)
        st.sidebar.download_button(
            label="Download Individual Excel",
            data=excel_data,
            file_name="Individual_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.sidebar.download_button(
            label="Download Aggregate By Deductor Excel",
            data=excel_data_agg,
            file_name="Aggregate_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {str(e)}")

elif uploaded_file is None:
    st.sidebar.write("Awaiting file upload...")
