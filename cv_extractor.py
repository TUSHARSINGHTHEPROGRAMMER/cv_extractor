import re
import xlsxwriter
import PyPDF2
from docx import Document
import zipfile
from io import BytesIO
import streamlit as st
import os

phone_number_pattern = r'\b(?:\+?(\d{1,3}))?[-. (]*(\d{3})[-. )]*(\d{3})[-. ]*(\d{4})\b'

# Function to extract useful name from filename without extension
def extract_useful_name(filename):
    # Remove the extension from the filename
    filename_without_extension = os.path.splitext(filename)[0]
    # Extract the part of the filename after the last '/'
    parts = filename_without_extension.split('/')
    # Check if there is a '/' in the filename
    if len(parts) > 1:
        # If there is, extract the last part and replace '/' with space
        name_with_surname = parts[-1].replace('/', ' ')
    else:
        # If there is no '/', use the whole filename
        name_with_surname = filename_without_extension
    return name_with_surname

# Function to write the results to an Excel file
def write_to_excel(data, output):
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet()

    # Write headers
    worksheet.write(0, 0, 'Name')
    worksheet.write(0, 1, 'Email')
    worksheet.write(0, 2, 'Mobile Number')
    worksheet.write(0, 3, 'Overall Text')

    # Set to keep track of already written names and emails
    written_names = set()
    written_emails = set()

    # Write data
    for i, (name, emails, contact_numbers, text) in enumerate(data, start=1):
        worksheet.write(i, 0, name if name not in written_names else '')
        worksheet.write(i, 1, ', '.join(email for email in emails if email not in written_emails))

        # Process contact numbers before writing
        formatted_numbers = []
        for num in contact_numbers:
            # Remove parentheses, single quotes, spaces, and commas
            formatted_number = ''.join(part.strip("' (),") for part in num if part.strip("' (),"))
            formatted_numbers.append(formatted_number)

        # Write the first formatted contact number (if any)
        worksheet.write(i, 2, formatted_numbers[0] if formatted_numbers else '')

        worksheet.write(i, 3, text)

        # Update written names and emails sets
        written_names.add(name)
        written_emails.update(emails)

    workbook.close()

# Function to extract emails from text
def extract_emails(text):
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    emails = re.findall(email_pattern, text)
    return emails

# Function to extract contact numbers from text
def extract_contact_numbers(text):
    return re.findall(phone_number_pattern, text)

# Function to extract text from a file
def extract_text(file, filename=None):
    if filename and filename.endswith('.pdf'):
        if isinstance(file, BytesIO):
            file = file.read()
        pdf_reader = PyPDF2.PdfReader(BytesIO(file))
        text = ''
        for page in range(len(pdf_reader.pages)):
            text += pdf_reader.pages[page].extract_text()
        return text
    elif filename and filename.endswith('.docx'):
        docx = Document(file)
        text = ""
        for paragraph in docx.paragraphs:
            text += paragraph.text
        return text
    else:
        return ''

# Function to extract information from a CV file
def extract_info_from_cv(file, filename=None):
    if isinstance(file, zipfile.ZipExtFile):
        file_content = file.read()
    else:
        file_content = file
    text = extract_text(BytesIO(file_content), filename)
    emails = extract_emails(text)
    contact_numbers = extract_contact_numbers(text)
    contact_numbers = [str(num) for num in contact_numbers]
    return extract_useful_name(filename), emails, contact_numbers, text

# Main function
def main():
    # Upload the CVs as a zip file
    uploaded_file = st.file_uploader("Upload CVs as a zip file", type="zip")

    if uploaded_file:
        # Extract the CVs from the zip file
        with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
            cv_files = [file for file in zip_ref.namelist() if file.endswith('.pdf') or file.endswith('.docx')]

            # Extract information from each CV
            data = []
            for file in cv_files:
                try:
                    with zip_ref.open(file) as cv_file:
                        name, emails, contact_numbers, text = extract_info_from_cv(cv_file, filename=file)
                        data.append((name, emails, contact_numbers, text))
                except Exception as e:
                    st.error(f"Error processing file '{file}': {e}")

        # Write the results to an Excel file
        output = 'output.xlsx'
        write_to_excel(data, output)

        # Display the results
        st.success("Extraction completed.Download your excel file")
      

        # Download the Excel file
        st.download_button(
            label="Download Excel file",
            data=open(output, 'rb').read(),
            file_name=output,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

# Run the main function
if __name__ == '__main__':
    main()
