import os
import pytesseract
from PIL import Image
import pandas as pd

# Set the path to the Tesseract executable if not in PATH
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

def extract_names_from_pdf(pdf_path):
    # Convert PDF to images (requires additional library like pdf2image)
    from pdf2image import convert_from_path
    images = convert_from_path(pdf_path)

    extracted_text = ""
    for image in images:
        extracted_text += pytesseract.image_to_string(image)

    # Extract names (assuming names are separated by newlines)
    names = [line.strip() for line in extracted_text.split('\n') if line.strip()]
    return names

def validate_names(pdf_names, excel_path, output_path):
    # Read Excel file
    df = pd.read_excel(excel_path)

    # Ensure the Excel file has a 'Name' column
    if 'Name' not in df.columns:
        raise ValueError("The Excel file must contain a 'Name' column.")

    # Add a new column for validation
    df['Exists in PDF'] = df['Name'].apply(lambda name: name in pdf_names)

    # Save the updated DataFrame to a new Excel file
    df.to_excel(output_path, index=False)

def main():
    pdf_path = 'SUA/input.pdf'  # Replace with the actual PDF path
    excel_path = 'EXCEL/input.xlsx'  # Replace with the actual Excel path
    output_path = 'OUTPUT/output.xlsx'  # Replace with the desired output path

    # Extract names from the PDF
    pdf_names = extract_names_from_pdf(pdf_path)

    # Validate names and generate output Excel
    validate_names(pdf_names, excel_path, output_path)

if __name__ == "__main__":
    main()