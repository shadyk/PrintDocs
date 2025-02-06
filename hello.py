import os
import pandas as pd
from docx import Document

def fill_template(template_path, data_row, output_path):
    """Fill the Word document template with data from the selected row."""
    if not os.path.exists(template_path):
        print(f"Template file not found: {template_path}")
        return

    try:
        doc = Document(template_path)
    except Exception as e:
        print(f"Error loading document: {e}")
        return

    for key, value in data_row.items():
        key = key.strip()
        for paragraph in doc.paragraphs:
            if f'{{{key}}}' in paragraph.text:
                paragraph.text = paragraph.text.replace(f'{{{key}}}', str(value))
    
    doc.save(output_path)

def generate_documents(excel_path, template_path, output_dir):
    """Generate Word documents from Excel data."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Read the Excel file
    df = pd.read_excel(excel_path, header=1)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    for index, row in df.iterrows():
        output_path = os.path.join(output_dir, f"Document_{index + 1}.docx")
        fill_template(template_path, row, output_path)
        print(f"Generated: {output_path}")

# Example usage
excel_file = 'Desktop/print/sample.xlsx'  # Path to your Excel file
word_template = 'Desktop/print/template.docx'  # Path to your Word template
output_directory = 'Desktop/print/output_docs'  # Directory to save generated documents

print("Absolute path to Excel file:", os.path.abspath(excel_file))
print("File exists?", os.path.exists(excel_file))
generate_documents(excel_file, word_template, output_directory)