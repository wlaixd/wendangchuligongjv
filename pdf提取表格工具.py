import os
import re
import pandas as pd
from docx import Document
from pdf2docx import Converter

def convert_pdf_to_docx(pdf_path, docx_path):
    """Convert PDF to DOCX format"""
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

def extract_tables_from_docx(docx_path, excel_path):
    """Extract tables from DOCX and save to Excel"""
    try:
        # Read all tables from the Word document
        doc = Document(docx_path)
        all_table_data = []
        
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                table_data.append([cell.text.strip() for cell in row.cells])
            all_table_data.append(table_data)

        # Create Excel file with tables
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            for i, table in enumerate(all_table_data):
                if not table:  # Skip empty tables
                    continue
                    
                # Get sheet name from first cell or use default
                if table and table[0] and table[0][0].strip():
                    sheet_name = table[0][0].strip()
                else:
                    sheet_name = f'Table_{i + 1}'

                # Clean sheet name
                sheet_name = re.sub(r'[\\/*?:"<>|]', '', sheet_name).replace(' ', '_')
                
                # Convert table to DataFrame and write to Excel
                if len(table) > 1:  # Ensure table has header and data
                    df = pd.DataFrame(table[1:], columns=table[0])
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

    except Exception as e:
        print(f"Error processing {docx_path}: {str(e)}")

def process_pdf_files(folder_path):
    """Process all PDF files in the folder"""
    print(f"Starting to process folder: {folder_path}")
    
    # Get all PDF files
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    print(f"Found {len(pdf_files)} PDF files")
    
    for pdf_file in pdf_files:
        print(f"\nProcessing: {pdf_file}")
        pdf_path = os.path.join(folder_path, pdf_file)
        
        # Create temporary DOCX file
        docx_filename = pdf_file.replace('.pdf', '.docx')
        docx_path = os.path.join(folder_path, docx_filename)
        
        # Create Excel file path
        excel_filename = pdf_file.replace('.pdf', '.xlsx')
        excel_path = os.path.join(folder_path, excel_filename)
        
        try:
            # Convert PDF to DOCX
            print(f"Converting {pdf_file} to DOCX...")
            convert_pdf_to_docx(pdf_path, docx_path)
            
            # Extract tables and save to Excel
            print(f"Extracting tables to {excel_filename}...")
            extract_tables_from_docx(docx_path, excel_path)
            
            # Clean up temporary DOCX file
            os.remove(docx_path)
            print(f"Successfully processed {pdf_file}")
            
        except Exception as e:
            print(f"Error processing {pdf_file}: {str(e)}")

if __name__ == "__main__":
    print("Starting PDF table extraction tool...")
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Working directory: {script_dir}")
    
    # Process all PDF files in the directory
    process_pdf_files(script_dir)
    print("Processing complete!")
