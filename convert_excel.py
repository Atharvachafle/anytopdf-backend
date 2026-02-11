# Save this as: convert_excel.py
import sys
import os
import pdfplumber
import pandas as pd

try:
    pdf_file = os.path.abspath(sys.argv[1])
    xlsx_file = os.path.abspath(sys.argv[2])
except IndexError:
    print("ERROR: Missing arguments")
    sys.exit(1)

if not os.path.exists(pdf_file):
    print(f"ERROR: File not found {pdf_file}")
    sys.exit(1)

print(f"Processing: {pdf_file}")

try:
    # Open the PDF
    with pdfplumber.open(pdf_file) as pdf:
        writer = pd.ExcelWriter(xlsx_file, engine='openpyxl')
        
        # Process each page
        for i, page in enumerate(pdf.pages):
            # 1. Extract Tables
            tables = page.extract_tables()
            
            if tables:
                # If tables found, save them cleanly
                for j, table in enumerate(tables):
                    df = pd.DataFrame(table)
                    # Clean up: Replace None with empty string
                    df.fillna('', inplace=True)
                    
                    # Save to a sheet (Page 1 Table 1, etc.)
                    sheet_name = f"Page{i+1}_Table{j+1}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            else:
                # 2. Fallback: If no strict table found, try to extract text lines
                # This mimics the "visual" layout
                text = page.extract_text()
                lines = text.split('\n') if text else []
                df = pd.DataFrame(lines)
                df.to_excel(writer, sheet_name=f"Page_{i+1}_Text", index=False, header=False)

        writer.close()
        print("SUCCESS")

except Exception as e:
    print(f"ERROR: {str(e)}")
    sys.exit(1)