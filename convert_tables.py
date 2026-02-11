# Save this as: convert_tables.py
import sys
import os
from pdf2docx import Converter

# 1. Get Arguments
try:
    pdf_file = os.path.abspath(sys.argv[1])
    docx_file = os.path.abspath(sys.argv[2])
except IndexError:
    print("ERROR: Missing arguments")
    sys.exit(1)

try:
    print(f"Processing: {pdf_file}")
    
    cv = Converter(pdf_file)
    
    # ⚙️ ADVANCED SETTINGS FOR FORMS
    # These settings force the script to "Snap" text into the grid lines
    # instead of letting it float loosely.
    settings = {
        "start": 0,
        "end": None,
        
        # 1. TABLE DETECTION (Crucial for your form)
        "detect_vertical_lines": True,    # Find vertical table borders
        "detect_horizontal_lines": True,  # Find horizontal table borders
        "connected_border_tolerance": 0.5,# Join lines that almost touch
        
        # 2. ALIGNMENT CORRECTION
        "snap_tolerance": 4.0,            # Snap text to nearest grid line (Fixes "floating" text)
        "join_tolerance": 3.0,            # Join words that are close together
        
        # 3. MARGINS (Reset to 0 to avoid shifting)
        "margin_bottom": 0,
        "margin_top": 0,
        "margin_left": 0,
        "margin_right": 0,
    }

    # Convert using these strict settings
    cv.convert(docx_file, **settings)
    
    cv.close()
    print("SUCCESS")

except Exception as e:
    print(f"ERROR: {str(e)}")
    sys.exit(1)