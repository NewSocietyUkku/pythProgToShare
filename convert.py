# This script converts the first sheet of an Excel (.xlsx) file into an HTML table.

import pandas as pd
import os
import sys

# --- CONFIGURATION ---
# 1. Input file: Use the simplified, clean filename 'test.xlsx'
XLSX_FILE = "test.xlsx"
# 2. Output file: The resulting HTML table
HTML_OUTPUT = "test.html"
# 3. Required packages (pandas uses openpyxl to read .xlsx)
REQUIRED_PACKAGES = ['pandas', 'openpyxl']
# ---------------------

def check_dependencies():
    """Checks if required libraries are installed and exits if not."""
    print("Checking required packages...")
    for package in REQUIRED_PACKAGES:
        try:
            __import__(package)
        except ImportError:
            print(f"Error: Python package '{package}' not found.")
            print(f"Please install it using: pip install {package}")
            sys.exit(1)
    print("All packages found.")

def convert_excel_to_html(xlsx_file, html_output):
    """Reads the Excel file and converts it to a styled HTML file."""

    # 1. Validate file existence
    if not os.path.exists(xlsx_file):
        print(f"\nFATAL ERROR: Input file not found.")
        print(f"Expected file: '{xlsx_file}'.")
        print("Please ensure this file is in the current directory and the name is correct.")
        return

    try:
        # 2. Read the first sheet of the Excel file into a pandas DataFrame
        print(f"Reading data from '{xlsx_file}'...")
        # We specify the engine explicitly, although pandas usually detects it.
        df = pd.read_excel(xlsx_file, engine='openpyxl')

        # 3. Convert the DataFrame to HTML string with custom styling
        # index=False prevents writing the pandas DataFrame row numbers to the HTML
        html_string = df.to_html(index=False, border=1)

        # 4. Add CSS styling for better readability (mobile friendly)
        # This styling is simple and keeps the HTML file self-contained.
        styled_html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Spreadsheet Data</title>
    <style>
        body {{ font-family: sans-serif; margin: 10px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #f2f2f2; color: #333; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
    </style>
</head>
<body>
{html_string}
</body>
</html>
"""

        # 5. Save the styled HTML to the output file
        with open(html_output, 'w', encoding='utf-8') as f:
            f.write(styled_html)

        print("-" * 50)
        print(f"SUCCESS: Data converted and saved to '{html_output}'")
        print("-" * 50)

    except Exception as e:
        print(f"\nAn unexpected error occurred during conversion:")
        print(f"Error details: {e}")
        print("Please check if the Excel file is corrupted or uses an unsupported format.")


if __name__ == "__main__":
    check_dependencies()
    convert_excel_to_html(XLSX_FILE, HTML_OUTPUT)
