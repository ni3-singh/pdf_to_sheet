import subprocess
import sys

# Automatically install required packages if they are missing
subprocess.check_call([sys.executable, "-m", "pip", "install", "pdfplumber", "pandas", "openpyxl"])

import pdfplumber
import pandas as pd

pdf_file = "bank_statement.pdf"
excel_file = "bank_statement.xlsx"
csv_file = "bank_statement.csv"

all_tables = []

# Open the PDF
with pdfplumber.open(pdf_file) as pdf:
    for page in pdf.pages:
        # Extract tables from the page
        tables = page.extract_tables()

        for table in tables:
            # table is a list of lists representing rows and columns
            if table:
                # Convert each table into a pandas DataFrame
                df = pd.DataFrame(table)
                all_tables.append(df)

if all_tables:
    # Combine all extracted tables across all pages into one DataFrame
    final_df = pd.concat(all_tables, ignore_index=True)

    # Export to Excel
    final_df.to_excel(excel_file, index=False, header=False)

    # Export to CSV
    final_df.to_csv(csv_file, index=False, header=False)

    print(f"Successfully saved to {excel_file} and {csv_file}")
else:
    print("No tables detected in the PDF.")