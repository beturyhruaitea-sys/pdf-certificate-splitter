import os
import pandas as pd
from pypdf import PdfReader, PdfWriter

# 1. Automatically find the Excel file
excel_files = [f for f in os.listdir(".") if f.endswith('.xlsx') or f.endswith('.xls')]
if not excel_files:
    raise FileNotFoundError("No Excel file (.xlsx) found in this folder!")
excel_file = excel_files[0]
print(f"Selected Excel file: {excel_file}")

# 2. Automatically find the LARGE multi-page PDF file
large_pdf_files = [f for f in os.listdir(".") if f.endswith('.pdf') and f != "rename_pdf.py"]
if not large_pdf_files:
    raise FileNotFoundError("No PDF certificate file found to split!")

# If you have multiple PDFs, it picks the first one. Otherwise, change this to your exact filename.
input_pdf_path = large_pdf_files[0]
print(f"Selected PDF to split: {input_pdf_path}")

# 3. Load Excel data and auto-detect columns
df = pd.read_excel(excel_file)
actual_columns = list(df.columns)
lower_columns = [str(col).lower().strip() for col in actual_columns]

name_column = None
class_column = None

for i, col_name in enumerate(lower_columns):
    if "name" in col_name:
        name_column = actual_columns[i]
    if "class" in col_name or "sec" in col_name or "course" in col_name:
        class_column = actual_columns[i]

# Fallbacks to Column B and Column C if headers aren't detected by text search
if not name_column:
    name_column = actual_columns[1]
if not class_column:
    class_column = actual_columns[2]

print(f"Using Name Column: '{name_column}'")
print(f"Using Class Column: '{class_column}'\n")

# 4. Read the large PDF file
reader = PdfReader(input_pdf_path)
total_pages = len(reader.pages)
print(f"Total certificate pages found in PDF: {total_pages}")

# 5. Process page by page matching with Excel rows
for index, row in df.iterrows():
    # Stop if we run out of pages in the PDF file
    if index >= total_pages:
        print("\nWarning: You have more rows in Excel than pages in your PDF file.")
        break

    # Extract data from spreadsheet
    student_name = str(row[name_column]).strip()
    student_class = str(row[class_column]).strip()
    combined_name = f"{student_name}_{student_class}"

    # Sanitize filename for Windows
    safe_name = "".join(c for c in combined_name if c.isalnum() or c in (' ', '_', '-')).strip()
    output_filename = f"{safe_name}.pdf"

    # Extract the single page and save it
    writer = PdfWriter()
    writer.add_page(reader.pages[index])

    with open(output_filename, "wb") as output_pdf:
        writer.write(output_pdf)

    print(f"Page {index + 1} Split & Saved -> {output_filename}")

print("\nAll certificates split and named successfully!")
