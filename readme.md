# PDF Certificate Splitter

A Python script that automatically splits a multi-page PDF into individual PDF files and renames each file using student information from an Excel spreadsheet.

## Features

- Automatically detects the Excel file
- Automatically detects the PDF file
- Matches each PDF page to an Excel row
- Renames files using Name and Class
- Supports .xlsx files
- Works on Windows

## Requirements

- Python 3.10+
- pandas
- openpyxl
- pypdf

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Place these files in the same folder:

```
list.xlsx
certificates.pdf
split_and_rename.py
```

Run:

```bash
python split_and_rename.py
```

## License

MIT License