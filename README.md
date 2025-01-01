# PDF to Excel Automation Tool

A Python script that processes data from a PDF file, converts it to CSV format, parses the data, and populates specific sections of a provided financial audit Excel spreadsheet (see "audit.xlsx" file) based on a given command.

## Features

- **PDF to CSV Conversion**: Extracts data from the PDF and saves it as a CSV file.
- **Data Parsing**: Analyzes and structures the extracted data.
- **Excel Population**: Fills specific sections of an input Excel spreadsheet based on parsed data and provided commands.

## Requirements

- Python 3.7+
- Required Python libraries:
  - `pandas`
  - `openpyxl`
  - `PyPDF2` or similar library for PDF processing

Install the dependencies:
```bash
pip install pandas openpyxl PyPDF2
```

## Usage

1. Run the script with the following arguments:
   ```bash
   python script.py <command> <data.pdf> <spreadsheet.xlsx>
   ```

   - `<command>`: Specifies the operation to perform: fillExecs or fillTransactions
   - `<data.pdf>`: Path to the input PDF file containing the data
   - `<spreadsheet.xlsx>`: Path to the Excel spreadsheet to be populated

2. The script will:
   - Convert the PDF to CSV
   - Parse the CSV data
   - Update the specified sections of the Excel spreadsheet

3. The output Excel file will be saved with updated content in the same directory

## Example

```bash
python script.py fillTransactions data.pdf audit.xlsx
```

This command converts `data.pdf` into a CSV, processes the data, and updates the `audit.xlsx` file based on the `fillTransactions` command.
