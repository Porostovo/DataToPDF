# DataToPDF


Here is a sample of the source code for the simple desctop application I developed in collaboration with my colleagues, which transfers data between Excel and original PDF protocols.

This Python script automates the process of generating and updating PDF forms using data from Excel files. It leverages `PyPDF2` for PDF manipulations and `pandas` for handling Excel data.

## Features

- **PDF Form Update**: Fills out and updates PDF form fields with data from an Excel spreadsheet.
- **Watermarks and Signatures**: Adds watermarks and signatures to PDFs.
- **Read-Only Fields**: Marks form fields as read-only after updating them.
- **PDF Merging**: Merges multiple generated PDFs into a single document.

## Usage

1. **Prepare Files**:
    - Excel data file: `./source/source_tables.xlsx`
    - PDF templates: `./source/source_forms/`
    - Watermark files: `./source/watermarks/`
    - Signature files: `./source/signs/`
    - Signature mapping: `./source/sign_names.txt`

2. **Run the Script**:
    - Execute the script using the command: `python script_name.py`

## Execution

Run the script with the following command:

```sh
python script_name.py
