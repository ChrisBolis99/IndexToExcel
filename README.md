# Folder Structure to Excel

This Python script scans a given folder and generates an Excel file containing the structure of the folder, including subfolders and files. Additionally, it records the type (folder or file) and the creation date of each entry.

---

## Features

- Scans a folder and its subfolders recursively.
- Outputs folder and file names in a hierarchical structure.
- Outputs the result to a Excel file.

---

## Requirements

- **Python 3.6+**
- **openpyxl** library for working with Excel files.

Install the required library using:
```bash
pip install openpyxl
```

## How to use

Edit the variables:

**folder_to_scan** = the folder to produce the index of

**output_file** = where the file will be saved

Then run the script with

```bash
python IndexToExcel.py
```
