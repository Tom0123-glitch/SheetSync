# Excel Sync GUI

A Python desktop application with a Tkinter-based GUI for synchronizing columns between two Excel spreadsheets.  
The tool allows users to upload a **source file** and a **destination file**, select which columns to copy, and update the destination spreadsheet while preserving formulas if required.

---

## Features

- **Upload Excel files** via a simple GUI (source and destination).
- **Select columns to copy** from the source to the destination.
- **Key-based overwrite**: updates are aligned by a shared ID column.
- **Formula protection**: prevents overwriting destination formulas (optional).
- **Corporate branding**: supports customization of colors and logo (e.g., Leeds Building Society).
- **Scrollable interface** for handling many columns.

---

## Requirements

- Python 3.9+
- Dependencies:
  - `pandas`
  - `openpyxl`
  - `tkinter` (usually bundled with Python)

Install required libraries with:

```bash
pip install pandas openpyxl
