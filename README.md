# RBC PDF Statement Reader
Converts RBC credit card PDF Statements to CSV format.

## Requirements
- Python 3.8+
- [PdfMiner](https://github.com/euske/pdfminer)

Ensure `python` is in your PATH.

## Use
Drop all PDF statements into the project directory. The program will read all transactions, sort them, and consolidate them into a single .csv file.

### Windows
Run (double-click) `convert.bat`.

### Linux
I should probably write a script for Linux...