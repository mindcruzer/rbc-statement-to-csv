# RBC Statement To CSV

This program will read transaction data out of RBC credit card statements and put them in a .csv file with the following columns:

- Transaction date
- Posting date
- Description
- Amount

## Requirements
- Python 3.8+
- [PdfMiner](https://github.com/euske/pdfminer)

Ensure `python` is in your PATH.

## Use
Drop all PDF statements into the project directory. The program will read all transactions, sort them, and consolidate them into a single .csv file.

### Windows
Run (double-click) `convert.bat`.

### Linux
Run `convert.sh` in terminal.
