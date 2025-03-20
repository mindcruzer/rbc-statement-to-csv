# RBC Statement To CSV

This program will read transaction data out of RBC credit card statements and put them in a .csv file with the following columns:

- Transaction Date
- Posting Date
- Description
- Amount
- Amount In Foreign Currency
- Foreign Currency
- Exchange Rate

## Requirements
- Python 3.8+
- [PdfMiner](https://github.com/pdfminer/pdfminer.six)

## Use
Drop all PDF statements into the project directory. The program will read all transactions, sort them, and consolidate them into a single .csv file.

### Windows
Ensure `.py` files are associated with Python.

Run (double-click) `convert.py`.

### Linux/macOS
Run `./convert.py` in a terminal.
