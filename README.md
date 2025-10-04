# RBC Statement To CSV

This program reads transaction data out of RBC statements and writes them to CSVs.

Supported statements (CSV created only if at least one matching statement is processed):
- Credit Card (Visa) → `credit_transactions.csv` with the following columns:

  - Transaction Date
  - Posting Date
  - Description
  - Credit
  - Debit
  - Amount Foreign Currency
  - Foreign Currency
  - Exchange Rate
  - Raw

- Chequing (identical layout to Savings) → `chequing_transactions.csv` with the following columns:
  - Date
  - Description
  - Withdrawls
  - Deposits
  - Balance

- Savings (identical layout to Chequing) → `savings_transactions.csv` with the same columns as chequing.

## Requirements
- Python 3.8+
- [PdfMiner](https://github.com/pdfminer/pdfminer.six)
- [python-dateutil](https://dateutil.readthedocs.io/en/stable/)

## Use
Drop all PDF statements into the project directory. The program will auto-discover PDFs in the current folder (or you can pass specific files on the command line), read all transactions, sort them, and write only the CSV file(s) that correspond to the statement types actually processed. Savings statements (e.g., `Savings Statement-4484 2025-01-15.pdf`) are processed the same as chequing but written to `savings_transactions.csv`.

Notes:
- Visa parsing supports both older and newer RBC layouts. FX details (Exchange rate and Foreign Currency) are extracted when present.
- CSV files are only created when at least one matching statement is successfully processed (no empty header-only files).
- If a CSV file is open in another program (e.g., Excel), the script will prompt you to close the file and press Enter, then retry writing.

### Windows
Ensure `.py` files are associated with Python.

Run (double-click) `convert.py`.

### Linux/macOS
Run `./convert.py` in a terminal.
