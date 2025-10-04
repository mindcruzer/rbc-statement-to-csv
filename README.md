# RBC Statement To CSV

This tool converts Royal Bank of Canada (RBC) statements - Credit Card (Visa), Chequing, and Savings - into CSV files.

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

## Installation

### Prerequisites
- Python 3.8 or higher

### macOS

1. Create a virtual environment:
```bash
python3 -m venv .venv
```

2. Activate the virtual environment:
```bash
source .venv/bin/activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

### Windows

1. Create a new Python virtual environment using PowerShell:
```powershell
python -m venv .venv
```

2. Activate the virtual environment:
```powershell
.venv\Scripts\activate
```

3. Install required packages:
```powershell
pip install -r requirements.txt
```

## Usage
Drop all PDF statements into the project directory. The program will auto-discover PDFs in the current folder (or you can pass specific files on the command line), read all transactions, sort them, and write only the CSV file(s) that correspond to the statement types actually processed. Savings statements (e.g., `Savings Statement-4484 2025-01-15.pdf`) are processed the same as chequing but written to `savings_transactions.csv`.

Notes:
- Visa parsing supports both older and newer RBC layouts. FX details (Exchange rate and Foreign Currency) are extracted when present.
- CSV files are only created when at least one matching statement is successfully processed (no empty header-only files).
- If a CSV file is open in another program (e.g., Excel), the script will prompt you to close the file and press Enter, then retry writing.

### macOS

Run the script from a terminal:
```bash
python3 convert.py
```

### Windows

To run the script, you can either:

Double-click `convert.py`.

> [!IMPORTANT]
> Ensure `.py` files are associated with Python.

or

Run from PowerShell:

```powershell
python convert.py
```
