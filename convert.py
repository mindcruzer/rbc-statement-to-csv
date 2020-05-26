from datetime import datetime
import sys
import xml.etree.ElementTree as ET
import re
import csv

output_file = sys.argv[1]
input_files = sys.argv[2:]
txns = []

for input_file in input_files:
    tree = ET.parse(input_file)
    root = tree.getroot()
    rows = []

    print(f'Processing {input_file}...')

    # Go through each page
    for page in root:
        # Txn rows are in the second figure
        figure = page[1]
        row = ''
        # A row is a list of <text> tags, each containing a character
        for tag in figure:
            if tag.tag == 'text':
                # Filter on text size to remove some of the noise
                size = float(tag.attrib['size'])
                if int(size) == 8:
                    row += tag.text
            elif tag.tag != 'text' and row != '':
                # Row is over, start a new one
                rows.append(row)
                row = ''

    # Get date range of the statement
    date_range_regex = re.compile(r'^.+STATEMENTFROM([A-Z]{3})\d{2},?(\d{4})?TO([A-Z]{3})\d{2},(\d{4})')
    date_range = {}

    for row in rows:
        if match := date_range_regex.match(row):
            # Year for start month may not be specified if it's the same 
            # as the end month
            date_range[match.group(1)] = match.group(2) or match.group(4)
            date_range[match.group(3)] = match.group(4)
            break

    # Filter down to rows that are for transactions
    MONTHS = {
        'JAN',
        'FEB',
        'MAR',
        'APR',
        'MAY',
        'JUN',
        'JUL',
        'AUG',
        'SEP',
        'OCT',
        'NOV',
        'DEC'
    }
    txn_rows = []

    for row in rows:
        # Match txn rows based on month of txn date and posting date
        if len(row) >= 10:
            month_1 = row[:3]
            month_2 = row[5:8]

            if month_1 in MONTHS and month_2 in MONTHS:
                txn_rows.append(row)

    # Parse and format the transaction data
    for row in txn_rows:
        date_1_month = row[:3]
        date_1_day = row[3:5]
        date_2_month = row[5:8]
        date_2_day = row[8:10]

        transaction_date = datetime.strptime(f'{date_1_month}-{date_1_day}-{date_range[date_1_month]}', '%b-%d-%Y')
        posting_date = datetime.strptime(f'{date_2_month}-{date_2_day}-{date_range[date_2_month]}', '%b-%d-%Y')
        
        description, amount = row[10:].split('$')

        if description.endswith('-'):
            description = description[:-1]
            amount = '-' + amount

        amount = amount.replace(',', '')
        
        txns.append({
            'transaction_date': transaction_date,
            'posting_date': posting_date,
            'description': description,
            'amount': amount
        })

txns = sorted(txns, key = lambda txn: txn['transaction_date'])

# Write as csv
with open(output_file, 'w', newline='') as csvfile:
    csv_writer = csv.writer(csvfile)
    csv_writer.writerow(['Transaction Date', 'Posting Date', 'Description', 'Amount'])

    for txn in txns:
        csv_writer.writerow([
            txn['transaction_date'].strftime('%Y-%m-%d'),
            txn['posting_date'].strftime('%Y-%m-%d'),
            txn['description'],
            txn['amount']
        ])