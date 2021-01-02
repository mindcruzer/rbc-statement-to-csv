from datetime import datetime
import sys
import xml.etree.ElementTree as ET
import re
import csv

output_file = sys.argv[1]
input_files = sys.argv[2:]
txns = []
re_exchange_rate = re.compile(r'Exchange rate-([0-9]+\.[0-9]+)', re.MULTILINE)
re_foreign_currency = re.compile(r'Foreign Currency-([A-Z]+) ([0-9]+\.[0-9]+)', re.MULTILINE)

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
        last_x2 = None
        # A row is a list of <text> tags, each containing a character
        for tag in figure:
            if tag.tag == 'text':
                # Filter on text size to remove some of the noise
                size = float(tag.attrib['size'])
                x_pos = float(tag.attrib["bbox"].split(",")[0])
                x2_pos = float(tag.attrib["bbox"].split(",")[2])

                if last_x2 is not None:
                    if x2_pos < last_x2:
                        row += "\n"

                    if len(row) > 10 and (x_pos - last_x2) > 0.7:
                        row += " "
                last_x2 = x2_pos
                if int(size) in [6, 8]:
                    row += tag.text
            elif tag.tag != 'text' and row != '':
                # Row is over, start a new one
                rows.append(row)
                row = ''
                last_x2 = None

    # Get date range of the statement
    date_range_regex = re.compile(r'^.*STATEMENT FROM ([A-Z]{3}) \d{2},? ?(\d{4})? TO ([A-Z]{3}) \d{2}, (\d{4})', re.MULTILINE)
    date_range = {}

    for row in rows:
        if match := date_range_regex.search(row):
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

        transaction_date = None
        try:
            transaction_date = datetime.strptime(f'{date_1_month}-{date_1_day}-{date_range[date_1_month]}', '%b-%d-%Y')
        except KeyError:
            # there is a strange case where the first date was before the days specified in date_range
            # so just use the first year
            first_year = min([int(year) for year in date_range.values()])
            transaction_date = datetime.strptime(f'{date_1_month}-{date_1_day}-{first_year}', '%b-%d-%Y')
        posting_date = datetime.strptime(f'{date_2_month}-{date_2_day}-{date_range[date_2_month]}', '%b-%d-%Y')
        
        description, amount = row[10:].split('$')

        if description.endswith('-'):
            description = description[:-1]
            amount = '-' + amount

        # split desc after negative check, otherwise `-` gets left behind
        description = description.split("\n")[0]
        raw = row.strip()

        amount = amount.replace(',', '').replace("\n", "")
        match_exchange_rate = re_exchange_rate.search(raw)
        match_foreign_currency = re_foreign_currency.search(raw)
        txns.append({
            'transaction_date': transaction_date,
            'posting_date': posting_date,
            'description': description,
            'amount': amount,
            'raw': raw,
            'exchange_rate': match_exchange_rate.group(1) if match_exchange_rate else None,
            'foreign_currency': match_foreign_currency.group(1) if match_foreign_currency else None,
            'amount_foreign': match_foreign_currency.group(2) if match_foreign_currency else None,
        })

txns = sorted(txns, key = lambda txn: txn['transaction_date'])

# Write as csv
with open(output_file, 'w', newline='') as csvfile:
    csv_writer = csv.writer(csvfile)
    csv_writer.writerow([
        'Transaction Date',
        'Posting Date',
        'Description',
        'Amount',
        'Amount Foreign Currency',
        'Foreign Currency',
        'Exchange Rate',
        'Raw',
    ])

    for txn in txns:
        csv_writer.writerow([
            txn['transaction_date'].strftime('%Y-%m-%d'),
            txn['posting_date'].strftime('%Y-%m-%d'),
            txn['description'],
            txn['amount'],
            txn['amount_foreign'],
            txn['foreign_currency'],
            txn['exchange_rate'],
            txn['raw'],
        ])
