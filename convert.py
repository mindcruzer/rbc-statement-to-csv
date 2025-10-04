#!/usr/bin/env python3
from datetime import datetime
from dateutil.parser import parse
import sys
import xml.etree.ElementTree as ET
import re
import csv
import os
import glob
from typing import List

# Lazy import of pdfminer when needed (keeps startup fast and avoids hard crash if not installed yet)
def _pdf_to_xml_root(pdf_path: str):
    from io import StringIO
    try:
        # Prefer high_level API with XML output
        from pdfminer.high_level import extract_text_to_fp
        from pdfminer.layout import LAParams
    except Exception as e:
        raise RuntimeError(
            "pdfminer.six is required to parse PDFs. Install dependencies first (see requirements.txt)."
        ) from e

    output = StringIO()
    with open(pdf_path, 'rb') as f:
        extract_text_to_fp(f, output, laparams=LAParams(), output_type='xml', codec=None)
    xml_text = output.getvalue()
    try:
        return ET.fromstring(xml_text)
    except ET.ParseError as e:
        # Surface a more actionable error with filename context
        raise RuntimeError(f"Failed to parse XML converted from PDF: {pdf_path}. Error: {e}") from e


def _get_xml_root(input_path: str):
    ext = os.path.splitext(input_path)[1].lower()
    if ext == '.xml':
        return ET.parse(input_path).getroot()
    if ext == '.pdf':
        return _pdf_to_xml_root(input_path)
    raise ValueError(f"Unsupported input type for '{input_path}'. Expected .pdf or .xml")


def _write_csv_with_retry(output_file: str, write_callback):
    """
    Open a CSV file for writing and execute write_callback(writer).
    If the file is locked (e.g., open in Excel), prompt the user to close it and press Enter to retry.
    """
    while True:
        try:
            with open(output_file, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                write_callback(writer)
            break
        except PermissionError:
            print(f"Cannot write to '{output_file}' because it's open in another program.")
            try:
                input("Please close the file, then press Enter to retry...")
            except KeyboardInterrupt:
                print("Write cancelled by user.")
                raise

output_file_credit = 'credit_transactions.csv'
output_file_chequing = 'chequing_transactions.csv'
output_file_savings = 'savings_transactions.csv'
input_files = sys.argv[1:]

font_header = "MetaBookLF-Roman"
font_txn = "MetaBoldLF-Roman"

class Block:
    def __init__(self, page, x, x2, y, text):
        self.page = page
        self.x = x
        self.x2 = x2
        self.text = text
        self.y = y
    
    def __repr__(self):
        return f"<Block page={self.page} x={self.x} x2={self.x2} y={self.y} text={self.text} />"

def process_credit_statements(input_files: List[str], output_file: str):
    txns = []
    re_exchange_rate = re.compile(r'Exchange rate-([0-9]+\.[0-9]+)', re.MULTILINE)
    re_foreign_currency = re.compile(r'Foreign Currency-([A-Z]+) ([0-9]+\.[0-9]+)', re.MULTILINE)

    for input_file in input_files:
        root = _get_xml_root(input_file)
        rows = []

        print(f'Processing {input_file}...')

        # Build line-oriented rows robustly by grouping text fragments by Y position
        # Do not assume a fixed child index (page[1]) as pdfminer XML can vary.
        pages = [el for el in list(root) if getattr(el, 'tag', None) == 'page']
        if not pages:
            # Fallback: find any nested page elements
            pages = list(root.findall('.//page'))

        for page in pages:
            # Collect candidate text segments with coordinates
            segments = []
            for tag in page.iter('text'):
                try:
                    size = float(tag.attrib.get('size', '0'))
                    if not (5 <= size <= 12):
                        continue
                    bbox = tag.attrib.get('bbox', '0,0,0,0').split(',')
                    x_pos = float(bbox[0])
                    y_pos = float(bbox[1])
                    x2_pos = float(bbox[2])
                    txt = tag.text or ''
                    if not txt:
                        continue
                except Exception:
                    continue
                segments.append((y_pos, x_pos, x2_pos, txt))

            if not segments:
                continue

            # Sort by Y descending (top to bottom), then X ascending (left to right)
            segments.sort(key=lambda s: (-s[0], s[1]))

            # Group into lines by Y with a small tolerance
            y_tol = 0.9
            line_items = []  # list of list of segments for each line
            current_y = None
            current_line = []
            for y_pos, x_pos, x2_pos, txt in segments:
                if current_y is None or abs(y_pos - current_y) > y_tol:
                    if current_line:
                        line_items.append(current_line)
                    current_line = []
                    current_y = y_pos
                current_line.append((x_pos, x2_pos, txt))
            if current_line:
                line_items.append(current_line)

            # For each grouped line, stitch text by X order, adding spaces on noticeable gaps
            for items in line_items:
                items.sort(key=lambda s: s[0])
                line_text = ''
                prev_x2 = None
                for x_pos, x2_pos, txt in items:
                    if prev_x2 is not None and (x_pos - prev_x2) > 0.7 and len(line_text) > 10:
                        line_text += ' '
                    line_text += txt
                    prev_x2 = x2_pos
                if line_text:
                    rows.append(line_text)

        date_range_regex = re.compile(r'^.*STATEMENT FROM ([A-Z]{3}) \d{2},? ?(\d{4})? TO ([A-Z]{3}) \d{2}, (\d{4})', re.MULTILINE)
        date_range = {}

        for row in rows:
            if match := date_range_regex.search(row):
                date_range[match.group(1)] = match.group(2) or match.group(4)
                date_range[match.group(3)] = match.group(4)
                break

        # If the statement header wasn't found (format change), infer year mapping from filename
        if not date_range:
            # Expect pattern like ...YYYY-MM-DD.pdf
            m = re.search(r'(\d{4})-(\d{2})-(\d{2})', os.path.basename(input_file))
            if m:
                end_year = int(m.group(1))
                end_month = int(m.group(2))
                months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
                for idx, mon in enumerate(months, start=1):
                    year_for_mon = end_year if idx <= end_month else end_year - 1
                    date_range[mon] = str(year_for_mon)

        MONTHS = {
            'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
            'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'
        }
        # Robust matcher for two date tokens at the beginning (optional spaces between month/day)
        mon_alt = '(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)'
        re_txn_prefix = re.compile(rf'^(?P<m1>{mon_alt})\s?(?P<d1>\d{{2}})(?P<m2>{mon_alt})\s?(?P<d2>\d{{2}})')
        txn_rows = []

        for row in rows:
            # Some rows may contain multiple lines; check line by line
            for line in row.splitlines():
                s = line.strip()
                if len(s) < 10:
                    continue
                m = re_txn_prefix.match(s)
                if m and m.group('m1') in MONTHS and m.group('m2') in MONTHS:
                    txn_rows.append(s)

        for row in txn_rows:
            m = re_txn_prefix.match(row)
            if not m:
                continue
            date_1_month = m.group('m1')
            date_1_day = m.group('d1')
            date_2_month = m.group('m2')
            date_2_day = m.group('d2')

            transaction_date = None
            try:
                transaction_date = datetime.strptime(f'{date_1_month}-{date_1_day}-{date_range[date_1_month]}', '%b-%d-%Y')
            except KeyError:
                first_year = min([int(year) for year in date_range.values()]) if date_range else datetime.now().year
                transaction_date = datetime.strptime(f'{date_1_month}-{date_1_day}-{first_year}', '%b-%d-%Y')
            try:
                posting_date = datetime.strptime(f'{date_2_month}-{date_2_day}-{date_range[date_2_month]}', '%b-%d-%Y')
            except KeyError:
                first_year = min([int(year) for year in date_range.values()]) if date_range else datetime.now().year
                posting_date = datetime.strptime(f'{date_2_month}-{date_2_day}-{first_year}', '%b-%d-%Y')

            # Description starts after the matched date tokens
            desc_and_amt = row[m.end():]
            # Extract the first monetary token (with optional $ and sign) to avoid trailing summary text
            m_amt = re.search(r'(-?\$?\d{1,3}(?:,\d{3})*(?:\.\d{2}))', desc_and_amt)
            if not m_amt:
                # If parsing fails, skip this line
                continue
            description = desc_and_amt[:m_amt.start()]
            amount = m_amt.group(1)

            # Normalize description and negative sign conventions
            if description.endswith('-'):
                description = description[:-1]
                if not amount.strip().startswith('-'):
                    amount = '-' + amount

            description = description.split("\n")[0].strip()
            raw = row.strip()

            amount = amount.replace('$', '').replace(',', '').replace("\n", "")
            match_exchange_rate = re_exchange_rate.search(raw)
            match_foreign_currency = re_foreign_currency.search(raw)
            
            if float(amount) > 0:
                txns.append({
                    'transaction_date': transaction_date,
                    'posting_date': posting_date,
                    'description': description,
                    'credit': '',
                    'debit': amount,
                    'raw': raw,
                    'exchange_rate': match_exchange_rate.group(1) if match_exchange_rate else None,
                    'foreign_currency': match_foreign_currency.group(1) if match_foreign_currency else None,
                    'amount_foreign': match_foreign_currency.group(2) if match_foreign_currency else None,
                })
            else:
                txns.append({
                    'transaction_date': transaction_date,
                    'posting_date': posting_date,
                    'description': description,
                    'credit': amount,
                    'debit': '',
                    'raw': raw,
                    'exchange_rate': match_exchange_rate.group(1) if match_exchange_rate else None,
                    'foreign_currency': match_foreign_currency.group(1) if match_foreign_currency else None,
                    'amount_foreign': match_foreign_currency.group(2) if match_foreign_currency else None,
                })

    txns = sorted(txns, key = lambda txn: txn['transaction_date'])

    # Only write the CSV if at least one transaction was parsed successfully
    if txns:
        def _write_credit(writer: csv.writer):
            writer.writerow([
                'Transaction Date',
                'Posting Date',
                'Description',
                'Credit',
                'Debit',
                'Amount Foreign Currency',
                'Foreign Currency',
                'Exchange Rate',
                'Raw',
            ])
            for txn in txns:
                writer.writerow([
                    txn['transaction_date'].strftime('%Y-%m-%d'),
                    txn['posting_date'].strftime('%Y-%m-%d'),
                    txn['description'],
                    txn['credit'],
                    txn['debit'],
                    txn['amount_foreign'],
                    txn['foreign_currency'],
                    txn['exchange_rate'],
                    txn['raw'],
                ])

        _write_csv_with_retry(output_file, _write_credit)
    else:
        print(f"No credit transactions detected. Not creating '{output_file}'.")

def process_chequing_statements(input_files: List[str], output_file: str):
    csv_rows = []

    for input_file in input_files:
        root = _get_xml_root(input_file)
        rows = []

        continue_input_loop = False
        for i_tag, tag in enumerate(root[0][1]):
            if i_tag > 10:
                continue_input_loop = True
                break
            if font := tag.get("font"):
                if font.endswith("MetaBoldLF-Roman") or font.endswith("Utopia-Bold"):
                    break
        
        if continue_input_loop:
            print(f"Skipping {input_file}...")
            continue

        print(f'Processing {input_file}...')

        blocks = []
        pages = set()
        for page_num, page in enumerate(root):
            pages.add(page_num)
            figure = page[1]
            text = ''
            last_x = None
            last_x2 = None
            block_x = None
            width = 0
            seen_text = False
            for tag in figure:
                clear_text = False
                append_block = ""
                if tag.tag == 'text':
                    seen_text = True
                    font = tag.attrib.get("font")
                    if font:
                        font = font.split("+")[1]
                    size = float(tag.attrib['size'])
                    x_pos = float(tag.attrib["bbox"].split(",")[0])
                    y_pos = float(tag.attrib["bbox"].split(",")[1])
                    x2_pos = float(tag.attrib["bbox"].split(",")[2])

                    if last_x2 is not None:
                        if x_pos - last_x2 > 5:
                            append_block = text
                            text = ""
                            width = 0
                        elif (x_pos - last_x2) > 0.7:
                            text += " "
                    last_x = x_pos
                    last_x2 = x2_pos
                    width += size
                    if font in (font_txn, font_header):
                        text += tag.text
                    if block_x is None:
                        block_x = x_pos
                elif tag.tag != 'text' and text != '':
                    if seen_text:
                        append_block = text
                    seen_text = False
                    clear_text = True
                    width = 0
                    last_x2 = None
                
                if append_block:
                    block = Block(page_num, block_x, x2_pos, y_pos, append_block.strip())
                    blocks.append(block)
                    block_x = None

                if clear_text:
                    text = ''

        open_balance_parts = [b.text for b in blocks if b.text.startswith("Your opening balance")][0].split(" ")[-3:]
        open_balance_date = parse(" ".join(open_balance_parts))
        start_year = int(open_balance_parts[2])

        header_sets = []
        for page in pages:
            page_blocks = [b for b in blocks if b.page == page]
            end_of_header_index = 0

            for i, block in enumerate(page_blocks):
                if block.text == "Date":
                    if other_blocks := page_blocks[i+1:i+5]:
                        header_sets.append([block, *other_blocks])
                        end_of_header_index = i + 4

            page_blocks = [block for i, block in enumerate(page_blocks) if i > end_of_header_index]

            if len(header_sets) <= page:
                break

            cell_pos = 0
            i = 0
            block_pos = 0
            row = []
            last_date = None

            while block_pos < len(page_blocks):
                row_pos = 0
                block = page_blocks[block_pos]
                block_consumed = False

                headers = header_sets[block.page]

                mid_point = (block.x2 - block.x) / 2 + block.x
                for header in headers:
                    if mid_point > header.x and mid_point < header.x2:
                        row_pos = headers.index(header)

                if i % 5 == row_pos:
                    if i % 5 == 0:
                        date = parse(f"{block.text} {start_year}")

                        if date < open_balance_date:
                            date = parse(f"{block.text} {start_year+1}")
                        block.text = str(date.date())
                        if block.text.strip():
                            last_date = block.text
                    block_consumed = True
                    row.append(block.text)
                elif i % 5 == 0 and page_blocks[block_pos].text == "Opening Balance":
                    row.append(str(open_balance_date.date()))
                elif last_date and i % 5 == 0:
                    row.append(last_date)
                else:
                    row.append("")
                if i % 5 == 4:
                    csv_rows.append(row)
                    row = []
                if block_consumed:
                    block_pos += 1
                i += 1

    def _write_chequing(writer: csv.writer):
        writer.writerow([
            "Date",
            "Description",
            "Withdrawls",
            "Deposits",
            "Balance"
        ])

        for row in csv_rows:
            if "Opening Balance" in row:
                description = "Opening Balance"
                deposits = ""
                withdrawals = ""
            else:
                description = row[1]
                deposits = row[2]
                withdrawals = row[3]

            writer.writerow([
                row[0],
                description,
                withdrawals,
                deposits,
                row[4]
            ])

    _write_csv_with_retry(output_file, _write_chequing)

if __name__ == "__main__":
    # If no arguments were provided, auto-discover PDFs in the current directory
    if not input_files:
        pdfs = []
        # Match PDFs in current directory; on Windows glob is case-insensitive,
        # so searching twice ("*.pdf" and "*.PDF") can create duplicates.
        for pattern in ("*.pdf",):
            pdfs.extend(glob.glob(os.path.join(os.getcwd(), pattern)))
        # Deduplicate while preserving order
        seen = set()
        deduped = []
        for p in pdfs:
            key = os.path.normcase(os.path.abspath(p))
            if key not in seen:
                seen.add(key)
                deduped.append(p)
        pdfs = deduped
        if not pdfs:
            print("No input files provided and no .pdf files found in the current directory.")
            print("Usage: python convert.py [optional files... (PDF or XML)]")
            sys.exit(1)
        input_files = pdfs

    # Narrow down to visa/chequing/savings by filename when possible, otherwise try both processors gracefully
    credit_files = [f for f in input_files if "visa" in os.path.basename(f).lower()]
    chequing_files = [f for f in input_files if "chequing" in os.path.basename(f).lower()]
    # Savings statements share the chequing layout but write to their own CSV
    savings_files = [f for f in input_files if "savings" in os.path.basename(f).lower()]

    # If nothing matched the patterns, be permissive and attempt to classify by structure
    other_files = [f for f in input_files if f not in credit_files + chequing_files + savings_files]
    for f in other_files:
        # Heuristic: try both, catching errors and proceeding
        try:
            process_credit_statements([f], output_file_credit)
        except Exception:
            try:
                process_chequing_statements([f], output_file_chequing)
            except Exception:
                print(f"Skipping unrecognized file type/format: {f}")

    if credit_files:
        process_credit_statements(credit_files, output_file_credit)

    if chequing_files:
        process_chequing_statements(chequing_files, output_file_chequing)

    if savings_files:
        process_chequing_statements(savings_files, output_file_savings)