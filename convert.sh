#!/bin/bash

# define output filename
OUTPUT_FILE='transactions.csv'

# find pdf2txt.py
PDF2TXT=$(which pdf2txt.py)

mkdir -p ./tmp

# -print0 will print the file name followed by a null character instead of newline
find . -name '*.pdf' -print0 | while IFS= read -r -d '' PDF; do
    # Use quotes around "$PDF" to handle spaces in filenames
    XML_FILE=$(basename -s .pdf "$PDF")
    python "$PDF2TXT" -o "./tmp/${XML_FILE}.xml" "$PDF"
done

# Collect XML filenames into an array, handling spaces properly
XML_FILES=()
while IFS=  read -r -d $'\0'; do
    XML_FILES+=("$REPLY")
done < <(find ./tmp -name '*.xml' -print0)

echo "${XML_FILES[@]}"
python convert.py "$OUTPUT_FILE" "${XML_FILES[@]}"

# Clean up XML files
rm -rf tmp/*.xml
