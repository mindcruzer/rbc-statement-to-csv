#!/bin/bash

# define output filename
OUTPUT_FILE='transactions.csv'

# find pdf2txt.py
PDF2TXT=$(which pdf2txt.py)

PDF_FILES=$(find -H . -type f -regextype posix-extended -regex '\./.+\.pdf' -print)

for PDF in ${PDF_FILES};
do
	XML_FILE=$(basename -s .pdf "${PDF}")
	python "${PDF2TXT}" -o "${XML_FILE}.xml" "${PDF}" 
done

XML_FILES=($(find -H . -maxdepth 1 -type f -regextype posix-extended -regex '\./.+\.xml' -print))

echo "${XML_FILES[@]}"
python convert.py "${OUTPUT_FILE}" "${XML_FILES[@]}"

rm -rf *.xml
