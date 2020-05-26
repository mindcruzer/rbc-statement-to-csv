@echo off
Rem Ensure python is in your PATH.
Rem Requires pdfminer Python package: pip install pdfminer

set output_file="transactions.csv"

Rem Find pdf2txt.py
for /f %%i in ('where pdf2txt.py') do set pdf2txt_path=%%i

Rem Remove any existing output file
if exist %output_file% (
    del %output_file%
)

Rem Convert each .pdf file in the directory to .xml using pdf2txt.py
for /r %%v in (*.pdf) do (
    IF not exist "%%v.xml" (
        echo Converting %%v...
        python %pdf2txt_path% -o "%%v.xml" "%%v"
    ) 
)

Rem Get a list of .xml files in the directory
set input_files=
for /f "delims=" %%a in ('dir "*.xml" /on /b /a-d ') do call set input_files=%%input_files%% "%%a"

Rem Extract the transactions from every .xml file and add to the output file
python convert.py %output_file% %input_files%

Rem Remove the .xml files that were generated
echo Cleaning up...
del "*.xml"

echo ------
echo Transaction data is available in %output_file%.

pause