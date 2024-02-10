# xlsxgrader

Converts exported CSV files to an xlsx format suitable for rapid grading and grade entry.

# Install

`pip install git+https://github.com/taobrienlbl/xlsxgrader.git`

# Usage

`xlsxgrader INFILE.csv --output OUTFILE.xlsx`

or for batch conversion (writes to the current directory with the original filename, but with xlsx extension for all files):

`xlsxgrader INFILE1.csv INFILE2.csv ...`
