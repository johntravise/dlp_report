# Email DLP Analysis Script

This Python script analyses a given email dataset (in .csv or .xls/.xlsx format) and provides the following information:

1. Top 10 email senders
2. Total number of PCI related emails and unique senders of PCI related emails
3. Total number of quarantined, released, and not released emails
4. Count of emails quarantined/released by brand (domain)

The results are outputted to a text file when using `calcdata.py` or an Excel file when using `datatoexcel.py`.

## Requirements

The script requires Python to be installed along with a few additional modules. These can be installed with pip using the provided requirements.txt file:

`pip install -r requirements.txt`

## Usage
### calcdata.py
The script takes multiple command-line arguments: the input files (CSV or Excel) and the last argument as the output file (text). If the output file is not specified, it defaults to results.txt.

You can run the script with the following command:

`python calcdata.py input_file1.csv input_file2.csv output_file.txt`

Replace input_file.csv or input_file.xls with the name of your input CSV or Excel files and output_file.txt with the name of your output text file. The script will combine the input files, perform the analysis, and write the results to the output text file.

### datatoexcel.py
This script is similar to calcdata.py, but instead outputs the results to an Excel file. It takes multiple command-line arguments: the input files (CSV or Excel) and the last argument as the output file (Excel).

You can run the script with the following command:

`python datatoexcel.py input_file1.xlsx input_file2.xlsx output_file.xlsx`

Replace input_file1.xlsx, input_file2.xlsx, etc. with the names of your input Excel files and output_file.xlsx with the name of your output Excel file. The script will combine all the input files, perform the analysis, and write the results to the output Excel file.
