# Email DLP Analysis Script

This Python script analyses a given email dataset (in .csv or .xls/.xlsx format) and provides the following information:

- Top 10 email senders
- Total number of PCI related emails and unique senders of PCI related emails
- Total number of quarantined, released, and not released emails
- Count of emails quarantined/released by brand (domain)

## Requirements

The script requires Python to be installed along with a few additional modules. These can be installed with pip using the provided `requirements.txt` file:

pip install -r requirements.txt

## Usage

The script takes two command-line arguments: the input file (CSV or Excel) and the output file (text). If the output file is not specified, it defaults to `results.txt`.

You can run the script with the following command:

python calcdata.py input_file.csv output_file.txt

Replace `input_file.csv` or `input_file.xls` with the name of your input CSV or Excel file and `output_file.txt` with the name of your output text file.
