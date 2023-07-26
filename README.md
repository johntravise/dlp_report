# Email DLP Analysis Script
This Python script analyses a given email dataset (in .csv or .xls/.xlsx format) and provides the following information:
 - Top 10 email senders
 - Total number of PCI related emails and unique senders of PCI related emails
 - Total number of quarantined, released, and not released emails
 - Count of emails quarantined/released by brand (domain)
 - The results are outputted to a text file when using calcdata.py or an Excel file when using datatoexcel.py.

##Requirements
The script requires Python to be installed along with a few additional modules. These can be installed with pip using the provided requirements.txt file:
`pip install -r requirements.txt`

## Data Format
The input data file should be in .csv or .xls/.xlsx format and must contain the following columns:
 - Sender Email Address: The email address of the sender
 - Matched Rules: Rules matched for the email (used to identify PCI related emails)
 - Quarantined State: The current state of quarantine for the email
 - Restored From Quarantine State: Whether the email has been released from quarantine
 - Subject: The subject of the email
 - Message ID: The email's Message ID

## Usage
### calcdata.py
This script takes multiple command-line arguments: the input files (CSV or Excel) and the last argument as the output file (text). If the output file is not specified, it defaults to results.txt.

You can run the script with the following command:
`python calcdata.py input_file1.csv input_file2.csv output_file.txt`
Replace `input_file1.csv`, `input_file2.csv`, etc. with the names of your input CSV or Excel files and `output_file.txt` with the name of your output text file. The script will combine the input files, perform the analysis, and write the results to the output text file.

### datatoexcel.py
This script is similar to calcdata.py, but instead outputs the results to an Excel file. It takes multiple command-line arguments: the input files (CSV or Excel) and the last argument as the output file (Excel).

You can run the script with the following command:
`python datatoexcel.py input_file1.xlsx input_file2.xlsx output_file.xlsx`
Replace `input_file1.xlsx`, `input_file2.xlsx`, etc. with the names of your input Excel files and `output_file.xlsx` with the name of your output Excel file. The script will combine all the input files, perform the analysis, and write the results to the output Excel file.

# Windows Executables
For Windows users, we've added executable versions of both scripts in the executable folder in the GitHub repository. These can be run without needing to install Python or any additional modules. Use calcdata.exe for text output and datatoexcel.exe for Excel output. Note that the data format and command line arguments remain the same as the Python scripts.
