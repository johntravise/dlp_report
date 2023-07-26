# Email DLP Analysis Script
This tool helps you understand your email data. It tells you things like who sends the most emails, how many emails contain PCI related information, and how many emails are being quarantined or released.

## What Your Data Should Look Like
Your data needs to be in a .csv or .xls/.xlsx (Excel) file. Your export from custom queries must have the following columns:
•	Sender Email Address: This is where the email came from.
•	Matched Rules: This tells us what rules the email matched (we use this to find emails related to PCI).
•	Quarantined State: This tells us if the email is currently quarantined.
•	Restored From Quarantine State: This tells us if the email has been released from quarantine.
•	Subject: This is the subject of the email.
•	Message ID: This is the Message ID of the email

## How to Install
1.	First, you need to install Python on your computer if you haven't already. You can download it from here.
2.	Next, download the script from this GitHub page. There should be a green button that says "Code". Click that and then click "Download ZIP".
3.	Unzip the downloaded file.
4.	Open your computer's command prompt or terminal. If you're not sure how to do this, you can search "command prompt" or "terminal" in your computer's search bar.
5.	Navigate to the unzipped folder using the cd command. For example, if you unzipped the folder to your desktop, you'd type something like cd Desktop/email-dlp-analysis-script-main.
6.	Now we need to install some extra tools the script uses. Type pip install -r requirements.txt and press Enter.

## How to Use
You'll use the command prompt or terminal to run the script.

### If You Want the Results in a Text File
1.	Type python calcdata.py, a space, and then the name of your data file. If your data file is in the same folder as the script, you can just type the name of the file (like mydata.csv). If it's in a different folder, you need to type the whole path to the file (like C:/Users/Me/Documents/mydata.csv).
2.	If you want the results to go in a specific text file, type another space and then the name of the text file you want to use (like myresults.txt). If you don't do this, the results will go in a file named results.txt.
3.	Press Enter.
For example, if your data is in a file named email_data.xlsx and you want the results in a file named analysis.txt, you'd type python calcdata.py email_data.xlsx analysis.txt.

### If You Want the Results in an Excel File
1.	Type python datatoexcel.py, a space, and then the name of your data file. If your data file is in the same folder as the script, you can just type the name of the file (like mydata.csv). If it's in a different folder, you need to type the whole path to the file (like C:/Users/Me/Documents/mydata.csv).
2.	If you want the results to go in a specific Excel file, type another space and then the name of the Excel file you want to use (like myresults.xlsx). If you don't do this, the results will go in a file named results.xlsx.
3.	Press Enter.
For example, if your data is in a file named email_data.xlsx and you want the results in a file named analysis.xlsx, you'd type python datatoexcel.py email_data.xlsx analysis.xlsx.
If You Have Any Problems
If something isn't working, try reading the instructions again to make sure you didn't miss anything. If it still doesn't work, you can ask for help in the "Issues" section of this GitHub page.


