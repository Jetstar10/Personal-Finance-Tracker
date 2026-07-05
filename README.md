# Bank Email Transactions

Bank Email Transactions is a Python project that helps automate personal finance tracking by reading DBS/POSB bank transaction alert emails from Gmail and converting them into organized Excel financial reports.

Instead of manually checking every bank email and typing transactions into a spreadsheet, this script connects to Gmail, finds bank alert emails, extracts the transaction details, categorizes them, and creates yearly Excel reports. The generated reports make it easier to review income, expenses, monthly spending, yearly totals, and category breakdowns.

## What This Project Does

This project is designed to save time when tracking bank transactions. It reads transaction alert emails from DBS/POSB, identifies important details such as the transaction amount, date, category, account, and description, then exports the data into Excel files.

The script also keeps track of emails that have already been processed, so running it multiple times will not duplicate the same transactions.

## Features

- Reads DBS/POSB transaction alert emails from Gmail
- Extracts transaction amount, date, type, party, account, and description
- Automatically categorizes transactions
- Separates income, expenses, and net balance
- Creates yearly Excel financial reports
- Generates monthly summaries
- Generates yearly summaries
- Includes category breakdowns for better spending analysis
- Tracks processed emails to avoid duplicate imports
- Keeps unparsed transactions for manual review

## Requirements

Before running this project, make sure you have:

- Python 3.10 or newer installed
- A Gmail account
- Gmail IMAP enabled
- A Google app password
- The `openpyxl` Python package installed

Install the required Python package with:

```powershell
pip install openpyxl
```

# Gmail Setup
This project uses Gmail IMAP to read bank transaction alert emails.
# Step 1: Enable IMAP in Gmail
- Open Gmail.
- Click the settings icon.
- Select See all settings.
- Go to the Forwarding and POP/IMAP tab.
- Enable IMAP access.
- Save the changes.
# Step 2: Create a Google App Password
- If your Google account uses 2-Step Verification, you need to create an app password.
- Go to your Google Account settings.
- Open Security.
- Enable 2-Step Verification if it is not already enabled.
- Go to App passwords.
- Create a new app password for Mail.
- Copy the generated password.
- Use this app password for the script. Do not use your normal Gmail password.
# Setup
- Before running the script, set your Gmail email and app password as environment variables.
For PowerShell:
- $env:JARVIS_EMAIL="your_email@gmail.com"
- $env:JARVIS_EMAIL_PASSWORD="your_google_app_password"
You can also set your DBS/POSB account ending. By default, the script uses account ending 5209.
- $env:JARVIS_DBS_ACCOUNT_ENDING="5209"
# How To Run
- Open PowerShell and go to the project folder.
Example:
cd "C:\path\to\BankEmailTransactions"
Then run:
- python bank_transactions.py
You can also run the project by double-clicking:
- run_transactions.bat
When the script runs, it will connect to Gmail, search for DBS/POSB transaction alert emails, process new transactions, and generate Excel reports.
Example terminal output:
Processed 5 transaction(s), 0 unparsed email(s). Income SGD 1000.00, expenses SGD 125.50, net SGD 874.50.
Updated workbook(s):
- financial_reports/financial_report_2026.xlsx
# Output
The generated financial reports are saved inside:
- financial_reports/
Example report files:
financial_report_2024.xlsx
financial_report_2025.xlsx
financial_report_2026.xlsx
Each report may include:
- Dashboard
- Index
- Yearly summary
- Category summary
- All transactions
- Monthly transaction sheets
- Monthly summary sheets
- The transaction sheets may include:
- UID
- Transaction Date
- Email Date
- Type
- Category
- Party
- Description
- Amount
- Income
- Expense
- Net
- Source Account
- Destination
- Subject
- How It Works
# The script follows this process:
It connects to Gmail using IMAP.
It searches for DBS/POSB bank transaction alert emails.
It checks which emails are new and have not been processed before.
It reads the email subject and body.
It extracts transaction details from the email content.
It categorizes each transaction using keyword rules.
It separates income and expenses.
It generates yearly Excel reports.
It saves the latest processed email UID so duplicate transactions are avoided.
# Important Files
- bank_transactions.py
This is the main Python script. It connects to Gmail, reads bank emails, parses transaction details, categorizes the transactions, and generates Excel reports.
- run_transactions.bat
This is a Windows batch file that runs the Python script. It is useful if you want to run the project by double-clicking instead of typing the command manually.
- last_uid.txt
This file stores the latest processed Gmail email UID. It helps the script know which emails were already processed so it does not import duplicates.
- financial_reports/
This folder stores the generated Excel reports.
