## Folder Structure
This folder should have a sub-folder for each account (Credit/Debit) card being tracked.
The naming format of each sub-folder should be {ACCOUNT_NAME}_{ACCOUNT_NUMBER_ENDING}.

Example: "TD_Rewards_Visa_1234"

## Statements
Within each account folders should be the account statement downloaded from bank website.
Script was only tested on the TD csv export format.

The expected columns in each statement are (in order):
- Date of transaction
- Transaction Description
- Debit (Money going out)
- Credit (Money coming in)
- Balance of account

There is no limit to the range of each statement and the `generate_finance.py` script should filter out duplicate transactions based on:
- Date
- Description
- Debit/Credit
- Balance
This allows for overlapping statements or if one is pulled mid-way through statement period and added again.

### Statement Naming format
Each `csv` file should have the naming structure of {START_YEAR_MONTH}_{END_YEAR_MONTH}.csv.
This will sort the files in order.
The dates should be in the YYYY_MM format
