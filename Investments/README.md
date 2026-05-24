## Folder Structure
This folder should have a sub-folder for each investment account being tracked.
The naming format of each sub-folder should be
{ACCOUNT_NAME}_{CURRENCY}\_{ACCOUNT_ID}.

Example: "TFSA_CAD_ABC123"

## Statements
Within each account folder should be 2 types of statements:
- Activity - Showing transactions within the account within some period
  - One file can span across a wide time range and show and transactions during that time
- Holdings - The latest positions of the accounts when the document was pulled
  - These holdings only show the performance and latest state of the account at the time it was pulled.
  - A new holdings record should be pulled at a regular interval to keep file updated as well as have a history of the performance over time.

**NOTE:** Script was only tested on the TD csv export format.

### Statement Naming Format
{ACCOUNT_ID}-{"activity"/"holdings"}-{DATE_STATEMENT_WAS_PULLED}.csv