## Projects

### 01. Internal Purchasing AR AP Reconcilation & AP Clearing Vendor
- find cleared AR invoice in SAP
- Find AP information in AR invoice
- Group AP invoice by subledger
- reconcile AP payment with bank statement
- genearate clean AP list for vendor clearing in SAP

### 02. Mapping File Reconciliation & ERP Balance Reconciliation
- Reconile bank account vs bank Gl from different ERP system
- Reconcile bank account vs bank GL between ERP and Treasury Record
- Generate updated conciliated bank mapping file after reconcilation
- Add new account into conciliated mapping file 
- Record any deleted bank account information.

### 03. Daily EFT Transactions Reconciliation and Posting
- Daily EFT trasnaction reconcilation between GL and Bank
- For any open items, find keyword in bank statment's description
- By regex, to find coding information through keyword
- Generate JE template to upload in SAP
- Generate requiry email for all open items missing coding info.

### 04. Monthly ZBA Transaction Posting
- In 50 bank accounts, find all relevant ZBA trasnaction in monthly bank statement
- Find and delete all duplicate lines in bank statement
- Generate pivot table by transfer pair - receiving bank account and sending bank account
- Find coding for each transfer pair
- Generate JE template to upload in SAP.

### 05. Email Attachment File Download and Combination
- Search outlook mail folder, to allocate certain email whose subject has specific keyword, for muitiple months
- Download attached csv files
- Combine all downloaded csv files, and select wanted columns

### 04. ACH Statement Transformation
- big ACH report with cheque nubmer hidden in descrption fileds with long string
- Use regex to find cheque number for each line transactions
