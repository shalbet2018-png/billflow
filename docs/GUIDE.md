# BILLFLOW - Business Guide

BILLFLOW prepares your monthly invoices and client ledgers from the files you already keep (Excel or CSV). It does the arithmetic, numbers the invoices, makes the documents, and shows you everything before anything is saved. Your original files are never changed.

## 1. What you need

- A Windows computer.
- Your client list: one row per client, with at least an ID and a name.
- The files that hold each month's figures, for example meter readings or extra bills.
- Your BILLFLOW licence key. Without one it runs in trial mode, limited to 5 clients per run; previews are unlimited.

## 2. Install

1. Run `billflow-setup-0.1.0.exe` and follow the steps.
2. Open BILLFLOW from the Start menu. You will see the main screen.
3. To enter your licence key, run `billflow license install YOUR-KEY`. The main screen shows the licence status.

## 3. Tell BILLFLOW about your business (the Setup Wizard)

Press SETUP WIZARD on the main screen. It asks six things, in plain language.

1. Your business: company name, currency, the payment details printed on invoices, and the month your example files are for.
2. What you need: describe in your own words what you bill and what should happen each month, then tick the kinds of charges you have. Your words are saved with the setup as a record. The ticked boxes and the questions that follow are what actually build your setup. Something unusual can be added as a custom rule - ask us.
3. Client list: choose your client file. BILLFLOW reads its column names and shows them in drop-downs. Pick the columns holding the client ID, name, email, and last invoice number.
4. Charges: add one line per charge, in the order they should appear on the invoice.
   - Fixed amount: the same for everyone, or taken from a column in your client list (rent, service charge).
   - Usage from meter readings: pick the readings file, the client-ID column, the current and previous reading columns, and a price per unit (or a rate column).
   - Extra bills: pick the file and the amount column. Several rows for one client are added together.
   - Percentage: for example VAT, applied to all charges above it or to the ones you name.
5. What to produce: PDF invoices, a Word invoice (your own template, or a starter we make for you), an Excel ledger per client, and a data file. Choose the folder for the results.
6. Check and save: press RUN TEST to see every client's invoice for a month. Nothing is written during the test. Then save.

## 4. How to link the folders and files you need

When you browse to a file, the whole path is stored, so the folder is part of the link. You can keep files anywhere: your computer, a network drive, or a synced OneDrive or Google Drive folder.

If your file or folder names contain the month, BILLFLOW turns that into a pattern. For example, if you pick `C:\Accounts\Water\Readings 2026-11.xlsx` while the month is 2026-11, it stores `Readings {period}.xlsx` and next month it will look for `Readings 2026-12.xlsx` by itself. Month names (November) and years (2026) work the same way, in folder names, file names and Excel sheet names.

So each month, you only need to save the new files with the month in the name, the same way you did before.

## 5. Your monthly routine

1. Save this month's files (readings, extra bills) into your folders.
2. Open BILLFLOW. Choose your setup (the wizard fills this in) and type the month, for example 2026-11.
3. Press PREVIEW. You see every client, their total, and anything flagged. Nothing is saved yet.
4. Fix anything flagged in your source file and preview again.
5. Press EXECUTE. BILLFLOW asks you to confirm, then writes the invoices, ledgers and a summary into a folder named for the month.
6. Check a few invoices, then send them.

## 6. Safety

- Nothing is written until you press EXECUTE and confirm.
- Original files are only read.
- A client already invoiced for the month is skipped, so running twice does not duplicate invoices.
- Emails are saved as drafts unless you deliberately switch on sending. Passwords are never stored in the setup file.
- Connections to accounting systems (Xero, QuickBooks) only send when you turn on live mode.
- Check the first month's invoices against what you would have done by hand.

## 7. If something goes wrong

- "File not found": the month pattern could not find a file for that month. Check the name and folder, including the month.
- "Column not found": a column was renamed in your file. Run the wizard again, or edit config.yml.
- A client is flagged: BILLFLOW lists the reason (for example a missing reading). Fix the source file and preview again.
- Anything else: send us the message shown and your config.yml (without client data).
