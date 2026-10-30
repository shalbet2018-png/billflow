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
3. Press Enter licence key (top right), paste the key we sent you (or press Open key file...), then Save key. The top right of the screen shows your licence.
4. Want to look around first? Press Try with sample data, then Check invoices. It uses made-up clients.

## 3. Tell BILLFLOW about your business (the Setup Wizard)

Press New setup on the main screen. It asks six things, in plain language.

1. Your business: company name, currency, the payment details printed on invoices, and the month of the files you are about to pick.
2. What you bill: tick the kinds of charges you have. There is also a box for notes to yourself; it is only kept as a reminder and does not change how invoices are worked out. Something unusual can be added as a custom rule - ask us.
3. Your client list: choose your client file. BILLFLOW reads its column names and shows them in drop-downs. Pick the columns holding the client ID, name and email. Only pick a "last invoice number" column if your list keeps each client's most recent invoice number (if H1's last invoice was 12, the next is H1/13); otherwise leave it as (none) and numbering starts at 1.
4. Charges: add one line per charge, in the order they should appear on the invoice.
   - Fixed amount: the same for everyone, or taken from a column in your client list (rent, service charge).
   - Usage from meter readings: pick the readings file, the client-ID column, the current and previous reading columns, and a price per unit (or a rate column).
   - Extra bills: pick the file and the amount column. Several rows for one client are added together.
   - Percentage: for example VAT, applied to every charge above it, or only to the charges you tick.
   Amounts, prices and percentages must be plain numbers (1500 or 1500.50). To fix a charge later, select it and press Change selected (or double-click it).
5. What to produce: PDF invoices, a Word invoice (a ready-made template, or your own), an Excel ledger per client, and a data file. Choose the folder for the results.
6. Test and save: choose a month and press Run test to see every client's invoice. Nothing is written during the test. Then press Save setup.

To change a setup later (a renamed column, a new charge, a different file), choose it in step 1 and press Change my setup. The wizard opens with your answers filled in; change what you need and save.

## 4. How to link the folders and files you need

When you browse to a file, the whole path is stored, so the folder is part of the link. You can keep files anywhere: your computer, a network drive, or a synced OneDrive or Google Drive folder.

If your file or folder names contain the month, BILLFLOW turns that into a pattern. For example, if you pick `C:\Accounts\Water\Readings 2026-11.xlsx` while the month is 2026-11, it stores `Readings {period}.xlsx` and next month it will look for `Readings 2026-12.xlsx` by itself. Month names (November) and years (2026) work the same way, in folder names, file names and Excel sheet names.

So each month, you only need to save the new files with the month in the name, the same way you did before.

## 5. Your monthly routine

The main screen is numbered in the order you use it, and the bar at the bottom always says what to do next.

1. Save this month's files (readings, extra bills) into your folders.
2. Open BILLFLOW. In step 1, choose your setup (it remembers the last one). In step 2, choose the month from the list.
3. Press Check invoices. You see every client, their total, and anything that needs attention, each with a "What to do" line and a link to open the file. Nothing is saved yet.
4. Fix anything that needs attention in your file and press Check invoices again.
5. Press Create invoices (it only works after a check). BILLFLOW asks you to confirm, then writes the invoices, ledgers and a summary into a folder named for the month, and offers to open it.
6. Check a few invoices, then send them.

Two looks: press Tron look (top right) for black, glowing-blue grid lines and orange accents; Neon look switches back to the original. BILLFLOW remembers your choice.

## 6. Safety

- Nothing is written until you press EXECUTE and confirm.
- Original files are only read.
- A client already invoiced for the month is skipped, so running twice does not duplicate invoices.
- Emails are saved as drafts unless you deliberately switch on sending. Passwords are never stored in the setup file.
- Connections to accounting systems (Xero, QuickBooks) only send when you turn on live mode.
- Check the first month's invoices against what you would have done by hand.

## 7. If something goes wrong

- Every problem is explained on screen with a "What to do" line. Most are fixed in your own file, then Check invoices again.
- "Could not find the file": BILLFLOW looked for this month's file and it was not there. The message shows the exact name and folder it expects.
- A column was renamed in your file, or a file moved: choose the setup and press Change my setup to pick it again.
- "Create invoices" is greyed out: press Check invoices first (and again after changing the month).
- Anything else: send us the message shown and your config.yml (without client data).
