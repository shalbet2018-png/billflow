# billflow

Configurable monthly invoicing and ledger automation. You describe your clients, where the monthly numbers come from,
how each charge is calculated and what to produce - billflow does the rest, and **asks for confirmation before it writes anything**.

No Microsoft Office needed. Reads CSV, Excel and any web address that returns CSV (e.g. a published Google Sheet);
produces PDF invoices, filled Word templates and per-client Excel/CSV ledgers.

## Quick start

```
pip install -e .
billflow init demo
billflow run --config demo/config.yml --period 2026-11 --dry-run   # preview, writes nothing
billflow run --config demo/config.yml --period 2026-11            # shows the preview, asks, then writes
```

## Safety built in

- **Preview and confirm** - nothing is written until you answer `y` (or pass `--yes` for scheduled runs).
- **Never guesses** - missing readings, readings that go down, non-numeric values and duplicate clients are reported per client and skipped.
- **No double billing** - a client already invoiced for a period is skipped; invoice numbers only advance for invoices actually written.
- **Your source files are only read**, never modified.

## Config (config.yml)

| Section | Purpose |
|---|---|
| `company` | name, currency, payment terms |
| `clients` | the client list (`csv`/`xlsx`/`url_csv`), id and name columns, `last_invoice_column` for numbering |
| `sources` | monthly data (meter readings, extra bills). Paths and sheet names can use `{period}`, `{year}`, `{month}`, `{month_name}` |
| `charges` | ordered list of rules: `fixed`, `metered`, `adjustment` (or your own) |
| `outputs` | `pdf`, `docx` (template with `{{placeholders}}`), `ledger_xlsx`, `ledger_csv` |
| `plugins` | your own Python files that add rules, sources or outputs |

Built-in rules: `fixed` (flat amount or client column), `metered` ((current - previous) x rate, rate from a client column or a number),
`adjustment` (extra bills; several rows for one client are summed).

Word templates use `{{company}} {{client_name}} {{client_<column>}} {{invoice_no}} {{date}} {{period}} {{total}} {{payment_terms}}`
and one table row containing `{{item}} {{description}} {{qty}} {{unit_price}} {{amount}}`, which is repeated per line.

## Email

Add an `email` output **after** the pdf/docx outputs so it can attach their files:

```yaml
outputs:
  - type: pdf
  - type: email
    mode: smtp                 # default is "draft"
    host: smtp.office365.com
    port: 587
    from: billing@yourcompany.com
    to_column: email           # client column holding the address
    subject: "Invoice {invoice_no} - {period}"
    body: "Dear {client_name}, please find attached invoice {invoice_no}. Total: {total}."
    attach: [".pdf"]
```

- By default every email is only **saved as an `.eml` draft** you can open and check.
- Emails are sent only when the config says `mode: smtp` **and** you run with `--send-emails`; you are asked to confirm again
  (`--yes` skips the question for scheduled runs).
- The password is never stored in the config: set the `BILLFLOW_SMTP_PASSWORD` environment variable.
- A client with a missing/invalid address or a failed send is reported and skipped; the others continue.
## Desktop screen

Start menu / desktop shortcut **BILLFLOW** (or `billflow gui`) opens a cyberpunk-styled window: choose a config
(or press **DEMO** for sample data), type the month, **PREVIEW**, then **EXECUTE** and confirm. Nothing is written until you
confirm, and the screen never sends emails or pushes to online systems - use the command line with `--send-emails` / `--push` for that.

## Accounting systems

| Output | What it does | Environment variables |
|---|---|---|
| `xero` | Creates an Accounts Receivable invoice (DRAFT by default). Options: `account_code`, `contact_id_column`, `status`, `due_days` | `XERO_ACCESS_TOKEN`, `XERO_TENANT_ID` |
| `quickbooks` | Creates an invoice. Needs `customer_id_column` (your client list's QuickBooks id) and `item_id`; `base_url` for the sandbox | `QUICKBOOKS_ACCESS_TOKEN`, `QUICKBOOKS_REALM_ID` |
| `webhook` | POSTs the invoice JSON to any URL (Zapier, Make, n8n, your own app); optional `bearer_env` | your choice |

Each writes its request as a JSON file first. It is **sent only when the output has `mode: api` and you run with `--push`**
(with a confirmation question). Put these outputs last so a later failure cannot re-send an invoice. Tokens are OAuth access
tokens you obtain from the provider; billflow does not store them. These connectors are tested with simulated servers -
try them against a Xero demo company or QuickBooks sandbox first.
## Licensing

billflow runs in **trial mode** (up to 5 clients per run; previews are always unlimited) until a licence key is installed:

```
billflow license install <key>      # key text, or a file containing it
billflow license status
```

Keys are signed offline (Ed25519) and name the customer, expiry date and client limit; they are checked on the customer's
machine with no internet. Vendors create keys with `python tools/make_license.py "Customer" --expires 2027-10-31 --max-clients 50`
using the private signing key, which is **never stored in this repository** (back it up safely - without it you cannot issue keys).
Set `BILLFLOW_LICENSE` to supply a key through the environment instead.

## More connectors

| Type | Name | Use |
|---|---|---|
| source | `json` | JSON file or web API (`url`, `bearer_env` for a token, `records_path` for nested lists) |
| source | `google_sheet` | Sheet shared by link: `sheet_id`, optional `tab` |
| source | `sqlite` | Read-only `SELECT` from a database; may use `:period :year :month` |
| rule | `percentage` | % of earlier charges, e.g. VAT (`percent`, optional `of: [names]`) |
| rule | `tiered` | usage in bands: `tiers: [{up_to: 10, rate: 100}, {rate: 200}]` |
| output | `json` | one JSON file per invoice, for importing elsewhere |
## Customising (plugins)

```python
# late_fee.py  - listed under `plugins:` in config.yml
from decimal import Decimal
from billflow.registry import rule
from billflow.models import Line

@rule("late_fee")
def late_fee(client, params, ctx):
    amt = Decimal(params["amount"])
    return [Line("Late fee", Decimal(1), amt, amt, "Late fee")]
```

Then use `- rule: late_fee` under `charges`. Sources (`@source`) and outputs (`@output`) work the same way - this is how
new data sources (other websites, accounting systems, email) are added without touching the core.

## Tests

```
pip install -e ".[dev]"
pytest
```



