"""Creates a ready-to-run demo workspace with made-up data."""
from pathlib import Path

CONFIG = """\
company:
  name: Sample Properties Ltd
  currency: KES
  payment_terms: |
    Payments are due on the 5th of the following month.
    Bank: Example Bank, Account 000-000-000

clients:
  type: csv
  path: clients.csv
  id_column: id
  name_column: name
  last_invoice_column: last_invoice

sources:
  water:
    type: csv
    path: "readings-{period}.csv"
    key: id
  power:
    type: csv
    path: "power-{period}.csv"
    key: id
    optional: true

charges:
  - rule: fixed
    name: Service Charge
    description: Security, cleaning and garbage collection
    column: service_charge
  - rule: metered
    name: Water Bill
    source: water
    current_column: current
    previous_column: previous
    rate_column: water_rate
  - rule: adjustment
    name: Power Bill
    source: power
    amount_column: amount
    description_column: description

outputs:
  - type: pdf
  - type: docx
    template: invoice_template.docx
  - type: ledger_xlsx

output_dir: output
"""

CLIENTS = """\
id,name,email,service_charge,water_rate,last_invoice
H1,Alice Example,alice@example.com,4000,180,12
H2,Bob Sample,bob@example.com,4000,180,7
A1,Carol Demo,carol@example.com,1000,30,3
"""

READINGS = """\
id,current,previous
H1,817,804
H2,1500,1480
A1,18,16
"""

POWER = """\
id,amount,description
H2,1000,Electricity bill A
H2,500.50,Electricity bill B
"""


def build_template(path, email_column="email"):
    from docx import Document

    d = Document()
    d.add_heading("{{company}}", 1)
    d.add_paragraph("INVOICE {{invoice_no}}")
    d.add_paragraph("Bill to: {{client_name}}" + (f"  ({{{{client_{email_column}}}}})" if email_column else ""))
    d.add_paragraph("Date: {{date}}   Period: {{period}}")
    t = d.add_table(rows=2, cols=5)
    t.style = "Table Grid"
    for i, h in enumerate(["Item", "Description", "Qty", "Unit price", "Amount"]):
        t.rows[0].cells[i].text = h
    for i, h in enumerate(["{{item}}", "{{description}}", "{{qty}}", "{{unit_price}}", "{{amount}}"]):
        t.rows[1].cells[i].text = h
    d.add_paragraph("TOTAL: {{total}}")
    d.add_paragraph("{{payment_terms}}")
    d.save(str(path))


def init_workspace(target):
    target = Path(target)
    target.mkdir(parents=True, exist_ok=True)
    (target / "config.yml").write_text(CONFIG, encoding="utf-8")
    (target / "clients.csv").write_text(CLIENTS, encoding="utf-8")
    (target / "readings-2026-11.csv").write_text(READINGS, encoding="utf-8")
    (target / "power-2026-11.csv").write_text(POWER, encoding="utf-8")
    build_template(target / "invoice_template.docx")
    return target
