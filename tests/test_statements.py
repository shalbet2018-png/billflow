"""Monthly routine of a property manager: readings carried over from last month, the company's own Word invoice
(with text boxes), and the invoice added to each client's existing Excel statement."""
import sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import parse_xml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from billflow import cli
from billflow.config import load_config
from billflow.connectors.statement import shift_rows_in_formula
from billflow.engine import plan

CONFIG = """
company: {name: Demo Estates, currency: KES}
clients: {type: csv, path: clients.csv, id_column: id, name_column: name, last_invoice_column: last_invoice}
sources:
  readings: {type: csv, path: "readings-{period}.csv", key: id}
charges:
  - {rule: fixed, name: Service Charge, description: Security and cleaning, column: service_charge}
  - {rule: metered, name: Water Bill, source: readings, current_column: water_current,
     previous_column: water_previous, previous_from_last_month: true, rate_column: water_rate}
  - {rule: metered, name: Electricity Bill, source: readings, current_column: power_current,
     previous_column: power_previous, previous_from_last_month: true, rate_column: power_rate, required: false}
outputs:
  - {type: docx, template: template.docx, path: "../units/{client_folder}/{seq}_{client_name}.docx"}
  - {type: statement_xlsx, path: "../units/{client_folder}/{client_statement}", as_at_cell: F4}
"""

TEXTBOX = """<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/>
  <wp:docPr id="9" name="box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>KES {{total_</w:t></w:r><w:r><w:t>amount}}</w:t></w:r></w:p>
  </w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"""


def _template(path):
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Bill To: ").bold = True
    p.add_run("{{client_name}}  Invoice Number: {{invoice_no}}  Date: {{date_long}}  Ending {{period_name}}")
    doc.element.body.append(parse_xml(TEXTBOX))
    t = doc.add_table(rows=3, cols=5)
    for c, v in zip(t.rows[0].cells, ["Item", "Description", "Qty", "Unit Price", "Amount"]):
        c.text = v
    for c, v in zip(t.rows[1].cells, ["{{item}}", "{{description}}", "{{qty}}", "{{unit_price}}", "{{amount}}"]):
        c.text = v
    t.rows[2].cells[3].text, t.rows[2].cells[4].text = "Total:", "{{total_amount}}"
    doc.save(str(path))


def _statement(path, last_balance=1000):
    """A small statement laid out like the real ones: header, entries, the total row and a footer below it."""
    wb = Workbook()
    ws = wb.active
    ws["A1"], ws["E4"], ws["F4"] = "STATEMENT", "Reconciliation as at:", datetime(2026, 10, 31)
    ws["E6"], ws["F6"] = "BALANCE:", "=F13"
    ws.append([])
    for r, row in enumerate([["Date", "Reference", "Deposit by", "Debit", "Credit", "Balance"]], start=10):
        for ci, v in enumerate(row, start=1):
            ws.cell(r, ci, v)
    ws["A11"], ws["B11"], ws["D11"], ws["F11"] = datetime(2026, 9, 30), "Invoice U1/1", 1500, 1500
    ws["A12"], ws["B12"], ws["E12"], ws["F12"] = datetime(2026, 10, 5), "PMT#1", 500, "=F11+D12-E12"
    for ci in range(1, 7):
        ws.cell(12, ci).font = Font(name="Calibri", italic=True)
    ws.row_dimensions[12].height = 24.95
    ws["A13"], ws["E13"], ws["F13"] = "Total must equal the General Ledger Balance above:", "Total:", "=F12"
    ws.merge_cells("A13:D13")
    ws["A15"], ws["F15"] = "Prepared By:", "=F4"
    ws.merge_cells("A16:F16")
    ws["A16"] = "Mpesa payments: Pay bill"
    ws.row_dimensions[15].height = 48.75
    ws.print_area = "A1:F16"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


@pytest.fixture
def estate(tmp_path):
    (tmp_path / "config.yml").write_text(CONFIG, encoding="utf-8")
    (tmp_path / "clients.csv").write_text(
        "id,name,service_charge,water_rate,power_rate,folder,statement,last_invoice\n"
        "U1,Ann Demo,4000,180,30,Ann Demo_U1,Block - U1 - Ann Demo.xlsx,1\n"
        "U2,Ben Demo,4000,180,30,Ben Demo_U2,Block - U2 - Ben Demo.xlsx,7\n", encoding="utf-8")
    (tmp_path / "readings-2026-10.csv").write_text(
        "id,water_previous,water_current,power_previous,power_current\nU1,1000,1010,500,600.5\nU2,20,25,,\n", encoding="utf-8")
    (tmp_path / "readings-2026-11.csv").write_text(
        "id,water_previous,water_current,power_previous,power_current\nU1,,1020,,700.5\nU2,,31,,\n", encoding="utf-8")
    _template(tmp_path / "template.docx")
    for cid, name in (("U1", "Ann Demo"), ("U2", "Ben Demo")):
        _statement(tmp_path / "units" / f"{name}_{cid}" / f"Block - {cid} - {name}.xlsx")
    return tmp_path


def _run(estate, *extra):
    return cli.main(["run", "--config", str(estate / "config.yml"), "--period", "2026-11", *extra])


def test_previous_reading_comes_from_last_month_and_missing_meter_is_skipped(estate):
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    u1, u2 = res["U1"].invoice, res["U2"].invoice
    assert [l.item for l in u1.lines] == ["Service Charge", "Water Bill", "Electricity Bill"]
    assert u1.lines[1].description == "Meter Reading (Last Reading: 1010 / Current Reading: 1020)"
    assert u1.total == Decimal("4000") + 10 * 180 + Decimal("100.0") * 30
    assert [l.item for l in u2.lines] == ["Service Charge", "Water Bill"]          # no electricity meter
    assert u2.total == Decimal("4000") + 6 * 180


def test_missing_last_month_reading_is_reported_not_guessed(estate):
    (estate / "readings-2026-10.csv").unlink()
    res = plan(load_config(estate / "config.yml"), "2026-11")
    assert {r.status for r in res} == {"ERROR"}
    assert "No previous Water Bill reading" in res[0].message and "2026-10" in res[0].message


def test_word_invoice_fills_text_boxes_and_keeps_formatting(estate):
    assert _run(estate, "--yes") == 0
    doc = Document(str(estate / "units" / "Ann Demo_U1" / "2_Ann Demo.docx"))
    first = doc.paragraphs[0]
    assert first.runs[0].bold and first.runs[0].text == "Bill To: "
    assert "Ann Demo  Invoice Number: U1/2  Date: December 1, 2026  Ending November 2026" in first.text
    xml = doc.element.xml
    assert "KES 8,800.00" in xml and "{{" not in xml                                  # the text box was filled
    rows = [[c.text for c in r.cells] for r in doc.tables[0].rows]
    assert rows[1][:2] == ["Service Charge", "Security and cleaning"] and rows[3][0] == "Electricity Bill"
    assert rows[-1][-1] == "8,800.00"


def test_invoice_is_added_to_the_existing_statement_keeping_its_layout(estate):
    path = estate / "units" / "Ann Demo_U1" / "Block - U1 - Ann Demo.xlsx"
    assert _run(estate, "--yes") == 0
    ws = load_workbook(path).active
    assert ws["A13"].value == datetime(2026, 11, 30) and ws["B13"].value == "Invoice U1/2"
    assert ws["D13"].value == 8800 and ws["F13"].value == "=F12+D13-E13"
    assert ws["F13"].font.italic and ws.row_dimensions[13].height == 24.95       # looks like the row above
    assert str(ws["A14"].value).startswith("Total must equal") and ws["F14"].value == "=F13"
    assert ws["F6"].value == "=F14"                                                # balance follows the total
    assert ws["F16"].value == "=F4" and ws["F4"].value == datetime(2026, 11, 30)  # footer moved, as-at updated
    assert {str(m) for m in ws.merged_cells.ranges} == {"A14:D14", "A17:F17"}
    assert ws.row_dimensions[16].height == 48.75 and ws.print_area.endswith("$F$17")
    assert (estate / "output" / "2026-11" / "statement-backups" / path.name).exists()


def test_statement_problems_show_up_in_the_preview_before_anything_is_written(estate):
    (estate / "units" / "Ben Demo_U2" / "Block - U2 - Ben Demo.xlsx").unlink()
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].status == "OK"
    assert res["U2"].status == "ERROR" and "Statement not found" in res["U2"].message


def test_statement_never_gets_the_same_invoice_twice(estate):
    assert _run(estate, "--yes") == 0
    (estate / "billflow-state.json").unlink()                     # BILLFLOW's own memory lost: the statement still knows
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].status == "ERROR" and "already has a row for Invoice U1/2" in res["U1"].message


def test_formula_rows_are_shifted_only_at_or_below_the_insert():
    assert shift_rows_in_formula("=F12+D13-$E$13", 13, 1) == "=F12+D14-$E$14"
    assert shift_rows_in_formula("=SUM(D11:D12)", 13, 1) == "=SUM(D11:D12)"
    assert shift_rows_in_formula('=IF(A20="B20",LOG10(C30),0)', 13, 2) == '=IF(A22="B20",LOG10(C32),0)'


def test_office_pdf_is_reported_when_office_is_not_available(estate, monkeypatch):
    cfg = (estate / "config.yml").read_text(encoding="utf-8") + "  - {type: office_pdf}\n"
    (estate / "config.yml").write_text(cfg, encoding="utf-8")
    monkeypatch.setattr(sys, "platform", "linux")
    res = plan(load_config(estate / "config.yml"), "2026-11")
    assert {r.status for r in res} == {"ERROR"} and "Microsoft Office" in res[0].message


@pytest.mark.skipif(sys.platform != "win32", reason="needs Microsoft Office on Windows")
def test_office_pdf_makes_pdfs_with_word_and_excel(estate):
    import shutil
    if not shutil.which("powershell"):
        pytest.skip("PowerShell not available")
    from billflow.connectors.office import convert_to_pdf
    from billflow.models import BillingError
    src = estate / "plain.docx"
    plain = Document()
    plain.add_paragraph("Invoice")
    plain.save(str(src))
    try:
        out = convert_to_pdf(src, estate / "plain.pdf")
    except BillingError as e:
        pytest.skip(f"Microsoft Word not available: {e}")
    assert Path(out).read_bytes()[:4] == b"%PDF"
