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


# ---------------------------------------------------------------- payments

def _with_payments(estate, rows):
    cfg = (estate / "config.yml").read_text(encoding="utf-8").replace(
        "charges:", "  payments: {type: csv, path: \"payments-{period}.csv\", key: id}\n"
        "payments: {source: payments}\ncharges:", 1)
    (estate / "config.yml").write_text(cfg, encoding="utf-8")
    (estate / "payments-2026-11.csv").write_text("id,date,reference,amount,deposit_by\n" + rows, encoding="utf-8")


def test_payments_go_into_the_statement_in_date_order_with_the_invoice(estate):
    _with_payments(estate, "U1,20/11/2026,PMT#900 & P050/01,3000,Ann\nU1,2026-11-03,PMT#880 & P041/02,\"1,500\",\n")
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].message == "2 payment(s) received: 4,500.00; balance after: 5,300.00"
    assert res["U2"].message == "balance after: 6,080.00"
    assert _run(estate, "--yes") == 0
    ws = load_workbook(estate / "units" / "Ann Demo_U1" / "Block - U1 - Ann Demo.xlsx").active
    got = [[ws.cell(r, c).value for c in range(1, 7)] for r in (13, 14, 15)]
    assert got[0] == [datetime(2026, 11, 3), "PMT#880 & P041/02", None, None, 1500, "=F12+D13-E13"]
    assert got[1] == [datetime(2026, 11, 20), "PMT#900 & P050/01", "Ann", None, 3000, "=F13+D14-E14"]
    assert got[2] == [datetime(2026, 11, 30), "Invoice U1/2", None, 8800, None, "=F14+D15-E15"]
    assert str(ws["A16"].value).startswith("Total must equal") and ws["F16"].value == "=F15" and ws["F6"].value == "=F16"
    assert ws["F18"].value == "=F4" and {str(m) for m in ws.merged_cells.ranges} == {"A16:D16", "A19:F19"}


def test_payment_already_typed_into_the_statement_is_not_added_again(estate):
    _with_payments(estate, "U1,2026-11-03,pmt#1,500\nU1,2026-11-04,PMT#2,700\n")
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].status == "OK" and "not added again: pmt#1" in res["U1"].message
    _run(estate, "--yes")
    ws = load_workbook(estate / "units" / "Ann Demo_U1" / "Block - U1 - Ann Demo.xlsx").active
    refs = [ws.cell(r, 2).value for r in range(11, 16)]
    assert refs == ["Invoice U1/1", "PMT#1", "PMT#2", "Invoice U1/2", None]


@pytest.mark.parametrize("row, problem", [
    ("U1,2026-10-31,PMT#1,500", "outside November 2026"),
    ("U1,2026-11-31,PMT#1,500", "not a real date"),
    ("U1,next week,PMT#1,500", "not understood"),
    ("U1,2026-11-05,,500", "has no reference"),
    ("U1,2026-11-05,PMT#1,0", "must be more than 0"),
    ("U1,2026-11-05,PMT#1,five", "not a number"),
    ("U1,2026-11-05,PMT#1,500\nU1,2026-11-06,PMT #1,500", "listed twice"),
])
def test_bad_payments_stop_that_client_and_nothing_is_guessed(estate, row, problem):
    _with_payments(estate, row + "\n")
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].status == "ERROR" and problem in res["U1"].message
    assert res["U2"].status == "OK"


def test_payment_for_a_unit_not_in_the_client_list_is_reported(estate):
    _with_payments(estate, "U9,2026-11-05,PMT#7,2500\nU9,2026-11-06,PMT#8,500\n")
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U9"].status == "ERROR" and "3,000.00" in res["U9"].message and "not in the client list" in res["U9"].message
    assert res["U1"].status == res["U2"].status == "OK"


def test_missing_payments_file_is_reported(estate):
    _with_payments(estate, "")
    (estate / "payments-2026-11.csv").unlink()
    res = plan(load_config(estate / "config.yml"), "2026-11")
    assert {r.status for r in res} == {"ERROR"} and all("payments-2026-11.csv" in r.message for r in res)


# ---------------------------------------------------------------- balances and arrears notices

def _with_notices(estate, limit=500, days=4):
    doc = Document()
    doc.add_paragraph("{{notice_date}}")
    doc.add_paragraph("TO:{{client_name_upper}}  UNIT NO. {{client_id}}")
    p = doc.add_paragraph("Dear {{first_name}}, ")
    p.add_run("RE: NOTICE TO CLEAR ARREARS OF KSH {{arrears}}.").bold = True
    doc.add_paragraph("Please clear on or before {{deadline}}.")
    doc.save(str(estate / "notice.docx"))
    cfg = (estate / "config.yml").read_text(encoding="utf-8") + (
        f"  - {{type: notice_docx, template: notice.docx, arrears_above: {limit}, pay_within_days: {days},\n"
        "     path: \"../units/{client_folder}/Notice to {client_name} {period}.docx\"}\n")
    (estate / "config.yml").write_text(cfg, encoding="utf-8")


def test_notice_is_written_when_earlier_months_are_unpaid(estate):
    _with_notices(estate)
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].message == "balance after: 9,800.00; arrears 1,000.00: notice will be written"
    assert _run(estate, "--yes") == 0
    doc = Document(str(estate / "units" / "Ann Demo_U1" / "Notice to Ann Demo 2026-11.docx"))
    text = [p.text for p in doc.paragraphs]
    assert text == ["1ST DECEMBER 2026", "TO:ANN DEMO  UNIT NO. U1",
                    "Dear Ann, RE: NOTICE TO CLEAR ARREARS OF KSH 1,000.", "Please clear on or before 5th December 2026."]
    assert doc.paragraphs[2].runs[1].bold


def test_no_notice_when_payments_cover_the_arrears_or_it_is_under_the_limit(estate):
    _with_notices(estate, limit=500)
    cfg = (estate / "config.yml").read_text(encoding="utf-8").replace(
        "charges:", "  payments: {type: csv, path: \"payments-{period}.csv\", key: id}\npayments: {source: payments}\ncharges:", 1)
    (estate / "config.yml").write_text(cfg, encoding="utf-8")
    (estate / "payments-2026-11.csv").write_text("id,date,reference,amount\nU1,2026-11-10,PMT#5,1000\n", encoding="utf-8")
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert "notice" not in res["U1"].message and "notice will be written" in res["U2"].message
    _run(estate, "--yes")
    assert not (estate / "units" / "Ann Demo_U1" / "Notice to Ann Demo 2026-11.docx").exists()
    assert (estate / "units" / "Ben Demo_U2" / "Notice to Ben Demo 2026-11.docx").exists()


def test_notice_needs_the_statement_first(estate):
    _with_notices(estate)
    cfg = (estate / "config.yml").read_text(encoding="utf-8").replace(
        '  - {type: statement_xlsx, path: "../units/{client_folder}/{client_statement}", as_at_cell: F4}\n', "")
    (estate / "config.yml").write_text(cfg, encoding="utf-8")
    res = plan(load_config(estate / "config.yml"), "2026-11")
    assert {r.status for r in res} == {"ERROR"} and "statement_xlsx output listed before" in res[0].message


def test_balance_uses_typed_figures_blank_rows_and_the_row_the_total_points_to(estate):
    from billflow.connectors.statement import statement_balance
    path = estate / "units" / "Ann Demo_U1" / "Block - U1 - Ann Demo.xlsx"
    wb = load_workbook(path)
    ws = wb.active
    ws["F11"], ws["D12"], ws["E12"], ws["F12"] = 2500, "-", "1,000", "=F11+D12-E12"   # typed balance, text amounts
    ws.insert_rows(13)                                                                # an empty row above the total
    ws["F14"] = "=F12"
    assert statement_balance(ws, 14) == Decimal("1500")
    ws["F12"] = "=SUM(D11:D12)"
    with pytest.raises(Exception, match="cannot work out"):
        statement_balance(ws, 14)


def test_new_rows_continue_from_the_real_balance_when_an_empty_row_sits_above_the_total(estate):
    path = estate / "units" / "Ann Demo_U1" / "Block - U1 - Ann Demo.xlsx"
    wb = load_workbook(path)
    wb.active.insert_rows(13)
    wb.active["F14"], wb.active["F6"] = "=F12", "=F14"
    wb.save(path)
    assert _run(estate, "--yes") == 0
    ws = load_workbook(path).active
    assert ws["B14"].value == "Invoice U1/2" and ws["F14"].value == "=F12+D14-E14" and ws["F15"].value == "=F14"
