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


# ---------------------------------------------------------------- month summary

def test_run_writes_a_month_summary_with_figures_clients_and_problems(estate, capsys):
    _with_notices(estate)
    cfg = (estate / "config.yml").read_text(encoding="utf-8").replace(
        "charges:", "  payments: {type: csv, path: \"payments-{period}.csv\", key: id}\npayments: {source: payments}\ncharges:", 1)
    (estate / "config.yml").write_text(cfg, encoding="utf-8")
    (estate / "payments-2026-11.csv").write_text(
        "id,date,reference,amount\nU1,2026-11-10,PMT#5,1000\nU9,2026-11-12,PMT#6,250\n", encoding="utf-8")
    assert _run(estate, "--yes") == 3                                       # finished, but something needs attention
    out = capsys.readouterr().out
    assert "Billed KES 13,880.00 on 2 invoice(s) | 1 payment(s) KES 1,000.00 | owed after this month KES 14,880.00 " \
           "| 1 arrears notice(s)" in out
    wb = load_workbook(estate / "output" / "2026-11" / "summary.xlsx")
    s = wb["Summary"]
    head = {s.cell(r, 1).value: s.cell(r, 3).value for r in range(4, 13)}
    assert head["Invoices made"] == 2 and head["Total billed (KES)"] == 13880 and head["Payments received (KES)"] == 1000
    assert head["Owed after this month (KES)"] == 14880 and head["Arrears notices written"] == 1
    assert head["Need attention (not invoiced)"] == 1
    rows = [[c.value for c in r] for r in s.iter_rows(min_row=15) if r[0].value]
    assert rows[0][:3] == ["Unit", "Name", "Status"]
    assert rows[1][:3] == ["U9", "(not in client list)", "Needs attention"]
    assert rows[2][:9] == ["U1", "Ann Demo", "Invoiced", "U1/2", 8800, 1000, 8800, None, None]
    assert rows[3][:9] == ["U2", "Ben Demo", "Invoiced", "U2/8", 5080, None, 6080, 1000, "Yes"]
    assert rows[4][0] == "Total" and rows[4][4:7] == [13880, 1000, 14880]
    att = wb["Needs attention"]
    assert att["A2"].value == "U9" and "not in your client list" in att["C2"].value and "unit number" in att["D2"].value
    csv_rows = (estate / "output" / "2026-11" / "summary.csv").read_text(encoding="utf-8").splitlines()
    assert csv_rows[0].endswith("Paid,Balance after,Notice") and csv_rows[2].endswith(",6080.00,yes")


def test_preview_prints_totals_and_writes_nothing(estate, capsys):
    assert _run(estate, "--dry-run") == 0
    assert "Billed KES 13,880.00 on 2 invoice(s) | owed after this month KES 15,880.00" in capsys.readouterr().out
    assert not (estate / "output").exists()


def test_fractions_of_a_cent_in_old_statement_figures_do_not_count_as_arrears(estate):
    _with_notices(estate, limit=0)
    for cid, name in (("U1", "Ann Demo"), ("U2", "Ben Demo")):
        path = estate / "units" / f"{name}_{cid}" / f"Block - {cid} - {name}.xlsx"
        wb = load_workbook(path)
        wb.active["E12"] = 1499.998                       # pays the 1,500 invoice, leaving 0.002 on the books
        wb.save(path)
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].invoice.balance == Decimal("8800.00") and not res["U1"].invoice.notice_due
    assert "notice" not in res["U1"].message


# ---------------------------------------------------------------- lessons from the real Moke statements

def _moved_layout(path):
    """A statement laid out like one real one: Deposit by in D, Debit over E:G, Credit in H, Balance over I:J,
    the 'as at' date in J4 and 'Total must equal' typed in column B."""
    wb = Workbook()
    ws = wb.active
    ws["C4"], ws["F4"], ws["G4"], ws["J4"] = "Phone:", "-", "Reconciliation as at:", datetime(2026, 10, 31)
    ws["G6"], ws["J6"] = "BALANCE:", "=I13"
    for col, v in zip("ABDEHI", ["Date", "Reference", "Deposit by", "Debit", "Credit", "Balance"]):
        ws[f"{col}10"] = v
    ws["A11"], ws["B11"], ws["E11"], ws["I11"] = datetime(2026, 9, 30), "Invoice U1/1", 1500, 1500
    ws["A12"], ws["B12"], ws["D12"], ws["H12"], ws["I12"] = datetime(2026, 10, 5), "PMT#1", "Ann", 500, "=I11+E12-H12"
    for r in (11, 12):
        ws.merge_cells(f"E{r}:G{r}")
        ws.merge_cells(f"I{r}:J{r}")
    ws["B13"], ws["H13"], ws["I13"] = "Total must equal the General Ledger Balance above:", "Total:", "=I12"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def test_statement_with_its_own_column_layout_is_read_from_its_heading_row(estate):
    cfg = (estate / "config.yml").read_text(encoding="utf-8").replace("as_at_cell: F4", "as_at_label: Reconciliation as at")
    (estate / "config.yml").write_text(cfg, encoding="utf-8")
    p = estate / "units" / "Ann Demo_U1" / "Block - U1 - Ann Demo.xlsx"
    _moved_layout(p)
    _with_payments(estate, "U1,2026-11-03,PMT#2,300,Ann\n")
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].status == "OK" and res["U1"].invoice.balance == Decimal(1000 - 300 + 8800)
    assert res["U2"].status == "OK"                     # the usual A-F layout still works next to it
    assert _run(estate, "--yes") == 0
    ws = load_workbook(p).active
    assert [ws[f"{c}13"].value for c in "ABDEHI"] == [datetime(2026, 11, 3), "PMT#2", "Ann", None, 300, "=I12+E13-H13"]
    assert [ws[f"{c}14"].value for c in "ABEHI"] == [datetime(2026, 11, 30), "Invoice U1/2", 8800, None, "=I13+E14-H14"]
    assert str(ws["B15"].value).startswith("Total must equal") and ws["I15"].value == "=I14" and ws["J6"].value == "=I15"
    assert ws["J4"].value == datetime(2026, 11, 30) and ws["F4"].value == "-"      # phone cell untouched
    assert {"E13:G13", "I13:J13", "E14:G14", "I14:J14"} <= {str(m) for m in ws.merged_cells.ranges}


def test_total_row_typed_in_another_column_is_found(estate):
    p = estate / "units" / "Ann Demo_U1" / "Block - U1 - Ann Demo.xlsx"
    wb = load_workbook(p)
    ws = wb.active
    ws.unmerge_cells("A13:D13")
    ws["B13"], ws["A13"] = ws["A13"].value, None
    wb.save(p)
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].status == "OK" and res["U1"].invoice.balance == Decimal(9800)


def test_reused_payment_reference_with_a_different_amount_is_a_new_payment(estate):
    # PMT#1 (500) is already on the statement; the same reference was used again for a new payment of 700
    _with_payments(estate, "U1,2026-11-03,PMT#1,700\n")
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].status == "OK" and "not added again" not in res["U1"].message
    assert res["U1"].invoice.balance == Decimal(1000 - 700 + 8800)
    # and in this month's file the same reference may appear twice when the amounts differ
    (estate / "payments-2026-11.csv").write_text("id,date,reference,amount,deposit_by\n"
                                                 "U1,2026-11-03,PMT#5,700,\nU1,2026-11-09,PMT#5,800,\n", encoding="utf-8")
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].status == "OK" and res["U1"].invoice.balance == Decimal(1000 - 1500 + 8800)


def test_template_value_with_no_column_stops_before_anything_is_written(estate):
    doc = Document(str(estate / "template.docx"))
    doc.add_paragraph("Phone: {{client_phone}}  Email: {{client_email}}")
    doc.save(str(estate / "template.docx"))
    res = plan(load_config(estate / "config.yml"), "2026-11")
    assert {r.status for r in res} == {"ERROR"}
    assert res[0].message == ("template.docx uses {{client_email}}, {{client_phone}}, which BILLFLOW cannot fill "
                              "- add a column called 'email', 'phone' to the client list")
    from billflow.friendly import explain
    assert "Nothing was written" in explain(res[0].message).text
    clients = (estate / "clients.csv").read_text(encoding="utf-8").splitlines()
    rows = [clients[0] + ",phone,email"] + [line + ",-,-" for line in clients[1:]]
    (estate / "clients.csv").write_text("\n".join(rows) + "\n", encoding="utf-8")
    assert {r.status for r in plan(load_config(estate / "config.yml"), "2026-11")} == {"OK"}


def test_bought_unit_pays_the_service_charge_only(estate):
    cfg = (estate / "config.yml").read_text(encoding="utf-8")
    cfg = cfg.replace("rate_column: water_rate}", "rate_column: water_rate, skip_if: service_only}")
    cfg = cfg.replace("rate_column: power_rate, required: false}", "rate_column: power_rate, required: false, skip_if: service_only}")
    (estate / "config.yml").write_text(cfg, encoding="utf-8")
    clients = (estate / "clients.csv").read_text(encoding="utf-8").splitlines()
    (estate / "clients.csv").write_text("\n".join([clients[0] + ",service_only", clients[1] + ",yes", clients[2] + ","]) + "\n",
                                        encoding="utf-8")
    # the bought unit has no readings at all
    for f in ("readings-2026-10.csv", "readings-2026-11.csv"):
        lines = (estate / f).read_text(encoding="utf-8").splitlines()
        (estate / f).write_text("\n".join(l for l in lines if not l.startswith("U1,")) + "\n", encoding="utf-8")
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert [l.item for l in res["U1"].invoice.lines] == ["Service Charge"] and res["U1"].invoice.total == 4000
    assert [l.item for l in res["U2"].invoice.lines] == ["Service Charge", "Water Bill"]
    (estate / "clients.csv").write_text("\n".join([clients[0] + ",service_only", clients[1] + ",maybe", clients[2] + ","]) + "\n",
                                        encoding="utf-8")
    res = {r.client_id: r for r in plan(load_config(estate / "config.yml"), "2026-11")}
    assert res["U1"].status == "ERROR" and "'service_only' must be yes or empty" in res["U1"].message
