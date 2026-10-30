from decimal import Decimal

import pytest
import yaml
from docx import Document
from openpyxl import Workbook

from billflow import cli, wizard
from billflow.config import ConfigError, load_config
from billflow.engine import plan
from billflow.models import BillingError
from billflow.sample import init_workspace


@pytest.fixture
def ws(tmp_path):
    init_workspace(tmp_path / "data")
    return tmp_path / "data"


def answers(ws, **over):
    d = str(ws)
    a = {"company": "Acme Ltd", "currency": "USD", "payment_terms": "Pay in 7 days",
         "needs": "We bill rent, water by meter and extra power bills.",
         "clients": {"file": d + r"\clients.csv", "id_column": "id", "name_column": "name", "last_invoice_column": "last_invoice",
                     "email_column": "email"},
         "charges": [
             {"kind": "fixed", "name": "Service Charge", "column": "service_charge"},
             {"kind": "metered", "name": "Water Bill", "file": d + r"\readings-{period}.csv", "key": "id",
              "current_column": "current", "previous_column": "previous", "rate_column": "water_rate"},
             {"kind": "extra", "name": "Power Bill", "file": d + r"\power-{period}.csv", "key": "id",
              "amount_column": "amount", "description_column": "description"}],
         "outputs": {"pdf": True, "starter_template": True, "ledger_xlsx": True}}
    a.update(over)
    return a


def test_suggest_template_makes_month_patterns():
    assert wizard.suggest_template(r"C:\x\Water Log 2026-11.xlsx", "2026-11") == r"C:\x\Water Log {period}.xlsx"
    assert wizard.suggest_template("Invoices November", "2026-11") == "Invoices {month_name}"
    assert wizard.suggest_template("October-2026", "2026-10") == "{month_name}-{year}"
    assert wizard.suggest_template("nothing here", "2026-11") == "nothing here"


def test_inspect_file_csv_and_xlsx(ws, tmp_path):
    assert wizard.inspect_file(ws / "clients.csv")["headers"][:3] == ["id", "name", "email"]
    wb = Workbook(); s = wb.active; s.title = "October-2026"
    s.append(["junk"]); s.append(["House", "Prev", "Curr"]); wb.create_sheet("Other"); wb.save(tmp_path / "w.xlsx")
    info = wizard.inspect_file(tmp_path / "w.xlsx", "October-2026", 2)
    assert info["headers"] == ["House", "Prev", "Curr"] and info["sheets"] == ["October-2026", "Other"]
    with pytest.raises(BillingError):
        wizard.inspect_file(tmp_path / "w.xlsx", "Nope")
    with pytest.raises(BillingError):
        wizard.inspect_file(tmp_path / "missing.csv")
    with pytest.raises(ConfigError):
        wizard.inspect_file(tmp_path / "old.xls")


def test_wizard_answers_produce_same_invoices_as_demo(ws):
    a = answers(ws)
    assert wizard.check_answers(a, "2026-11") == []
    path = wizard.write_config(a, ws / "mine")
    res = {r.client_id: r for r in plan(load_config(path), "2026-11")}
    assert all(r.status == "OK" for r in res.values())
    assert res["H1"].invoice.total == Decimal("6340.00") and res["H1"].invoice.number == "H1/13"
    assert res["H2"].invoice.total == Decimal("9100.50")                  # power rows added together
    assert "needs" in yaml.safe_load(path.read_text())                    # their description is kept with the setup


def test_generated_config_runs_end_to_end_with_starter_template(ws):
    path = wizard.write_config(answers(ws), ws / "mine")
    assert cli.main(["run", "--config", str(path), "--period", "2026-11", "--yes"]) == 0
    doc = Document(str(ws / "mine/output/2026-11/H1/invoice-H1.docx"))
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "{{" not in text and "alice@example.com" in text and "USD 6,340.00" in text
    assert (ws / "mine/output/ledgers/H1.xlsx").exists()


def test_month_pattern_follows_the_month(ws):
    (ws / "readings-2026-12.csv").write_text("id,current,previous\nH1,830,817\nH2,1510,1500\nA1,20,18\n")
    path = wizard.write_config(answers(ws), ws / "mine")
    res = {r.client_id: r for r in plan(load_config(path), "2026-12")}
    assert res["H1"].invoice.total == Decimal("4000") + 13 * 180
    assert wizard.check_answers(answers(ws), "2027-01") and "not found" in wizard.check_answers(answers(ws), "2027-01")[0].lower()


def test_check_answers_flags_wrong_columns_and_files(ws):
    a = answers(ws)
    a["charges"][1]["current_column"] = "reading_now"
    a["charges"][2]["file"] = str(ws / "nope.csv")
    a["charges"][0]["column"] = "rent"
    problems = "\n".join(wizard.check_answers(a, "2026-11"))
    assert "reading_now" in problems and "nope.csv" in problems and "rent" in problems


@pytest.mark.parametrize("change, msg", [
    ({"company": ""}, "company name"),
    ({"clients": {"file": "x.csv"}}, "ID column"),
    ({"charges": []}, "at least one charge"),
    ({"outputs": {}}, "at least one thing"),
])
def test_incomplete_answers_get_plain_messages(ws, change, msg):
    with pytest.raises(ConfigError, match=msg):
        wizard.build_config(answers(ws, **change))


def test_charge_validation_messages(ws):
    a = answers(ws)
    a["charges"] = [{"kind": "fixed", "name": "Rent"}]
    with pytest.raises(ConfigError, match="enter an amount"):
        wizard.build_config(a)
    a["charges"] = [{"kind": "metered", "name": "Water", "file": "w.csv", "key": "id", "current_column": "c", "previous_column": "p"}]
    with pytest.raises(ConfigError, match="price per unit"):
        wizard.build_config(a)


def test_cli_setup_with_answers_file(ws, tmp_path):
    import json
    f = tmp_path / "a.json"
    f.write_text(json.dumps(answers(ws)))
    assert cli.main(["setup", "--answers", str(f), "--folder", str(tmp_path / "s")]) == 0
    assert (tmp_path / "s/config.yml").exists()
    assert cli.main(["setup", "--answers", str(f)]) == 2


def test_wizard_screen_collects_answers(ws):
    tk = pytest.importorskip("tkinter")
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    from billflow.wizard_gui import Wizard
    w = Wizard(root)
    try:
        w.company.set("Acme"); w.next()
        w.next(); assert w.step == 2
        w.cfp.file.set(str(ws / "clients.csv")); w.next(); assert w.step == 3
        assert w.id_menu.value() == "id" and w.last_menu.value() == "last_invoice" and w.email_menu.value() == "email"
        assert w.next() is None and w.step == 3   # needs a charge first
        w.charges.append({"kind": "fixed", "name": "Rent", "amount": "100"}); w.next(); w.next()
        assert w.step == 5
        a = w.collect()
        assert a["clients"]["id_column"] == "id" and a["company"] == "Acme"
        assert wizard.check_answers(a, "2026-11") == []
    finally:
        w.win.destroy(); root.destroy()


def test_flat_amount_vat_and_xlsx_sheet_pattern(ws, tmp_path):
    wb = Workbook(); s = wb.active; s.title = "November-2026"
    s.append(["House", "Last", "Now"]); s.append(["H1", 800, 810]); s.append(["H2", 1000, 1010]); s.append(["A1", 5, 6])
    wb.save(tmp_path / "log.xlsx")
    a = answers(ws, charges=[
        {"kind": "fixed", "name": "Rent", "amount": "1000"},
        {"kind": "metered", "name": "Water", "file": str(tmp_path / "log.xlsx"), "sheet": "{month_name}-{year}", "key": "House",
         "current_column": "Now", "previous_column": "Last", "rate": "10"},
        {"kind": "percentage", "name": "VAT", "percent": "16"}])
    assert wizard.check_answers(a, "2026-11") == []
    cfg = wizard.write_config(a, ws / "mine")
    res = {r.client_id: r for r in plan(load_config(cfg), "2026-11")}
    assert res["H1"].invoice.total == Decimal("1276.00")                  # (1000 + 10 x 10) x 1.16
