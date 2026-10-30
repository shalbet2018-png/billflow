import csv
import shutil
from decimal import Decimal

import pytest
from docx import Document
from openpyxl import load_workbook

from billflow import cli
from billflow.config import ConfigError, load_config
from billflow.engine import commit, plan
from billflow.sample import init_workspace


@pytest.fixture
def ws(tmp_path):
    init_workspace(tmp_path)
    return tmp_path


def run(ws, *extra):
    return cli.main(["run", "--config", str(ws / "config.yml"), "--period", "2026-11", *extra])


def test_plan_calculates_amounts(ws):
    res = {r.client_id: r for r in plan(load_config(ws / "config.yml"), "2026-11")}
    assert all(r.status == "OK" for r in res.values())
    assert res["H1"].invoice.total == Decimal("6340.00")          # 4000 + 13 x 180
    assert res["H1"].invoice.number == "H1/13"                    # last invoice 12 + 1
    assert res["A1"].invoice.total == Decimal("1060.00")          # 1000 + 2 x 30
    assert res["H2"].invoice.total == Decimal("4000") + Decimal("3600") + Decimal("1500.50")  # power rows added together


def test_dry_run_writes_nothing(ws):
    assert run(ws, "--dry-run") == 0
    assert not (ws / "output").exists() and not (ws / "billflow-state.json").exists()


def test_full_run_creates_all_outputs_and_ledger(ws):
    assert run(ws, "--yes") == 0
    out = ws / "output"
    assert (out / "2026-11" / "H1" / "invoice-H1.pdf").stat().st_size > 1000
    doc = Document(str(out / "2026-11" / "H1" / "invoice-H1.docx"))
    text = "\n".join(p.text for p in doc.paragraphs) + "\n".join(c.text for t in doc.tables for r in t.rows for c in r.cells)
    assert "{{" not in text and "H1/13" in text and "KES 6,340.00" in text
    assert "Water Bill" in text and "Meter Reading (Last Reading: 804 / Current Reading: 817)" in text
    ws_l = load_workbook(out / "ledgers" / "H1.xlsx").active
    assert [c.value for c in ws_l[2]][1:3] == ["Invoice H1/13", 6340.0]
    rows = list(csv.DictReader(open(out / "2026-11" / "summary.csv")))
    assert {r["Status"] for r in rows} == {"OK"}


def test_second_run_does_not_double_bill_and_numbers_advance(ws):
    run(ws, "--yes")
    res = plan(load_config(ws / "config.yml"), "2026-11")
    assert {r.status for r in res} == {"SKIPPED"}
    shutil.copy(ws / "readings-2026-11.csv", ws / "readings-2026-12.csv")
    res = plan(load_config(ws / "config.yml"), "2026-12")
    assert {r.invoice.number for r in res if r.client_id == "H1"} == {"H1/14"}


def test_bad_data_is_reported_not_guessed(ws):
    (ws / "readings-2026-11.csv").write_text("id,current,previous\nH1,700,804\nH2,abc,1480\n", encoding="utf-8")
    res = {r.client_id: r for r in plan(load_config(ws / "config.yml"), "2026-11")}
    assert res["H1"].status == "ERROR" and "went down" in res["H1"].message
    assert res["H2"].status == "ERROR" and "not a number" in res["H2"].message
    assert res["A1"].status == "ERROR" and "No 'water' reading" in res["A1"].message


def test_error_clients_do_not_burn_invoice_numbers(ws):
    (ws / "readings-2026-11.csv").write_text("id,current,previous\nH1,700,804\nH2,1500,1480\nA1,18,16\n", encoding="utf-8")
    cfg = load_config(ws / "config.yml")
    commit(cfg, "2026-11", plan(cfg, "2026-11"))
    (ws / "readings-2026-11.csv").write_text("id,current,previous\nH1,817,804\nH2,1500,1480\nA1,18,16\n", encoding="utf-8")
    res = {r.client_id: r for r in plan(cfg, "2026-11")}
    assert res["H1"].status == "OK" and res["H1"].invoice.number == "H1/13"
    assert res["H2"].status == "SKIPPED"


def test_confirmation_declined_writes_nothing(ws, monkeypatch):
    monkeypatch.setattr("builtins.input", lambda *_: "n")
    assert run(ws) == 1
    assert not (ws / "output").exists()


def test_xlsx_source_and_sheet_name_template(ws, tmp_path):
    from openpyxl import Workbook
    wb = Workbook()
    s = wb.active
    s.title = "November-2026"
    s.append(["id", "current", "previous"])
    s.append(["H1", 900, 800]); s.append(["H2", 1500, 1480]); s.append(["A1", 18, 16])
    wb.save(ws / "log.xlsx")
    cfg = (ws / "config.yml").read_text()
    cfg = cfg.replace('type: csv\n    path: "readings-{period}.csv"', 'type: xlsx\n    path: log.xlsx\n    sheet: "{month_name}-{year}"')
    (ws / "config.yml").write_text(cfg)
    res = {r.client_id: r for r in plan(load_config(ws / "config.yml"), "2026-11")}
    assert res["H1"].invoice.total == Decimal("4000") + Decimal("18000")


def test_url_source_reads_web_csv(ws):
    uri = (ws / "readings-2026-11.csv").resolve().as_uri()
    cfg = (ws / "config.yml").read_text().replace('type: csv\n    path: "readings-{period}.csv"', f'type: url_csv\n    url: "{uri}"')
    (ws / "config.yml").write_text(cfg)
    res = plan(load_config(ws / "config.yml"), "2026-11")
    assert all(r.status == "OK" for r in res)


def test_company_plugin_adds_custom_rule(ws):
    (ws / "late_fee.py").write_text(
        "from decimal import Decimal\nfrom billflow.registry import rule\nfrom billflow.models import Line\n"
        "@rule('late_fee')\ndef late_fee(client, params, ctx):\n"
        "    return [Line('Late fee', Decimal(1), Decimal(params['amount']), Decimal(params['amount']), 'Late fee')]\n")
    cfg = (ws / "config.yml").read_text().replace("outputs:", "  - rule: late_fee\n    name: Late fee\n    amount: 250\n\noutputs:", 1)
    cfg = "plugins: [late_fee.py]\n" + cfg
    (ws / "config.yml").write_text(cfg)
    res = {r.client_id: r for r in plan(load_config(ws / "config.yml"), "2026-11")}
    assert res["H1"].invoice.total == Decimal("6590.00")


def test_config_errors_are_clear(ws):
    (ws / "bad.yml").write_text("company: {name: X}\n")
    with pytest.raises(ConfigError, match="missing required section"):
        load_config(ws / "bad.yml")
    assert cli.main(["run", "--config", str(ws / "bad.yml"), "--period", "2026-11"]) == 2
    assert cli.main(["run", "--config", str(ws / "config.yml"), "--period", "2026-13"]) == 2
