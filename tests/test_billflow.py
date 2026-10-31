import json
import csv
import shutil
from decimal import Decimal

import pytest
from docx import Document
from openpyxl import load_workbook

from billflow import cli
from billflow.config import ConfigError, load_config
from billflow.engine import commit, plan
from billflow.models import BillingError
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


def _with_email(ws, **params):
    cfg = (ws / "config.yml").read_text(encoding="utf-8")
    block = "  - type: email\n" + "".join(f"    {k}: {v}\n" for k, v in params.items())
    (ws / "config.yml").write_text(cfg.replace("\noutput_dir:", "\n" + block + "\noutput_dir:", 1)
                                   .replace("  - type: ledger_xlsx\n", "  - type: ledger_xlsx\n", 1), encoding="utf-8")


def test_email_saves_draft_with_pdf_and_never_sends_by_default(ws, monkeypatch):
    import smtplib
    monkeypatch.setattr(smtplib, "SMTP", lambda *a, **k: pytest.fail("must not send"))
    _with_email(ws, mode="smtp", host="smtp.example.com", **{"from": "billing@acme.test"})
    assert run(ws, "--yes") == 0
    import email
    from email import policy
    msg = email.message_from_bytes((ws / "output/2026-11/H1/email-H1.eml").read_bytes(), policy=policy.default)
    assert "H1/13" in msg["Subject"]
    assert [a.get_filename() for a in msg.iter_attachments()] == ["invoice-H1.pdf"]


def test_email_sends_only_with_flag_and_mode_smtp(ws, monkeypatch):
    import smtplib
    sent = []

    class FakeSMTP:
        def __init__(self, *a, **k): pass
        def __enter__(self): return self
        def __exit__(self, *a): pass
        def starttls(self): pass
        def login(self, u, p): pass
        def send_message(self, m): sent.append(m["To"])

    monkeypatch.setattr(smtplib, "SMTP", FakeSMTP)
    _with_email(ws, mode="smtp", host="smtp.example.com", **{"from": "billing@acme.test"})
    assert run(ws, "--yes", "--send-emails") == 0
    assert len(sent) == 3


def test_email_missing_address_is_reported_not_guessed(ws):
    (ws / "clients.csv").write_text((ws / "clients.csv").read_text().replace("@", "_at_", 1), encoding="utf-8")
    _with_email(ws)
    run(ws, "--yes")
    rows = {r["Client"]: r for r in csv.DictReader(open(ws / "output/2026-11/summary.csv"))}
    assert sum(r["Status"] == "ERROR" for r in rows.values()) == 1


# ---- licensing -------------------------------------------------------------------------------------------------

@pytest.fixture
def keys(monkeypatch, tmp_path):
    from cryptography.hazmat.primitives import serialization as s
    from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
    from billflow import licensing
    k = Ed25519PrivateKey.generate()
    monkeypatch.setattr(licensing, "PUBLIC_KEY_HEX", k.public_key().public_bytes(s.Encoding.Raw, s.PublicFormat.Raw).hex())
    monkeypatch.setenv("APPDATA", str(tmp_path / "appdata"))
    monkeypatch.delenv("BILLFLOW_LICENSE", raising=False)
    pem = k.private_bytes(s.Encoding.PEM, s.PrivateFormat.PKCS8, s.NoEncryption())
    return lambda customer="Acme", expires="2099-01-01", max_clients=0: licensing.sign(pem, customer, "standard", expires, max_clients)


def _many_clients(ws, n):
    rows = ["id,name,email,service_charge,water_rate,last_invoice"] + [f"C{i},Client {i},c{i}@x.test,1000,10,0" for i in range(n)]
    (ws / "clients.csv").write_text("\n".join(rows), encoding="utf-8")
    (ws / "readings-2026-11.csv").write_text("id,current,previous\n" + "\n".join(f"C{i},10,5" for i in range(n)), encoding="utf-8")


def test_trial_blocks_runs_over_the_client_limit_but_allows_preview(ws, keys, capsys):
    _many_clients(ws, 6)
    assert run(ws, "--dry-run") == 0
    assert run(ws, "--yes") == 4
    assert not (ws / "output").exists()
    assert "Trial mode" in capsys.readouterr().out


def test_valid_key_lifts_limit_and_tampered_or_expired_keys_do_not(ws, keys, monkeypatch):
    from billflow import licensing
    _many_clients(ws, 6)
    k = keys(max_clients=10)
    assert licensing.install(k).customer == "Acme"
    assert run(ws, "--yes") == 0
    with pytest.raises(licensing.LicenseError):
        licensing.verify(k[:-4] + "AAAA")
    with pytest.raises(licensing.LicenseError):
        licensing.verify(keys(expires="2020-01-01"))
    assert licensing.current().valid


def test_key_limit_is_enforced(ws, keys, monkeypatch):
    monkeypatch.setenv("BILLFLOW_LICENSE", keys(max_clients=2))
    assert run(ws, "--yes") == 4


def test_license_cli_install_and_status(keys, capsys):
    assert cli.main(["license", "install", keys("Beta Ltd")]) == 0
    assert cli.main(["license", "status"]) == 0
    assert "Beta Ltd" in capsys.readouterr().out
    assert cli.main(["license", "install", "garbage"]) == 2


# ---- more connectors -------------------------------------------------------------------------------------------

def _edit(ws, fn):
    import yaml
    cfg = yaml.safe_load((ws / "config.yml").read_text(encoding="utf-8"))
    fn(cfg)
    (ws / "config.yml").write_text(yaml.safe_dump(cfg, sort_keys=False), encoding="utf-8")


def _totals(ws):
    return {r.client_id: r.invoice.total for r in plan(load_config(ws / "config.yml"), "2026-11")}


def test_percentage_rule_adds_vat_on_earlier_charges(ws):
    _edit(ws, lambda c: c["charges"].append({"rule": "percentage", "name": "VAT", "percent": 16}))
    assert _totals(ws)["H1"] == Decimal("7354.40")        # 6340 x 1.16


def test_tiered_rule_prices_usage_in_bands(ws):
    def f(c):
        c["charges"][1] = {"rule": "tiered", "name": "Water", "source": "water", "current_column": "current",
                           "previous_column": "previous", "tiers": [{"up_to": 10, "rate": 100}, {"rate": 200}]}
    _edit(ws, f)
    assert _totals(ws)["H1"] == Decimal("4000") + 1000 + 600     # 13 units: 10 x 100 + 3 x 200


def test_json_source_file_and_nested_records(ws):
    import json
    rows = list(csv.DictReader(open(ws / "readings-2026-11.csv")))
    (ws / "r.json").write_text(json.dumps({"data": {"readings": rows}}), encoding="utf-8")
    _edit(ws, lambda c: c["sources"].__setitem__("water", {"type": "json", "path": "r.json", "records_path": "data.readings", "key": "id"}))
    assert _totals(ws)["H1"] == Decimal("6340.00")


def test_sqlite_source_is_read_only_select(ws):
    import sqlite3
    con = sqlite3.connect(ws / "m.db")
    con.execute("create table r(id text, current int, previous int)")
    con.executemany("insert into r values (?,?,?)", [("H1", 817, 804), ("H2", 1500, 1480), ("A1", 18, 16)])
    con.commit(); con.close()
    _edit(ws, lambda c: c["sources"].__setitem__("water", {"type": "sqlite", "path": "m.db", "key": "id",
                                                           "query": "select id, current, previous from r"}))
    assert _totals(ws)["A1"] == Decimal("1060.00")
    _edit(ws, lambda c: c["sources"]["water"].__setitem__("query", "delete from r"))
    res = plan(load_config(ws / "config.yml"), "2026-11")
    assert all(r.status == "ERROR" and "SELECT" in r.message for r in res)
    assert sqlite3.connect(ws / "m.db").execute("select count(*) from r").fetchone()[0] == 3


def test_google_sheet_source_builds_export_url(ws, monkeypatch):
    from billflow.connectors import sources
    seen = []
    monkeypatch.setattr(sources, "_fetch", lambda url, p: seen.append(url) or (ws / "readings-2026-11.csv").read_text())
    _edit(ws, lambda c: c["sources"].__setitem__("water", {"type": "google_sheet", "sheet_id": "abc123", "tab": "{month_name} {year}", "key": "id"}))
    assert _totals(ws)["H1"] == Decimal("6340.00")
    assert "abc123" in seen[0] and "sheet=November%202026" in seen[0]


def test_json_output_file(ws):
    import json
    _edit(ws, lambda c: c["outputs"].append({"type": "json"}))
    assert run(ws, "--yes") == 0
    d = json.loads((ws / "output/2026-11/H1/invoice-H1.json").read_text())
    assert d["invoice_no"] == "H1/13" and d["total"] == "6340.00"




# ---- accounting connectors -------------------------------------------------------------------------------------

def _accounting(ws, **params):
    _edit(ws, lambda c: c["outputs"].append(params))


def test_xero_and_quickbooks_save_payload_and_never_post_without_push(ws, monkeypatch):
    from billflow.connectors import accounting
    monkeypatch.setattr(accounting, "_post_json", lambda *a: pytest.fail("must not post"))
    (ws / "clients.csv").write_text((ws / "clients.csv").read_text().replace("last_invoice", "last_invoice,qb_id")
                                    .replace(",12\n", ",12,101\n").replace(",7\n", ",7,102\n").replace(",3\n", ",3,103\n"), encoding="utf-8")
    _accounting(ws, type="xero", mode="api", account_code="200")
    _accounting(ws, type="quickbooks", mode="api", customer_id_column="qb_id", item_id="9")
    assert run(ws, "--yes") == 0
    x = json.loads((ws / "output/2026-11/H1/xero-H1.json").read_text())["Invoices"][0]
    assert x["InvoiceNumber"] == "H1/13" and x["Status"] == "DRAFT" and x["LineItems"][0]["AccountCode"] == "200"
    q = json.loads((ws / "output/2026-11/H1/quickbooks-H1.json").read_text())
    assert q["CustomerRef"] == {"value": "101"} and q["DocNumber"] == "H1/13"


def test_push_posts_with_tokens_from_environment(ws, monkeypatch):
    from billflow.connectors import accounting
    calls = []
    monkeypatch.setattr(accounting, "_post_json", lambda url, h, p: calls.append((url, h)) or {})
    monkeypatch.setenv("XERO_ACCESS_TOKEN", "tok"); monkeypatch.setenv("XERO_TENANT_ID", "ten")
    _accounting(ws, type="xero", mode="api")
    assert run(ws, "--yes", "--push") == 0
    assert len(calls) == 3 and calls[0][0].endswith("/Invoices")
    assert calls[0][1] == {"Authorization": "Bearer tok", "xero-tenant-id": "ten"}


def test_push_without_token_is_reported_per_client(ws, monkeypatch):
    from billflow.connectors import accounting
    monkeypatch.setattr(accounting, "_post_json", lambda *a: {})
    monkeypatch.delenv("XERO_ACCESS_TOKEN", raising=False)
    _accounting(ws, type="xero", mode="api")
    run(ws, "--yes", "--push")
    rows = list(csv.DictReader(open(ws / "output/2026-11/summary.csv")))
    assert all(r["Status"] == "ERROR" and "XERO_ACCESS_TOKEN" in r["Notes"] for r in rows)


def test_webhook_posts_invoice_json(ws, monkeypatch):
    from billflow.connectors import accounting
    calls = []
    monkeypatch.setattr(accounting, "_post_json", lambda url, h, p: calls.append((url, p)) or {})
    _accounting(ws, type="webhook", mode="api", url="https://hooks.example/abc")
    assert run(ws, "--yes", "--push") == 0
    assert calls[0][0] == "https://hooks.example/abc" and calls[0][1]["invoice_no"] == "H1/13"


def test_quickbooks_missing_customer_id_is_reported(ws):
    _accounting(ws, type="quickbooks", item_id="9", customer_id_column="nope")
    run(ws, "--yes")
    assert all(r["Status"] == "ERROR" for r in csv.DictReader(open(ws / "output/2026-11/summary.csv")))


# ---- desktop screen --------------------------------------------------------------------------------------------

def _pump(root, seconds=3.0, until=None):
    import time
    end = time.time() + seconds
    while time.time() < end and not (until and until()):
        root.update()
        time.sleep(0.02)


def test_gui_preview_execute_and_belt_moves(ws, monkeypatch, keys):
    import tkinter as tk
    from billflow import gui
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    monkeypatch.setattr(gui, "save_settings", lambda d: None)
    monkeypatch.setattr(gui, "load_settings", lambda: {})
    monkeypatch.setattr(gui.messagebox, "askyesno", lambda *a, **k: True)
    monkeypatch.setattr(gui.messagebox, "showinfo", lambda *a, **k: None)
    opened = []
    monkeypatch.setattr(gui, "open_path", opened.append)
    try:
        app = gui.App(root)
        lane = app.belt.lanes[0]
        x_before = app.c.coords(lane["ids"][0])[0]
        app.cfg_var.set(str(ws / "config.yml")); app.period_var.set("2026-11")
        assert str(app.b_run["state"]) == "disabled"                      # cannot create before checking
        app.preview()
        _pump(root, 10, lambda: app.last is not None)
        assert app.last and app.stat["ready"].get() == "3"
        assert "16,500.50" in app.stat["total"].get()
        assert str(app.b_run["state"]) == "normal" and "look good" in app.hint.get()
        app.execute()
        _pump(root, 10, lambda: (ws / "output/2026-11/summary.csv").exists() and not app.busy and opened)
        assert (ws / "output/2026-11/H1/invoice-H1.pdf").exists()
        assert opened and str(opened[0]).endswith("2026-11")                # offered to open the month's folder
        assert app.c.coords(lane["ids"][0])[0] != x_before          # the belt kept moving
    finally:
        root.destroy()


def test_gui_execute_requires_preview(ws, monkeypatch, keys):
    import tkinter as tk
    from billflow import gui
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    shown = []
    monkeypatch.setattr(gui, "load_settings", lambda: {})
    monkeypatch.setattr(gui.messagebox, "showinfo", lambda *a, **k: shown.append(a))
    try:
        app = gui.App(root)
        app.cfg_var.set(str(ws / "config.yml"))
        app.execute()
        assert shown and "Check invoices" in shown[0][1] and not (ws / "output").exists()
    finally:
        root.destroy()


def test_gui_explains_problems_in_plain_words_and_changing_month_needs_a_new_check(ws, monkeypatch, keys):
    import tkinter as tk
    from billflow import gui
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    monkeypatch.setattr(gui, "save_settings", lambda d: None)
    monkeypatch.setattr(gui, "load_settings", lambda: {"theme": "plain"})
    (ws / "readings-2026-11.csv").write_text("id,current,previous\nH1,817,830\nH2,1500,1450\n", encoding="utf-8")
    try:
        app = gui.App(root)
        assert app.belt is not None and gui.THEME == "tron" and app.grid is not None   # an old "plain" choice opens as Tron
        app.cfg_var.set(str(ws / "config.yml")); app.period_var.set("2026-11")
        app.preview()
        _pump(root, 10, lambda: app.last is not None)
        text = app.log.get("1.0", "end")
        assert "meter reading went down" in text and "No Water Bill reading for A1 in readings-2026-11.csv" in text
        assert "What to do" in text and "Open readings-2026-11.csv" in text
        assert app.stat["errors"].get() == "2" and "need attention" in app.hint.get()
        app.period_var.set("2026-12")
        assert str(app.b_run["state"]) == "disabled"                      # a different month must be checked first
        app.preview()
        _pump(root, 10, lambda: not app.busy and "Something needs fixing" in app.log.get("1.0", "end"))
        text = app.log.get("1.0", "end")
        assert "Could not find the file" in text and "December 2026" in text
    finally:
        root.destroy()
        gui.apply_theme("neon")


def test_gui_licence_box_installs_key_without_command_line(monkeypatch, keys):
    import tkinter as tk
    from billflow import gui, licensing
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    monkeypatch.setattr(gui, "save_settings", lambda d: None)
    monkeypatch.setattr(gui, "load_settings", lambda: {})
    try:
        app = gui.App(root)
        assert "Trial" in app.licence_text
        d = app.enter_licence()
        d.box.insert("1.0", "not-a-key")
        assert d.save(quiet=True) is False and "not accepted" in d.err["text"]
        d.box.delete("1.0", "end")
        d.box.insert("1.0", keys("Moke Gardens", max_clients=50))
        assert d.save(quiet=True) is True
        assert licensing.current().valid and "Moke Gardens" in app.licence_text
    finally:
        root.destroy()


def test_gui_change_setup_reopens_the_wizard_with_saved_answers(ws, monkeypatch):
    import tkinter as tk
    from billflow import gui
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    monkeypatch.setattr(gui, "save_settings", lambda d: None)
    monkeypatch.setattr(gui, "load_settings", lambda: {})
    try:
        app = gui.App(root)
        app.cfg_var.set(str(ws / "config.yml"))                           # the hand-written demo config
        w = app.change_setup()
        assert w.editing and w.company.get() and len(w.charges) == 3
        assert {c["kind"] for c in w.charges} == {"fixed", "metered", "extra"}
        w.win.destroy()
    finally:
        root.destroy()


def test_app_icon_ships_with_the_package_and_loads():
    import tkinter as tk
    from importlib.resources import files
    from billflow import gui
    ico = files("billflow") / "assets" / "billflow.ico"
    assert ico.is_file() and (files("billflow") / "assets" / "billflow.png").is_file()
    from PIL import Image
    sizes = Image.open(str(ico)).info.get("sizes", set())
    assert {(16, 16), (32, 32), (48, 48), (256, 256)} <= set(sizes)        # Windows needs the small and large ones
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    try:
        gui.set_window_icon(root)
        assert root._billflow_icon.width() == 256
    finally:
        root.destroy()
