"""Setup wizard core: turns a business's answers into a working config.yml, and checks the linked files.

The screen (wizard_gui.py) collects the answers; this module has no UI so it can be tested and scripted.
"""
import calendar
import csv
import re
from io import StringIO
from pathlib import Path

import yaml

from .config import ConfigError, parse_period, period_vars
from .models import BillingError

SOURCE_TYPES = {".csv": "csv", ".xlsx": "xlsx", ".xlsm": "xlsx"}
KINDS = ("fixed", "metered", "extra", "percentage")


def source_type(location):
    loc = str(location).strip()
    if loc.lower().startswith(("http://", "https://")):
        return "url_csv"
    ext = Path(loc).suffix.lower()
    if ext not in SOURCE_TYPES:
        raise ConfigError(f"'{loc}' is not a supported file. Use CSV or Excel (.xlsx), or a web address that returns CSV "
                          "(an old .xls file: open it in Excel and Save As .xlsx).")
    return SOURCE_TYPES[ext]


def suggest_template(text, period):
    """Turns a name that contains the month into one that follows every month, e.g.
    'Water 2026-11.xlsx' -> 'Water {period}.xlsx', 'November' -> '{month_name}'."""
    y, m = parse_period(period)
    out = str(text).replace("{", "{{").replace("}", "}}")
    out = out.replace(period, "{period}")
    out = re.sub(calendar.month_name[m], "{month_name}", out, flags=re.I)
    return out.replace(str(y), "{year}")


def fill(template, period):
    try:
        return str(template).format(**period_vars(period))
    except (KeyError, IndexError, ValueError) as e:
        raise BillingError(f"Bad month pattern in '{template}': {e}")


def inspect_file(location, sheet=None, header_row=1):
    """Column names (and sheet names) of a data file, so the wizard can offer them as choices."""
    loc = str(location).strip()
    kind = source_type(loc)
    header_row = max(1, int(header_row or 1))
    if kind == "url_csv":
        from .connectors.sources import _fetch
        rows = list(csv.reader(StringIO(_fetch(loc, {}))))
        heads = rows[header_row - 1] if len(rows) >= header_row else []
        return {"type": kind, "sheets": [], "headers": [h.strip() for h in heads if h.strip()]}
    path = Path(loc)
    if not path.is_file():
        raise BillingError(f"File not found: {path}")
    if kind == "csv":
        with open(path, newline="", encoding="utf-8-sig") as fh:
            rows = [r for _, r in zip(range(header_row), csv.reader(fh))]
        heads = rows[header_row - 1] if len(rows) >= header_row else []
        return {"type": kind, "sheets": [], "headers": [h.strip() for h in heads if h.strip()]}
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        names = list(wb.sheetnames)
        name = sheet or names[0]
        if name not in names:
            raise BillingError(f"Sheet '{name}' not found in {path.name} (it has: {', '.join(names)})")
        rows = list(wb[name].iter_rows(min_row=header_row, max_row=header_row, values_only=True))
        heads = [str(v).strip() for v in (rows[0] if rows else []) if v is not None and str(v).strip()]
        return {"type": kind, "sheets": names, "headers": heads}
    finally:
        wb.close()


def guess_column(headers, words):
    for w in words:
        for h in headers:
            if w in h.lower():
                return h
    return ""


def _loc(file):
    return {"url": file} if source_type(file) == "url_csv" else {"path": file}


def _need(d, key, what):
    v = str(d.get(key, "")).strip()
    if not v:
        raise ConfigError(what)
    return v


def build_config(a):
    """answers dict -> config dict (the same structure as a hand-written config.yml)."""
    company = _need(a, "company", "Enter your company name.")
    cl = a.get("clients") or {}
    cfile = _need(cl, "file", "Choose the file that lists your clients.")
    clients = {"type": source_type(cfile), **_loc(cfile), "id_column": _need(cl, "id_column", "Choose the client ID column."),
               "name_column": cl.get("name_column") or cl["id_column"]}
    if clients["type"] == "xlsx":
        if cl.get("sheet"):
            clients["sheet"] = cl["sheet"]
        if int(cl.get("header_row") or 1) != 1:
            clients["header_row"] = int(cl["header_row"])
    if cl.get("last_invoice_column"):
        clients["last_invoice_column"] = cl["last_invoice_column"]
    if not a.get("charges"):
        raise ConfigError("Add at least one charge (what do you bill your clients for?).")

    sources, names, charges = {}, {}, []

    def source_for(item, optional=False):
        file = _need(item, "file", f"Choose the data file for '{item.get('name', 'charge')}'.")
        key = _need(item, "key", f"Choose the client ID column in the file for '{item.get('name', 'charge')}'.")
        sig = (file, item.get("sheet") or "", int(item.get("header_row") or 1), key, optional)
        if sig not in names:
            s = {"type": source_type(file), **_loc(file), "key": key}
            if s["type"] == "xlsx":
                if item.get("sheet"):
                    s["sheet"] = item["sheet"]
                if sig[2] != 1:
                    s["header_row"] = sig[2]
            if optional:
                s["optional"] = True
            names[sig] = f"data{len(names) + 1}"
            sources[names[sig]] = s
        return names[sig]

    for ch in a["charges"]:
        kind, name = ch.get("kind"), str(ch.get("name", "")).strip()
        if kind not in KINDS:
            raise ConfigError(f"Unknown charge type '{kind}'.")
        if not name:
            raise ConfigError("Every charge needs a name (it is printed on the invoice).")
        if kind == "fixed":
            r = {"rule": "fixed", "name": name}
            if ch.get("column"):
                r["column"] = ch["column"]
            elif str(ch.get("amount", "")).strip():
                r["amount"] = str(ch["amount"]).strip()
            else:
                raise ConfigError(f"'{name}': enter an amount or choose the column that holds it.")
            if ch.get("description"):
                r["description"] = ch["description"]
        elif kind == "metered":
            r = {"rule": "metered", "name": name, "source": source_for(ch),
                 "current_column": _need(ch, "current_column", f"'{name}': choose the current-reading column."),
                 "previous_column": _need(ch, "previous_column", f"'{name}': choose the previous-reading column.")}
            if ch.get("rate_column"):
                r["rate_column"] = ch["rate_column"]
            elif str(ch.get("rate", "")).strip():
                r["rate"] = str(ch["rate"]).strip()
            else:
                raise ConfigError(f"'{name}': enter the price per unit or choose the column that holds it.")
            if ch.get("required") is False:
                r["required"] = False
        elif kind == "extra":
            r = {"rule": "adjustment", "name": name, "source": source_for(ch, optional=not ch.get("required", False)),
                 "amount_column": _need(ch, "amount_column", f"'{name}': choose the amount column.")}
            if ch.get("description_column"):
                r["description_column"] = ch["description_column"]
            if ch.get("required"):
                r["required"] = True
        else:
            r = {"rule": "percentage", "name": name, "percent": _need(ch, "percent", f"'{name}': enter the percentage.")}
            of = [x.strip() for x in str(ch.get("of", "")).split(",") if x.strip()]
            if of:
                r["of"] = of
        charges.append(r)

    outs = a.get("outputs") or {}
    outputs = []
    if outs.get("pdf"):
        outputs.append({"type": "pdf"})
    if outs.get("docx_template"):
        outputs.append({"type": "docx", "template": outs["docx_template"]})
    if outs.get("ledger_xlsx"):
        outputs.append({"type": "ledger_xlsx"})
    if outs.get("json"):
        outputs.append({"type": "json"})
    if not outputs:
        raise ConfigError("Choose at least one thing to produce (PDF invoice, Word invoice, ledger...).")

    cfg = {"company": {"name": company, "currency": str(a.get("currency", "")).strip()}}
    if str(a.get("payment_terms", "")).strip():
        cfg["company"]["payment_terms"] = a["payment_terms"].strip()
    cfg["clients"], cfg["sources"], cfg["charges"], cfg["outputs"] = clients, sources, charges, outputs
    cfg["output_dir"] = (a.get("output_dir") or "output").strip()
    if str(a.get("needs", "")).strip():
        cfg["needs"] = a["needs"].strip()
    return cfg


def check_answers(a, period, base=None):
    """Opens every linked file for the given month and reports anything missing. Returns a list of problems (empty = fine)."""
    problems = []
    base = Path(base) if base else Path.cwd()
    try:
        parse_period(period)
    except ConfigError as e:
        return [str(e)]

    def look(label, item, wanted):
        try:
            file = fill(item["file"], period)
            sheet = fill(item["sheet"], period) if item.get("sheet") else None
            p = Path(file)
            if not file.lower().startswith("http") and not p.is_absolute():
                file = str(base / p)
            info = inspect_file(file, sheet, item.get("header_row") or 1)
        except (BillingError, ConfigError, OSError, ValueError) as e:
            problems.append(f"{label}: {e}")
            return
        for col in wanted:
            if col and col not in info["headers"]:
                problems.append(f"{label}: column '{col}' was not found in {Path(file).name} "
                                f"(columns there: {', '.join(info['headers']) or 'none'})")

    cl = a.get("clients") or {}
    if cl.get("file"):
        look("Client list", cl, [cl.get("id_column"), cl.get("name_column"), cl.get("last_invoice_column")])
    for ch in a.get("charges") or []:
        label = f"Charge '{ch.get('name', '')}'"
        if ch.get("kind") == "metered":
            look(label, ch, [ch.get("key"), ch.get("current_column"), ch.get("previous_column")])
        elif ch.get("kind") == "extra":
            look(label, ch, [ch.get("key"), ch.get("amount_column"), ch.get("description_column")])
        elif ch.get("kind") == "fixed" and ch.get("column") and cl.get("file"):
            look(label + " (client list)", cl, [ch["column"]])
    return problems


def write_config(a, folder):
    """Writes config.yml (and a starter Word template if asked) into the folder. Returns the config path."""
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    a = {**a, "outputs": dict(a.get("outputs") or {})}
    if a["outputs"].get("starter_template"):
        from .sample import build_template
        build_template(folder / "invoice_template.docx", (a.get("clients") or {}).get("email_column") or None)
        a["outputs"]["docx_template"] = "invoice_template.docx"
    cfg = build_config(a)
    header = ("# Created by the billflow setup wizard. You can edit this file; see the README for every option.\n"
              "# {period} {month_name} {year} {month} in a file name or sheet name follow the month being billed.\n")
    path = folder / "config.yml"
    path.write_text(header + yaml.safe_dump(cfg, sort_keys=False, allow_unicode=True), encoding="utf-8")
    return path
