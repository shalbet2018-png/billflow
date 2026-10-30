import csv
import urllib.request
from io import StringIO

from ..config import resolve_path
from ..models import BillingError
from ..registry import source


def _clean(row):
    return {str(k).strip(): ("" if v is None else str(v).strip()) for k, v in row.items() if k is not None}


@source("csv")
def csv_source(params, ctx):
    path = resolve_path(ctx.cfg, params["path"], **ctx.vars)
    if not path.exists():
        raise BillingError(f"File not found: {path}")
    with open(path, newline="", encoding="utf-8-sig") as fh:
        return [_clean(r) for r in csv.DictReader(fh)]


@source("url_csv")
def url_csv_source(params, ctx):
    """Any web address that returns CSV, e.g. a Google Sheet published to the web as CSV."""
    url = params["url"].format(**ctx.vars)
    with urllib.request.urlopen(url, timeout=30) as resp:
        text = resp.read().decode("utf-8-sig")
    return [_clean(r) for r in csv.DictReader(StringIO(text))]


@source("xlsx")
def xlsx_source(params, ctx):
    from openpyxl import load_workbook

    path = resolve_path(ctx.cfg, params["path"], **ctx.vars)
    if not path.exists():
        raise BillingError(f"File not found: {path}")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = params.get("sheet")
        name = sheet.format(**ctx.vars) if sheet else wb.sheetnames[0]
        if name not in wb.sheetnames:
            raise BillingError(f"Sheet '{name}' not found in {path.name} (has: {', '.join(wb.sheetnames)})")
        rows = list(wb[name].iter_rows(values_only=True))
    finally:
        wb.close()
    header_row = int(params.get("header_row", 1)) - 1
    headers = [str(h).strip() if h is not None else "" for h in rows[header_row]]
    out = []
    for r in rows[header_row + 1:]:
        if all(v is None or str(v).strip() == "" for v in r):
            continue
        out.append(_clean(dict(zip(headers, r))))
    return out


def _fetch(url, params):
    import os
    headers = {"User-Agent": "billflow"}
    if params.get("bearer_env"):
        token = os.environ.get(params["bearer_env"], "")
        if not token:
            raise BillingError(f"Environment variable {params['bearer_env']} is not set")
        headers["Authorization"] = f"Bearer {token}"
    try:
        with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=30) as resp:
            return resp.read().decode("utf-8-sig")
    except OSError as e:
        raise BillingError(f"Could not fetch data: {e}")


def _flatten(v):
    if isinstance(v, (dict, list)):
        import json
        return json.dumps(v)
    return "" if v is None else str(v).strip()


@source("json")
def json_source(params, ctx):
    """A JSON file or web API returning a list of objects (or an object holding that list at 'records_path').
    API tokens come from the environment variable named in 'bearer_env'."""
    import json
    if "url" in params:
        text = _fetch(params["url"].format(**ctx.vars), params)
    else:
        path = resolve_path(ctx.cfg, params["path"], **ctx.vars)
        if not path.exists():
            raise BillingError(f"File not found: {path}")
        text = path.read_text(encoding="utf-8-sig")
    try:
        data = json.loads(text)
    except ValueError as e:
        raise BillingError(f"Invalid JSON: {e}")
    for part in filter(None, str(params.get("records_path", "")).split(".")):
        if not isinstance(data, dict) or part not in data:
            raise BillingError(f"'{params['records_path']}' not found in the JSON data")
        data = data[part]
    if not isinstance(data, list) or not all(isinstance(r, dict) for r in data):
        raise BillingError("JSON data must be a list of objects (set 'records_path' if it is nested)")
    return [{str(k).strip(): _flatten(v) for k, v in r.items()} for r in data]


@source("google_sheet")
def google_sheet_source(params, ctx):
    """A Google Sheet shared as 'anyone with the link can view'. Give the sheet id from its address, and optionally the tab name."""
    from urllib.parse import quote
    url = f"https://docs.google.com/spreadsheets/d/{params['sheet_id']}/gviz/tq?tqx=out:csv"
    if params.get("tab"):
        url += "&sheet=" + quote(params["tab"].format(**ctx.vars))
    return [_clean(r) for r in csv.DictReader(StringIO(_fetch(url, params)))]


@source("sqlite")
def sqlite_source(params, ctx):
    """Rows from a SQLite database (opened read-only). The query may use :period, :year, :month."""
    import sqlite3
    path = resolve_path(ctx.cfg, params["path"], **ctx.vars)
    if not path.exists():
        raise BillingError(f"Database not found: {path}")
    q = params["query"]
    if not q.lstrip().lower().startswith("select"):
        raise BillingError("Only SELECT queries are allowed")
    con = sqlite3.connect(f"file:{path.as_posix()}?mode=ro", uri=True)
    try:
        con.row_factory = sqlite3.Row
        cur = con.execute(q, {k: str(v) for k, v in ctx.vars.items()})
        return [_clean(dict(r)) for r in cur.fetchall()]
    except sqlite3.Error as e:
        raise BillingError(f"Database query failed: {e}")
    finally:
        con.close()
