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
