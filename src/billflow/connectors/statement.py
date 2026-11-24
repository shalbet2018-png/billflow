"""Adds each month's invoice to the client's existing statement workbook (a running-balance ledger).

The statement keeps its own look: the new row copies the style of the row above it, everything below
(the total, signatures, payment details) moves down, and formulas, merged cells, row heights, pictures
and the print area are adjusted to match. A copy of the workbook is saved before it is changed."""
import re
import shutil
from copy import copy
from datetime import datetime

from ..config import period_end, resolve_path
from ..models import BillingError
from ..registry import output, output_check
from .outputs import _fmt_path

_REF = re.compile(r"(?<![A-Za-z_\d.])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])")


def shift_rows_in_formula(formula, at, by):
    """Moves every cell reference at row `at` or below down by `by` rows (text in quotes is left alone)."""
    def fix(part):
        return _REF.sub(lambda m: f"{m.group(1)}{m.group(2)}{m.group(3)}"
                        f"{int(m.group(4)) + by if int(m.group(4)) >= at else m.group(4)}", part)
    pieces = re.split(r'("[^"]*")', formula)
    return "".join(p if p.startswith('"') else fix(p) for p in pieces)


def insert_rows_keep_layout(ws, at, by=1):
    """Inserts `by` empty rows before row `at`, moving everything below down and keeping the sheet intact."""
    heights = {r: ws.row_dimensions[r].height for r in range(at, ws.max_row + 1)}
    merged = [mr for mr in ws.merged_cells.ranges]
    for mr in merged:
        ws.merged_cells.remove(mr)
    ws.insert_rows(at, by)
    for mr in merged:
        if mr.min_row >= at:
            mr.shift(0, by)
        elif mr.max_row >= at:
            mr.expand(down=by)
        ws.merged_cells.add(mr)
    for r in sorted(heights, reverse=True):
        ws.row_dimensions[r + by].height = heights[r]
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = shift_rows_in_formula(c.value, at, by)
    for img in getattr(ws, "_images", []):
        anchor = img.anchor
        for end in ("_from", "to"):
            marker = getattr(anchor, end, None)
            if marker is not None and marker.row >= at - 1:
                marker.row += by
    if ws.print_area:
        ws.print_area = shift_rows_in_formula(ws.print_area.split("!")[-1], at, by)


def _cols(params):
    names = {"date": "A", "reference": "B", "deposit_by": "C", "debit": "D", "credit": "E", "balance": "F"}
    return {k: str(params.get(f"{k}_column", d)).upper() for k, d in names.items()}


def _reference(invoice, params):
    return str(params.get("reference", "Invoice {invoice_no}")).format(invoice_no=invoice.number, seq=invoice.seq,
                                                                       client_id=invoice.client_id)


def _norm(text):
    return re.sub(r"\s+", "", str(text or "")).lower()


def _open(invoice, params, ctx):
    from openpyxl import load_workbook
    if not params.get("path"):
        raise BillingError("The statement output needs a 'path' (for example \"{client_statement}\")")
    path = _fmt_path(ctx, params, "path", invoice, "")
    if not path.exists():
        raise BillingError(f"Statement not found: {path}")
    try:
        wb = load_workbook(path)
    except Exception as e:  # damaged or not an Excel workbook
        raise BillingError(f"Statement could not be opened ({path.name}): {e}")
    sheet = params.get("sheet")
    if sheet and sheet not in wb.sheetnames:
        raise BillingError(f"Sheet '{sheet}' not found in {path.name}")
    ws = wb[sheet] if sheet else wb.active
    cols = _cols(params)
    marker = _norm(params.get("total_marker", "Total must equal"))
    total_row = next((r for r in range(1, ws.max_row + 1)
                      if _norm(ws[f"{cols['date']}{r}"].value).startswith(marker)), None)
    if total_row is None:
        raise BillingError(f"{path.name}: could not find the row starting "
                           f"'{params.get('total_marker', 'Total must equal')}' (the end of the statement)")
    ref = _reference(invoice, params)
    existing = {_norm(ws[f"{cols['reference']}{r}"].value) for r in range(1, total_row)}
    if _norm(ref) in existing:
        raise BillingError(f"{path.name} already has a row for {ref}")
    return path, wb, ws, cols, total_row, ref, existing


def _entries(invoice, params, ref, existing):
    """The rows to add, in date order (on the same day the invoice comes first), and payments already listed."""
    when = period_end(invoice.period) if params.get("date", "period_end") == "period_end" else invoice.issue_date
    rows = [(when, 0, ref, "", invoice.total, None)]
    already = []
    for p in invoice.payments:
        if _norm(p.reference) in existing:
            already.append(p.reference)
        else:
            rows.append((p.date, 1, p.reference, p.deposit_by, None, p.amount))
    return sorted(rows, key=lambda e: (e[0], e[1])), already


@output_check("statement_xlsx")
def check_statement(invoice, params, ctx):
    path, wb, ws, cols, total_row, ref, existing = _open(invoice, params, ctx)
    _, already = _entries(invoice, params, ref, existing)
    if already:
        return [f"already on the statement, not added again: {', '.join(already)}"]
    return []


@output("statement_xlsx")
def statement_xlsx(invoice, params, ctx):
    path, wb, ws, cols, total_row, ref, existing = _open(invoice, params, ctx)
    entries, _ = _entries(invoice, params, ref, existing)
    backup = resolve_path(ctx.cfg, ctx.cfg.get("output_dir", "output")) / ctx.period / "statement-backups" / path.name
    backup.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, backup)

    last, n = total_row - 1, len(entries)
    insert_rows_keep_layout(ws, total_row, n)
    bal, deb, cre = cols["balance"], cols["debit"], cols["credit"]
    for i, (when, _, reference, deposit_by, debit, credit) in enumerate(entries):
        r = total_row + i
        for ci in range(1, ws.max_column + 1):
            src = ws.cell(last, ci)
            if src.has_style:
                ws.cell(r, ci)._style = copy(src._style)
        ws.row_dimensions[r].height = ws.row_dimensions[last].height
        ws[f"{cols['date']}{r}"] = datetime.combine(when, datetime.min.time())
        ws[f"{cols['reference']}{r}"] = reference
        if deposit_by:
            ws[f"{cols['deposit_by']}{r}"] = deposit_by
        if debit is not None:
            ws[f"{deb}{r}"] = float(debit)
        if credit is not None:
            ws[f"{cre}{r}"] = float(credit)
        ws[f"{bal}{r}"] = f"={bal}{r - 1}+{deb}{r}-{cre}{r}"
    total_cell = ws[f"{bal}{total_row + n}"]
    if _norm(total_cell.value) == _norm(f"={bal}{last}"):
        total_cell.value = f"={bal}{total_row + n - 1}"
    if params.get("as_at_cell"):
        ws[params["as_at_cell"]] = datetime.combine(period_end(invoice.period), datetime.min.time())
    wb.save(path)
    return [path]
