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


_TERM = re.compile(r"([+-]?)\s*(?:(\$?[A-Za-z]{1,3}\$?\d+)|(\d+(?:\.\d+)?))\s*")


def _plain_number(value, where):
    from decimal import Decimal, InvalidOperation
    if value is None or isinstance(value, bool):
        return Decimal(0)
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).strip().replace(",", "")
    if text in ("", "-", "--"):
        return Decimal(0)
    try:
        return Decimal(text)
    except InvalidOperation:
        raise BillingError(f"cell {where} holds {value!r}, which is not an amount")


def _cell_amount(ws, ref, seen):
    """Value of a cell, working out simple +/- formulas such as =F75+D76-E76 the way Excel would."""
    from decimal import Decimal
    ref = ref.replace("$", "").upper()
    if ref in seen:
        raise BillingError(f"cell {ref} refers to itself")
    v = ws[ref].value
    if isinstance(v, str) and v.startswith("="):
        body, pos, total = v[1:].replace(" ", ""), 0, Decimal(0)
        if not body:
            return Decimal(0)
        while pos < len(body):
            m = _TERM.match(body, pos)
            if not m or m.end() == pos:
                raise BillingError(f"cell {ref} has a formula BILLFLOW cannot work out: {v}")
            part = _cell_amount(ws, m.group(2), seen | {ref}) if m.group(2) else Decimal(m.group(3))
            total += -part if m.group(1) == "-" else part
            pos = m.end()
        return total
    return _plain_number(v, ref)


def balance_anchor(ws, total_row, balance_col="F"):
    """The row whose balance the statement total shows (normally the last entry above the total row)."""
    m = re.fullmatch(r"=\$?%s\$?(\d+)" % balance_col, str(ws[f"{balance_col}{total_row}"].value or "").replace(" ", ""),
                     re.IGNORECASE)
    if m and int(m.group(1)) < total_row:
        return int(m.group(1))
    for r in range(total_row - 1, 0, -1):
        v = ws[f"{balance_col}{r}"].value
        if isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=")):
            return r
    raise BillingError("no balance found above the total row")


def statement_balance(ws, total_row, balance_col="F"):
    """The statement's current balance, worked out from its own figures."""
    return _cell_amount(ws, f"{balance_col}{balance_anchor(ws, total_row, balance_col)}", frozenset())


_DEFAULT_COLS = {"date": "A", "reference": "B", "deposit_by": "C", "debit": "D", "credit": "E", "balance": "F"}
_HEADINGS = {"date": "date", "reference": "reference", "deposit_by": "depositby", "debit": "debit", "credit": "credit",
             "balance": "balance"}


def _heading_cols(ws, before_row):
    """Columns read from the statement's own heading row (Date | Reference | Deposit by | Debit | Credit | Balance),
    so a statement laid out differently from the others is still read correctly."""
    from openpyxl.utils import get_column_letter
    for r in range(1, before_row):
        found = {}
        for c in range(1, ws.max_column + 1):
            text = _norm(ws.cell(r, c).value)
            for key, word in _HEADINGS.items():
                if text == word and key not in found:
                    found[key] = get_column_letter(c)
        if {"date", "reference", "debit", "credit", "balance"} <= set(found):
            return found
    return {}


def _cols(params, ws=None, total_row=None):
    """Statement columns: set in the config, else read from the heading row, else the usual A to F."""
    found = _heading_cols(ws, total_row) if ws is not None and total_row else {}
    return {k: str(params.get(f"{k}_column") or found.get(k) or d).upper() for k, d in _DEFAULT_COLS.items()}


def _total_row(ws, marker):
    """The row whose text starts with the marker ('Total must equal...'), in whichever column it was typed."""
    for r in range(1, ws.max_row + 1):
        for c in range(1, min(ws.max_column, 12) + 1):
            if _norm(ws.cell(r, c).value).startswith(marker):
                return r
    return None


def _as_at_cell(ws, params):
    """Where the 'Reconciliation as at' date goes: the cell named in as_at_cell, or the date cell to the right
    of the label named in as_at_label (found on each statement, since layouts differ)."""
    from openpyxl.utils import get_column_letter
    label = _norm(params.get("as_at_label", ""))
    if label:
        for r in range(1, min(ws.max_row, 40) + 1):
            for c in range(1, ws.max_column + 1):
                if _norm(ws.cell(r, c).value).startswith(label):
                    right = [ws.cell(r, cc) for cc in range(c + 1, ws.max_column + 1)]
                    dated = next((x for x in right if isinstance(x.value, datetime)), None)
                    target = dated or next((x for x in right if x.value in (None, "")), None)
                    if target is None:
                        raise BillingError(f"no empty or date cell to the right of '{params['as_at_label']}'")
                    return f"{get_column_letter(target.column)}{r}"
        raise BillingError(f"could not find '{params['as_at_label']}' on the statement")
    return params.get("as_at_cell")


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
    marker = _norm(params.get("total_marker", "Total must equal"))
    total_row = _total_row(ws, marker)
    if total_row is None:
        raise BillingError(f"{path.name}: could not find the row starting "
                           f"'{params.get('total_marker', 'Total must equal')}' (the end of the statement)")
    cols = _cols(params, ws, total_row)
    ref = _reference(invoice, params)
    refs = {_norm(ws[f"{cols['reference']}{r}"].value) for r in range(1, total_row)}
    if _norm(ref) in refs:
        raise BillingError(f"{path.name} already has a row for {ref}")
    # payments already typed in: the same reference AND the same amount (references are sometimes reused)
    existing = set()
    for r in range(1, total_row):
        credit = ws[f"{cols['credit']}{r}"].value
        try:
            amount = _plain_number(credit, f"{cols['credit']}{r}")
        except BillingError:
            continue
        if amount:
            existing.add((_norm(ws[f"{cols['reference']}{r}"].value), amount))
    return path, wb, ws, cols, total_row, ref, existing


def _entries(invoice, params, ref, existing):
    """The rows to add, in date order (on the same day the invoice comes first), and payments already listed."""
    when = period_end(invoice.period) if params.get("date", "period_end") == "period_end" else invoice.issue_date
    rows = [(when, 0, ref, "", invoice.total, None)]
    already = []
    for p in invoice.payments:
        if (_norm(p.reference), p.amount) in existing:
            already.append(p.reference)
        else:
            rows.append((p.date, 1, p.reference, p.deposit_by, None, p.amount))
    return sorted(rows, key=lambda e: (e[0], e[1])), already


@output_check("statement_xlsx")
def check_statement(invoice, params, ctx):
    path, wb, ws, cols, total_row, ref, existing = _open(invoice, params, ctx)
    entries, already = _entries(invoice, params, ref, existing)
    try:
        before = statement_balance(ws, total_row, cols["balance"])
    except BillingError as e:
        raise BillingError(f"{path.name}: balance cannot be worked out - {e}")
    try:
        _as_at_cell(ws, params)
    except BillingError as e:
        raise BillingError(f"{path.name}: {e}")
    from decimal import Decimal
    # Excel keeps amounts as binary fractions (e.g. 23771.510000000002), so balances are rounded to cents
    invoice.balance = (before + sum((e[4] or 0) - (e[5] or 0) for e in entries)).quantize(Decimal("0.01"))
    notes = [f"balance after: {invoice.balance:,.2f}"]
    if already:
        notes.append(f"already on the statement, not added again: {', '.join(already)}")
    return notes


@output("statement_xlsx")
def statement_xlsx(invoice, params, ctx):
    path, wb, ws, cols, total_row, ref, existing = _open(invoice, params, ctx)
    entries, _ = _entries(invoice, params, ref, existing)
    backup = resolve_path(ctx.cfg, ctx.cfg.get("output_dir", "output")) / ctx.period / "statement-backups" / path.name
    backup.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, backup)

    last, n = total_row - 1, len(entries)
    anchor = balance_anchor(ws, total_row, cols["balance"])
    as_at = _as_at_cell(ws, params)
    # cells merged across columns on the row above (e.g. Debit spread over E:G) are merged the same way
    from openpyxl.utils import column_index_from_string
    written = {column_index_from_string(c) for c in cols.values()}
    row_merges = [(mr.min_col, mr.max_col) for mr in ws.merged_cells.ranges
                  if mr.min_row == mr.max_row == last and mr.max_col > mr.min_col
                  and not any(mr.min_col < c <= mr.max_col for c in written)]
    insert_rows_keep_layout(ws, total_row, n)
    if as_at:
        m = re.fullmatch(r"([A-Za-z]+)(\d+)", as_at)
        if m and int(m.group(2)) >= total_row:
            as_at = f"{m.group(1)}{int(m.group(2)) + n}"
    bal, deb, cre = cols["balance"], cols["debit"], cols["credit"]
    for i, (when, _, reference, deposit_by, debit, credit) in enumerate(entries):
        r = total_row + i
        for ci in range(1, ws.max_column + 1):
            src = ws.cell(last, ci)
            if src.has_style:
                ws.cell(r, ci)._style = copy(src._style)
        for c1, c2 in row_merges:
            ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        ws.row_dimensions[r].height = ws.row_dimensions[last].height
        ws[f"{cols['date']}{r}"] = datetime.combine(when, datetime.min.time())
        ws[f"{cols['reference']}{r}"] = reference
        if deposit_by:
            ws[f"{cols['deposit_by']}{r}"] = deposit_by
        if debit is not None:
            ws[f"{deb}{r}"] = float(debit)
        if credit is not None:
            ws[f"{cre}{r}"] = float(credit)
        ws[f"{bal}{r}"] = f"={bal}{anchor if i == 0 else r - 1}+{deb}{r}-{cre}{r}"
    total_cell = ws[f"{bal}{total_row + n}"]
    if _norm(total_cell.value) in (_norm(f"={bal}{last}"), _norm(f"={bal}{anchor}"), ""):
        total_cell.value = f"={bal}{total_row + n - 1}"
    if as_at:
        ws[as_at] = datetime.combine(period_end(invoice.period), datetime.min.time())
    wb.save(path)
    return [path]
