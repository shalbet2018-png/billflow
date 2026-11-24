"""The month at a glance: one workbook per run with the headline figures, a row per client and a plain-words
list of everything that needs attention. Also the short totals printed after a preview or a run."""
from datetime import datetime
from decimal import Decimal

from .friendly import explain, month_label


def _paid(inv):
    return sum((p.amount for p in inv.payments), Decimal(0))


def figures(results):
    ok = [r for r in results if r.status == "OK" and r.invoice]
    owed = [r.invoice.balance for r in ok if r.invoice.balance is not None]
    earlier = [r.invoice.balance - r.invoice.total for r in ok if r.invoice.balance is not None]
    return {
        "invoices": len(ok),
        "billed": sum((r.invoice.total for r in ok), Decimal(0)),
        "payments": sum(len(r.invoice.payments) for r in ok),
        "paid": sum((_paid(r.invoice) for r in ok), Decimal(0)),
        "owed": sum(owed, Decimal(0)) if owed else None,
        "owing_clients": sum(1 for x in earlier if x > 0),
        "notices": sum(1 for r in ok if r.invoice.notice_due),
        "attention": sum(1 for r in results if r.status == "ERROR"),
        "skipped": sum(1 for r in results if r.status == "SKIPPED"),
    }


def totals_line(results, currency=""):
    f = figures(results)
    cur = f"{currency} " if currency else ""
    parts = [f"Billed {cur}{f['billed']:,.2f} on {f['invoices']} invoice(s)"]
    if f["payments"]:
        parts.append(f"{f['payments']} payment(s) {cur}{f['paid']:,.2f}")
    if f["owed"] is not None:
        parts.append(f"owed after this month {cur}{f['owed']:,.2f}")
    if f["notices"]:
        parts.append(f"{f['notices']} arrears notice(s)")
    return " | ".join(parts)


def write_workbook(path, cfg, period, results):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    company = cfg["company"]
    cur = company.get("currency", "")
    f = figures(results)
    bold, head_fill = Font(bold=True), PatternFill("solid", fgColor="DDE4EE")
    thin = Border(bottom=Side(style="thin", color="999999"))
    money = '#,##0.00;[Red]-#,##0.00'

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = f"{company['name']} - {month_label(period)} billing"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Made by BILLFLOW on {datetime.now():%d %B %Y at %H:%M}"
    ws["A2"].font = Font(italic=True, color="666666")
    headline = [("Invoices made", f["invoices"], None), (f"Total billed ({cur})", f["billed"], money),
                ("Payments recorded", f["payments"], None), (f"Payments received ({cur})", f["paid"], money),
                (f"Owed after this month ({cur})", f["owed"] if f["owed"] is not None else "-", money),
                ("Clients owing from earlier months", f["owing_clients"], None),
                ("Arrears notices written", f["notices"], None),
                ("Need attention (not invoiced)", f["attention"], None),
                ("Already invoiced earlier (skipped)", f["skipped"], None)]
    for i, (label, value, fmt) in enumerate(headline, start=4):
        ws.cell(i, 1, label)
        c = ws.cell(i, 3, float(value) if isinstance(value, Decimal) else value)
        c.font = bold
        if fmt and not isinstance(value, str):
            c.number_format = fmt
    if f["attention"]:
        ws.cell(4 + len(headline), 1, "See the 'Needs attention' sheet for what to do.").font = Font(color="C00000", bold=True)

    top = 6 + len(headline)
    cols = ["Unit", "Name", "Status", "Invoice", f"Billed ({cur})", f"Paid this month ({cur})",
            f"Balance after ({cur})", f"Owed from earlier ({cur})", "Notice", "Notes"]
    for ci, name in enumerate(cols, start=1):
        c = ws.cell(top, ci, name)
        c.font, c.fill, c.border = bold, head_fill, thin
        c.alignment = Alignment(wrap_text=True, vertical="center")
    status_words = {"OK": "Invoiced", "ERROR": "Needs attention", "SKIPPED": "Already done"}
    order = {"ERROR": 0, "OK": 1, "SKIPPED": 2}
    r = top
    for res in sorted(results, key=lambda x: order.get(x.status, 3)):
        r += 1
        inv = res.invoice if res.status == "OK" else None
        bal = inv.balance if inv is not None else None
        row = [res.client_id, res.client_name, status_words.get(res.status, res.status), inv.number if inv else "",
               float(inv.total) if inv else None, float(_paid(inv)) if inv and inv.payments else None,
               float(bal) if bal is not None else None,
               float(bal - inv.total) if bal is not None and bal - inv.total > 0 else None,
               "Yes" if inv and inv.notice_due else "", res.message]
        for ci, v in enumerate(row, start=1):
            c = ws.cell(r, ci, v)
            if 5 <= ci <= 8:
                c.number_format = money
        if res.status == "ERROR":
            ws.cell(r, 3).font = Font(color="C00000", bold=True)
    r += 1
    ws.cell(r, 1, "Total").font = bold
    for ci, val in ((5, f["billed"]), (6, f["paid"]), (7, f["owed"])):
        if val is not None:
            c = ws.cell(r, ci, float(val))
            c.font, c.number_format = bold, money
    ws.freeze_panes = ws.cell(top + 1, 1)
    for ci, width in enumerate([10, 28, 16, 12, 14, 16, 16, 16, 9, 60], start=1):
        ws.column_dimensions[get_column_letter(ci)].width = width

    att = wb.create_sheet("Needs attention")
    for ci, name in enumerate(["Unit", "Name", "What happened", "What to do", "File to open"], start=1):
        c = att.cell(1, ci, name)
        c.font, c.fill, c.border = bold, head_fill, thin
    problems = [x for x in results if x.status == "ERROR"]
    for i, res in enumerate(problems, start=2):
        e = explain(res.message, cfg, period, res.client_id)
        for ci, v in enumerate([res.client_id, res.client_name, e.text, e.fix, e.file], start=1):
            att.cell(i, ci, v).alignment = Alignment(wrap_text=True, vertical="top")
    if not problems:
        att["A2"] = "Nothing - every client was invoiced or had been invoiced already."
    for ci, width in enumerate([10, 26, 50, 60, 40], start=1):
        att.column_dimensions[get_column_letter(ci)].width = width

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
