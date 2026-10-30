import copy
import csv
from decimal import Decimal
from pathlib import Path

from docx.table import _Row

from ..config import resolve_path, safe_name
from ..models import BillingError
from ..registry import output


def _money(v):
    return f"{Decimal(v):,.2f}"


def _fmt_path(ctx, params, key, invoice, default):
    template = str(params.get(key, default)).format(client_id=safe_name(invoice.client_id),
                                                    client_name=safe_name(invoice.client_name), **ctx.vars)
    if Path(template).is_absolute():
        return Path(template)
    return resolve_path(ctx.cfg, ctx.cfg.get("output_dir", "output")) / template


@output("pdf")
def pdf_output(invoice, params, ctx):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    company = ctx.cfg["company"]
    cur = company.get("currency", "")
    path = _fmt_path(ctx, params, "path", invoice, "{period}/{client_id}/invoice-{client_id}.pdf")
    path.parent.mkdir(parents=True, exist_ok=True)
    st = getSampleStyleSheet()
    story = [Paragraph(f"<b>{company['name']}</b>", st["Title"]), Paragraph("INVOICE", st["Heading2"]),
             Paragraph(f"Bill to: <b>{invoice.client_name}</b>", st["Normal"]),
             Paragraph(f"Invoice number: {invoice.number}", st["Normal"]),
             Paragraph(f"Date: {invoice.issue_date:%d %B %Y}", st["Normal"]),
             Paragraph(f"For period: {invoice.period}", st["Normal"]), Spacer(1, 8 * mm)]
    data = [["Item", "Description", "Qty", "Unit price", "Amount"]]
    for l in invoice.lines:
        data.append([l.item, Paragraph(l.description, st["BodyText"]), f"{l.qty:g}", _money(l.unit_price), _money(l.amount)])
    data.append(["", "", "", "Total " + cur, _money(invoice.total)])
    t = Table(data, colWidths=[30 * mm, 70 * mm, 20 * mm, 30 * mm, 30 * mm])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey), ("GRID", (0, 0), (-1, -2), 0.25, colors.grey),
                           ("FONTNAME", (3, -1), (-1, -1), "Helvetica-Bold"), ("ALIGN", (2, 0), (-1, -1), "RIGHT")]))
    story += [t, Spacer(1, 8 * mm)]
    if company.get("payment_terms"):
        story.append(Paragraph(str(company["payment_terms"]).replace("\n", "<br/>"), st["Normal"]))
    SimpleDocTemplate(str(path), pagesize=A4).build(story)
    return [path]


def _replace_in_paragraph(par, mapping):
    text = "".join(r.text for r in par.runs)
    if "{{" not in text or not par.runs:
        return
    for k, v in mapping.items():
        text = text.replace("{{" + k + "}}", v)
    par.runs[0].text = text
    for r in par.runs[1:]:
        r.text = ""


def _replace_everywhere(doc, mapping):
    def walk(pars, tables):
        for p in pars:
            _replace_in_paragraph(p, mapping)
        for t in tables:
            for row in t.rows:
                for cell in row.cells:
                    walk(cell.paragraphs, cell.tables)
    walk(doc.paragraphs, doc.tables)
    for s in doc.sections:
        walk(s.header.paragraphs, s.header.tables)
        walk(s.footer.paragraphs, s.footer.tables)


@output("docx")
def docx_output(invoice, params, ctx):
    """Fills a Word template. {{placeholders}} are replaced; the table row holding {{item}} is repeated per line."""
    from docx import Document

    template = resolve_path(ctx.cfg, params["template"])
    if not template.exists():
        raise BillingError(f"Word template not found: {template}")
    doc = Document(str(template))
    company = ctx.cfg["company"]
    mapping = {"company": company["name"], "client_name": invoice.client_name, "invoice_no": invoice.number,
               "date": f"{invoice.issue_date:%B %d, %Y}", "period": invoice.period,
               "total": f"{company.get('currency', '')} {_money(invoice.total)}".strip(),
               "payment_terms": str(company.get("payment_terms", ""))}
    mapping.update({f"client_{k}": v for k, v in invoice.client.items() if not k.startswith("_")})
    for table in doc.tables:
        for ri, row in enumerate(table.rows):
            if any("{{item}}" in c.text or "{{amount}}" in c.text for c in row.cells):
                proto = copy.deepcopy(row._tr)
                anchor = row._tr
                for l in invoice.lines:
                    new = copy.deepcopy(proto)
                    anchor.addnext(new)
                    anchor = new
                    r = _Row(new, table)
                    _replace_everywhere_row(r, {"item": l.item, "description": l.description, "qty": f"{l.qty:g}",
                                                "unit_price": _money(l.unit_price), "amount": _money(l.amount)})
                row._tr.getparent().remove(row._tr)
                break
    _replace_everywhere(doc, mapping)
    path = _fmt_path(ctx, params, "path", invoice, "{period}/{client_id}/invoice-{client_id}.docx")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    return [path]


def _replace_everywhere_row(row, mapping):
    for cell in row.cells:
        for p in cell.paragraphs:
            _replace_in_paragraph(p, mapping)


@output("ledger_csv")
def ledger_csv(invoice, params, ctx):
    path = _fmt_path(ctx, params, "path", invoice, "ledgers/{client_id}.csv")
    path.parent.mkdir(parents=True, exist_ok=True)
    new = not path.exists()
    with open(path, "a", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        if new:
            w.writerow(["Date", "Reference", "Debit", "Credit", "Balance"])
        bal = Decimal(0)
        if not new:
            with open(path, newline="", encoding="utf-8") as rf:
                for r in list(csv.DictReader(rf)):
                    bal += Decimal(r["Debit"] or 0) - Decimal(r["Credit"] or 0)
        w.writerow([invoice.issue_date.isoformat(), f"Invoice {invoice.number}", f"{invoice.total:.2f}", "",
                    f"{bal + invoice.total:.2f}"])
    return [path]


@output("ledger_xlsx")
def ledger_xlsx(invoice, params, ctx):
    """Appends the invoice to the client's ledger workbook. Payments are typed in by hand in the Credit column."""
    from openpyxl import Workbook, load_workbook

    path = _fmt_path(ctx, params, "path", invoice, "ledgers/{client_id}.xlsx")
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        wb = load_workbook(path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Date", "Reference", "Debit", "Credit", "Balance"])
    r = ws.max_row + 1
    ws.append([invoice.issue_date, f"Invoice {invoice.number}", float(invoice.total), None,
               f"=SUM(C$2:C{r})-SUM(D$2:D{r})"])
    ws.cell(r, 1).number_format = "yyyy-mm-dd"
    wb.save(path)
    return [path]
