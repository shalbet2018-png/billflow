import copy
import csv
from decimal import Decimal
from pathlib import Path

from docx.table import _Row

from ..config import resolve_path, safe_name
from ..models import BillingError
from ..registry import output, output_check


def _money(v):
    return f"{Decimal(v):,.2f}"


def path_fields(invoice, ctx):
    """Values a file path may use: {client_id} {client_name} {client_<column>} {seq} {invoice_no} and the period ones."""
    fields = {f"client_{k}": safe_name(v) for k, v in invoice.client.items() if not str(k).startswith("_")}
    fields.update(ctx.vars)
    fields.update(client_id=safe_name(invoice.client_id), client_name=safe_name(invoice.client_name),
                  seq=invoice.seq, invoice_no=safe_name(invoice.number))
    return fields


def _fmt_path(ctx, params, key, invoice, default):
    try:
        template = str(params.get(key, default)).format(**path_fields(invoice, ctx))
    except KeyError as e:
        raise BillingError(f"File name uses {{{e.args[0]}}}, which is not a column in the client list")
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


_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _own_texts(p_el):
    """The w:t pieces that belong to this paragraph itself (not to text boxes nested inside it)."""
    out = []
    for t in p_el.iter(_W + "t"):
        anc = t.getparent()
        while anc is not None and anc.tag != _W + "p":
            anc = anc.getparent()
        if anc is p_el:
            out.append(t)
    return out


def _replace_in_element(p_el, mapping):
    """Replaces {{placeholders}} in one paragraph. Word often splits text into pieces with their own
    formatting; the value goes into the piece where the placeholder starts, so every other piece keeps its look."""
    texts = _own_texts(p_el)
    full = "".join(t.text or "" for t in texts)
    if "{{" not in full:
        return
    import re
    starts, pos = [], 0
    for t in texts:
        starts.append(pos)
        pos += len(t.text or "")
    for m in reversed(list(re.finditer(r"\{\{([A-Za-z0-9_]+)\}\}", full))):
        if m.group(1) not in mapping:
            continue
        a, b = m.span()
        first = max(i for i, st in enumerate(starts) if st <= a)
        for i in range(len(texts) - 1, first - 1, -1):
            st, txt = starts[i], texts[i].text or ""
            end = st + len(txt)
            if end <= a or st >= b:
                continue
            lo, hi = max(a, st) - st, min(b, end) - st
            repl = str(mapping[m.group(1)]) if i == first else ""
            texts[i].text = txt[:lo] + repl + txt[hi:]
            texts[i].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _replace_in_paragraph(par, mapping):
    _replace_in_element(par._p, mapping)


def _replace_everywhere(doc, mapping):
    """Body, tables, text boxes, headers and footers."""
    parts = [doc.element.body]
    for s in doc.sections:
        for hf in (s.header, s.footer, s.first_page_header, s.first_page_footer, s.even_page_header, s.even_page_footer):
            if not hf.is_linked_to_previous:
                parts.append(hf._element)
    for part in parts:
        for p_el in list(part.iter(_W + "p")):
            _replace_in_element(p_el, mapping)


_LINE_FIELDS = ("item", "description", "qty", "unit_price", "amount")


def template_placeholders(template):
    """Every {{name}} in a Word template: body, tables, text boxes, headers and footers."""
    import re
    from docx import Document
    doc = Document(str(template))
    parts = [doc.element.body]
    for s in doc.sections:
        for hf in (s.header, s.footer, s.first_page_header, s.first_page_footer, s.even_page_header, s.even_page_footer):
            if not hf.is_linked_to_previous:
                parts.append(hf._element)
    names = set()
    for part in parts:
        for p_el in part.iter(_W + "p"):
            names.update(re.findall(r"\{\{([A-Za-z0-9_]+)\}\}", "".join(t.text or "" for t in _own_texts(p_el))))
    return names


def check_placeholders(template, mapping, extra=()):
    """Stops before anything is written if the template asks for a value BILLFLOW does not have,
    so a raw {{placeholder}} is never printed on a client's document."""
    missing = sorted(template_placeholders(template) - set(mapping) - set(extra))
    if missing:
        cols = [m[len("client_"):] for m in missing if m.startswith("client_")]
        hint = f" - add a column called {', '.join(repr(c) for c in cols)} to the client list" if cols else ""
        raise BillingError(f"{template.name} uses {', '.join('{{' + m + '}}' for m in missing)}, "
                           f"which BILLFLOW cannot fill{hint}")


def invoice_mapping(invoice, ctx):
    company = ctx.cfg["company"]
    mapping = {"company": company["name"], "client_name": invoice.client_name, "invoice_no": invoice.number,
               "date": f"{invoice.issue_date:%B %d, %Y}", "period": invoice.period,
               "date_long": f"{invoice.issue_date:%B} {invoice.issue_date.day}, {invoice.issue_date.year}",
               "period_name": ctx.vars["period_name"], "seq": str(invoice.seq),
               "total": f"{company.get('currency', '')} {_money(invoice.total)}".strip(),
               "total_amount": _money(invoice.total), "subtotal": _money(invoice.total),
               "payment_terms": str(company.get("payment_terms", ""))}
    mapping.update({f"client_{k}": v for k, v in invoice.client.items() if not k.startswith("_")})
    return mapping


@output_check("docx")
def check_docx(invoice, params, ctx):
    template = resolve_path(ctx.cfg, params.get("template", ""))
    if not params.get("template") or not template.exists():
        raise BillingError(f"Word template not found: {template}")
    check_placeholders(template, invoice_mapping(invoice, ctx), _LINE_FIELDS)
    return []


@output("docx")
def docx_output(invoice, params, ctx):
    """Fills a Word template. {{placeholders}} are replaced; the table row holding {{item}} is repeated per line."""
    from docx import Document

    template = resolve_path(ctx.cfg, params["template"])
    if not template.exists():
        raise BillingError(f"Word template not found: {template}")
    doc = Document(str(template))
    mapping = invoice_mapping(invoice, ctx)
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
    for p_el in list(row._tr.iter(_W + "p")):
        _replace_in_element(p_el, mapping)


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


@output("email")
def email_output(invoice, params, ctx):
    """Emails the files made earlier for this client. Saves an .eml draft unless mode is 'smtp'
    and the run was started with --send-emails. The SMTP password is read from an environment variable."""
    import mimetypes
    import os
    import smtplib
    from email.message import EmailMessage

    company = ctx.cfg["company"]
    to = str(invoice.client.get(params.get("to_column", "email"), "")).strip()
    if "@" not in to:
        raise BillingError(f"No valid email address for client {invoice.client_id}")
    fields = {"company": company["name"], "client_name": invoice.client_name, "invoice_no": invoice.number,
              "period": invoice.period, "total": f"{company.get('currency', '')} {_money(invoice.total)}".strip(),
              **{f"client_{k}": v for k, v in invoice.client.items() if not k.startswith("_")}}
    subject = str(params.get("subject", "Invoice {invoice_no} - {period}")).format(**fields)
    body = str(params.get("body", "Dear {client_name},\n\nPlease find attached invoice {invoice_no} "
                                  "for {period}. Total: {total}.\n\nKind regards,\n{company}")).format(**fields)
    sender = params.get("from") or os.environ.get(params.get("from_env", "BILLFLOW_FROM"), "")
    msg = EmailMessage()
    msg["To"], msg["Subject"] = to, subject
    msg["From"] = sender or "billing@example.invalid"
    if params.get("cc"):
        msg["Cc"] = params["cc"]
    msg.set_content(body)
    wanted = [e.lower() for e in params.get("attach", [".pdf"])]
    for f in ctx.files:
        p = Path(f)
        if p.suffix.lower() in wanted:
            ctype = (mimetypes.guess_type(p.name)[0] or "application/octet-stream").split("/", 1)
            msg.add_attachment(p.read_bytes(), maintype=ctype[0], subtype=ctype[1], filename=p.name)

    draft = _fmt_path(ctx, params, "draft_path", invoice, "{period}/{client_id}/email-{client_id}.eml")
    draft.parent.mkdir(parents=True, exist_ok=True)
    draft.write_bytes(bytes(msg))
    if params.get("mode", "draft") != "smtp" or not ctx.cfg.get("_send_emails"):
        return [draft]
    if not sender:
        raise BillingError("Email 'from' address is not set")
    host = params.get("host")
    if not host:
        raise BillingError("Email 'host' is not set")
    user = params.get("username") or sender
    password = os.environ.get(params.get("password_env", "BILLFLOW_SMTP_PASSWORD"), "")
    try:
        with smtplib.SMTP(host, int(params.get("port", 587)), timeout=30) as s:
            s.starttls()
            if password:
                s.login(user, password)
            s.send_message(msg)
    except (smtplib.SMTPException, OSError) as e:
        raise BillingError(f"Email to {to} failed: {e}")
    return [draft]


@output("json")
def json_output(invoice, params, ctx):
    """One JSON file per invoice, for importing into other systems."""
    import json
    path = _fmt_path(ctx, params, "path", invoice, "{period}/{client_id}/invoice-{client_id}.json")
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {"company": ctx.cfg["company"]["name"], "currency": ctx.cfg["company"].get("currency", ""),
            "invoice_no": invoice.number, "client_id": invoice.client_id, "client_name": invoice.client_name,
            "issue_date": invoice.issue_date.isoformat(), "period": invoice.period,
            "lines": [{"item": l.item, "description": l.description, "qty": str(l.qty),
                       "unit_price": str(l.unit_price), "amount": str(l.amount)} for l in invoice.lines],
            "total": str(invoice.total)}
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return [path]
