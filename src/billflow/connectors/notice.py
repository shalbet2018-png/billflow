"""Arrears notices: a letter from the company's Word template for every client whose unpaid balance from
earlier months (the statement balance after this month, less this month's new invoice) is above a limit.

Needs `statement_xlsx` listed before it, since the balance comes from the client's statement."""
from datetime import timedelta
from decimal import Decimal

from ..config import resolve_path
from ..models import BillingError
from ..registry import output, output_check
from .outputs import _fmt_path, _replace_everywhere


def arrears(invoice):
    if invoice.balance is None:
        raise BillingError("Arrears notices need the statement_xlsx output listed before notice_docx")
    return invoice.balance - invoice.total


def _limit(params):
    try:
        return Decimal(str(params.get("arrears_above", 0)).replace(",", ""))
    except Exception:
        raise BillingError(f"arrears_above must be an amount, not {params.get('arrears_above')!r}")


def _ordinal(n):
    return f"{n}{'th' if 10 <= n % 100 <= 20 else {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')}"


def _amount(x):
    text = f"{x:,.2f}"
    return text[:-3] if text.endswith(".00") else text


@output_check("notice_docx")
def check_notice(invoice, params, ctx):
    template = resolve_path(ctx.cfg, params.get("template", ""))
    if not params.get("template") or not template.exists():
        raise BillingError(f"Notice template not found: {template}")
    owed = arrears(invoice)
    if owed > _limit(params):
        return [f"arrears {owed:,.2f}: notice will be written"]
    return []


@output("notice_docx")
def notice_docx(invoice, params, ctx):
    from docx import Document

    owed = arrears(invoice)
    if owed <= _limit(params):
        return []
    day = invoice.issue_date
    deadline = day + timedelta(days=int(params.get("pay_within_days", 14)))
    name = str(invoice.client_name).strip()
    mapping = {"company": ctx.cfg["company"]["name"], "client_name": name, "client_name_upper": name.upper(),
               "first_name": name.split()[0] if name else "", "client_id": invoice.client_id,
               "notice_date": f"{_ordinal(day.day)} {day:%B %Y}".upper(),
               "deadline": f"{_ordinal(deadline.day)} {deadline:%B %Y}",
               "arrears": _amount(owed), "balance": _amount(invoice.balance), "period_name": ctx.vars["period_name"]}
    mapping.update({f"client_{k}": str(v) for k, v in invoice.client.items() if not str(k).startswith("_")})
    doc = Document(str(resolve_path(ctx.cfg, params["template"])))
    _replace_everywhere(doc, mapping)
    path = _fmt_path(ctx, params, "path", invoice, "{period}/{client_id}/notice-{client_id}.docx")
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    return [path]
