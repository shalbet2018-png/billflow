"""Payments received during the billing month, read from a source (e.g. payments-2026-12.csv) and
written into each client's statement together with the month's invoice.

Every payment is checked before anything is written: a missing or unreadable date, amount or reference,
a payment dated outside the month, the same payment listed twice, or a payment for a unit that is not in the
client list is reported, never guessed or dropped."""
import re
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal

from .config import parse_period, period_end
from .models import BillingError, to_money


@dataclass
class Payment:
    date: date
    reference: str
    amount: Decimal
    deposit_by: str = ""


def parse_day(value, order="dmy"):
    """Dates as 2026-12-03, 03/12/2026 (day first, as written in Kenya) or an Excel date."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if not text:
        raise BillingError("Payment date is empty")
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]00:00(?::00)?)?", text)
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.fullmatch(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})", text)
        if not m:
            raise BillingError(f"Payment date not understood: {text!r} (use 2026-12-03 or 03/12/2026)")
        a, b, y = map(int, m.groups())
        d, mo = (a, b) if order == "dmy" else (b, a)
    try:
        return date(y, mo, d)
    except ValueError:
        raise BillingError(f"Payment date is not a real date: {text!r}")


def _settings(cfg):
    p = cfg.get("payments") or {}
    return {"source": p.get("source"), "date": p.get("date_column", "date"),
            "reference": p.get("reference_column", "reference"), "amount": p.get("amount_column", "amount"),
            "deposit_by": p.get("deposit_by_column", "deposit_by"), "order": p.get("date_order", "dmy")}


def enabled(cfg):
    return bool(_settings(cfg)["source"])


def client_payments(cfg, ctx, client_id):
    """This client's payments for the month, checked and sorted by date."""
    s = _settings(cfg)
    key = ctx.source_key(s["source"])
    rows = [r for r in ctx.source_rows(s["source"]) if r.get(key, "").strip() == client_id]
    year, month = parse_period(ctx.period)
    out, refs = [], set()
    for r in rows:
        ref = str(r.get(s["reference"], "")).strip()
        if not ref:
            raise BillingError("A payment has no reference (the receipt / PMT number is needed to avoid entering it twice)")
        when = parse_day(r.get(s["date"]), s["order"])
        if (when.year, when.month) != (year, month):
            raise BillingError(f"Payment {ref} is dated {when:%d %b %Y}, outside {period_end(ctx.period):%B %Y}")
        amount = to_money(r.get(s["amount"]), f"amount of payment {ref}")
        if amount <= 0:
            raise BillingError(f"Payment {ref} has amount {amount} - it must be more than 0")
        # the same receipt typed twice (references alone are sometimes reused for different payments)
        norm = (re.sub(r"\s+", "", ref).lower(), amount)
        if norm in refs:
            raise BillingError(f"Payment {ref} is listed twice")
        refs.add(norm)
        out.append(Payment(when, ref, amount, str(r.get(s["deposit_by"], "") or "").strip()))
    return sorted(out, key=lambda p: p.date)


def unknown_clients(cfg, ctx, known_ids):
    """Payments for units that are not in the client list, as (unit, total) pairs."""
    s = _settings(cfg)
    key = ctx.source_key(s["source"])
    found = {}
    for r in ctx.source_rows(s["source"]):
        cid = r.get(key, "").strip()
        if cid and cid not in known_ids:
            try:
                amt = to_money(r.get(s["amount"]), "amount")
            except BillingError:
                amt = Decimal(0)
            found[cid] = found.get(cid, Decimal(0)) + amt
    return sorted(found.items())
