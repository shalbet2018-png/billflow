from decimal import Decimal

from ..models import BillingError, Line, to_money
from ..registry import rule


def _client_rows(ctx, params, client):
    src = params["source"]
    key = ctx.source_key(src)
    rows = [r for r in ctx.source_rows(src) if r.get(key, "").strip() == client["_id"]]
    return rows


@rule("fixed")
def fixed(client, params, ctx):
    """A flat amount per client, e.g. a service charge held in the client list."""
    if "column" in params:
        amount = to_money(client.get(params["column"]), f"{params['column']} for {client['_id']}")
    else:
        amount = to_money(params["amount"], "amount")
    return [Line(params.get("description", params["name"]), Decimal(1), amount, amount, params["name"])]


@rule("metered")
def metered(client, params, ctx):
    """(current reading - previous reading) x rate, from a usage source."""
    rows = _client_rows(ctx, params, client)
    if not rows:
        if params.get("required", True):
            raise BillingError(f"No '{params['source']}' reading found for {client['_id']}")
        return []
    if len(rows) > 1:
        raise BillingError(f"{len(rows)} '{params['source']}' rows found for {client['_id']} - expected one")
    row = rows[0]
    cur = to_money(row.get(params["current_column"]), "current reading")
    prev = to_money(row.get(params["previous_column"]), "previous reading")
    if cur < prev:
        raise BillingError(f"Reading went down ({prev} -> {cur})")
    if "rate_column" in params:
        rate = to_money(client.get(params["rate_column"]) or row.get(params["rate_column"]), "rate")
    else:
        rate = to_money(params["rate"], "rate")
    units = cur - prev
    desc = f"Meter Reading (Last Reading: {prev} / Current Reading: {cur})"
    return [Line(desc, units, rate, (units * rate).quantize(Decimal("0.01")), params["name"])]


@rule("adjustment")
def adjustment(client, params, ctx):
    """Extra charges from a list (e.g. electricity bills). Several rows for one client are added together."""
    rows = _client_rows(ctx, params, client)
    if not rows:
        if params.get("required", False):
            raise BillingError(f"No '{params['source']}' entry found for {client['_id']}")
        return []
    amounts = [to_money(r.get(params["amount_column"]), "amount") for r in rows]
    total = sum(amounts, Decimal(0))
    dcol = params.get("description_column")
    descs = [r.get(dcol, "") for r in rows if dcol and r.get(dcol, "")]
    desc = ", ".join(descs) if descs else params["name"]
    return [Line(desc, Decimal(1), total, total, params["name"])]


@rule("percentage")
def percentage(client, params, ctx):
    """A percentage of earlier charges (e.g. VAT or a late fee). 'of' limits it to named charges; default is all so far."""
    base = [l for l in client.get("_lines", []) if not params.get("of") or l.item in params["of"]]
    if not base:
        return []
    subtotal = sum((l.amount for l in base), Decimal(0))
    pct = to_money(params["percent"], "percent")
    amount = (subtotal * pct / 100).quantize(Decimal("0.01"))
    desc = params.get("description", f"{pct:g}% of {subtotal:,.2f}")
    return [Line(desc, Decimal(1), amount, amount, params["name"])]


@rule("tiered")
def tiered(client, params, ctx):
    """Usage priced in bands, e.g. tiers: [{up_to: 10, rate: 50}, {rate: 80}] (the last band has no limit)."""
    rows = _client_rows(ctx, params, client)
    if not rows:
        if params.get("required", True):
            raise BillingError(f"No '{params['source']}' reading found for {client['_id']}")
        return []
    if len(rows) > 1:
        raise BillingError(f"{len(rows)} '{params['source']}' rows found for {client['_id']} - expected one")
    cur = to_money(rows[0].get(params["current_column"]), "current reading")
    prev = to_money(rows[0].get(params["previous_column"]), "previous reading")
    if cur < prev:
        raise BillingError(f"Reading went down ({prev} -> {cur})")
    left, lower, lines = cur - prev, Decimal(0), []
    for t in params["tiers"]:
        up = t.get("up_to")
        band = left if up is None else min(left, Decimal(str(up)) - lower)
        if band > 0:
            rate = to_money(t["rate"], "tier rate")
            lines.append(Line(f"{params['name']} ({band:g} units)", band, rate,
                              (band * rate).quantize(Decimal("0.01")), params["name"]))
            left -= band
        if up is not None:
            lower = Decimal(str(up))
    if left > 0:
        raise BillingError("Usage exceeds the last tier - give the last tier no 'up_to'")
    return lines
