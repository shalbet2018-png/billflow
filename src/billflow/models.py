import re
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Dict, List


class BillingError(Exception):
    """A problem with one client's data. The client is skipped and reported, never guessed."""


def to_money(value, what="value"):
    if isinstance(value, Decimal):
        return value
    if value is None or str(value).strip() == "":
        raise BillingError(f"{what} is empty")
    try:
        return Decimal(re.sub(r"[,\s]|[A-Za-z]{3}$", "", str(value)))
    except InvalidOperation:
        raise BillingError(f"{what} is not a number: {value!r}")


@dataclass
class Line:
    description: str
    qty: Decimal
    unit_price: Decimal
    amount: Decimal
    item: str = ""


@dataclass
class Invoice:
    client: Dict[str, str]
    client_id: str
    client_name: str
    number: str
    seq: int
    issue_date: date
    period: str
    lines: List[Line] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)

    @property
    def total(self) -> Decimal:
        return sum((l.amount for l in self.lines), Decimal("0"))
