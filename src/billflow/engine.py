import csv
import importlib.util
import json
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path

from . import connectors, payments  # noqa: F401
from .config import ConfigError, parse_period, period_vars, resolve_path, safe_name
from .models import BillingError, Invoice
from .registry import CHECKS, OUTPUTS, RULES, SOURCES


@dataclass
class Result:
    client_id: str
    client_name: str
    status: str            # OK | ERROR | SKIPPED
    invoice: Invoice = None
    message: str = ""
    files: list = field(default_factory=list)


class Context:
    def __init__(self, cfg, period):
        self.cfg = cfg
        self.period = period
        self.vars = period_vars(period)
        self._cache = {}
        self.files = []

    def _src_cfg(self, name):
        try:
            return self.cfg["sources"][name]
        except KeyError:
            raise ConfigError(f"Charge refers to unknown source '{name}'")

    def source_key(self, name):
        return self._src_cfg(name).get("key", "id")

    def source_rows(self, name, period=None):
        """Rows of a source for this run's month, or for another month (e.g. last month's readings)."""
        period = period or self.period
        key = (name, period)
        if key not in self._cache:
            sc = self._src_cfg(name)
            try:
                fn = SOURCES[sc["type"]]
            except KeyError:
                raise ConfigError(f"Unknown source type '{sc.get('type')}'")
            try:
                self._cache[key] = fn(sc, self if period == self.period else Context(self.cfg, period))
            except BillingError:
                if sc.get("optional") and period == self.period:
                    self._cache[key] = []
                else:
                    raise
        return self._cache[key]


def load_plugins(cfg):
    """Company-specific rules/sources/outputs: python files listed under 'plugins' in the config."""
    for p in cfg.get("plugins", []) or []:
        path = resolve_path(cfg, p)
        if not path.exists():
            raise ConfigError(f"Plugin not found: {path}")
        spec = importlib.util.spec_from_file_location(f"billflow_plugin_{path.stem}", path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)


def _state_path(cfg):
    return resolve_path(cfg, cfg.get("state_file", "billflow-state.json"))


def _load_state(cfg):
    p = _state_path(cfg)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {"last_seq": {}, "done": {}}


def _issue_date(period):
    y, m = parse_period(period)
    return date(y + (m == 12), 1 if m == 12 else m + 1, 1)


def plan(cfg, period, only=None, force=False):
    """Works everything out without writing a single file."""
    load_plugins(cfg)
    ctx = Context(cfg, period)
    state = _load_state(cfg)
    cc = cfg["clients"]
    cs = dict(cc, **{"type": cc.get("type", "csv")})
    clients = SOURCES[cs["type"]](cs, ctx)
    idc, namec = cc.get("id_column", "id"), cc.get("name_column", "name")
    results, seen = [], set()
    for row in clients:
        cid = row.get(idc, "").strip()
        if not cid:
            continue
        name = row.get(namec, cid)
        if only and cid not in only:
            continue
        if cid in seen:
            results.append(Result(cid, name, "ERROR", message="Client appears twice in the client list"))
            continue
        seen.add(cid)
        if not force and cid in state["done"].get(period, []):
            results.append(Result(cid, name, "SKIPPED", message=f"Already invoiced for {period}"))
            continue
        client = dict(row, _id=cid)
        try:
            lines = []
            client["_lines"] = lines
            for ch in cfg["charges"]:
                fn = RULES.get(ch["rule"])
                if not fn:
                    raise ConfigError(f"Unknown charge rule '{ch['rule']}'")
                lines += fn(client, ch, ctx)
            if not lines:
                raise BillingError("No charges calculated")
            last = state["last_seq"].get(cid)
            if last is None:
                lc = cc.get("last_invoice_column")
                last = int(row.get(lc) or 0) if lc else 0
            seq = int(last) + 1
            inv = Invoice(row, cid, name, f"{cid}/{seq}", seq, _issue_date(period), period, lines)
            if payments.enabled(cfg):
                inv.payments = payments.client_payments(cfg, ctx, cid)
            notes = []
            if inv.payments:
                paid = sum((p.amount for p in inv.payments), Decimal(0))
                notes.append(f"{len(inv.payments)} payment(s) received: {paid:,.2f}")
            for out in cfg["outputs"]:
                check = CHECKS.get(out.get("type"))
                if check:
                    notes += check(inv, out, ctx) or []
            results.append(Result(cid, name, "OK", inv, "; ".join(notes)))
        except BillingError as e:
            results.append(Result(cid, name, "ERROR", message=str(e)))
    if payments.enabled(cfg) and not only:
        try:
            for cid, amt in payments.unknown_clients(cfg, ctx, seen):
                results.append(Result(cid, "(not in client list)", "ERROR",
                                      message=f"Payments of {amt:,.2f} for unit {cid}, which is not in the client list"))
        except BillingError:
            pass  # the payments file itself could not be read: every client already reports why
    return results


def commit(cfg, period, results):
    """Writes the files for every OK result and records the invoice numbers used."""
    ctx = Context(cfg, period)
    state = _load_state(cfg)
    for res in results:
        if res.status != "OK":
            continue
        try:
            ctx.files = []
            for out in cfg["outputs"]:
                fn = OUTPUTS.get(out["type"])
                if not fn:
                    raise ConfigError(f"Unknown output type '{out['type']}'")
                made = [str(p) for p in fn(res.invoice, out, ctx)]
                ctx.files += made
                res.files += made
        except (BillingError, OSError) as e:
            res.status, res.message = "ERROR", f"Writing files failed: {e}"
            continue
        state["last_seq"][res.client_id] = res.invoice.seq
        state["done"].setdefault(period, []).append(res.client_id)
    _state_path(cfg).write_text(json.dumps(state, indent=2), encoding="utf-8")
    out_dir = resolve_path(cfg, cfg.get("output_dir", "output"))
    summary = out_dir / period / "summary.csv"
    summary.parent.mkdir(parents=True, exist_ok=True)
    with open(summary, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["Client", "Name", "Status", "Invoice", "Total", "Notes", "Paid", "Balance after", "Notice"])
        for r in results:
            inv = r.invoice if r.status == "OK" else None
            w.writerow([r.client_id, r.client_name, r.status, r.invoice.number if r.invoice else "",
                        f"{r.invoice.total:.2f}" if r.invoice else "", r.message,
                        f"{sum(p.amount for p in inv.payments):.2f}" if inv and inv.payments else "",
                        f"{inv.balance:.2f}" if inv and inv.balance is not None else "",
                        "yes" if inv and inv.notice_due else ""])
    from .report import write_workbook
    write_workbook(summary.with_suffix(".xlsx"), cfg, period, results)
    return summary
