"""Plain-language explanations of billing problems, for people who are not technical.

Every message the engine produces is turned into: what happened, what to do about it, and (when known) the file
to open. Nothing here changes how billing works - it only explains it.
"""
import calendar
import re
from dataclasses import dataclass
from pathlib import Path

from .config import ConfigError, parse_period, period_vars, resolve_path


@dataclass
class Explanation:
    text: str                 # what happened, in plain words
    fix: str = ""             # what to do about it
    file: str = ""            # a file or folder worth opening, if known


def month_label(period):
    """'2026-11' -> 'November 2026' (unchanged if it is not a valid month)."""
    try:
        y, m = parse_period(period)
    except ConfigError:
        return str(period)
    return f"{calendar.month_name[m]} {y}"


def _source_file(cfg, name, period):
    """The real file a source reads for this month, or '' if it is not a local file."""
    sc = (cfg or {}).get("sources", {}).get(name) or {}
    loc = sc.get("path")
    if not loc or "_dir" not in cfg:
        return ""
    try:
        return str(resolve_path(cfg, loc, **period_vars(period)))
    except (ConfigError, KeyError, IndexError, ValueError):
        return ""


def _client_file(cfg, period):
    cl = (cfg or {}).get("clients") or {}
    if not cl.get("path") or "_dir" not in cfg:
        return ""
    try:
        return str(resolve_path(cfg, cl["path"], **period_vars(period)))
    except (ConfigError, KeyError, IndexError, ValueError):
        return ""


def _charge_for_source(cfg, name):
    for ch in (cfg or {}).get("charges") or []:
        if ch.get("source") == name:
            return ch.get("name") or name
    return name


def _reading_files(cfg, period):
    """Files of meter-reading charges (used when a message does not say which reading it was)."""
    out = []
    for ch in (cfg or {}).get("charges") or []:
        if ch.get("rule") in ("metered", "tiered") and ch.get("source"):
            f = _source_file(cfg, ch["source"], period)
            if f and f not in out:
                out.append(f)
    return out


def explain(message, cfg=None, period="", client_id=""):
    """Turns one engine message into an Explanation."""
    msg = str(message or "").strip()
    who = client_id or "this client"
    month = month_label(period)

    m = re.fullmatch(r"Already invoiced for (.+)", msg)
    if m:
        return Explanation(f"Already invoiced for {month_label(m.group(1))}. Skipped so nobody is billed twice.")

    if msg == "Client appears twice in the client list":
        f = _client_file(cfg, period)
        return Explanation(f"{who} is listed twice in your client list.",
                           f"Delete or merge the extra row for {who} in your client list, save it, then press Check invoices again.", f)

    m = re.fullmatch(r"No '(.+)' (reading|entry) found for (.+)", msg)
    if m:
        src, kind, cid = m.groups()
        f = _source_file(cfg, src, period)
        charge = _charge_for_source(cfg, src)
        where = f" in {Path(f).name}" if f else ""
        return Explanation(f"No {charge} {kind} for {cid}{where}.",
                           f"Add the {kind} for {cid} to the file, save it, then press Check invoices again.", f)

    m = re.fullmatch(r"(\d+) '(.+)' rows found for (.+) - expected one", msg)
    if m:
        n, src, cid = m.groups()
        f = _source_file(cfg, src, period)
        charge = _charge_for_source(cfg, src)
        where = f" in {Path(f).name}" if f else ""
        return Explanation(f"{cid} has {n} {charge} rows{where}, but there should be only one.",
                           f"Keep one row for {cid} (delete the extra ones), save, then press Check invoices again.", f)

    m = re.fullmatch(r"Reading went down \((.+) -> (.+)\)", msg)
    if m:
        prev, cur = m.groups()
        files = _reading_files(cfg, period)
        return Explanation(f"The meter reading went down: last month {prev}, this month {cur}.",
                           f"Check {who}'s readings - the current reading should be the same or higher than the previous one. "
                           "Correct the file, save, then press Check invoices again.", files[0] if len(files) == 1 else "")

    m = re.fullmatch(r"(.+) is empty", msg)
    if m:
        what = m.group(1).replace(f" for {client_id}", "") if client_id else m.group(1)
        files = _reading_files(cfg, period) if "reading" in what else []
        return Explanation(f"The {what} for {who} is blank.",
                           f"Fill it in in your file, save, then press Check invoices again.", files[0] if len(files) == 1 else "")

    m = re.fullmatch(r"(.+) is not a number: (.+)", msg)
    if m:
        what = m.group(1).replace(f" for {client_id}", "") if client_id else m.group(1)
        files = _reading_files(cfg, period) if "reading" in what else []
        return Explanation(f"The {what} for {who} should be a number, but it says {m.group(2)}.",
                           "Change it to a plain number such as 1500 or 1500.50, save, then press Check invoices again.",
                           files[0] if len(files) == 1 else "")

    m = re.fullmatch(r"File not found: (.+)", msg)
    if m:
        p = Path(m.group(1))
        return Explanation(f"Could not find the file '{p.name}' for {month}.",
                           f"Save {month}'s file in the folder {p.parent} with exactly this name: {p.name}. "
                           "If the file has a different name or lives somewhere else, press Change my setup.",
                           str(p.parent) if p.parent.exists() else "")

    m = re.fullmatch(r"Sheet '(.+)' not found in (.+) \((?:has|it has): (.+)\)", msg)
    if m:
        sheet, fname, has = m.groups()
        return Explanation(f"The file {fname} has no sheet called '{sheet}' (its sheets are: {has}).",
                           f"Rename the sheet to '{sheet}', or press Change my setup and pick the right sheet.")

    if msg.startswith("Writing files failed"):
        return Explanation("This client's invoice could not be saved.",
                           "An invoice or ledger may be open in Word or Excel - close it and press Create invoices again. "
                           "Other clients were not affected.")

    m = re.fullmatch(r"Word template not found: (.+)", msg)
    if m:
        return Explanation(f"The Word invoice template '{Path(m.group(1)).name}' is missing.",
                           "Put the template back, or press Change my setup and choose it again.")

    m = re.fullmatch(r"Statement not found: (.+)", msg)
    if m:
        p = Path(m.group(1))
        return Explanation(f"{who}'s statement '{p.name}' could not be found.",
                           f"Check that the statement file name in your client list matches the file in {who}'s folder exactly, "
                           "save the client list, then press Check invoices again.", str(p.parent) if p.parent.exists() else "")

    m = re.fullmatch(r"(.+) already has a row for (.+)", msg)
    if m:
        return Explanation(f"{who}'s statement already shows {m.group(2)}.",
                           "If this month was already entered by hand, nothing needs doing. Otherwise check the invoice number "
                           "in your client list (last invoice column).")

    m = re.fullmatch(r"(.+): balance cannot be worked out - (.+)", msg)
    if m:
        return Explanation(f"The balance on {who}'s statement ({m.group(1)}) cannot be worked out: {m.group(2)}.",
                           "Open the statement and check the Balance column above the 'Total must equal' row, "
                           "then press Check invoices again.")

    m = re.fullmatch(r"(.+\.docx) uses (.+), which BILLFLOW cannot fill(?: - add a column called (.+) to the client list)?", msg)
    if m:
        fix = (f"Add a column called {m.group(3)} to your client list (type - where there is nothing), "
               "save it, then press Check invoices again." if m.group(3) else
               "Remove or correct that {{...}} in the Word template, then press Check invoices again.")
        return Explanation(f"The template {m.group(1)} has {m.group(2)}, which BILLFLOW has no value for. "
                           "Nothing was written, so no document shows a raw {{...}}.", fix,
                           _client_file(cfg, period) if m.group(3) else "")

    m = re.fullmatch(r"The charge '(.+)' uses skip_if: (.+), but the client list has no '(.+)' column", msg)
    if m:
        return Explanation(f"The {m.group(1)} is set to be left off some units using a '{m.group(3)}' column, "
                           "but your client list has no such column.",
                           f"Add a '{m.group(3)}' column to the client list: yes for units that pay the service charge "
                           "only, empty for the rest.", _client_file(cfg, period))

    m = re.fullmatch(r"'(.+)' must be yes or empty, not (.+)", msg)
    if m:
        return Explanation(f"The '{m.group(1)}' column for {who} says {m.group(2)}.",
                           f"Type yes in '{m.group(1)}' for a unit that pays the service charge only, or leave it empty.",
                           _client_file(cfg, period))

    m = re.fullmatch(r"(.+\.xlsx): could not find '(.+)' on the statement", msg)
    if m:
        return Explanation(f"{who}'s statement ({m.group(1)}) has no '{m.group(2)}' label, so the date cannot be updated.",
                           "Open the statement and check the label at the top, then press Check invoices again.")

    m = re.fullmatch(r"No previous (.+) reading: (.+)", msg)
    if m:
        return Explanation(f"Last month's {m.group(1)} reading for {who} is missing ({m.group(2)}).",
                           "Make sure last month's readings file is in the folder and has this unit's current reading, "
                           "or type the previous reading into this month's file, then press Check invoices again.")

    m = re.fullmatch(r"Payments of (.+) for unit (.+), which is not in the client list", msg)
    if m:
        return Explanation(f"This month's payments file has {m.group(1)} paid for unit {m.group(2)}, "
                           "which is not in your client list.",
                           "Check the unit number on those payments (it may be a typing mistake), or add the client to "
                           "your client list, then press Check invoices again.", _source_file(cfg, ((cfg or {}).get("payments") or {}).get("source", "payments"), period))

    if re.search(r"(?i)payment", msg):
        return Explanation(f"A payment for {who} has a problem: {msg}",
                           "Correct the row in this month's payments file, save it, then press Check invoices again.",
                           _source_file(cfg, ((cfg or {}).get("payments") or {}).get("source", "payments"), period))

    if msg.startswith("Arrears notices need") or msg.startswith("Notice template not found"):
        return Explanation(msg, "Press Change my setup, or ask whoever set up BILLFLOW to fix the notice settings.")

    if msg == "No charges calculated":
        return Explanation(f"Nothing to charge {who} this month (every charge came to nothing).",
                           "If that is wrong, check this client's rows in your files.")

    m = re.fullmatch(r"No valid email address for client (.+)", msg)
    if m:
        return Explanation(f"{m.group(1)} has no valid email address.", "Add their email in your client list.",
                           _client_file(cfg, period))

    if msg.startswith(("Config file not found", "Config is missing", "Config needs", "Unknown ", "Charge refers to")):
        return Explanation("Your setup file is missing or damaged.",
                           "Press Change my setup to repair it, or New setup to make it again.")

    if msg.startswith("Period must look like"):
        return Explanation("Please choose a month from the list.")

    return Explanation(msg, "If this keeps happening, send us this message and your setup file (config.yml).")


def describe_results(results, cfg=None, period="", currency=""):
    """Results -> list of dicts the screen shows, problems first:
    {status, client_id, client_name, text, fix, file, number, total}."""
    order = {"ERROR": 0, "SKIPPED": 1, "OK": 2}
    out = []
    for r in sorted(results, key=lambda r: order.get(r.status, 3)):
        d = {"status": r.status, "client_id": r.client_id, "client_name": r.client_name, "text": "", "fix": "", "file": "",
             "number": r.invoice.number if r.invoice else "", "total": r.invoice.total if r.invoice else None}
        if r.status == "OK":
            d["text"] = f"invoice {r.invoice.number}  -  {(currency + ' ') if currency else ''}{r.invoice.total:,.2f}"
        else:
            e = explain(r.message, cfg, period, r.client_id)
            d.update(text=e.text, fix=e.fix, file=e.file)
        out.append(d)
    return out
