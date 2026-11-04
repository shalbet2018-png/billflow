"""Read the text of documents the user picks: Word, Excel, PowerPoint and PDF.

BILLFLOW never searches the computer by itself. It only opens files the user explicitly chooses
(through the operating system's file dialog, or a path typed on the command line)."""
import re
import zipfile
from pathlib import Path

from .models import BillingError

SUPPORTED = (".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm", ".pdf", ".csv", ".txt")
LEGACY = {".doc": "Word", ".xls": "Excel", ".ppt": "PowerPoint"}
FILE_TYPES = [
    ("Office and PDF documents", "*.docx *.docm *.xlsx *.xlsm *.pptx *.pptm *.pdf *.doc *.xls *.ppt"),
    ("Word", "*.docx *.docm *.doc"),
    ("Excel", "*.xlsx *.xlsm *.xls"),
    ("PowerPoint", "*.pptx *.pptm *.ppt"),
    ("PDF", "*.pdf"),
    ("All files", "*.*"),
]


def _word(path):
    from docx import Document
    doc = Document(str(path))
    lines = [p.text for p in doc.paragraphs if p.text.strip()]
    for t in doc.tables:
        for row in t.rows:
            cells = [c.text.strip() for c in row.cells]
            if any(cells):
                lines.append("\t".join(cells))
    fields = []
    for el in doc.element.body.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText"):
        code = (el.text or "").strip()
        if code.startswith("="):
            fields.append(f"{code}\n    means: {explain_formula(code)}")
    if fields:
        lines.append(f"\n[Formula fields found: {len(fields)}]")
        lines.extend(fields)
    return "\n".join(lines)


FUNCTIONS = {
    "SUM": "adds up", "SUMIF": "adds up the cells that meet a condition", "SUMIFS": "adds up cells meeting several conditions",
    "SUMPRODUCT": "multiplies lists together and adds the results", "AVERAGE": "averages", "AVERAGEIF": "averages cells meeting a condition",
    "COUNT": "counts numbers in", "COUNTA": "counts non-empty cells in", "COUNTIF": "counts cells meeting a condition",
    "COUNTIFS": "counts cells meeting several conditions", "MIN": "finds the smallest of", "MAX": "finds the largest of",
    "IF": "chooses between two results depending on a test", "IFS": "checks several tests in turn", "IFERROR": "replaces an error with another value",
    "AND": "checks that all tests are true", "OR": "checks that any test is true", "NOT": "reverses a true/false test",
    "VLOOKUP": "looks a value up in the first column of a table and returns a value from the same row",
    "HLOOKUP": "looks a value up in the first row of a table", "XLOOKUP": "looks a value up and returns the matching result",
    "INDEX": "returns the value at a position in a range", "MATCH": "finds the position of a value in a range",
    "ROUND": "rounds a number", "ROUNDUP": "rounds a number up", "ROUNDDOWN": "rounds a number down", "ABS": "takes the absolute value of",
    "CONCATENATE": "joins text together", "CONCAT": "joins text together", "TEXTJOIN": "joins text with a separator",
    "LEFT": "takes characters from the start of text", "RIGHT": "takes characters from the end of text", "MID": "takes characters from the middle of text",
    "LEN": "counts the characters in", "TRIM": "removes extra spaces from", "UPPER": "makes text upper case", "LOWER": "makes text lower case",
    "TEXT": "formats a value as text", "VALUE": "turns text into a number", "TODAY": "gives today's date", "NOW": "gives the current date and time",
    "DATE": "builds a date", "YEAR": "takes the year from a date", "MONTH": "takes the month from a date", "DAY": "takes the day from a date",
    "EOMONTH": "finds the last day of a month", "EDATE": "moves a date by whole months", "DATEDIF": "measures the time between two dates",
    "PMT": "works out a loan payment", "NPV": "works out net present value", "IRR": "works out the internal rate of return",
    "SUBTOTAL": "gives a subtotal of", "ISBLANK": "checks whether a cell is empty", "ISNUMBER": "checks whether a value is a number",
    "ISERROR": "checks whether a value is an error", "SUMX": "adds up", "TRUNC": "cuts off the decimals of", "MOD": "gives the remainder of a division",
    "FILTER": "keeps only the rows that meet a condition", "SORT": "sorts", "UNIQUE": "lists the distinct values in", "LET": "names intermediate results",
}
_FUNC_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\(")
_REF_RE = re.compile(r"(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?")
_OPS = {"+": "add", "-": "subtract", "*": "multiply", "/": "divide", "^": "raise to a power", "&": "join text"}


def explain_formula(formula):
    """A plain-English summary of an Excel formula: its functions, the cells it reads and its arithmetic."""
    body = re.sub(r'"[^"]*"', '""', formula.lstrip("="))
    parts = []
    funcs = []
    for name in _FUNC_RE.findall(body):
        n = name.upper().replace("_XLFN.", "")
        if n not in funcs:
            funcs.append(n)
    for n in funcs:
        parts.append(f"{n} {FUNCTIONS[n]}" if n in FUNCTIONS else f"{n} (function)")
    refs = list(dict.fromkeys(_REF_RE.findall(body)))
    if refs:
        parts.append("reads " + ", ".join(refs))
    ops = [o for o in _OPS if o in body.replace("->", "") and (o != "-" or re.search(r"[\w)]\s*-", body))]
    if ops:
        parts.append("uses: " + ", ".join(_OPS[o] for o in ops))
    return "; ".join(parts) or "a fixed value"


def _excel(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=False)
    vals = load_workbook(path, data_only=True)
    try:
        lines = []
        formulas = []
        for ws in wb.worksheets:
            vws = vals[ws.title]
            lines.append(f"[Sheet: {ws.title}]")
            for row in ws.iter_rows():
                cells = []
                for c in row:
                    v = c.value
                    text = getattr(v, "text", v)
                    if isinstance(text, str) and text.startswith("="):
                        shown = vws[c.coordinate].value
                        cells.append(f"{text} -> {'(not calculated yet)' if shown is None else shown}")
                        formulas.append(f"{ws.title}!{c.coordinate}: {text} -> "
                                        f"{'(not calculated yet)' if shown is None else shown}\n    means: {explain_formula(text)}")
                    else:
                        cells.append("" if v is None else str(v).strip())
                if any(cells):
                    lines.append("\t".join(cells))
            lines.append("")
        if formulas:
            lines.append(f"\n[Formulas found: {len(formulas)}]")
            lines.extend(formulas)
        return "\n".join(lines)
    finally:
        wb.close()
        vals.close()


def _old_excel(path):
    try:
        import xlrd
    except ImportError:
        raise BillingError(f"{path.name} is an old-style Excel file. Open it in Excel and use "
                           "File > Save As > Excel Workbook (.xlsx), then choose that file.")
    book = xlrd.open_workbook(str(path))
    lines = []
    for sh in book.sheets():
        lines.append(f"[Sheet: {sh.name}]")
        for i in range(sh.nrows):
            vals = [str(v).strip() for v in sh.row_values(i)]
            if any(vals):
                lines.append("\t".join(vals))
    return "\n".join(lines)


def _powerpoint(path):
    from xml.etree import ElementTree as ET
    ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    with zipfile.ZipFile(path) as z:
        slides = sorted((n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")),
                        key=lambda n: int("".join(c for c in n if c.isdigit()) or 0))
        out = []
        for i, name in enumerate(slides, 1):
            root = ET.fromstring(z.read(name))
            paras = ["".join(t.text or "" for t in p.iter(ns + "t")) for p in root.iter(ns + "p")]
            out.append(f"[Slide {i}]\n" + "\n".join(p for p in paras if p.strip()))
        return "\n".join(out)


def _pdf(path):
    try:
        from pypdf import PdfReader
    except ImportError:
        raise BillingError("PDF reading needs the 'pypdf' package (pip install pypdf).")
    reader = PdfReader(str(path))
    if reader.is_encrypted:
        raise BillingError(f"{path.name} is password-protected and cannot be read.")
    text = "\n".join((page.extract_text() or "") for page in reader.pages).strip()
    if not text:
        raise BillingError(f"No readable text in {path.name} (it may be a scanned picture).")
    return text


def read_document(path):
    """Return the plain text of a chosen document. Raises BillingError with a friendly message on failure."""
    path = Path(path)
    if not path.is_file():
        raise BillingError(f"File not found: {path}")
    ext = path.suffix.lower()
    try:
        if ext in (".docx", ".docm"):
            return _word(path)
        if ext in (".xlsx", ".xlsm"):
            return _excel(path)
        if ext in (".pptx", ".pptm"):
            return _powerpoint(path)
        if ext == ".pdf":
            return _pdf(path)
        if ext == ".xls":
            return _old_excel(path)
        if ext in (".csv", ".txt"):
            return path.read_text(encoding="utf-8-sig", errors="replace")
    except BillingError:
        raise
    except PermissionError:
        raise BillingError(f"Permission denied opening {path.name}. Is it open in another program?")
    except (OSError, ValueError, KeyError, zipfile.BadZipFile) as e:
        raise BillingError(f"Could not read {path.name}: {e}")
    except Exception as e:  # parser libraries raise many different errors on damaged files
        raise BillingError(f"Could not read {path.name}: {e}")
    if ext in LEGACY:
        raise BillingError(f"{path.name} is an old-style {LEGACY[ext]} file. Open it and use File > Save As to "
                           "save it in the newer format (.docx / .xlsx / .pptx), then choose that file.")
    raise BillingError(f"Unsupported file type '{ext}'. Supported: Word, Excel, PowerPoint, PDF.")
