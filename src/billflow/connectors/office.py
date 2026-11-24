"""PDFs made by Microsoft Word / Excel themselves, so they look exactly like the Word invoice and Excel statement.

Put `office_pdf` after the outputs whose files it should turn into PDFs. Windows with Microsoft Office only;
on other computers each client is reported, never skipped silently."""
import os
import subprocess
import sys
from pathlib import Path

from ..models import BillingError
from ..registry import output, output_check

_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$src = $env:BILLFLOW_PDF_SRC; $dst = $env:BILLFLOW_PDF_DST
if ($src -match '\.(docx|docm|doc)$') {
  $app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
  try { $d = $app.Documents.Open($src, $false, $true); $d.ExportAsFixedFormat($dst, 17); $d.Close($false) } finally { $app.Quit() }
} else {
  $app = New-Object -ComObject Excel.Application; $app.Visible = $false; $app.DisplayAlerts = $false
  try { $b = $app.Workbooks.Open($src, 0, $true); $b.ExportAsFixedFormat(0, $dst); $b.Close($false) } finally { $app.Quit() }
}
"""

_KINDS = {".docx": "Word", ".docm": "Word", ".doc": "Word", ".xlsx": "Excel", ".xlsm": "Excel", ".xls": "Excel"}


def _wanted(params):
    return [e.lower() for e in params.get("from", [".docx", ".xlsx"])]


def convert_to_pdf(src, dst, timeout=180):
    if sys.platform != "win32":
        raise BillingError("Making PDFs with Word/Excel needs Windows with Microsoft Office")
    env = dict(os.environ, BILLFLOW_PDF_SRC=str(src), BILLFLOW_PDF_DST=str(dst))
    try:
        done = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                               "-Command", _SCRIPT], env=env, capture_output=True, text=True, timeout=timeout,
                              creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.TimeoutExpired) as e:
        raise BillingError(f"PDF of {Path(src).name} failed: {e}")
    if done.returncode != 0 or not Path(dst).exists():
        msg = (done.stderr or done.stdout).strip().splitlines()
        raise BillingError(f"PDF of {Path(src).name} failed ({_KINDS.get(Path(src).suffix.lower(), 'Office')} "
                           f"is needed): {msg[0] if msg else 'no PDF was made'}")
    return Path(dst)


@output_check("office_pdf")
def check_office_pdf(invoice, params, ctx):
    if sys.platform != "win32":
        raise BillingError("Making PDFs with Word/Excel needs Windows with Microsoft Office")


@output("office_pdf")
def office_pdf(invoice, params, ctx):
    made = []
    for f in list(ctx.files):
        p = Path(f)
        if p.suffix.lower() in _wanted(params):
            made.append(convert_to_pdf(p, p.with_suffix(".pdf")))
    return made
