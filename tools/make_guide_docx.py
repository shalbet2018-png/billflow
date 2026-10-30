"""Turn docs/GUIDE.md into a Word document that can be emailed to businesses."""
import re
import sys
from pathlib import Path

from docx import Document

root = Path(__file__).resolve().parent.parent
src = root / "docs/GUIDE.md"
out = Path(sys.argv[1]) if len(sys.argv) > 1 else root / "docs/BILLFLOW-Business-Guide.docx"


def runs(par, text):
    for i, part in enumerate(re.split(r"`([^`]*)`", text)):
        r = par.add_run(part)
        r.font.name = "Consolas" if i % 2 else None


doc = Document()
for line in src.read_text(encoding="utf-8").splitlines():
    if not line.strip():
        continue
    if line.startswith("# "):
        doc.add_heading(line[2:], 0)
    elif line.startswith("## "):
        doc.add_heading(line[3:], 1)
    elif re.match(r"\s*- ", line):
        runs(doc.add_paragraph(style="List Bullet 2" if line.startswith(" ") else "List Bullet"), line.strip()[2:])
    elif re.match(r"\d+\. ", line):
        runs(doc.add_paragraph(style="List Number"), re.sub(r"^\d+\. ", "", line))
    else:
        runs(doc.add_paragraph(), line)
doc.save(out)
print(out)
