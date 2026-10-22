"""Regenerates "Scraping Sources and Advisors Reference.docx" from SCRAPING-SOURCES.md.

SCRAPING-SOURCES.md is the single source of truth (edited by hand when a website or advisor is
added/removed) - this script just mirrors it into a Word doc, so the two can never drift apart as
long as this is re-run after every edit to the .md (a PostToolUse hook does this automatically,
see .claude/settings.json).

Neither file is read by the scraper itself - both are standalone reference documents.

Usage: python build_scraping_sources_docx.py
"""

import re
import sys
from pathlib import Path

try:
    import docx
    from docx.shared import Pt, Inches, RGBColor
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("SKIPPED: python-docx not installed (pip install python-docx) - "
          "SCRAPING-SOURCES.md was still updated, just not the .docx mirror.", file=sys.stderr)
    sys.exit(0)

HERE = Path(__file__).parent
MD_PATH = HERE / "SCRAPING-SOURCES.md"
DOCX_PATH = HERE / "Scraping Sources and Advisors Reference.docx"

HEADER_FILL = "2E2E2E"
HEADER_TEXT = RGBColor(0xFF, 0xFF, 0xFF)

_INLINE_LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
_INLINE_CODE_RE = re.compile(r"`([^`]+)`")


def unescape_pipes(text: str) -> str:
    return text.replace(r"\|", "|")


def parse_markdown(md_text: str):
    """Minimal parser for this file's own fixed shape: a title, intro paragraphs, '## '/'### '
    headings, and '|'-delimited tables. Deliberately not a general Markdown parser - just enough
    to mirror this one document faithfully."""
    lines = md_text.splitlines()
    blocks = []  # list of ("h1"|"h2"|"p"|"table", content)
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue
        if line.startswith("# "):
            blocks.append(("h1", line[2:].strip()))
            i += 1
        elif line.startswith("## "):
            blocks.append(("h2", line[3:].strip()))
            i += 1
        elif line.strip().startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            header = [c.strip() for c in table_lines[0].strip("|").split("|")]
            rows = []
            for row_line in table_lines[2:]:  # skip header + separator
                cells = []
                cell = ""
                escaped = False
                for ch in row_line.strip("|"):
                    if escaped:
                        cell += ch
                        escaped = False
                    elif ch == "\\":
                        cell += ch
                        escaped = True
                    elif ch == "|":
                        cells.append(cell.strip())
                        cell = ""
                    else:
                        cell += ch
                cells.append(cell.strip())
                rows.append(cells)
            blocks.append(("table", (header, rows)))
        else:
            para_lines = [line]
            i += 1
            while i < len(lines) and lines[i].strip() and not lines[i].startswith(("#", "|")):
                para_lines.append(lines[i].rstrip())
                i += 1
            blocks.append(("p", " ".join(para_lines)))
    return blocks


def set_cell_shading(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)


def add_hyperlink_run(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '1155CC')
    rPr.append(color)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_rich_text(paragraph, text, base_size=Pt(9)):
    """Renders inline [label](url) links and `code` spans; everything else as plain text."""
    text = unescape_pipes(text)
    pos = 0
    tokens = []
    combined = []
    for m in _INLINE_LINK_RE.finditer(text):
        combined.append((m.start(), m.end(), "link", m.group(1), m.group(2)))
    for m in _INLINE_CODE_RE.finditer(text):
        combined.append((m.start(), m.end(), "code", m.group(1), None))
    combined.sort(key=lambda t: t[0])

    pos = 0
    for start, end, kind, content, url in combined:
        if start < pos:
            continue
        if start > pos:
            run = paragraph.add_run(text[pos:start])
            run.font.size = base_size
        if kind == "link":
            add_hyperlink_run(paragraph, url, content)
        else:
            run = paragraph.add_run(content)
            run.font.name = 'Consolas'
            run.font.size = base_size
        pos = end
    if pos < len(text):
        run = paragraph.add_run(text[pos:])
        run.font.size = base_size


def make_table(doc, header, rows):
    n_cols = len(header)
    total_width_in = 9.2  # fits within 0.6in margins on US Letter
    # Give the longest-header columns proportionally more room; simple heuristic based on
    # header label length, good enough for this document's two known tables.
    weights = [max(len(h), 6) for h in header]
    total_weight = sum(weights)
    col_widths = [total_width_in * w / total_weight for w in weights]

    table = doc.add_table(rows=1, cols=n_cols)
    table.style = 'Table Grid'
    table.autofit = False

    hdr_cells = table.rows[0].cells
    for i, h in enumerate(header):
        hdr_cells[i].width = Inches(col_widths[i])
        p = hdr_cells[i].paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.color.rgb = HEADER_TEXT
        run.font.size = Pt(10)
        set_cell_shading(hdr_cells[i], HEADER_FILL)

    for row in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].width = Inches(col_widths[i])
            p = cells[i].paragraphs[0]
            add_rich_text(p, val)
    return table


def build(md_text: str, out_path: Path):
    blocks = parse_markdown(md_text)

    doc = docx.Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)

    section = doc.sections[0]
    section.left_margin = Inches(0.6)
    section.right_margin = Inches(0.6)
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)

    for kind, content in blocks:
        if kind == "h1":
            doc.add_heading(content, level=1)
        elif kind == "h2":
            doc.add_heading(content, level=2)
        elif kind == "p":
            p = doc.add_paragraph()
            add_rich_text(p, content, base_size=Pt(10))
        elif kind == "table":
            header, rows = content
            make_table(doc, header, rows)
            doc.add_paragraph()

    doc.save(str(out_path))


if __name__ == "__main__":
    if not MD_PATH.exists():
        print(f"SKIPPED: {MD_PATH} not found.", file=sys.stderr)
        sys.exit(0)
    build(MD_PATH.read_text(encoding="utf-8"), DOCX_PATH)
    print(f"Regenerated: {DOCX_PATH}")
