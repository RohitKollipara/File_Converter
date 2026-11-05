"""
FileShift - File Converter Backend  (v3)
High-fidelity conversion engine — preserves tables, styles, formatting.

Formats : PDF, DOCX, PPTX, XLSX, CSV, MD, TXT, HTML, PNG, JPG, WEBP
Legacy  : .doc / .ppt / .xls input (needs LibreOffice)
Tools   : merge, split, compress, protect, unlock, OCR   (see GET /tools  or  /docs)

Strategy per route:
  DOCX → PDF   : docx2pdf (MS Word) → LibreOffice → html-chain fallback
  PPTX → PDF   : PowerPoint (Windows) → LibreOffice → text-only fallback
  XLSX/CSV→PDF : HTML table render (landscape for wide sheets)
  DOCX → MD/TXT: python-docx (structure-aware extraction)
  MD   → PDF   : MD→HTML→PDF via weasyprint (or pymupdf html rendering)
  MD   → PPTX  : python-pptx (# / ## headings become slides)
  PDF  → *     : pymupdf (best-in-class text/layout extraction, page images)
  HTML → *     : beautifulsoup4 structure-aware extraction
  IMG  → *     : Pillow / pymupdf

Run: python server.py   →  http://localhost:8080   (set PORT / HOST env vars to change)
"""

import html as html_mod
import contextvars
import csv
import importlib
import importlib.util
import io
import mimetypes
import os
import platform
import re
import secrets
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.parse
import uuid
import zipfile
from datetime import date, datetime
from pathlib import Path
from typing import List, Optional

# pyrefly: ignore [missing-import]
import fitz  # pymupdf
# pyrefly: ignore [missing-import]
import markdown as md_lib
# pyrefly: ignore [missing-import]
from bs4 import BeautifulSoup
# pyrefly: ignore [missing-import]
from docx import Document
# pyrefly: ignore [missing-import]
from docx.shared import Pt, RGBColor, Inches
# pyrefly: ignore [missing-import]
from docx.oxml.ns import qn
# pyrefly: ignore [missing-import]
from docx.oxml import OxmlElement
# pyrefly: ignore [missing-import]
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
# pyrefly: ignore [missing-import]
from fastapi.middleware.cors import CORSMiddleware
# pyrefly: ignore [missing-import]
from fastapi.responses import FileResponse, HTMLResponse
# pyrefly: ignore [missing-import]
import uvicorn
# pyrefly: ignore [missing-import]
from starlette.background import BackgroundTask
# pyrefly: ignore [missing-import]
from starlette.concurrency import run_in_threadpool

app = FastAPI(title="FileShift")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"],
                   expose_headers=["X-Output-Filename", "X-Conversion-Note", "X-Original-Size", "X-Output-Size",
                                   "X-Pages", "X-Parts", "X-Pages-OCR", "X-Pages-Kept"])

UPLOAD_DIR = Path(tempfile.gettempdir()) / "fileshift"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

CONVERSIONS = {
    "pdf":  ["md", "docx", "txt", "html", "png", "jpg"],
    "md":   ["pdf", "docx", "txt", "html", "pptx"],
    "docx": ["pdf", "md", "txt", "html"],
    "txt":  ["pdf", "md", "docx", "html"],
    "html": ["pdf", "md", "docx", "txt"],
    "pptx": ["pdf", "txt", "md"],
    "xlsx": ["pdf", "csv", "md", "html"],
    "csv":  ["xlsx", "pdf", "md", "html"],
    "png":  ["pdf", "jpg", "webp", "txt"],
    "jpg":  ["pdf", "png", "webp", "txt"],
    "webp": ["pdf", "png", "jpg", "txt"],
    # legacy Office formats (read via LibreOffice)
    "doc":  ["docx", "pdf", "md", "txt", "html"],
    "ppt":  ["pptx", "pdf", "txt", "md"],
    "xls":  ["xlsx", "pdf", "csv", "md", "html"],
}
EXT_ALIASES = {"jpeg": "jpg", "htm": "html"}

# ─── CSS for HTML→PDF rendering ───────────────────────────────────────────────
# NOTE: pymupdf Story has limited CSS support. Rules that work reliably:
#   - font-family, font-size, font-weight, font-style, color
#   - margin, padding, line-height
#   - border (simple), background-color ON BLOCK elements only
#   - width on table/td/th
# Rules that BREAK pymupdf rendering:
#   - thead tr { background } / tbody tr:nth-child → applies to char spans → black boxes
#   - @import (google fonts) → silently ignored
#   - complex selectors
# Solution: inject inline styles on th/td directly in HTML post-processing

PDF_CSS = """
body { font-family: Helvetica, Arial, sans-serif; font-size: 10.5pt; line-height: 1.65; color: #1a1a1a; }
h1 { font-size: 20pt; font-weight: bold; color: #111111; margin-top: 24pt; margin-bottom: 8pt; }
h2 { font-size: 15pt; font-weight: bold; color: #c0392b; margin-top: 18pt; margin-bottom: 6pt; }
h3 { font-size: 12pt; font-weight: bold; color: #2c3e50; margin-top: 14pt; margin-bottom: 4pt; }
h4 { font-size: 11pt; font-weight: bold; color: #2c3e50; margin-top: 10pt; margin-bottom: 3pt; }
h5, h6 { font-size: 10pt; font-weight: bold; margin-top: 8pt; margin-bottom: 2pt; }
p { margin: 0 0 7pt; }
ul { margin: 4pt 0 6pt; padding-left: 18pt; }
ol { margin: 4pt 0 6pt; padding-left: 18pt; }
li { margin-bottom: 3pt; line-height: 1.55; }
/* NOTE: background-color for code/pre/blockquote is intentionally NOT set here.
   pymupdf's Story renderer applies stylesheet-level background-color to individual
   character spans instead of the element block, painting solid dark bars over text.
   These elements get their background-color injected as an INLINE style instead —
   see _apply_table_inline_styles(). Only non-background properties are safe to keep
   at the stylesheet level, as a font/spacing fallback for any element the inliner misses. */
code { font-family: Courier, monospace; font-size: 9pt; color: #c0392b; padding: 1pt 3pt; }
pre { font-family: Courier, monospace; font-size: 8.5pt; color: #333333; padding: 10pt 12pt; margin: 8pt 0; white-space: pre-wrap; }
pre code { color: #333333; padding: 0; }
blockquote { border-left: 3pt solid #c0392b; margin: 6pt 0; padding: 4pt 10pt; color: #555555; font-style: italic; }
strong { font-weight: bold; }
em { font-style: italic; }
hr { border-top: 1pt solid #cccccc; margin: 14pt 0; }
a { color: #2563eb; }
"""

WEASYPRINT_AVAILABLE = False
try:
    
    # pyrefly: ignore [missing-import]
    from weasyprint import HTML as WP_HTML, CSS as WP_CSS
    import warnings as _warnings
    with _warnings.catch_warnings():
        _warnings.simplefilter("ignore")
        _test = io.BytesIO()
        WP_HTML(string='<p>test</p>').write_pdf(_test)
        if _test.tell() > 0:
            WEASYPRINT_AVAILABLE = True
except Exception:
    WEASYPRINT_AVAILABLE = False

# ─── Helpers ──────────────────────────────────────────────────────────────────

def style_name(para) -> str:
    try:
        return (para.style.name or '').lower()
    except Exception:
        return ''

def _apply_table_inline_styles(html: str) -> str:
    """
    Inject inline styles on every element that carries a background-color.

    Root cause of the "black box" bug: pymupdf's Story renderer does not apply
    CSS selector rules (thead tr{background}, code{background}, blockquote{background}, etc.)
    to the element block — it applies them to each individual character span instead,
    which paints a filled rectangle behind every glyph. Depending on stacking/order
    this shows up as solid dark bars over the text (headers, inline code, code
    blocks, and blockquotes were all affected — anything with a stylesheet-level
    background-color).

    Fix: never rely on the <style> stylesheet for background-color. Inject the
    color as an inline `style="..."` attribute directly on each element instead —
    pymupdf (and weasyprint) both apply inline styles correctly as a single block
    fill. Done with BeautifulSoup rather than string replace so nested elements
    (e.g. <code> inside <pre>) are handled correctly and cannot double-fill.
    """
    soup = BeautifulSoup(html, 'html.parser')

    for table in soup.find_all('table'):
        table['style'] = 'width:100%;border-collapse:collapse;margin:10pt 0 14pt;font-size:9.5pt;'
    for thead in soup.find_all('thead'):
        thead['style'] = 'background-color:#2c3e50;color:#ffffff;'
    for th in soup.find_all('th'):
        th['style'] = ('background-color:#2c3e50;color:#ffffff;padding:7pt 10pt;'
                        'text-align:left;font-weight:bold;border:1px solid #2c3e50;')
    # Alternate row coloring must be done via explicit style since nth-child is unsupported
    for td in soup.find_all('td'):
        td['style'] = 'padding:6pt 10pt;border:1px solid #dee2e6;vertical-align:top;'

    # Code blocks (<pre><code>...</code></pre>) — style the <pre>, and make the
    # inner <code> transparent so it never paints a second, misaligned fill.
    for pre in soup.find_all('pre'):
        pre['style'] = ('font-family:Courier, monospace;font-size:8.5pt;background-color:#f4f4f5;'
                         'color:#333333;padding:10pt 12pt;margin:8pt 0;white-space:pre-wrap;')
        for code in pre.find_all('code'):
            code['style'] = 'background-color:transparent;color:#333333;padding:0;font-family:Courier, monospace;'

    # Inline code (not inside a <pre>)
    for code in soup.find_all('code'):
        if code.find_parent('pre') is None:
            code['style'] = ('font-family:Courier, monospace;font-size:9pt;background-color:#f4f4f5;'
                              'color:#c0392b;padding:1pt 3pt;')

    # Blockquotes
    for bq in soup.find_all('blockquote'):
        bq['style'] = ('border-left:3pt solid #c0392b;margin:6pt 0;padding:4pt 10pt;'
                        'background-color:#fdf2f2;color:#555555;font-style:italic;')

    return str(soup)

def md_to_html_string(md_text: str) -> str:
    """Convert markdown to a full HTML document with PDF-safe CSS."""
    # Do NOT use codehilite (adds dark div wrappers that break pymupdf)
    # Do NOT use nl2br (inserts <br> inside table cells, breaking layout)
    # Do NOT use toc (adds anchor IDs that clutter output)
    body = md_lib.markdown(
        md_text,
        extensions=['tables', 'fenced_code', 'sane_lists']
    )
    body = _apply_table_inline_styles(body)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<style>{PDF_CSS}</style>
</head>
<body>{body}</body>
</html>"""

def html_file_to_pdf_weasyprint(html_path: Path, out: Path):
    html_content = html_path.read_text(encoding='utf-8', errors='replace')
    if '<style>' not in html_content:
        html_content = html_content.replace('</head>', f'<style>{PDF_CSS}</style></head>', 1)
    html_content = _apply_table_inline_styles(html_content)
    WP_HTML(string=html_content, base_url=str(html_path.parent)).write_pdf(
        str(out),
        stylesheets=[WP_CSS(string='@page { size: A4; margin: 20mm 22mm; }')]
    )

def html_string_to_pdf(html_str: str, out: Path, base_url: str = None, landscape: bool = False):
    if WEASYPRINT_AVAILABLE:
        page_css = ('@page { size: A4 landscape; margin: 12mm 12mm; }' if landscape
                    else '@page { size: A4; margin: 20mm 22mm; }')
        WP_HTML(string=html_str, base_url=base_url).write_pdf(
            str(out),
            stylesheets=[WP_CSS(string=page_css)]
        )
    else:
        # Fallback: pymupdf Story (fitz) — good table support
        _html_to_pdf_pymupdf(html_str, out, landscape=landscape)

def _html_to_pdf_pymupdf(html_str: str, out: Path, landscape: bool = False):
    """Pymupdf Story renderer — decent HTML→PDF without system libs."""
    try:
        # pyrefly: ignore [missing-import]
        import fitz
        story = fitz.Story(html=html_str, user_css=PDF_CSS)
        writer = fitz.DocumentWriter(str(out))
        mediabox = fitz.paper_rect("a4-l" if landscape else "a4")
        where = mediabox + (36, 36, -36, -36)  # margins
        more = True
        while more:
            dev = writer.begin_page(mediabox)
            more, _ = story.place(where)
            story.draw(dev)
            writer.end_page()
        writer.close()
    except Exception as e:
        raise RuntimeError(f"pymupdf Story rendering failed: {e}")

def docx_to_pdf_word(src: Path, out: Path) -> bool:
    """Try to use MS Word via docx2pdf (Windows/Mac only)."""
    try:
        # pyrefly: ignore [missing-import]
        from docx2pdf import convert
        convert(str(src), str(out))
        return out.exists() and out.stat().st_size > 0
    except Exception:
        return False

def docx_to_pdf_libreoffice(src: Path, out: Path) -> bool:
    """Try LibreOffice headless (Windows/Linux/Mac) — see libreoffice_convert()."""
    return libreoffice_convert(src, out)

def docx_to_html_string(src: Path) -> str:
    """Convert docx to HTML preserving tables, headings, lists, formatting."""
    doc = Document(str(src))
    parts = ['<!DOCTYPE html><html><head><meta charset="utf-8"/>',
             f'<style>{PDF_CSS}</style></head><body>']

    def run_html(run):
        text = html_mod.escape(run.text or '')
        if not text:
            return ''
        if run.bold and run.italic:
            text = f'<strong><em>{text}</em></strong>'
        elif run.bold:
            text = f'<strong>{text}</strong>'
        elif run.italic:
            text = f'<em>{text}</em>'
        if run.underline:
            text = f'<u>{text}</u>'
        try:
            if run.font.name and 'mono' in run.font.name.lower() or \
               run.font.name and 'courier' in run.font.name.lower() or \
               run.font.name and 'code' in run.font.name.lower():
                text = f'<code>{text}</code>'
        except Exception:
            pass
        return text

    # Track table processing
    processed_tables = set()

    def para_html(para):
        sn = style_name(para)
        runs_html = ''.join(run_html(r) for r in para.runs)
        if not runs_html.strip() and not para.text.strip():
            return '<p>&nbsp;</p>'
        if 'heading 1' in sn: return f'<h1>{runs_html}</h1>'
        if 'heading 2' in sn: return f'<h2>{runs_html}</h2>'
        if 'heading 3' in sn: return f'<h3>{runs_html}</h3>'
        if 'heading 4' in sn: return f'<h4>{runs_html}</h4>'
        if 'heading 5' in sn: return f'<h5>{runs_html}</h5>'
        if 'list bullet' in sn: return f'<li>{runs_html}</li>'
        if 'list number' in sn: return f'<li>{runs_html}</li>'
        return f'<p>{runs_html}</p>'

    def table_html(table):
        rows = table.rows
        if not rows:
            return ''
        out = ['<table>']
        for i, row in enumerate(rows):
            out.append('<tr>')
            for cell in row.cells:
                tag = 'th' if i == 0 else 'td'
                cell_text = html_mod.escape(cell.text or '')
                out.append(f'<{tag}>{cell_text}</{tag}>')
            out.append('</tr>')
        out.append('</table>')
        return '\n'.join(out)

    # Walk body elements in order
    # pyrefly: ignore [missing-import]
    from docx.oxml.ns import qn as _qn
    body = doc.element.body
    in_ul = False
    in_ol = False

    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            # Find matching paragraph object
            para = None
            for p in doc.paragraphs:
                if p._element is child:
                    para = p
                    break
            if para is None:
                continue
            sn = style_name(para)
            is_bullet = 'list bullet' in sn
            is_number = 'list number' in sn

            if is_bullet:
                if not in_ul:
                    if in_ol: parts.append('</ol>'); in_ol = False
                    parts.append('<ul>'); in_ul = True
            elif is_number:
                if not in_ol:
                    if in_ul: parts.append('</ul>'); in_ul = False
                    parts.append('<ol>'); in_ol = True
            else:
                if in_ul: parts.append('</ul>'); in_ul = False
                if in_ol: parts.append('</ol>'); in_ol = False

            parts.append(para_html(para))

        elif tag == 'tbl':
            if in_ul: parts.append('</ul>'); in_ul = False
            if in_ol: parts.append('</ol>'); in_ol = False
            # Find matching table
            for tbl in doc.tables:
                if tbl._element is child:
                    parts.append(table_html(tbl))
                    break

    if in_ul: parts.append('</ul>')
    if in_ol: parts.append('</ol>')
    parts.append('</body></html>')
    return '\n'.join(parts)


# ─── CONVERSION FUNCTIONS ──────────────────────────────────────────────────────

# ── PDF → * ───────────────────────────────────────────────────────────────────
def pdf_to_md(src: Path, out: Path):
    doc = fitz.open(str(src))
    pages = []
    for page in doc:
        # Use dict extraction for better structure
        blocks = page.get_text("blocks")
        lines = []
        for b in sorted(blocks, key=lambda b: (b[1], b[0])):  # sort by y then x
            text = b[4].strip()
            if text:
                lines.append(text)
        pages.append('\n\n'.join(lines))
    note = None
    if not any(p.strip() for p in pages):                      # no text layer → scanned PDF → OCR
        pages = [p.strip() for p in _pdf_ocr_fallback_pages(doc)]
        note = 'Scanned PDF detected - text was recognised with OCR. Check for mistakes.'
    doc.close()
    out.write_text('\n\n---\n\n'.join(pages), encoding='utf-8')
    return Result(note=note) if note else None

def pdf_to_txt(src: Path, out: Path):
    doc = fitz.open(str(src))
    pages = [page.get_text('text') for page in doc]
    note = None
    if not any(p.strip() for p in pages):                      # no text layer → scanned PDF → OCR
        pages = _pdf_ocr_fallback_pages(doc)
        note = 'Scanned PDF detected - text was recognised with OCR. Check for mistakes.'
    doc.close()
    out.write_text('\n\n'.join(pages), encoding='utf-8')
    return Result(note=note) if note else None

def pdf_to_docx(src: Path, out: Path):
    probe = fitz.open(str(src))
    try:
        has_text = any(_page_has_text(p, 1) for p in probe)
        pages = None if has_text else _pdf_ocr_fallback_pages(probe)   # scanned → OCR text
    finally:
        probe.close()
    if pages is not None:
        d = Document()
        for i, text in enumerate(pages):
            if i:
                d.add_page_break()
            for para in re.split(r'\n\s*\n', text.strip()):
                if para.strip():
                    d.add_paragraph(para.strip())
        d.save(str(out))
        return Result(note='Scanned PDF detected - text was recognised with OCR into a plain Word document '
                           '(original layout is not preserved).')
    # pyrefly: ignore [missing-import]
    from pdf2docx import Converter
    cv = Converter(str(src))
    cv.convert(str(out))
    cv.close()


# ── MD → * ────────────────────────────────────────────────────────────────────
def md_to_pdf(src: Path, out: Path):
    md_text = src.read_text(encoding='utf-8', errors='replace')
    html_str = md_to_html_string(md_text)
    html_string_to_pdf(html_str, out)

def md_to_docx(src: Path, out: Path):
    md_text = src.read_text(encoding='utf-8', errors='replace')
    # Convert MD→HTML→parse→docx for proper table and formatting support
    body_html = md_lib.markdown(md_text, extensions=['tables', 'fenced_code', 'sane_lists', 'nl2br'])
    soup = BeautifulSoup(body_html, 'html.parser')
    doc = Document()

    # Set default font

    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    def add_run_formatted(para, text, bold=False, italic=False, code=False):
        run = para.add_run(text)
        run.bold = bold
        run.italic = italic
        if code:
            run.font.name = 'Courier New'
            run.font.size = Pt(9)

    def process_inline(para, element):
        """Process inline elements (bold, italic, code, text)."""
        if isinstance(element, str):
            if element:
                para.add_run(element)
        elif element.name == 'strong':
            run = para.add_run(element.get_text())
            run.bold = True
        elif element.name == 'em':
            run = para.add_run(element.get_text())
            run.italic = True
        elif element.name == 'code':
            run = para.add_run(element.get_text())
            run.font.name = 'Courier New'
            run.font.size = Pt(9)
        elif element.name == 'a':
            para.add_run(element.get_text())
        else:
            for child in element.children:
                process_inline(para, child)

    def process_element(el):
        if isinstance(el, str):
            return

        if el.name in ('h1','h2','h3','h4','h5','h6'):
            level = int(el.name[1])
            doc.add_heading(el.get_text(), level=level)

        elif el.name == 'p':
            para = doc.add_paragraph()
            for child in el.children:
                process_inline(para, child)

        elif el.name == 'ul':
            for li in el.find_all('li', recursive=False):
                para = doc.add_paragraph(style='List Bullet')
                for child in li.children:
                    if hasattr(child, 'name') and child.name in ('ul','ol'):
                        continue  # skip nested for now
                    process_inline(para, child)

        elif el.name == 'ol':
            for li in el.find_all('li', recursive=False):
                para = doc.add_paragraph(style='List Number')
                for child in li.children:
                    if hasattr(child, 'name') and child.name in ('ul','ol'):
                        continue
                    process_inline(para, child)

        elif el.name == 'pre':
            code_text = el.get_text()
            para = doc.add_paragraph()
            run = para.add_run(code_text)
            run.font.name = 'Courier New'
            run.font.size = Pt(9)
            pPr = para._p.get_or_add_pPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), 'F4F4F5')
            pPr.append(shd)

        elif el.name == 'blockquote':
            para = doc.add_paragraph(el.get_text().strip())
            para.style = doc.styles['Normal']
            run = para.runs[0] if para.runs else para.add_run('')
            run.italic = True

        elif el.name == 'hr':
            para = doc.add_paragraph()
            pPr = para._p.get_or_add_pPr()
            pBdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '6')
            bottom.set(qn('w:space'), '1')
            bottom.set(qn('w:color'), 'AAAAAA')
            pBdr.append(bottom)
            pPr.append(pBdr)

        elif el.name == 'table':
            rows = el.find_all('tr')
            if not rows:
                return
            max_cols = max(len(r.find_all(['td','th'])) for r in rows)
            if max_cols == 0:
                return
            table = doc.add_table(rows=len(rows), cols=max_cols)
            table.style = 'Table Grid'
            for r_idx, row in enumerate(rows):
                cells = row.find_all(['td','th'])
                for c_idx, cell in enumerate(cells):
                    if c_idx >= max_cols:
                        break
                    tc = table.rows[r_idx].cells[c_idx]
                    tc.text = cell.get_text(strip=True)
                    # Bold header row
                    if cell.name == 'th' or r_idx == 0:
                        for run in tc.paragraphs[0].runs:
                            run.bold = True
                        # pyrefly: ignore [missing-import]
                        from docx.oxml import OxmlElement as OE
                        tcPr = tc._tc.get_or_add_tcPr()
                        shd = OE('w:shd')
                        shd.set(qn('w:val'), 'clear')
                        shd.set(qn('w:color'), 'auto')
                        shd.set(qn('w:fill'), '2C3E50')
                        tcPr.append(shd)
                        for para in tc.paragraphs:
                            for run in para.runs:
                                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for element in soup.children:
        process_element(element)

    doc.save(str(out))

def md_to_html(src: Path, out: Path):
    md_text = src.read_text(encoding='utf-8', errors='replace')
    out.write_text(md_to_html_string(md_text), encoding='utf-8')

def md_to_txt(src: Path, out: Path):
    out.write_text(src.read_text(encoding='utf-8', errors='replace'), encoding='utf-8')


# ── DOCX → * ──────────────────────────────────────────────────────────────────
def docx_to_pdf(src: Path, out: Path):
    # Try MS Word first (Windows) → LibreOffice → HTML chain fallback
    if docx_to_pdf_word(src, out):
        return
    if docx_to_pdf_libreoffice(src, out):
        return
    # HTML chain fallback
    html_str = docx_to_html_string(src)
    html_string_to_pdf(html_str, out)

def docx_to_md(src: Path, out: Path):
    doc = Document(str(src))
    lines = []
    for child in doc.element.body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            para = next((p for p in doc.paragraphs if p._element is child), None)
            if para is None: continue
            sn = style_name(para)
            text = para.text or ''
            if not text.strip(): lines.append(''); continue
            if 'heading 1' in sn: lines.append(f'# {text}')
            elif 'heading 2' in sn: lines.append(f'## {text}')
            elif 'heading 3' in sn: lines.append(f'### {text}')
            elif 'heading 4' in sn: lines.append(f'#### {text}')
            elif 'list bullet' in sn: lines.append(f'- {text}')
            elif 'list number' in sn: lines.append(f'1. {text}')
            else: lines.append(text)
        elif tag == 'tbl':
            tbl = next((t for t in doc.tables if t._element is child), None)
            if tbl is None: continue
            rows = tbl.rows
            if not rows: continue
            header = '| ' + ' | '.join(_md_cell(c.text) for c in rows[0].cells) + ' |'
            sep    = '| ' + ' | '.join('---' for _ in rows[0].cells) + ' |'
            lines.append(header)
            lines.append(sep)
            for row in rows[1:]:
                lines.append('| ' + ' | '.join(_md_cell(c.text) for c in row.cells) + ' |')
            lines.append('')
    out.write_text('\n'.join(lines), encoding='utf-8')

def docx_to_txt(src: Path, out: Path):
    doc = Document(str(src))
    lines = []
    for para in doc.paragraphs:
        lines.append(para.text)
    # Also extract table text
    for table in doc.tables:
        for row in table.rows:
            lines.append('\t'.join(c.text for c in row.cells))
    out.write_text('\n'.join(lines), encoding='utf-8')


# ── TXT → * ───────────────────────────────────────────────────────────────────
def txt_to_pdf(src: Path, out: Path):
    text = src.read_text(encoding='utf-8', errors='replace')
    # Wrap as preformatted HTML for clean rendering
    html_str = f"""<!DOCTYPE html><html><head><meta charset="utf-8"/>
<style>{PDF_CSS} body{{white-space:pre-wrap;font-family:'JetBrains Mono','Courier New',monospace;font-size:9.5pt;}}</style>
</head><body>{html_mod.escape(text)}</body></html>"""
    html_string_to_pdf(html_str, out)

def txt_to_md(src: Path, out: Path):
    out.write_text(src.read_text(encoding='utf-8', errors='replace'), encoding='utf-8')

def txt_to_docx(src: Path, out: Path):
    text = src.read_text(encoding='utf-8', errors='replace')
    doc = Document()
    for line in text.splitlines():
        doc.add_paragraph(line)
    doc.save(str(out))


# ── HTML → * ──────────────────────────────────────────────────────────────────
def html_to_pdf(src: Path, out: Path):
    if WEASYPRINT_AVAILABLE:
        html_file_to_pdf_weasyprint(src, out)
    else:
        _html_to_pdf_pymupdf(src.read_text(encoding='utf-8', errors='replace'), out)

def html_to_md(src: Path, out: Path):
    soup = BeautifulSoup(src.read_text(encoding='utf-8', errors='replace'), 'html.parser')
    lines = []
    for el in soup.find_all(['h1','h2','h3','h4','h5','h6','p','li','pre','hr','blockquote','table']):
        n = el.name
        text = el.get_text(' ', strip=True)
        if   n == 'h1': lines += [f'# {text}', '']
        elif n == 'h2': lines += [f'## {text}', '']
        elif n == 'h3': lines += [f'### {text}', '']
        elif n == 'h4': lines += [f'#### {text}', '']
        elif n == 'h5': lines += [f'##### {text}', '']
        elif n == 'li': lines.append(f'- {text}')
        elif n == 'pre': lines += [f'```\n{el.get_text()}\n```', '']
        elif n == 'blockquote': lines += [f'> {text}', '']
        elif n == 'hr': lines += ['---', '']
        elif n == 'table':
            rows = el.find_all('tr')
            if rows:
                header = '| ' + ' | '.join(c.get_text(strip=True) for c in rows[0].find_all(['th','td'])) + ' |'
                sep    = '| ' + ' | '.join('---' for _ in rows[0].find_all(['th','td'])) + ' |'
                lines += [header, sep]
                for row in rows[1:]:
                    lines.append('| ' + ' | '.join(c.get_text(strip=True) for c in row.find_all('td')) + ' |')
                lines.append('')
        elif text: lines += [text, '']
    out.write_text('\n'.join(lines), encoding='utf-8')

def html_to_docx(src: Path, out: Path):
    # Reuse md_to_docx logic via html soup
    html_content = src.read_text(encoding='utf-8', errors='replace')
    soup = BeautifulSoup(html_content, 'html.parser')
    body = soup.find('body') or soup
    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(11)

    for el in body.children:
        if not hasattr(el, 'name') or el.name is None:
            continue
        n = el.name
        text = el.get_text(strip=True)
        if   n == 'h1': doc.add_heading(text, 1)
        elif n == 'h2': doc.add_heading(text, 2)
        elif n == 'h3': doc.add_heading(text, 3)
        elif n == 'h4': doc.add_heading(text, 4)
        elif n == 'p':  doc.add_paragraph(text)
        elif n == 'ul':
            for li in el.find_all('li', recursive=False):
                doc.add_paragraph(li.get_text(strip=True), style='List Bullet')
        elif n == 'ol':
            for li in el.find_all('li', recursive=False):
                doc.add_paragraph(li.get_text(strip=True), style='List Number')
        elif n == 'pre':
            p = doc.add_paragraph()
            run = p.add_run(el.get_text())
            run.font.name = 'Courier New'
            run.font.size = Pt(9)
        elif n == 'table':
            rows = el.find_all('tr')
            if not rows: continue
            max_cols = max(len(r.find_all(['td','th'])) for r in rows)
            if max_cols == 0: continue
            table = doc.add_table(rows=len(rows), cols=max_cols)
            table.style = 'Table Grid'
            for r_i, row in enumerate(rows):
                cells = row.find_all(['td','th'])
                for c_i, cell in enumerate(cells):
                    if c_i >= max_cols: break
                    table.rows[r_i].cells[c_i].text = cell.get_text(strip=True)
    doc.save(str(out))


# ═══════════════════════════════════════════════════════════════════════════════
#  NEW IN v2: shared utilities, PPTX / spreadsheet / image converters, PDF merge
# ═══════════════════════════════════════════════════════════════════════════════

class Result:
    """Optional return value for converters that need to change the output file
    (e.g. a multi-page PDF → images produces a .zip) or attach a warning note."""
    def __init__(self, path: Optional[Path] = None, note: Optional[str] = None, headers: Optional[dict] = None):
        self.path = path
        self.note = note
        self.headers = headers or {}


def _need(module: str, pip_name: str):
    """Import an optional dependency lazily with a friendly error message."""
    try:
        return importlib.import_module(module)
    except ImportError:
        raise RuntimeError(f"Missing dependency '{pip_name}'. Run:  pip install {pip_name}")


def _md_cell(text) -> str:
    return (str(text) if text is not None else '').replace('|', '\\|').replace('\r', ' ').replace('\n', ' ').strip()


def _safe_name(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', '_', name).strip(' .')
    return name or 'file'


def _zip_files(files: List[Path], zip_path: Path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, f.name)


def _strip_md_inline(text: str) -> str:
    text = re.sub(r'!\[[^\]]*\]\([^)]*\)', '', text)          # images
    text = re.sub(r'\[([^\]]+)\]\([^)]*\)', r'\1', text)        # links
    text = re.sub(r'(\*\*|__)(.+?)\1', r'\2', text)             # bold
    text = re.sub(r'(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])', r'\1', text)  # italic
    text = re.sub(r'`([^`]*)`', r'\1', text)                    # inline code
    return text.strip()


# ── LibreOffice (free, cross-platform) ────────────────────────────────────────
_LO_LOCK = threading.Lock()

def find_libreoffice() -> Optional[str]:
    """Find soffice even when it is not on PATH (the default on Windows/macOS)."""
    for name in ('soffice', 'libreoffice', 'soffice.exe'):
        p = shutil.which(name)
        if p:
            return p
    candidates = [
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        '/usr/bin/soffice', '/usr/bin/libreoffice',
        '/usr/local/bin/soffice', '/snap/bin/libreoffice',
    ]
    return next((c for c in candidates if os.path.exists(c)), None)


def libreoffice_convert(src: Path, out: Path) -> bool:
    """Convert any LibreOffice-readable file (docx/pptx/xlsx/...) to PDF.
    Uses a private profile dir so it still works if a LibreOffice window is
    already open, and a lock so concurrent requests don't collide."""
    lo = find_libreoffice()
    if not lo:
        return False
    profile = (UPLOAD_DIR / '_lo_profile').resolve()
    outdir = out.parent / '_lo_out'
    outdir.mkdir(exist_ok=True)
    try:
        with _LO_LOCK:
            subprocess.run(
                [lo, f'-env:UserInstallation={profile.as_uri()}', '--headless', '--norestore',
                 '--convert-to', 'pdf', '--outdir', str(outdir), str(src)],
                capture_output=True, timeout=180
            )
        produced = outdir / (src.stem + '.pdf')
        if produced.exists() and produced.stat().st_size > 0:
            shutil.move(str(produced), str(out))
            return True
        return False
    except Exception:
        return False
    finally:
        shutil.rmtree(outdir, ignore_errors=True)


# ── Spreadsheets (XLSX / CSV) ─────────────────────────────────────────────────
def _cell_str(v) -> str:
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime):
        return v.strftime('%Y-%m-%d') if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime('%Y-%m-%d %H:%M:%S')
    if isinstance(v, date):
        return v.strftime('%Y-%m-%d')
    return str(v)


def _trim_rows(rows: List[List[str]]) -> List[List[str]]:
    while rows and not any(c != '' for c in rows[-1]):
        rows.pop()
    if not rows:
        return rows
    width = max((max((i for i, c in enumerate(r) if c != ''), default=-1) + 1) for r in rows)
    return [(r[:width] + [''] * (width - len(r[:width]))) for r in rows]


def read_tables(src: Path):
    """Return [(sheet_name, rows)] for .csv or .xlsx. Every cell is a string."""
    ext = src.suffix.lower()
    if ext == '.csv':
        raw = src.read_bytes()
        if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
            text = raw.decode('utf-16')
        else:
            text = None
            for enc in ('utf-8-sig', 'cp1252', 'latin-1'):
                try:
                    text = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
        csv.field_size_limit(10 ** 7)
        try:
            dialect = csv.Sniffer().sniff(text[:4096], delimiters=',;\t|')
        except csv.Error:
            dialect = csv.excel
        rows = [list(r) for r in csv.reader(io.StringIO(text), dialect)]
        return [('Sheet1', _trim_rows(rows))]

    openpyxl = _need('openpyxl', 'openpyxl')
    wb = openpyxl.load_workbook(str(src), data_only=True, read_only=True)
    tables = []
    try:
        for ws in wb.worksheets:
            if getattr(ws, 'sheet_state', 'visible') != 'visible':
                continue
            rows = [[_cell_str(c) for c in row] for row in ws.iter_rows(values_only=True)]
            tables.append((ws.title, _trim_rows(rows)))
    finally:
        wb.close()
    return tables or [('Sheet1', [])]


def tables_to_md(tables) -> str:
    out = []
    for name, rows in tables:
        if len(tables) > 1:
            out += [f'## {name}', '']
        if not rows:
            out += ['*(empty sheet)*', '']
            continue
        width = max(len(r) for r in rows)
        rows = [r + [''] * (width - len(r)) for r in rows]
        out.append('| ' + ' | '.join(_md_cell(c) for c in rows[0]) + ' |')
        out.append('| ' + ' | '.join('---' for _ in rows[0]) + ' |')
        for r in rows[1:]:
            out.append('| ' + ' | '.join(_md_cell(c) for c in r) + ' |')
        out.append('')
    return '\n'.join(out)


def tables_to_html_string(tables, title: str = 'Spreadsheet') -> str:
    max_cols = max((len(r) for _, rows in tables for r in rows), default=0)
    fs = '9.5pt' if max_cols <= 6 else '8pt' if max_cols <= 10 else '6.5pt'
    th = ('background-color:#2c3e50;color:#ffffff;padding:4pt 6pt;text-align:left;'
          'font-weight:bold;border:1px solid #2c3e50;')
    parts = [f'<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"/>'
             f'<title>{html_mod.escape(title)}</title><style>{PDF_CSS}</style></head><body>']
    for name, rows in tables:
        if len(tables) > 1:
            parts.append(f'<h2>{html_mod.escape(name)}</h2>')
        if not rows:
            parts.append('<p><em>(empty sheet)</em></p>')
            continue
        width = max(len(r) for r in rows)
        parts.append(f'<table style="width:100%;border-collapse:collapse;font-size:{fs};margin:6pt 0 14pt;">')
        for i, row in enumerate(rows):
            row = row + [''] * (width - len(row))
            if i == 0:
                cells = ''.join(f'<th style="{th}">{html_mod.escape(c)}</th>' for c in row)
            else:
                bg = '#f8f9fa' if i % 2 == 0 else '#ffffff'
                td = (f'padding:3pt 6pt;border:1px solid #dee2e6;vertical-align:top;'
                      f'background-color:{bg};')
                cells = ''.join(f'<td style="{td}">{html_mod.escape(c)}</td>' for c in row)
            parts.append(f'<tr>{cells}</tr>')
        parts.append('</table>')
    parts.append('</body></html>')
    return '\n'.join(parts)


def sheet_to_pdf(src: Path, out: Path):
    tables = read_tables(src)
    max_cols = max((len(r) for _, rows in tables for r in rows), default=0)
    html_string_to_pdf(tables_to_html_string(tables, src.stem), out, landscape=max_cols > 6)

def sheet_to_md(src: Path, out: Path):
    out.write_text(tables_to_md(read_tables(src)), encoding='utf-8')

def sheet_to_html(src: Path, out: Path):
    out.write_text(tables_to_html_string(read_tables(src), src.stem), encoding='utf-8')

def xlsx_to_csv(src: Path, out: Path):
    tables = read_tables(src)
    if len(tables) == 1:
        with open(out, 'w', newline='', encoding='utf-8-sig') as f:   # BOM so Excel opens UTF-8 correctly
            csv.writer(f).writerows(tables[0][1])
        return
    files = []
    for name, rows in tables:
        p = out.parent / f"{src.stem}_{_safe_name(name)}.csv"
        with open(p, 'w', newline='', encoding='utf-8-sig') as f:
            csv.writer(f).writerows(rows)
        files.append(p)
    zip_path = out.with_suffix('.zip')
    _zip_files(files, zip_path)
    return Result(path=zip_path, note=f'Workbook has {len(files)} sheets - one CSV per sheet, delivered as a .zip')

def _infer_value(v: str):
    s = v.strip()
    if re.fullmatch(r'-?(0|[1-9]\d{0,14})', s):
        return int(s)
    if re.fullmatch(r'-?(0|[1-9]\d*)\.\d+', s):
        return float(s)
    return v

def csv_to_xlsx(src: Path, out: Path):
    openpyxl = _need('openpyxl', 'openpyxl')
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    rows = read_tables(src)[0][1]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = _safe_name(src.stem)[:31] or 'Sheet1'
    for r in rows:
        ws.append([_infer_value(c) if c != '' else None for c in r])
    if rows:
        for c in ws[1]:
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='2C3E50')
        ws.freeze_panes = 'A2'
        for i, col in enumerate(zip(*rows), 1):
            ws.column_dimensions[get_column_letter(i)].width = min(max(len(str(v)) for v in col) + 2, 50)
    wb.save(str(out))


# ── Images (PNG / JPG / WEBP) and PDF → images ────────────────────────────────
IMG_EXTS = ('png', 'jpg', 'jpeg', 'webp')

def _load_image(src: Path):
    _need('PIL', 'Pillow')
    from PIL import Image, ImageOps
    img = Image.open(str(src))
    return ImageOps.exif_transpose(img)          # honour phone-camera rotation

def _flatten_rgb(img):
    from PIL import Image
    if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
        rgba = img.convert('RGBA')
        bg = Image.new('RGB', rgba.size, (255, 255, 255))
        bg.paste(rgba, mask=rgba.split()[3])
        return bg
    return img.convert('RGB')

def _image_to_pdf_page(doc, src: Path):
    """Append one A4 page (portrait/landscape to match the image) holding src."""
    rgb = _flatten_rgb(_load_image(src))
    buf = io.BytesIO()
    if src.suffix.lower() == '.png':
        rgb.save(buf, 'PNG')
    else:
        rgb.save(buf, 'JPEG', quality=92)
    w, h = rgb.size
    pw, ph = (842, 595) if w > h else (595, 842)
    page = doc.new_page(width=pw, height=ph)
    m = 24
    page.insert_image(fitz.Rect(m, m, pw - m, ph - m), stream=buf.getvalue(), keep_proportion=True)

def img_to_pdf(src: Path, out: Path):
    doc = fitz.open()
    try:
        _image_to_pdf_page(doc, src)
        doc.save(str(out), garbage=3, deflate=True)
    finally:
        doc.close()

def make_img_converter(fmt: str):
    def conv(src: Path, out: Path):
        img = _load_image(src)
        if fmt == 'jpg':
            _flatten_rgb(img).save(str(out), 'JPEG', quality=92, optimize=True)
            return
        if img.mode not in ('RGB', 'RGBA', 'L', 'LA', 'P'):
            img = img.convert('RGBA' if 'A' in img.mode else 'RGB')
        if fmt == 'png':
            img.save(str(out), 'PNG')
        elif fmt == 'webp':
            img.save(str(out), 'WEBP', quality=90)
    return conv

def make_pdf_to_images(fmt: str):
    def conv(src: Path, out: Path):
        doc = fitz.open(str(src))
        try:
            if doc.needs_pass:
                raise ValueError('PDF is password protected')
            mat = fitz.Matrix(150 / 72, 150 / 72)              # 150 DPI
            enc = 'png' if fmt == 'png' else 'jpeg'
            if len(doc) == 1:
                out.write_bytes(doc[0].get_pixmap(matrix=mat, alpha=False).tobytes(enc))
                return
            pages_dir = out.parent / '_pages'
            pages_dir.mkdir(exist_ok=True)
            files = []
            for i, page in enumerate(doc, 1):
                p = pages_dir / f"{src.stem}_page{i:03d}.{fmt}"
                p.write_bytes(page.get_pixmap(matrix=mat, alpha=False).tobytes(enc))
                files.append(p)
            zip_path = out.with_suffix('.zip')
            _zip_files(files, zip_path)
            return Result(path=zip_path, note=f'{len(files)} pages - one image per page, delivered as a .zip')
        finally:
            doc.close()
    return conv


# ── PDF → HTML, DOCX → HTML, TXT → HTML, HTML → TXT ───────────────────────────
def pdf_to_html(src: Path, out: Path):
    doc = fitz.open(str(src))
    try:
        body = '\n'.join(page.get_text('xhtml') for page in doc)
    finally:
        doc.close()
    out.write_text(f'<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"/>'
                   f'<title>{html_mod.escape(src.stem)}</title>'
                   f'<style>body{{font-family:Helvetica,Arial,sans-serif;max-width:820px;margin:2rem auto;'
                   f'line-height:1.6;padding:0 1rem}}img{{max-width:100%}}</style></head>'
                   f'<body>{body}</body></html>', encoding='utf-8')

def docx_to_html(src: Path, out: Path):
    out.write_text(docx_to_html_string(src), encoding='utf-8')

def txt_to_html(src: Path, out: Path):
    text = src.read_text(encoding='utf-8', errors='replace')
    paras = [p for p in re.split(r'\n\s*\n', text.strip()) if p.strip()]
    body = '\n'.join('<p>' + html_mod.escape(p).replace('\n', '<br>\n') + '</p>' for p in paras)
    out.write_text(f'<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"/>'
                   f'<title>{html_mod.escape(src.stem)}</title><style>{PDF_CSS}</style></head>'
                   f'<body>{body}</body></html>', encoding='utf-8')

def html_to_txt(src: Path, out: Path):
    soup = BeautifulSoup(src.read_text(encoding='utf-8', errors='replace'), 'html.parser')
    for tag in soup(['script', 'style', 'head']):
        tag.decompose()
    text = soup.get_text('\n')
    text = re.sub(r'[ \t]+\n', '\n', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    out.write_text(text.strip() + '\n', encoding='utf-8')


# ── PPTX → * ──────────────────────────────────────────────────────────────────
_PPTX_SKIP_PLACEHOLDERS = {13, 15, 16}      # slide number, footer, date

def _pptx_iter_shapes(shapes):
    for sh in sorted(shapes, key=lambda s: ((s.top or 0), (s.left or 0))):
        if hasattr(sh, 'shapes'):               # group shape → recurse
            yield from _pptx_iter_shapes(sh.shapes)
        else:
            yield sh

def pptx_read(src: Path):
    """Return [{'n', 'title', 'items': [('p', text, level) | ('t', rows)], 'notes'}]"""
    pptx = _need('pptx', 'python-pptx')
    prs = pptx.Presentation(str(src))
    slides = []
    for n, slide in enumerate(prs.slides, 1):
        title_shape = None
        try:
            title_shape = slide.shapes.title
        except Exception:
            pass
        title = (title_shape.text_frame.text.strip().replace('\v', ' ') if title_shape is not None else '')
        items = []
        for sh in _pptx_iter_shapes(slide.shapes):
            try:
                if title_shape is not None and sh.shape_id == title_shape.shape_id:
                    continue
                if sh.is_placeholder and int(sh.placeholder_format.type) in _PPTX_SKIP_PLACEHOLDERS:
                    continue
                if getattr(sh, 'has_table', False) and sh.has_table:
                    rows = [[c.text.strip().replace('\n', ' ') for c in r.cells] for r in sh.table.rows]
                    items.append(('t', rows))
                elif getattr(sh, 'has_text_frame', False) and sh.has_text_frame:
                    for p in sh.text_frame.paragraphs:
                        text = p.text.replace('\v', ' ').strip()
                        if text:
                            items.append(('p', text, p.level))
            except Exception:
                continue
        notes = ''
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
                notes = slide.notes_slide.notes_text_frame.text.strip()
        except Exception:
            pass
        slides.append({'n': n, 'title': title, 'items': items, 'notes': notes})
    return slides

def pptx_to_txt(src: Path, out: Path):
    lines = []
    for s in pptx_read(src):
        lines.append(f"=== Slide {s['n']}" + (f": {s['title']}" if s['title'] else '') + ' ===')
        for it in s['items']:
            if it[0] == 'p':
                lines.append('    ' * it[2] + it[1])
            else:
                lines += ['\t'.join(r) for r in it[1]]
        if s['notes']:
            lines += ['', f"Notes: {s['notes']}"]
        lines.append('')
    out.write_text('\n'.join(lines), encoding='utf-8')

def pptx_to_md(src: Path, out: Path):
    lines = []
    for s in pptx_read(src):
        lines += [f"## {s['title'] or 'Slide ' + str(s['n'])}", '']
        for it in s['items']:
            if it[0] == 'p':
                lines.append('  ' * it[2] + '- ' + it[1])
            else:
                rows = it[1]
                if not rows:
                    continue
                w = max(len(r) for r in rows)
                rows = [r + [''] * (w - len(r)) for r in rows]
                lines += ['', '| ' + ' | '.join(_md_cell(c) for c in rows[0]) + ' |',
                          '| ' + ' | '.join('---' for _ in rows[0]) + ' |']
                lines += ['| ' + ' | '.join(_md_cell(c) for c in r) + ' |' for r in rows[1:]]
                lines.append('')
        if s['notes']:
            lines += ['', '> **Notes:** ' + s['notes'].replace('\n', ' ')]
        lines += ['', '---', '']
    out.write_text('\n'.join(lines), encoding='utf-8')

def pptx_to_html_string(src: Path) -> str:
    esc = html_mod.escape
    parts = [f'<!DOCTYPE html><html><head><meta charset="utf-8"/><style>{PDF_CSS}</style></head><body>']
    for i, s in enumerate(pptx_read(src)):
        brk = ' style="page-break-before:always"' if i else ''
        parts.append(f'<div{brk}><p style="color:#999999;font-size:8pt;">Slide {s["n"]}</p>')
        parts.append(f'<h1>{esc(s["title"] or "Slide " + str(s["n"]))}</h1>')
        for it in s['items']:
            if it[0] == 'p':
                parts.append(f'<p style="margin-left:{it[2] * 18}pt">&#8226; {esc(it[1])}</p>')
            else:
                parts.append(tables_to_html_string([('', it[1])]).split('<body>')[1].split('</body>')[0])
        if s['notes']:
            parts.append(f'<blockquote>Notes: {esc(s["notes"])}</blockquote>')
        parts.append('</div>')
    parts.append('</body></html>')
    return '\n'.join(parts)

def pptx_to_pdf_powerpoint(src: Path, out: Path) -> bool:
    """Best-effort: drive Microsoft PowerPoint through COM (Windows only)."""
    if platform.system() != 'Windows':
        return False
    try:
        import pythoncom                      # pyrefly: ignore [missing-import]
        import win32com.client                # pyrefly: ignore [missing-import]
    except ImportError:
        return False
    pythoncom.CoInitialize()
    app = pres = None
    started_here = False
    try:
        app = win32com.client.Dispatch('PowerPoint.Application')
        started_here = app.Presentations.Count == 0
        pres = app.Presentations.Open(str(src.resolve()), True, False, False)  # read-only, no window
        pres.SaveAs(str(out.resolve()), 32)                                    # 32 = ppSaveAsPDF
        pres.Close()
        pres = None
        return out.exists() and out.stat().st_size > 0
    except Exception:
        return False
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if app is not None and started_here:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

def pptx_to_pdf(src: Path, out: Path):
    # 1) PowerPoint (Windows)  2) LibreOffice (free, any OS)  3) text-only fallback
    if pptx_to_pdf_powerpoint(src, out):
        return
    if libreoffice_convert(src, out):
        return
    html_string_to_pdf(pptx_to_html_string(src), out)
    return Result(note='LibreOffice/PowerPoint not found - made a simplified text-only PDF (layout, images and '
                       'styling are not preserved). Install LibreOffice (free) for an exact copy of your slides.')


# ── MD → PPTX ─────────────────────────────────────────────────────────────────
def md_to_pptx(src: Path, out: Path):
    """# first heading → title slide; each # / ## heading → a new slide;
    bullets, paragraphs, tables and code blocks become slide content."""
    _need('pptx', 'python-pptx')
    from pptx import Presentation
    from pptx.util import Inches, Pt

    lines = src.read_text(encoding='utf-8', errors='replace').splitlines()
    slides = []            # {'kind': 'title'|'content', 'title', 'blocks': [...]}
    cur = None
    i = 0
    while i < len(lines):
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()
        i += 1
        if not stripped or re.fullmatch(r'(-{3,}|\*{3,}|_{3,})', stripped):
            continue
        if stripped.startswith('```'):
            code = []
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code.append(lines[i].rstrip())
                i += 1
            i += 1
            if cur is None:
                cur = {'kind': 'content', 'title': '', 'blocks': []}
                slides.append(cur)
            cur['blocks'].append(('code', code))
            continue
        m = re.match(r'^(#{1,6})\s+(.*)$', stripped)
        if m:
            level, text = len(m.group(1)), _strip_md_inline(m.group(2))
            if level <= 2 or cur is None:
                kind = 'title' if (not slides and level == 1) else 'content'
                cur = {'kind': kind, 'title': text, 'blocks': []}
                slides.append(cur)
            else:
                cur['blocks'].append(('h', text, 0))
            continue
        if cur is None:
            cur = {'kind': 'content', 'title': '', 'blocks': []}
            slides.append(cur)
        if stripped.startswith('|'):
            rows = []
            i -= 1
            while i < len(lines) and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                i += 1
                if all(re.fullmatch(r':?-{2,}:?', c) for c in cells if c):
                    continue
                rows.append([_strip_md_inline(c) for c in cells])
            if rows:
                cur['blocks'].append(('t', rows))
            continue
        bm = re.match(r'^(\s*)([-*+]|\d+[.)])\s+(.*)$', line)
        if bm:
            lvl = min(len(bm.group(1).replace('\t', '    ')) // 2, 4)
            prev = next((b[2] for b in reversed(cur['blocks']) if b[0] == 'b'), -1)
            cur['blocks'].append(('b', _strip_md_inline(bm.group(3)), min(lvl, prev + 1)))
            continue
        cur['blocks'].append(('p', _strip_md_inline(stripped), 0))

    prs = Presentation()                                  # 4:3 default template, placeholders line up
    L_TITLE, L_CONTENT, L_TITLE_ONLY = prs.slide_layouts[0], prs.slide_layouts[1], prs.slide_layouts[5]

    def drop(placeholder):
        el = placeholder._element
        el.getparent().remove(el)

    def line_cost(text):
        return max(1, -(-len(text) // 70))

    def add_text_slide(title, blocks):
        s = prs.slides.add_slide(L_CONTENT)
        s.shapes.title.text = title
        body = s.placeholders[1]
        if not blocks:
            drop(body)
            return
        total = sum(line_cost(b[1]) for b in blocks)
        size = 24 if total <= 6 else 20 if total <= 9 else 17
        tf = body.text_frame
        tf.word_wrap = True
        for n, b in enumerate(blocks):
            p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
            run = p.add_run()
            run.text = b[1]
            run.font.size = Pt(max(size - 2 * b[2], 12))
            run.font.bold = (b[0] == 'h')
            p.level = b[2]

    def add_table_slide(title, rows):
        s = prs.slides.add_slide(L_TITLE_ONLY)
        s.shapes.title.text = title
        n_cols = max(len(r) for r in rows)
        gf = s.shapes.add_table(len(rows), n_cols, Inches(0.5), Inches(1.7), Inches(9), Inches(0.4 * len(rows)))
        for r, row in enumerate(rows):
            for c in range(n_cols):
                cell = gf.table.cell(r, c)
                cell.text = row[c] if c < len(row) else ''
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(14 if n_cols <= 4 else 11)

    def add_code_slide(title, code):
        s = prs.slides.add_slide(L_TITLE_ONLY)
        s.shapes.title.text = title
        tb = s.shapes.add_textbox(Inches(0.5), Inches(1.7), Inches(9), Inches(5.3))
        tf = tb.text_frame
        tf.word_wrap = True
        for n, ln in enumerate(code):
            p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
            run = p.add_run()
            run.text = ln
            run.font.name = 'Courier New'
            run.font.size = Pt(12)

    for sl in slides:
        title = sl['title'] or ' '
        if sl['kind'] == 'title':
            s = prs.slides.add_slide(L_TITLE)
            s.shapes.title.text = title
            sub = [b[1] for b in sl['blocks'] if b[0] in ('p', 'b', 'h')][:3]
            if sub:
                s.placeholders[1].text = '\n'.join(sub)
            else:
                drop(s.placeholders[1])
            for b in sl['blocks']:          # tables/code under a title slide → follow-up slides
                if b[0] == 't':
                    add_table_slide(title, b[1])
                elif b[0] == 'code':
                    add_code_slide(title, b[1])
            continue
        if not sl['blocks']:
            add_text_slide(title, [])
            continue
        state = {'group': [], 'cost': 0, 'part': 0}

        def flush():
            if state['group']:
                add_text_slide(title + (' (cont.)' if state['part'] else ''), state['group'])
                state['part'] += 1
            state['group'], state['cost'] = [], 0

        for b in sl['blocks']:
            if b[0] in ('p', 'b', 'h'):
                c = line_cost(b[1])
                if state['group'] and state['cost'] + c > 12:
                    flush()
                state['group'].append(b)
                state['cost'] += c
            else:
                flush()
                label = title + (' (cont.)' if state['part'] else '')
                if b[0] == 't':
                    header, body_rows = b[1][0], b[1][1:]
                    for k in range(0, max(len(body_rows), 1), 11):
                        add_table_slide(label if k == 0 else label + ' (cont.)', [header] + body_rows[k:k + 11])
                else:
                    for k in range(0, len(b[1]), 18):
                        add_code_slide(label if k == 0 else label + ' (cont.)', b[1][k:k + 18])
                state['part'] += 1
        flush()
    if not slides:
        s = prs.slides.add_slide(L_TITLE)
        s.shapes.title.text = src.stem
        drop(s.placeholders[1])
    prs.save(str(out))


# ── Merge PDFs / images into one PDF ──────────────────────────────────────────
def merge_to_pdf(paths: List[Path], out: Path):
    merged = fitz.open()
    try:
        for p in paths:
            ext = p.suffix.lower().lstrip('.')
            if ext == 'pdf':
                d = fitz.open(str(p))
                try:
                    if d.needs_pass:
                        raise ValueError(f"'{p.name}' is password protected")
                    merged.insert_pdf(d)
                finally:
                    d.close()
            elif ext in IMG_EXTS:
                _image_to_pdf_page(merged, p)
            else:
                raise ValueError(f"'{p.name}' is not a PDF or image")
        merged.save(str(out), garbage=3, deflate=True)
    finally:
        merged.close()


# ═══════════════════════════════════════════════════════════════════════════════
#  NEW IN v3: legacy Office input, OCR, split / compress / protect / unlock PDF
# ═══════════════════════════════════════════════════════════════════════════════

class ToolUnavailable(RuntimeError):
    """A required external engine (LibreOffice / Tesseract) is not installed → HTTP 503."""


OCR_LANG: contextvars.ContextVar = contextvars.ContextVar('ocr_lang', default='eng')
_LANG_RE = re.compile(r'^[A-Za-z_]{3,12}(\+[A-Za-z_]{3,12}){0,3}$')


def _open_pdf(path: Path, password: str = ''):
    """Open a PDF, authenticating if needed. Raises ValueError (→ HTTP 400) with a friendly message."""
    try:
        doc = fitz.open(str(path))
    except Exception:
        raise ValueError('Could not read this file as a PDF - it may be corrupted or not a real PDF.')
    if doc.needs_pass:
        if not password:
            doc.close()
            raise ValueError('This PDF is password protected. Enter its password, or unlock it first.')
        if not doc.authenticate(password):
            doc.close()
            raise ValueError('Incorrect password.')
    return doc


# ── Legacy Office input (.doc / .ppt / .xls) via LibreOffice ──────────────────
def libreoffice_convert_to(src: Path, out_dir: Path, fmt: str) -> Optional[Path]:
    """Convert src to another Office format (docx / pptx / xlsx) with LibreOffice."""
    lo = find_libreoffice()
    if not lo:
        return None
    profile = (UPLOAD_DIR / '_lo_profile').resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        with _LO_LOCK:
            subprocess.run(
                [lo, f'-env:UserInstallation={profile.as_uri()}', '--headless', '--norestore',
                 '--convert-to', fmt, '--outdir', str(out_dir), str(src)],
                capture_output=True, timeout=180
            )
        produced = out_dir / f'{src.stem}.{fmt}'
        return produced if produced.exists() and produced.stat().st_size > 0 else None
    except Exception:
        return None


_LO_MISSING = ('Old Office files (.doc / .ppt / .xls) need LibreOffice to be read. '
               'Install it free from libreoffice.org, then restart FileShift.')

def make_legacy_converter(modern_fmt: str, modern_fn):
    """Upgrade the legacy file to its modern twin first, then reuse the normal converter."""
    def conv(src: Path, out: Path):
        modern = libreoffice_convert_to(src, out.parent / '_up', modern_fmt)
        if not modern:
            if not find_libreoffice():
                raise ToolUnavailable(_LO_MISSING)
            raise RuntimeError(f"LibreOffice could not open '{src.name}' - the file may be corrupted or password protected.")
        return modern_fn(modern, out)
    return conv

def make_legacy_upgrader(modern_fmt: str):
    def conv(src: Path, out: Path):
        modern = libreoffice_convert_to(src, out.parent / '_up', modern_fmt)
        if not modern:
            if not find_libreoffice():
                raise ToolUnavailable(_LO_MISSING)
            raise RuntimeError(f"LibreOffice could not open '{src.name}'.")
        shutil.move(str(modern), str(out))
    return conv

def legacy_to_pdf(src: Path, out: Path):
    if not libreoffice_convert(src, out):
        if not find_libreoffice():
            raise ToolUnavailable(_LO_MISSING)
        raise RuntimeError(f"LibreOffice could not open '{src.name}' - the file may be corrupted or password protected.")


# ── OCR (Tesseract, via PyMuPDF) ──────────────────────────────────────────────
_OCR_INSTALL_HELP = ('OCR needs Tesseract (free). Windows: install from github.com/UB-Mannheim/tesseract/wiki  |  '
                     'macOS: brew install tesseract  |  Linux: sudo apt install tesseract-ocr  -  then restart FileShift.')

def find_tesseract() -> Optional[str]:
    p = shutil.which('tesseract')
    if p:
        return p
    candidates = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        '/opt/homebrew/bin/tesseract', '/usr/local/bin/tesseract', '/usr/bin/tesseract',
    ]
    return next((c for c in candidates if os.path.exists(c)), None)

def _setup_tessdata():
    """Point PyMuPDF at Tesseract's language files if the installer did not set TESSDATA_PREFIX."""
    if os.environ.get('TESSDATA_PREFIX'):
        return
    exe = find_tesseract()
    if exe and (Path(exe).parent / 'tessdata').is_dir():
        os.environ['TESSDATA_PREFIX'] = str(Path(exe).parent / 'tessdata')

def ocr_languages() -> List[str]:
    exe = find_tesseract()
    if not exe:
        return []
    _setup_tessdata()
    try:
        r = subprocess.run([exe, '--list-langs'], capture_output=True, text=True, timeout=15)
        langs = [l.strip() for l in r.stdout.splitlines()[1:] if l.strip() and l.strip() != 'osd']
        return sorted(langs)
    except Exception:
        return []

def _ocr_check(lang: str) -> str:
    if not find_tesseract():
        raise ToolUnavailable(_OCR_INSTALL_HELP)
    _setup_tessdata()
    avail = ocr_languages()
    for l in lang.split('+'):
        if avail and l not in avail:
            raise ValueError(f"OCR language '{l}' is not installed. Installed languages: {', '.join(avail) or 'none'}.")
    return lang

def _page_has_text(page, min_chars: int = 20) -> bool:
    return len(page.get_text('text').strip()) >= min_chars

def _ocr_pages_text(doc, lang: str, force: bool = False, dpi: int = 300) -> List[str]:
    """Plain text per page; runs OCR on pages that have no text layer (or all, if force)."""
    out = []
    for page in doc:
        if not force and _page_has_text(page):
            out.append(page.get_text('text'))
        else:
            tp = page.get_textpage_ocr(language=lang, dpi=dpi, full=True)
            out.append(page.get_text('text', textpage=tp))
    return out

def ocr_pdf_document(doc, lang: str, force: bool = False, dpi: int = 300):
    """Return (new_doc, pages_ocred, pages_kept). Pages that already have text are copied untouched."""
    new = fitz.open()
    n_ocr = n_kept = 0
    try:
        for page in doc:
            if not force and _page_has_text(page):
                new.insert_pdf(doc, from_page=page.number, to_page=page.number)
                n_kept += 1
                continue
            pix = page.get_pixmap(dpi=dpi, alpha=False)
            one = fitz.open('pdf', pix.pdfocr_tobytes(compress=True, language=lang))
            try:
                new.insert_pdf(one)
            finally:
                one.close()
            n_ocr += 1
    except Exception:
        new.close()
        raise
    return new, n_ocr, n_kept

def ocr_tool(src: Path, job: Path, language: str = 'eng', output: str = 'pdf', force: bool = False, password: str = ''):
    lang = _ocr_check(language)
    if output not in ('pdf', 'txt'):
        raise ValueError("output must be 'pdf' or 'txt'")
    if src.suffix.lower() == '.pdf':
        doc = _open_pdf(src, password)
    else:                                              # image → one A4 page → OCR
        doc = fitz.open()
        _image_to_pdf_page(doc, src)
        force = True
    try:
        if output == 'txt':
            pages = _ocr_pages_text(doc, lang, force)
            out = job / f'{src.stem}_ocr.txt'
            out.write_text('\n\n'.join(p.strip() for p in pages).strip() + '\n', encoding='utf-8')
            return Result(path=out, note=f'Text recognised from {len(pages)} page(s). Check for OCR mistakes.')
        new, n_ocr, n_kept = ocr_pdf_document(doc, lang, force)
        out = job / f'{src.stem}_searchable.pdf'
        try:
            new.save(str(out), garbage=3, deflate=True)
        finally:
            new.close()
        note = f'OCR applied to {n_ocr} page(s)'
        if n_kept:
            note += f'; {n_kept} page(s) already had text and were kept as-is'
        return Result(path=out, note=note, headers={'X-Pages-OCR': str(n_ocr), 'X-Pages-Kept': str(n_kept)})
    finally:
        doc.close()

def image_to_txt(src: Path, out: Path):
    lang = _ocr_check(OCR_LANG.get())
    doc = fitz.open()
    try:
        _image_to_pdf_page(doc, src)
        text = _ocr_pages_text(doc, lang, force=True)[0]
    finally:
        doc.close()
    out.write_text(text.strip() + '\n', encoding='utf-8')
    return Result(note='Text recognised with OCR. Check for mistakes.')

def _pdf_ocr_fallback_pages(doc) -> List[str]:
    """Used when a PDF has no text layer at all (a scan)."""
    return _ocr_pages_text(doc, _ocr_check(OCR_LANG.get()), force=True)


# ── Split PDF ─────────────────────────────────────────────────────────────────
def parse_page_ranges(spec: str, total: int):
    """'1-3, 5, 8-' → [(1,3),(5,5),(8,total)]  (1-based, inclusive). Accepts 'end'/'last'."""
    spec = (spec or '').strip().lower().replace('end', str(total)).replace('last', str(total))
    if not spec:
        raise ValueError('Enter the pages, e.g.  1-3, 5, 8-10')
    out = []
    for tok in re.split(r'\s*[,;]\s*', spec):
        if not tok:
            continue
        m = re.fullmatch(r'(\d+)(?:\s*(-)\s*(\d+)?)?', tok)
        if not m:
            raise ValueError(f"Could not understand '{tok}'. Use numbers and ranges like 1-3, 5, 8-10")
        a = int(m.group(1))
        b = int(m.group(3)) if m.group(3) else (total if m.group(2) else a)
        if a > b:
            raise ValueError(f"'{tok}' is backwards - write it as {b}-{a}.")
        if a < 1 or b > total:
            raise ValueError(f"Pages '{tok}' are outside this document (it has {total} page{'s' if total != 1 else ''}).")
        out.append((a, b))
    if not out:
        raise ValueError('Enter the pages, e.g.  1-3, 5, 8-10')
    return out

def split_pdf(src: Path, job: Path, mode: str = 'extract', pages: str = '', every: int = 1, password: str = ''):
    """mode: extract (selected pages → 1 PDF) | delete (drop pages) | ranges (1 PDF per range → zip)
             | every (chunks of N pages → zip) | each (1 PDF per page → zip)"""
    if mode not in ('extract', 'delete', 'ranges', 'every', 'each'):
        raise ValueError("mode must be one of: extract, delete, ranges, every, each")
    doc = _open_pdf(src, password)
    try:
        total = len(doc)
        stem = src.stem
        out_dir = job / '_split'
        out_dir.mkdir(exist_ok=True)

        def write(ranges, path: Path):
            new = fitz.open()
            try:
                for a, b in ranges:
                    new.insert_pdf(doc, from_page=a - 1, to_page=b - 1)
                new.save(str(path), garbage=3, deflate=True)
            finally:
                new.close()

        if mode == 'extract':
            ranges = parse_page_ranges(pages, total)
            out = job / f'{stem}_extract.pdf'
            write(ranges, out)
            n = sum(b - a + 1 for a, b in ranges)
            return Result(path=out, note=f'Extracted {n} page(s) from {total}', headers={'X-Pages': str(n)})

        if mode == 'delete':
            drop = {p for a, b in parse_page_ranges(pages, total) for p in range(a, b + 1)}
            keep = [p for p in range(1, total + 1) if p not in drop]
            if not keep:
                raise ValueError('That would delete every page.')
            out = job / f'{stem}_edited.pdf'
            write([(p, p) for p in keep], out)
            return Result(path=out, note=f'Removed {len(drop)} page(s); {len(keep)} remain', headers={'X-Pages': str(len(keep))})

        if mode == 'each':
            groups = [(p, p) for p in range(1, total + 1)]
        elif mode == 'every':
            if every < 1:
                raise ValueError('Pages per file must be at least 1.')
            groups = [(s, min(s + every - 1, total)) for s in range(1, total + 1, every)]
        else:                                                   # ranges
            groups = parse_page_ranges(pages, total)

        files = []
        for a, b in groups:
            name = f'{stem}_page{a:03d}.pdf' if a == b else f'{stem}_p{a}-{b}.pdf'
            p, k = out_dir / name, 2
            while p.exists():
                p = out_dir / name.replace('.pdf', f'_{k}.pdf')
                k += 1
            write([(a, b)], p)
            files.append(p)
        if len(files) == 1:
            single = job / files[0].name
            shutil.move(str(files[0]), str(single))
            return Result(path=single, note='Only one part was produced', headers={'X-Parts': '1'})
        zip_path = job / f'{stem}_split.zip'
        _zip_files(files, zip_path)
        return Result(path=zip_path, note=f'Split into {len(files)} PDFs, delivered as a .zip', headers={'X-Parts': str(len(files))})
    finally:
        doc.close()


# ── Compress PDF ──────────────────────────────────────────────────────────────
COMPRESS_LEVELS = {
    'low':    {'dpi_threshold': None, 'dpi_target': 0,  'quality': 0},    # lossless: clean + deflate only
    'medium': {'dpi_threshold': 170,  'dpi_target': 130, 'quality': 75},  # screen/email quality
    'high':   {'dpi_threshold': 110,  'dpi_target': 80,  'quality': 50},  # smallest, noticeably softer images
}

def _shrink_images(doc, dpi_threshold: int, dpi_target: int, quality: int):
    if hasattr(doc, 'rewrite_images'):                       # PyMuPDF >= 1.24.2
        doc.rewrite_images(dpi_threshold=dpi_threshold, dpi_target=dpi_target, quality=quality)
        return
    _need('PIL', 'Pillow')                                    # manual fallback for older PyMuPDF
    from PIL import Image
    seen = set()
    for page in doc:
        for img in page.get_images(full=True):
            xref, smask = img[0], img[1]
            if xref in seen or smask:
                continue
            seen.add(xref)
            try:
                rects = page.get_image_rects(xref)
                if not rects:
                    continue
                width_in = max(r.width for r in rects) / 72
                info = doc.extract_image(xref)
                w, h = info['width'], info['height']
                dpi = w / width_in if width_in > 0 else 0
                if dpi <= dpi_threshold:
                    continue
                scale = dpi_target / dpi
                im = Image.open(io.BytesIO(info['image'])).convert('RGB')
                im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
                buf = io.BytesIO()
                im.save(buf, 'JPEG', quality=quality, optimize=True)
                page.replace_image(xref, stream=buf.getvalue())
            except Exception:
                continue

def compress_pdf(src: Path, job: Path, level: str = 'medium', password: str = ''):
    if level not in COMPRESS_LEVELS:
        raise ValueError("level must be one of: low, medium, high")
    prm = COMPRESS_LEVELS[level]
    orig = src.stat().st_size
    out = job / f'{src.stem}_compressed.pdf'
    doc = _open_pdf(src, password)
    try:
        if prm['dpi_threshold']:
            _shrink_images(doc, prm['dpi_threshold'], prm['dpi_target'], prm['quality'])
        if level == 'high':
            try:
                doc.subset_fonts()
            except Exception:
                pass
        doc.save(str(out), garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True)
    finally:
        doc.close()
    new = out.stat().st_size
    hdr = {'X-Original-Size': str(orig), 'X-Output-Size': str(min(new, orig))}
    if new >= orig:
        shutil.copyfile(str(src), str(out))
        return Result(path=out, note='This PDF is already well optimised - returned it unchanged.', headers=hdr)
    pct = round((1 - new / orig) * 100)
    note = f'{_fmt_bytes(orig)} → {_fmt_bytes(new)} ({pct}% smaller)'
    if pct < 2 and level == 'low':
        note += ' - lossless mode can only do so much; try Medium or High to shrink the images.'
    return Result(path=out, note=note, headers=hdr)

def _fmt_bytes(n: int) -> str:
    return f'{n} B' if n < 1024 else f'{n / 1024:.1f} KB' if n < 1048576 else f'{n / 1048576:.2f} MB'


# ── Protect / unlock PDF ──────────────────────────────────────────────────────
def protect_pdf(src: Path, job: Path, password: str = '', owner_password: str = '',
                allow_print: bool = True, allow_copy: bool = True, allow_edit: bool = True):
    if not password:
        raise ValueError('Enter a password.')
    doc = _open_pdf(src)
    perm = fitz.PDF_PERM_ACCESSIBILITY
    if allow_print:
        perm |= fitz.PDF_PERM_PRINT | fitz.PDF_PERM_PRINT_HQ
    if allow_copy:
        perm |= fitz.PDF_PERM_COPY
    if allow_edit:
        perm |= fitz.PDF_PERM_MODIFY | fitz.PDF_PERM_ANNOTATE | fitz.PDF_PERM_FORM | fitz.PDF_PERM_ASSEMBLE
    restricted = not (allow_print and allow_copy and allow_edit)
    owner = owner_password or (secrets.token_urlsafe(16) if restricted else password)
    out = job / f'{src.stem}_protected.pdf'
    try:
        doc.save(str(out), encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=password, owner_pw=owner,
                 permissions=perm, garbage=3, deflate=True)
    finally:
        doc.close()
    note = 'Protected with AES-256. You need the password to open this file - keep it safe, it cannot be recovered.'
    return Result(path=out, note=note)

def unlock_pdf(src: Path, job: Path, password: str = ''):
    try:
        doc = fitz.open(str(src))
    except Exception:
        raise ValueError('Could not read this file as a PDF - it may be corrupted or not a real PDF.')
    note = 'Password removed.'
    try:
        if doc.needs_pass:
            if not password:
                raise ValueError('This PDF needs its password to be unlocked.')
            if not doc.authenticate(password):
                raise ValueError('Incorrect password.')
        elif not doc.is_encrypted:
            note = 'This PDF was not password protected - saved a clean copy.'
        else:
            note = 'Opening restrictions removed.'
        out = job / f'{src.stem}_unlocked.pdf'
        doc.save(str(out), encryption=fitz.PDF_ENCRYPT_NONE, garbage=3, deflate=True)
    finally:
        doc.close()
    return Result(path=out, note=note)


# ── Self-describing tool catalogue (handy for building the UI) ─────────────────
TOOLS_INFO = [
    {'id': 'merge', 'endpoint': '/merge', 'title': 'Merge PDFs', 'accepts': ['pdf', 'png', 'jpg', 'webp'], 'multiple': True,
     'params': []},
    {'id': 'split', 'endpoint': '/tools/split', 'title': 'Split / extract / delete pages', 'accepts': ['pdf'],
     'params': [{'name': 'mode', 'type': 'enum', 'values': ['extract', 'delete', 'ranges', 'every', 'each'], 'default': 'extract'},
                {'name': 'pages', 'type': 'string', 'example': '1-3, 5, 8-', 'needed_for': ['extract', 'delete', 'ranges']},
                {'name': 'every', 'type': 'int', 'default': 1, 'needed_for': ['every']},
                {'name': 'password', 'type': 'string', 'optional': True}]},
    {'id': 'compress', 'endpoint': '/tools/compress', 'title': 'Compress PDF', 'accepts': ['pdf'],
     'params': [{'name': 'level', 'type': 'enum', 'values': list(COMPRESS_LEVELS), 'default': 'medium'},
                {'name': 'password', 'type': 'string', 'optional': True}]},
    {'id': 'protect', 'endpoint': '/tools/protect', 'title': 'Password-protect PDF', 'accepts': ['pdf'],
     'params': [{'name': 'password', 'type': 'string', 'required': True},
                {'name': 'owner_password', 'type': 'string', 'optional': True},
                {'name': 'allow_print', 'type': 'bool', 'default': True},
                {'name': 'allow_copy', 'type': 'bool', 'default': True},
                {'name': 'allow_edit', 'type': 'bool', 'default': True}]},
    {'id': 'unlock', 'endpoint': '/tools/unlock', 'title': 'Unlock PDF', 'accepts': ['pdf'],
     'params': [{'name': 'password', 'type': 'string', 'optional': True}]},
    {'id': 'ocr', 'endpoint': '/tools/ocr', 'title': 'OCR - make scans searchable / extract text', 'accepts': ['pdf', 'png', 'jpg', 'webp'],
     'params': [{'name': 'language', 'type': 'string', 'default': 'eng', 'note': "e.g. 'eng' or 'eng+hin' - see GET /ocr/languages"},
                {'name': 'output', 'type': 'enum', 'values': ['pdf', 'txt'], 'default': 'pdf'},
                {'name': 'force', 'type': 'bool', 'default': False, 'note': 'OCR pages even if they already contain text'},
                {'name': 'password', 'type': 'string', 'optional': True}]},
]


# ─── DISPATCH TABLE ────────────────────────────────────────────────────────────
DISPATCH = {
    # PDF
    ('pdf',  'md'):   pdf_to_md,
    ('pdf',  'txt'):  pdf_to_txt,
    ('pdf',  'docx'): pdf_to_docx,
    ('pdf',  'html'): pdf_to_html,
    ('pdf',  'png'):  make_pdf_to_images('png'),
    ('pdf',  'jpg'):  make_pdf_to_images('jpg'),
    # Markdown
    ('md',   'pdf'):  md_to_pdf,
    ('md',   'docx'): md_to_docx,
    ('md',   'html'): md_to_html,
    ('md',   'txt'):  md_to_txt,
    ('md',   'pptx'): md_to_pptx,
    # Word
    ('docx', 'pdf'):  docx_to_pdf,
    ('docx', 'md'):   docx_to_md,
    ('docx', 'txt'):  docx_to_txt,
    ('docx', 'html'): docx_to_html,
    # Text
    ('txt',  'pdf'):  txt_to_pdf,
    ('txt',  'md'):   txt_to_md,
    ('txt',  'docx'): txt_to_docx,
    ('txt',  'html'): txt_to_html,
    # HTML
    ('html', 'pdf'):  html_to_pdf,
    ('html', 'md'):   html_to_md,
    ('html', 'docx'): html_to_docx,
    ('html', 'txt'):  html_to_txt,
    # PowerPoint
    ('pptx', 'pdf'):  pptx_to_pdf,
    ('pptx', 'txt'):  pptx_to_txt,
    ('pptx', 'md'):   pptx_to_md,
    # Spreadsheets
    ('xlsx', 'pdf'):  sheet_to_pdf,
    ('xlsx', 'csv'):  xlsx_to_csv,
    ('xlsx', 'md'):   sheet_to_md,
    ('xlsx', 'html'): sheet_to_html,
    ('csv',  'xlsx'): csv_to_xlsx,
    ('csv',  'pdf'):  sheet_to_pdf,
    ('csv',  'md'):   sheet_to_md,
    ('csv',  'html'): sheet_to_html,
    # Images
    ('png',  'pdf'):  img_to_pdf,
    ('png',  'jpg'):  make_img_converter('jpg'),
    ('png',  'webp'): make_img_converter('webp'),
    ('jpg',  'pdf'):  img_to_pdf,
    ('jpg',  'png'):  make_img_converter('png'),
    ('jpg',  'webp'): make_img_converter('webp'),
    ('webp', 'pdf'):  img_to_pdf,
    ('webp', 'png'):  make_img_converter('png'),
    ('webp', 'jpg'):  make_img_converter('jpg'),
    ('png',  'txt'):  image_to_txt,
    ('jpg',  'txt'):  image_to_txt,
    ('webp', 'txt'):  image_to_txt,
    # Legacy Office → modern twin, PDF, and everything the modern format supports
    ('doc',  'docx'): make_legacy_upgrader('docx'),
    ('doc',  'pdf'):  legacy_to_pdf,
    ('doc',  'md'):   make_legacy_converter('docx', docx_to_md),
    ('doc',  'txt'):  make_legacy_converter('docx', docx_to_txt),
    ('doc',  'html'): make_legacy_converter('docx', docx_to_html),
    ('ppt',  'pptx'): make_legacy_upgrader('pptx'),
    ('ppt',  'pdf'):  legacy_to_pdf,
    ('ppt',  'txt'):  make_legacy_converter('pptx', pptx_to_txt),
    ('ppt',  'md'):   make_legacy_converter('pptx', pptx_to_md),
    ('xls',  'xlsx'): make_legacy_upgrader('xlsx'),
    ('xls',  'pdf'):  make_legacy_converter('xlsx', sheet_to_pdf),
    ('xls',  'csv'):  make_legacy_converter('xlsx', xlsx_to_csv),
    ('xls',  'md'):   make_legacy_converter('xlsx', sheet_to_md),
    ('xls',  'html'): make_legacy_converter('xlsx', sheet_to_html),
}


# ─── API ───────────────────────────────────────────────────────────────────────
def _norm_ext(name: str) -> str:
    e = Path(name).suffix.lstrip('.').lower()
    return EXT_ALIASES.get(e, e)


@app.get('/formats')
def get_formats():
    return CONVERSIONS


@app.get('/capabilities')
def get_capabilities():
    """Which optional engines are available on this machine (shown in the UI)."""
    def has(mod):
        try:
            return importlib.util.find_spec(mod) is not None
        except Exception:
            return False
    return {
        'libreoffice': find_libreoffice() is not None,
        'weasyprint':  WEASYPRINT_AVAILABLE,
        'tesseract':   find_tesseract() is not None,
        'ocr_languages': ocr_languages(),
        'python-pptx': has('pptx'),
        'openpyxl':    has('openpyxl'),
        'Pillow':      has('PIL'),
    }


@app.post('/convert')
async def convert(file: UploadFile = File(...), target_format: str = Form(...), ocr_lang: str = Form('eng')):
    filename = _safe_name(Path(file.filename or 'upload').name)
    src_ext = _norm_ext(filename)
    target_format = target_format.lower().lstrip('.')
    if not _LANG_RE.match(ocr_lang):
        raise HTTPException(400, "Invalid OCR language code (examples: 'eng', 'eng+hin')")
    OCR_LANG.set(ocr_lang)                       # used by OCR-capable converters in the worker thread
    if src_ext not in CONVERSIONS:
        raise HTTPException(400, f"Format '.{src_ext}' not supported")
    if target_format not in CONVERSIONS[src_ext]:
        raise HTTPException(400, f"Cannot convert .{src_ext} → .{target_format}")
    fn = DISPATCH.get((src_ext, target_format))
    if not fn:
        raise HTTPException(400, "Conversion not available")

    job = UPLOAD_DIR / uuid.uuid4().hex
    job.mkdir(parents=True)
    src_path = job / filename
    out_path = job / f"{Path(filename).stem}.{target_format}"

    try:
        with open(src_path, 'wb') as f:
            shutil.copyfileobj(file.file, f)
        result = await run_in_threadpool(fn, src_path, out_path)   # keep the server responsive
    except ToolUnavailable as e:
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(503, str(e))
    except Exception as e:
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(500, f"Conversion failed: {e}")

    final, note, extra = out_path, None, {}
    if isinstance(result, Result):
        final, note, extra = (result.path or out_path), result.note, result.headers
    elif isinstance(result, str):
        note = result

    if not final.exists():
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(500, "Output file was not created")

    headers = {'X-Output-Filename': urllib.parse.quote(final.name)}
    if note:
        headers['X-Conversion-Note'] = urllib.parse.quote(note)
    headers.update(extra)
    return FileResponse(str(final), filename=final.name, media_type='application/octet-stream',
                        headers=headers,
                        background=BackgroundTask(shutil.rmtree, job, ignore_errors=True))   # clean temp files


@app.post('/merge')
async def merge(files: List[UploadFile] = File(...)):
    """Combine several PDFs and/or images (in the order sent) into one PDF."""
    if len(files) < 2:
        raise HTTPException(400, "Select at least two files to merge")
    job = UPLOAD_DIR / uuid.uuid4().hex
    job.mkdir(parents=True)
    try:
        paths = []
        for i, f in enumerate(files):
            name = _safe_name(Path(f.filename or f'file{i}').name)
            ext = _norm_ext(name)
            if ext != 'pdf' and ext not in IMG_EXTS:
                raise HTTPException(400, f"'{name}' is not a PDF or image (png / jpg / webp)")
            p = job / f"{i:03d}_{name}"
            with open(p, 'wb') as fh:
                shutil.copyfileobj(f.file, fh)
            paths.append(p)
        out_path = job / 'merged.pdf'
        await run_in_threadpool(merge_to_pdf, paths, out_path)
    except HTTPException:
        shutil.rmtree(job, ignore_errors=True)
        raise
    except Exception as e:
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(500, f"Merge failed: {e}")
    return FileResponse(str(out_path), filename='merged.pdf', media_type='application/pdf',
                        background=BackgroundTask(shutil.rmtree, job, ignore_errors=True))


# ─── PDF TOOLS API ─────────────────────────────────────────────────────────────
async def _run_tool(upload: UploadFile, allowed: tuple, work, **kwargs):
    """Save the upload, run work(src, job, **kwargs) in a worker thread, stream the result back, clean up."""
    filename = _safe_name(Path(upload.filename or 'upload').name)
    ext = _norm_ext(filename)
    if ext not in allowed:
        raise HTTPException(400, 'This tool accepts: ' + ', '.join('.' + a for a in allowed))
    job = UPLOAD_DIR / uuid.uuid4().hex
    job.mkdir(parents=True)
    src_path = job / filename
    try:
        with open(src_path, 'wb') as f:
            shutil.copyfileobj(upload.file, f)
        result = await run_in_threadpool(lambda: work(src_path, job, **kwargs))
    except ToolUnavailable as e:                  # engine not installed (Tesseract / LibreOffice)
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(503, str(e))
    except ValueError as e:                       # problems the user can fix (bad range, wrong password...)
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(400, str(e))
    except Exception as e:
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(500, f"Failed: {e}")
    final = result.path
    if not final or not final.exists():
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(500, "Output file was not created")
    headers = {'X-Output-Filename': urllib.parse.quote(final.name)}
    if result.note:
        headers['X-Conversion-Note'] = urllib.parse.quote(result.note)
    headers.update(result.headers)
    return FileResponse(str(final), filename=final.name,
                        media_type=mimetypes.guess_type(final.name)[0] or 'application/octet-stream',
                        headers=headers,
                        background=BackgroundTask(shutil.rmtree, job, ignore_errors=True))


@app.get('/tools', tags=['PDF tools'])
def list_tools():
    """Catalogue of every tool, its endpoint, accepted file types and parameters."""
    return TOOLS_INFO


@app.get('/ocr/languages', tags=['PDF tools'])
def get_ocr_languages():
    return {'available': find_tesseract() is not None, 'languages': ocr_languages()}


@app.post('/tools/split', tags=['PDF tools'])
async def tool_split(file: UploadFile = File(...), mode: str = Form('extract'), pages: str = Form(''),
                     every: int = Form(1), password: str = Form('')):
    """mode: extract | delete | ranges | every | each.  pages like '1-3, 5, 8-'."""
    return await _run_tool(file, ('pdf',), split_pdf, mode=mode, pages=pages, every=every, password=password)


@app.post('/tools/compress', tags=['PDF tools'])
async def tool_compress(file: UploadFile = File(...), level: str = Form('medium'), password: str = Form('')):
    """level: low (lossless) | medium | high. Returns X-Original-Size / X-Output-Size headers."""
    return await _run_tool(file, ('pdf',), compress_pdf, level=level, password=password)


@app.post('/tools/protect', tags=['PDF tools'])
async def tool_protect(file: UploadFile = File(...), password: str = Form(''), owner_password: str = Form(''),
                       allow_print: bool = Form(True), allow_copy: bool = Form(True), allow_edit: bool = Form(True)):
    """AES-256 encrypt a PDF. The password is never stored or logged."""
    return await _run_tool(file, ('pdf',), protect_pdf, password=password, owner_password=owner_password,
                           allow_print=allow_print, allow_copy=allow_copy, allow_edit=allow_edit)


@app.post('/tools/unlock', tags=['PDF tools'])
async def tool_unlock(file: UploadFile = File(...), password: str = Form('')):
    """Remove the password from a PDF you have the password for."""
    return await _run_tool(file, ('pdf',), unlock_pdf, password=password)


@app.post('/tools/ocr', tags=['PDF tools'])
async def tool_ocr(file: UploadFile = File(...), language: str = Form('eng'), output: str = Form('pdf'),
                   force: bool = Form(False), password: str = Form('')):
    """Scanned PDF or image → searchable PDF (output=pdf) or plain text (output=txt)."""
    if not _LANG_RE.match(language):
        raise HTTPException(400, "Invalid OCR language code (examples: 'eng', 'eng+hin')")
    return await _run_tool(file, ('pdf', 'png', 'jpg', 'webp'), ocr_tool,
                           language=language, output=output, force=force, password=password)


@app.get('/', response_class=HTMLResponse)
def serve_ui():
    return (Path(__file__).parent / 'index.html').read_text(encoding='utf-8')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    # 127.0.0.1 = only this computer can reach it (matches "your files never leave your machine").
    # Set HOST=0.0.0.0 if you deliberately want other devices on your network to use it.
    host = os.environ.get('HOST', '127.0.0.1')
    print(f"\n  FileShift is running ->  http://localhost:{port}\n"
          f"  Open that address in your browser (not 0.0.0.0).  Press Ctrl+C to stop.\n")
    uvicorn.run(app, host=host, port=port)