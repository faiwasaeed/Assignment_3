#!/usr/bin/env python3
"""
nb2pdf_agent.py — Jupyter Notebook to Professional PDF Agent
Converts .ipynb files to polished, formatted PDF reports.
"""

import json
import sys
import os
import re
import base64
import io
import textwrap
import argparse
from datetime import datetime

# ReportLab
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Preformatted,
    Table, TableStyle, PageBreak, HRFlowable, Image,
    KeepTogether, ListFlowable, ListItem
)
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas

# Pygments for syntax highlighting
from pygments import highlight
from pygments.lexers import PythonLexer, get_lexer_by_name, TextLexer
from pygments.formatters import HtmlFormatter

# PIL for images
from PIL import Image as PILImage


# ─────────────────────────────────────────────
# COLOUR PALETTE
# ─────────────────────────────────────────────
class Palette:
    PRIMARY      = colors.HexColor("#1a1a2e")   # deep navy
    ACCENT       = colors.HexColor("#0f3460")   # rich blue
    ACCENT2      = colors.HexColor("#e94560")   # coral red
    CODE_BG      = colors.HexColor("#f8f9fa")   # light gray
    CODE_BORDER  = colors.HexColor("#dee2e6")   # border gray
    OUTPUT_BG   = colors.HexColor("#fff8e7")   # warm cream
    OUTPUT_BORDER= colors.HexColor("#ffc107")   # amber
    ERROR_BG    = colors.HexColor("#fff0f0")   # light red
    ERROR_BORDER = colors.HexColor("#dc3545")   # red
    H1           = colors.HexColor("#1a1a2e")
    H2           = colors.HexColor("#0f3460")
    H3           = colors.HexColor("#16213e")
    TEXT         = colors.HexColor("#212529")
    SUBTEXT      = colors.HexColor("#6c757d")
    WHITE        = colors.white
    RULE         = colors.HexColor("#dee2e6")
    TOC_LINK     = colors.HexColor("#0f3460")


# ─────────────────────────────────────────────
# STYLE FACTORY
# ─────────────────────────────────────────────
def build_styles():
    base = getSampleStyleSheet()

    def S(name, parent="Normal", **kw):
        return ParagraphStyle(name, parent=base[parent], **kw)

    styles = {
        # ── Headings ──────────────────────────
        "H1": S("NB_H1", "Heading1",
                fontSize=22, textColor=Palette.H1,
                spaceAfter=10, spaceBefore=18,
                fontName="Helvetica-Bold",
                borderPad=4),

        "H2": S("NB_H2", "Heading2",
                fontSize=16, textColor=Palette.H2,
                spaceAfter=8, spaceBefore=14,
                fontName="Helvetica-Bold",
                borderPad=3),

        "H3": S("NB_H3", "Heading3",
                fontSize=13, textColor=Palette.H3,
                spaceAfter=6, spaceBefore=10,
                fontName="Helvetica-Bold"),

        "H4": S("NB_H4", "Heading4",
                fontSize=11, textColor=Palette.H3,
                spaceAfter=5, spaceBefore=8,
                fontName="Helvetica-BoldOblique"),

        # ── Body ──────────────────────────────
        "Body": S("NB_Body", "Normal",
                  fontSize=10.5, textColor=Palette.TEXT,
                  leading=16, spaceAfter=6,
                  fontName="Helvetica",
                  alignment=TA_JUSTIFY),

        "Bold": S("NB_Bold", "Normal",
                  fontSize=10.5, textColor=Palette.TEXT,
                  leading=16, fontName="Helvetica-Bold"),

        "Italic": S("NB_Italic", "Normal",
                    fontSize=10.5, textColor=Palette.TEXT,
                    leading=16, fontName="Helvetica-Oblique"),

        "BulletItem": S("NB_Bullet", "Normal",
                        fontSize=10.5, textColor=Palette.TEXT,
                        leading=15, leftIndent=16,
                        spaceAfter=3, fontName="Helvetica"),

        # ── Code ──────────────────────────────
        "Code": ParagraphStyle("NB_Code",
                fontSize=8.8, fontName="Courier",
                textColor=Palette.PRIMARY,
                backColor=Palette.CODE_BG,
                leading=13, leftIndent=8, rightIndent=8,
                spaceAfter=0, spaceBefore=0),

        "CodeLabel": S("NB_CodeLabel", "Normal",
                       fontSize=8, fontName="Helvetica-Bold",
                       textColor=Palette.SUBTEXT,
                       spaceAfter=2, spaceBefore=8),

        "Output": ParagraphStyle("NB_Output",
                fontSize=8.5, fontName="Courier",
                textColor=Palette.TEXT,
                leading=12.5,
                spaceAfter=0, spaceBefore=0),

        # ── TOC ───────────────────────────────
        "TOCTitle": S("NB_TOCTitle", "Heading1",
                      fontSize=18, textColor=Palette.PRIMARY,
                      fontName="Helvetica-Bold",
                      spaceAfter=12, spaceBefore=6),

        "TOC1": S("NB_TOC1", "Normal",
                  fontSize=10.5, textColor=Palette.TOC_LINK,
                  fontName="Helvetica-Bold",
                  leftIndent=0, spaceAfter=4),

        "TOC2": S("NB_TOC2", "Normal",
                  fontSize=9.5, textColor=Palette.ACCENT,
                  fontName="Helvetica",
                  leftIndent=16, spaceAfter=3),

        "TOC3": S("NB_TOC3", "Normal",
                  fontSize=9, textColor=Palette.SUBTEXT,
                  fontName="Helvetica",
                  leftIndent=28, spaceAfter=2),

        # ── Cover ─────────────────────────────
        "CoverTitle": S("NB_CoverTitle", "Title",
                        fontSize=28, textColor=Palette.WHITE,
                        fontName="Helvetica-Bold",
                        alignment=TA_CENTER, leading=34),

        "CoverSubtitle": S("NB_CoverSub", "Normal",
                           fontSize=12, textColor=colors.HexColor("#adb5bd"),
                           fontName="Helvetica",
                           alignment=TA_CENTER, leading=18),

        "CoverMeta": S("NB_CoverMeta", "Normal",
                       fontSize=10, textColor=colors.HexColor("#ced4da"),
                       fontName="Helvetica-Oblique",
                       alignment=TA_CENTER),
    }
    return styles


# ─────────────────────────────────────────────
# NOTEBOOK PARSER
# ─────────────────────────────────────────────
class NotebookParser:
    def __init__(self, path: str):
        with open(path, "r", encoding="utf-8") as f:
            self.nb = json.load(f)
        self.path = path
        self.filename = os.path.basename(path)
        self.kernel = (self.nb.get("metadata", {})
                           .get("kernelspec", {})
                           .get("display_name", "Python"))
        self.nbformat = self.nb.get("nbformat", 4)

    def cells(self):
        return self.nb.get("cells", [])

    def get_title(self):
        for cell in self.cells():
            if cell.get("cell_type") == "markdown":
                src = "".join(cell.get("source", []))
                m = re.match(r"^#\s+(.+)", src.strip())
                if m:
                    return m.group(1)
        return os.path.splitext(self.filename)[0].replace("_", " ").title()


# ─────────────────────────────────────────────
# MARKDOWN → REPORTLAB FLOWABLES
# ─────────────────────────────────────────────
class MarkdownConverter:
    """Converts markdown text into ReportLab Paragraph flowables."""

    def __init__(self, styles):
        self.S = styles

    def _escape(self, txt):
        txt = txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return txt

    def _inline(self, txt):
        """Handle bold, italic, inline code, links."""
        txt = self._escape(txt)
        # inline code
        txt = re.sub(r"`([^`]+)`",
                     r'<font name="Courier" color="#e94560">\1</font>', txt)
        # bold+italic
        txt = re.sub(r"\*\*\*(.+?)\*\*\*",
                     r"<b><i>\1</i></b>", txt)
        # bold
        txt = re.sub(r"\*\*(.+?)\*\*|__(.+?)__",
                     lambda m: f"<b>{m.group(1) or m.group(2)}</b>", txt)
        # italic
        txt = re.sub(r"\*(.+?)\*|_(.+?)_",
                     lambda m: f"<i>{m.group(1) or m.group(2)}</i>", txt)
        # links
        txt = re.sub(r"\[(.+?)\]\((.+?)\)",
                     r'<a href="\2" color="#0f3460">\1</a>', txt)
        return txt

    def convert(self, source: str):
        """Return list of flowables."""
        lines = source.split("\n")
        flowables = []
        i = 0

        while i < len(lines):
            line = lines[i]

            # Blank line
            if not line.strip():
                i += 1
                continue

            # Headings
            hm = re.match(r"^(#{1,4})\s+(.*)", line)
            if hm:
                level = len(hm.group(1))
                text = self._inline(hm.group(2))
                style_map = {1: "H1", 2: "H2", 3: "H3", 4: "H4"}
                sname = style_map.get(level, "H3")
                flowables.append(Paragraph(text, self.S[sname]))
                if level == 1:
                    flowables.append(HRFlowable(
                        width="100%", thickness=2,
                        color=Palette.ACCENT2, spaceAfter=6))
                elif level == 2:
                    flowables.append(HRFlowable(
                        width="100%", thickness=1,
                        color=Palette.RULE, spaceAfter=4))
                i += 1
                continue

            # Fenced code block
            if line.strip().startswith("```"):
                lang = line.strip()[3:].strip() or "text"
                code_lines = []
                i += 1
                while i < len(lines) and not lines[i].strip().startswith("```"):
                    code_lines.append(lines[i])
                    i += 1
                i += 1  # skip closing ```
                code_text = "\n".join(code_lines)
                flowables.extend(self._code_block(code_text, lang, inline=True))
                continue

            # Horizontal rule
            if re.match(r"^[-*_]{3,}$", line.strip()):
                flowables.append(HRFlowable(
                    width="100%", thickness=1,
                    color=Palette.RULE, spaceAfter=6, spaceBefore=6))
                i += 1
                continue

            # Unordered list
            if re.match(r"^[\*\-\+]\s+", line):
                items = []
                while i < len(lines) and re.match(r"^[\*\-\+]\s+", lines[i]):
                    item_text = re.sub(r"^[\*\-\+]\s+", "", lines[i])
                    items.append(ListItem(
                        Paragraph(self._inline(item_text), self.S["BulletItem"]),
                        bulletColor=Palette.ACCENT2, bulletType="bullet"))
                    i += 1
                flowables.append(ListFlowable(items,
                    bulletType="bullet", leftIndent=16, spaceAfter=6))
                continue

            # Ordered list
            if re.match(r"^\d+\.\s+", line):
                items = []
                num = 1
                while i < len(lines) and re.match(r"^\d+\.\s+", lines[i]):
                    item_text = re.sub(r"^\d+\.\s+", "", lines[i])
                    items.append(ListItem(
                        Paragraph(self._inline(item_text), self.S["BulletItem"]),
                        value=num))
                    num += 1
                    i += 1
                flowables.append(ListFlowable(items,
                    bulletType="1", leftIndent=16, spaceAfter=6))
                continue

            # Blockquote
            if line.startswith(">"):
                bq_lines = []
                while i < len(lines) and lines[i].startswith(">"):
                    bq_lines.append(lines[i].lstrip("> "))
                    i += 1
                bq_text = " ".join(bq_lines)
                bq_style = ParagraphStyle("BQ", parent=self.S["Body"],
                    leftIndent=20, rightIndent=10,
                    borderPad=6, borderColor=Palette.ACCENT2,
                    borderWidth=3, borderRadius=2,
                    textColor=Palette.SUBTEXT,
                    fontName="Helvetica-Oblique")
                flowables.append(Paragraph(self._inline(bq_text), bq_style))
                flowables.append(Spacer(1, 4))
                continue

            # Simple markdown table
            if "|" in line and i + 1 < len(lines) and re.match(r"[\|\-\s:]+", lines[i+1]):
                table_lines = []
                while i < len(lines) and "|" in lines[i]:
                    table_lines.append(lines[i])
                    i += 1
                flowables.extend(self._md_table(table_lines))
                continue

            # Normal paragraph — collect consecutive body lines
            para_lines = []
            while i < len(lines) and lines[i].strip() and \
                  not re.match(r"^#{1,4}\s", lines[i]) and \
                  not lines[i].startswith("```") and \
                  not re.match(r"^[\*\-\+]\s+", lines[i]) and \
                  not re.match(r"^\d+\.\s+", lines[i]) and \
                  "|" not in lines[i]:
                para_lines.append(lines[i])
                i += 1
            if para_lines:
                text = " ".join(para_lines)
                flowables.append(Paragraph(self._inline(text), self.S["Body"]))
                flowables.append(Spacer(1, 4))

        return flowables

    def _code_block(self, code: str, lang: str = "python", inline: bool = False):
        """Render a code block with monospace font in a shaded box."""
        flowables = []
        # Clean and wrap
        code = code.rstrip()
        lines = code.split("\n")
        max_width = 90
        wrapped = []
        for ln in lines:
            if len(ln) > max_width:
                wrapped.extend(textwrap.wrap(ln, max_width, subsequent_indent="  "))
            else:
                wrapped.append(ln)
        code_text = "\n".join(wrapped)

        # Build table with left accent bar
        code_para = Preformatted(code_text, self.S["Code"])
        data = [[code_para]]
        t = Table(data, colWidths=["100%"])
        t.setStyle(TableStyle([
            ("BACKGROUND",  (0,0), (-1,-1), Palette.CODE_BG),
            ("BOX",         (0,0), (-1,-1), 0.75, Palette.CODE_BORDER),
            ("LEFTPADDING",  (0,0), (-1,-1), 10),
            ("RIGHTPADDING", (0,0), (-1,-1), 8),
            ("TOPPADDING",   (0,0), (-1,-1), 8),
            ("BOTTOMPADDING",(0,0), (-1,-1), 8),
            ("LINEBEFORE",   (0,0), (0,-1), 3, Palette.ACCENT),
        ]))
        flowables.append(t)
        flowables.append(Spacer(1, 6))
        return flowables

    def _md_table(self, table_lines):
        """Convert markdown table to ReportLab Table."""
        flowables = []
        rows = []
        for line in table_lines:
            if re.match(r"[\|\-\s:]+$", line.replace("|", "").replace("-", "").replace(":", "").replace(" ", "") + "x"):
                continue  # separator row
            cells = [c.strip() for c in line.split("|") if c.strip()]
            if cells:
                rows.append(cells)

        if not rows:
            return flowables

        col_count = max(len(r) for r in rows)
        table_data = []
        styles_ts = [
            ("BACKGROUND",   (0,0), (-1,0), Palette.ACCENT),
            ("TEXTCOLOR",    (0,0), (-1,0), Palette.WHITE),
            ("FONTNAME",     (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",     (0,0), (-1,-1), 9),
            ("GRID",         (0,0), (-1,-1), 0.5, Palette.RULE),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, Palette.CODE_BG]),
            ("ALIGN",        (0,0), (-1,-1), "LEFT"),
            ("LEFTPADDING",  (0,0), (-1,-1), 6),
            ("RIGHTPADDING", (0,0), (-1,-1), 6),
            ("TOPPADDING",   (0,0), (-1,-1), 4),
            ("BOTTOMPADDING",(0,0), (-1,-1), 4),
        ]
        for ri, row in enumerate(rows):
            # Pad row to col_count
            padded = row + [""] * (col_count - len(row))
            style_key = "Bold" if ri == 0 else "Body"
            table_data.append([Paragraph(self._inline(c), self.S[style_key]) for c in padded])

        col_w = [None] * col_count  # auto
        t = Table(table_data, repeatRows=1)
        t.setStyle(TableStyle(styles_ts))
        flowables.append(t)
        flowables.append(Spacer(1, 8))
        return flowables


# ─────────────────────────────────────────────
# PDF BUILDER
# ─────────────────────────────────────────────
class PDFBuilder:
    def __init__(self, output_path: str, title: str, kernel: str, notebook_file: str):
        self.output_path = output_path
        self.title = title
        self.kernel = kernel
        self.notebook_file = notebook_file
        self.styles = build_styles()
        self.md_conv = MarkdownConverter(self.styles)
        self.story = []
        self.toc = TableOfContents()
        self._setup_toc()
        self.page_width, self.page_height = A4
        self._cell_counter = 0
        self._heading_bookmarks = []

    def _setup_toc(self):
        self.toc.levelStyles = [
            self.styles["TOC1"],
            self.styles["TOC2"],
            self.styles["TOC3"],
        ]
        self.toc.dotsMinLevel = 0

    # ── Cover page ──────────────────────────
    def _draw_cover(self, canvas, doc):
        canvas.saveState()
        w, h = A4

        # Background gradient simulation with rectangles
        canvas.setFillColor(Palette.PRIMARY)
        canvas.rect(0, 0, w, h, fill=1, stroke=0)

        # Accent stripe
        canvas.setFillColor(Palette.ACCENT)
        canvas.rect(0, h * 0.38, w, h * 0.62, fill=1, stroke=0)

        # Decorative circles
        canvas.setFillColor(colors.HexColor("#16213e"))
        canvas.circle(w * 0.85, h * 0.80, 80, fill=1, stroke=0)
        canvas.circle(w * 0.10, h * 0.15, 50, fill=1, stroke=0)
        canvas.setFillColor(Palette.ACCENT2)
        canvas.circle(w * 0.85, h * 0.80, 20, fill=1, stroke=0)

        # Title box
        canvas.setFillColor(colors.HexColor("#0f3460"))
        canvas.roundRect(2*cm, h*0.45, w - 4*cm, h*0.20, 8, fill=1, stroke=0)

        # Title text
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 26)
        max_w = w - 5*cm
        # Word-wrap title
        words = self.title.split()
        lines = []
        cur = ""
        for word in words:
            test = (cur + " " + word).strip()
            if canvas.stringWidth(test, "Helvetica-Bold", 26) < max_w:
                cur = test
            else:
                lines.append(cur)
                cur = word
        lines.append(cur)
        y_start = h * 0.62 - (len(lines) - 1) * 30
        for li, ln in enumerate(lines):
            canvas.drawCentredString(w/2, y_start - li*32, ln)

        # Subtitle bar
        canvas.setFillColor(Palette.ACCENT2)
        canvas.rect(2*cm, h*0.44, w-4*cm, 3, fill=1, stroke=0)

        # Meta info
        canvas.setFont("Helvetica", 11)
        canvas.setFillColor(colors.HexColor("#adb5bd"))
        canvas.drawCentredString(w/2, h*0.40, f"Kernel: {self.kernel}")
        canvas.drawCentredString(w/2, h*0.37,
                                 f"Generated: {datetime.now().strftime('%B %d, %Y at %H:%M')}")
        canvas.drawCentredString(w/2, h*0.34, f"Source: {self.notebook_file}")

        # Bottom badge
        canvas.setFillColor(Palette.ACCENT2)
        canvas.roundRect(w/2 - 60, 1.5*cm, 120, 22, 5, fill=1, stroke=0)
        canvas.setFont("Helvetica-Bold", 9)
        canvas.setFillColor(colors.white)
        canvas.drawCentredString(w/2, 1.5*cm + 7, "nb2pdf  AI Agent")

        canvas.restoreState()

    # ── Header / Footer ─────────────────────
    def _on_page(self, canvas, doc):
        canvas.saveState()
        w, h = A4

        # Header bar
        canvas.setFillColor(Palette.PRIMARY)
        canvas.rect(0, h - 1.5*cm, w, 1.5*cm, fill=1, stroke=0)
        canvas.setFillColor(Palette.ACCENT2)
        canvas.rect(0, h - 1.5*cm - 2, w, 2, fill=1, stroke=0)

        # Header text
        canvas.setFont("Helvetica-Bold", 9)
        canvas.setFillColor(colors.white)
        canvas.drawString(1.5*cm, h - 1.0*cm, self.title)
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(w - 1.5*cm, h - 1.0*cm,
                               f"Kernel: {self.kernel}")

        # Footer
        canvas.setStrokeColor(Palette.RULE)
        canvas.setLineWidth(0.5)
        canvas.line(1.5*cm, 1.5*cm, w - 1.5*cm, 1.5*cm)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(Palette.SUBTEXT)
        canvas.drawString(1.5*cm, 1.0*cm,
                          datetime.now().strftime("%Y-%m-%d"))
        canvas.drawCentredString(w/2, 1.0*cm, self.notebook_file)
        canvas.drawRightString(w - 1.5*cm, 1.0*cm,
                               f"Page {doc.page}")

        canvas.restoreState()

    def _on_cover_page(self, canvas, doc):
        self._draw_cover(canvas, doc)

    # ── Cell rendering ──────────────────────
    def _label_tag(self, text: str, color):
        style = ParagraphStyle("CellLabel",
            fontSize=7.5, fontName="Helvetica-Bold",
            textColor=colors.white,
            backColor=color,
            leftIndent=4, rightIndent=4,
            spaceAfter=2, spaceBefore=6,
            borderPad=3)
        return Paragraph(text, style)

    def add_markdown_cell(self, source: str):
        self.story.extend(self.md_conv.convert(source))
        self.story.append(Spacer(1, 4))

    def add_code_cell(self, source: str, cell_idx: int, outputs: list):
        self._cell_counter += 1
        n = self._cell_counter

        # Code block label
        label_data = [[Paragraph(
            f'<font color="white"><b>In [{n}]</b></font>',
            ParagraphStyle("CL", fontSize=8, fontName="Helvetica-Bold",
                           textColor=colors.white, leading=12))]]
        label_t = Table(label_data, colWidths=["100%"])
        label_t.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), Palette.ACCENT),
            ("LEFTPADDING",  (0,0),(-1,-1), 8),
            ("TOPPADDING",   (0,0),(-1,-1), 3),
            ("BOTTOMPADDING",(0,0),(-1,-1), 3),
            ("RIGHTPADDING", (0,0),(-1,-1), 8),
        ]))

        # Code content
        code_text = source.rstrip()
        lines = code_text.split("\n")
        max_w = 88
        wrapped = []
        for ln in lines:
            if len(ln) > max_w:
                wrapped.extend(textwrap.wrap(ln, max_w, subsequent_indent="    "))
            else:
                wrapped.append(ln)
        code_text = "\n".join(wrapped)

        code_pre = Preformatted(code_text, self.styles["Code"])
        code_data = [[code_pre]]
        code_t = Table(code_data, colWidths=["100%"])
        code_t.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), Palette.CODE_BG),
            ("BOX",          (0,0),(-1,-1), 0.5, Palette.CODE_BORDER),
            ("LEFTPADDING",  (0,0),(-1,-1), 10),
            ("RIGHTPADDING", (0,0),(-1,-1), 8),
            ("TOPPADDING",   (0,0),(-1,-1), 8),
            ("BOTTOMPADDING",(0,0),(-1,-1), 8),
            ("LINEBEFORE",   (0,0),(0,-1),  3, Palette.ACCENT),
        ]))

        parts = [label_t, code_t]

        # Outputs
        for out in outputs:
            out_flowable = self._render_output(out, n)
            if out_flowable:
                parts.extend(out_flowable)

        self.story.append(KeepTogether(parts[:3]))
        if len(parts) > 3:
            self.story.extend(parts[3:])
        self.story.append(Spacer(1, 8))

    def _render_output(self, output: dict, cell_n: int):
        otype = output.get("output_type", "")
        flowables = []

        # Output label
        def out_label(text, color):
            data = [[Paragraph(
                f'<font color="white"><b>{text}</b></font>',
                ParagraphStyle("OL", fontSize=8, fontName="Helvetica-Bold",
                               textColor=colors.white, leading=12))]]
            t = Table(data, colWidths=["100%"])
            t.setStyle(TableStyle([
                ("BACKGROUND",   (0,0),(-1,-1), color),
                ("LEFTPADDING",  (0,0),(-1,-1), 8),
                ("TOPPADDING",   (0,0),(-1,-1), 3),
                ("BOTTOMPADDING",(0,0),(-1,-1), 3),
                ("RIGHTPADDING", (0,0),(-1,-1), 8),
            ]))
            return t

        def text_box(lines_text, bg, border, is_error=False):
            txt = lines_text.rstrip()
            if not txt:
                return None
            # Truncate very long outputs
            lines_list = txt.split("\n")
            if len(lines_list) > 60:
                lines_list = lines_list[:30] + ["... (output truncated) ..."] + lines_list[-10:]
                txt = "\n".join(lines_list)
            # Wrap long lines
            wrapped = []
            for ln in lines_list:
                if len(ln) > 90:
                    wrapped.extend(textwrap.wrap(ln, 90, subsequent_indent="  "))
                else:
                    wrapped.append(ln)
            txt = "\n".join(wrapped)

            pre = Preformatted(txt, self.styles["Output"])
            data = [[pre]]
            t = Table(data, colWidths=["100%"])
            t.setStyle(TableStyle([
                ("BACKGROUND",   (0,0),(-1,-1), bg),
                ("BOX",          (0,0),(-1,-1), 0.5, border),
                ("LEFTPADDING",  (0,0),(-1,-1), 10),
                ("RIGHTPADDING", (0,0),(-1,-1), 8),
                ("TOPPADDING",   (0,0),(-1,-1), 6),
                ("BOTTOMPADDING",(0,0),(-1,-1), 6),
                ("LINEBEFORE",   (0,0),(0,-1),  3,
                 Palette.ERROR_BORDER if is_error else Palette.OUTPUT_BORDER),
            ]))
            return t

        # ── stream output ──
        if otype == "stream":
            text = "".join(output.get("text", []))
            if text.strip():
                name = output.get("name", "stdout")
                color = Palette.ACCENT if name == "stdout" else Palette.ACCENT2
                flowables.append(out_label(f"Out [{cell_n}] — {name}", color))
                t = text_box(text, Palette.OUTPUT_BG, Palette.OUTPUT_BORDER)
                if t:
                    flowables.append(t)

        # ── execute_result / display_data ──
        elif otype in ("execute_result", "display_data"):
            data = output.get("data", {})
            # Image first
            for mime in ("image/png", "image/jpeg", "image/gif"):
                if mime in data:
                    flowables.append(out_label(f"Out [{cell_n}] — image", Palette.ACCENT))
                    img_f = self._render_image(data[mime])
                    if img_f:
                        flowables.append(img_f)
                    break
            else:
                # HTML table or text
                text = None
                if "text/plain" in data:
                    text = "".join(data["text/plain"])
                elif "text/html" in data:
                    html = "".join(data["text/html"])
                    text = re.sub("<[^>]+>", "", html)
                    text = re.sub(r"\s+", " ", text).strip()
                if text and text.strip():
                    flowables.append(out_label(f"Out [{cell_n}]", Palette.ACCENT))
                    t = text_box(text, Palette.OUTPUT_BG, Palette.OUTPUT_BORDER)
                    if t:
                        flowables.append(t)

        # ── error ──
        elif otype == "error":
            ename = output.get("ename", "Error")
            evalue = output.get("evalue", "")
            tb = output.get("traceback", [])
            # Strip ANSI escape codes
            ansi_re = re.compile(r"\x1b\[[0-9;]*m")
            clean_tb = [ansi_re.sub("", line) for line in tb]
            text = f"{ename}: {evalue}\n" + "\n".join(clean_tb[-10:])
            flowables.append(out_label(f"Error — {ename}", Palette.ACCENT2))
            t = text_box(text, Palette.ERROR_BG, Palette.ERROR_BORDER, is_error=True)
            if t:
                flowables.append(t)

        return flowables if flowables else None

    def _render_image(self, b64_data):
        try:
            if isinstance(b64_data, list):
                b64_data = "".join(b64_data)
            raw = base64.b64decode(b64_data)
            img = PILImage.open(io.BytesIO(raw))
            # Resize to fit page
            max_w = self.page_width - 4*cm
            max_h = 12*cm
            iw, ih = img.size
            ratio = min(max_w / iw, max_h / ih, 1.0)
            new_w, new_h = iw * ratio, ih * ratio
            img_buf = io.BytesIO()
            img.save(img_buf, format="PNG")
            img_buf.seek(0)
            rl_img = Image(img_buf, width=new_w, height=new_h)
            rl_img.hAlign = "LEFT"
            return rl_img
        except Exception as e:
            return Paragraph(f"[Image could not be rendered: {e}]",
                             self.styles["Body"])

    # ── TOC heading registration ─────────────
    def _register_heading(self, text: str, level: int):
        """Add a bookmark for TOC."""
        clean = re.sub(r"<[^>]+>", "", text)
        style_map = {1: "NB_TOC1", 2: "NB_TOC2", 3: "NB_TOC3"}
        sname = style_map.get(level, "NB_TOC3")
        # Notify TOC
        self.story.append(
            Paragraph(text, self.styles.get(f"H{level}", self.styles["H2"])))
        if level <= 3:
            self.toc.notify("TOCEntry", (level - 1, clean, self.story[-1]))

    # ── Build ────────────────────────────────
    def build(self, parser: NotebookParser):
        doc = SimpleDocTemplate(
            self.output_path,
            pagesize=A4,
            topMargin=2.2*cm,
            bottomMargin=2.5*cm,
            leftMargin=2*cm,
            rightMargin=2*cm,
            title=self.title,
            author="nb2pdf AI Agent",
            subject="Jupyter Notebook Report",
        )

        # ── Cover ──
        self.story.append(PageBreak())   # cover is page 1

        # ── TOC page ──
        self.story.append(Paragraph("Table of Contents", self.styles["TOCTitle"]))
        self.story.append(HRFlowable(width="100%", thickness=2,
                                     color=Palette.ACCENT2, spaceAfter=10))
        self.story.append(self.toc)
        self.story.append(PageBreak())

        # ── Notebook cells ──
        for cell_idx, cell in enumerate(parser.cells()):
            ctype = cell.get("cell_type", "")
            src = "".join(cell.get("source", []))

            if not src.strip():
                continue

            if ctype == "markdown":
                self.add_markdown_cell(src)

            elif ctype == "code":
                outputs = cell.get("outputs", [])
                self.add_code_cell(src, cell_idx, outputs)

            elif ctype == "raw":
                raw_pre = Preformatted(src, self.styles["Code"])
                self.story.append(raw_pre)
                self.story.append(Spacer(1, 6))

        # ── Build PDF ──
        doc.multiBuild(
            self.story,
            onFirstPage=self._on_cover_page,
            onLaterPages=self._on_page,
        )
        print(f"✅  PDF saved to: {self.output_path}")


# ─────────────────────────────────────────────
# SAMPLE NOTEBOOK GENERATOR (for demo)
# ─────────────────────────────────────────────
def create_sample_notebook(path: str):
    nb = {
        "nbformat": 4,
        "nbformat_minor": 5,
        "metadata": {
            "kernelspec": {
                "display_name": "Python 3",
                "language": "python",
                "name": "python3"
            }
        },
        "cells": [
            {
                "cell_type": "markdown",
                "source": [
                    "# Data Science Lab Report\n",
                    "## Exploratory Data Analysis & Machine Learning Pipeline\n\n",
                    "This notebook demonstrates a complete data science workflow — from raw data ingestion through feature engineering, model training, and evaluation. All results are reproducible.\n\n",
                    "**Author:** Data Science Team  \n",
                    "**Date:** 2024-01-15  \n",
                    "**Dataset:** Synthetic Sales Data\n"
                ],
                "outputs": [],
                "metadata": {}
            },
            {
                "cell_type": "markdown",
                "source": [
                    "## 1. Environment Setup\n\n",
                    "We begin by importing the necessary libraries and configuring the analysis environment.\n"
                ],
                "outputs": [],
                "metadata": {}
            },
            {
                "cell_type": "code",
                "source": [
                    "import numpy as np\n",
                    "import pandas as pd\n",
                    "import matplotlib.pyplot as plt\n",
                    "from sklearn.model_selection import train_test_split\n",
                    "from sklearn.preprocessing import StandardScaler\n",
                    "from sklearn.linear_model import LinearRegression\n",
                    "from sklearn.metrics import mean_squared_error, r2_score\n\n",
                    "# Configuration\n",
                    "np.random.seed(42)\n",
                    "pd.set_option('display.max_columns', 20)\n",
                    "pd.set_option('display.float_format', '{:.4f}'.format)\n\n",
                    "print('Environment ready.')\n",
                    "print(f'NumPy  {np.__version__}')\n",
                    "print(f'Pandas {pd.__version__}')"
                ],
                "outputs": [
                    {
                        "output_type": "stream",
                        "name": "stdout",
                        "text": [
                            "Environment ready.\n",
                            "NumPy  1.24.3\n",
                            "Pandas 2.0.3\n"
                        ]
                    }
                ],
                "metadata": {}
            },
            {
                "cell_type": "markdown",
                "source": [
                    "## 2. Data Generation & Exploration\n\n",
                    "We generate a synthetic dataset that models sales performance across multiple regions with the following features:\n\n",
                    "- **advertising_spend** — marketing budget in USD\n",
                    "- **region_score** — regional market index (0–10)\n",
                    "- **seasonality** — month-of-year encoded factor\n",
                    "- **sales** — target variable (units sold)\n"
                ],
                "outputs": [],
                "metadata": {}
            },
            {
                "cell_type": "code",
                "source": [
                    "# Generate synthetic sales data\n",
                    "n = 500\n",
                    "advertising = np.random.uniform(1000, 50000, n)\n",
                    "region_score = np.random.uniform(1, 10, n)\n",
                    "seasonality = np.sin(np.linspace(0, 4 * np.pi, n)) + 1\n\n",
                    "noise = np.random.normal(0, 500, n)\n",
                    "sales = (0.8 * advertising + 200 * region_score +\n",
                    "         1500 * seasonality + 3000 + noise)\n\n",
                    "df = pd.DataFrame({\n",
                    "    'advertising_spend': advertising,\n",
                    "    'region_score': region_score,\n",
                    "    'seasonality': seasonality,\n",
                    "    'sales': sales\n",
                    "})\n\n",
                    "print(f'Dataset shape: {df.shape}')\n",
                    "print('\\nFirst 5 rows:')\n",
                    "print(df.head().to_string())"
                ],
                "outputs": [
                    {
                        "output_type": "stream",
                        "name": "stdout",
                        "text": [
                            "Dataset shape: (500, 4)\n\n",
                            "First 5 rows:\n",
                            "   advertising_spend  region_score  seasonality         sales\n",
                            "0       37454.011884      9.318197     1.000000  34323.851043\n",
                            "1        9507.143064      6.702972     1.025168   9874.224862\n",
                            "2       73199.394531      5.956366     1.100083  61924.371085\n",
                            "3       59865.847929      7.781567     1.224745  51098.334892\n",
                            "4       15601.864745      3.197022     1.397127  15887.992043\n"
                        ]
                    }
                ],
                "metadata": {}
            },
            {
                "cell_type": "code",
                "source": [
                    "# Descriptive statistics\n",
                    "print('Summary Statistics:')\n",
                    "print(df.describe().to_string())"
                ],
                "outputs": [
                    {
                        "output_type": "stream",
                        "name": "stdout",
                        "text": [
                            "Summary Statistics:\n",
                            "       advertising_spend  region_score  seasonality         sales\n",
                            "count         500.000000    500.000000   500.000000    500.000000\n",
                            "mean        25723.818929      5.512401     1.000000  24078.956791\n",
                            "std         14346.791435      2.580907     0.707213  12936.854017\n",
                            "min           104.689044      1.001858    -0.000000   1542.273089\n",
                            "25%         13144.781960      3.220455     0.329168  13196.287952\n",
                            "50%         25765.632618      5.563875     1.000000  23892.143094\n",
                            "75%         38462.044273      7.797256     1.670832  35001.847123\n",
                            "max         49979.133174      9.998947     2.000000  46781.012034\n"
                        ]
                    }
                ],
                "metadata": {}
            },
            {
                "cell_type": "markdown",
                "source": [
                    "## 3. Feature Engineering & Preprocessing\n\n",
                    "Before training, we apply standard preprocessing steps:\n\n",
                    "1. Split into train/test sets (80/20)\n",
                    "2. Standardise features with `StandardScaler`\n",
                    "3. Verify shapes\n"
                ],
                "outputs": [],
                "metadata": {}
            },
            {
                "cell_type": "code",
                "source": [
                    "X = df[['advertising_spend', 'region_score', 'seasonality']]\n",
                    "y = df['sales']\n\n",
                    "X_train, X_test, y_train, y_test = train_test_split(\n",
                    "    X, y, test_size=0.2, random_state=42\n",
                    ")\n\n",
                    "scaler = StandardScaler()\n",
                    "X_train_scaled = scaler.fit_transform(X_train)\n",
                    "X_test_scaled  = scaler.transform(X_test)\n\n",
                    "print(f'Training samples : {X_train_scaled.shape[0]}')\n",
                    "print(f'Test samples     : {X_test_scaled.shape[0]}')\n",
                    "print(f'Features         : {X_train_scaled.shape[1]}')"
                ],
                "outputs": [
                    {
                        "output_type": "stream",
                        "name": "stdout",
                        "text": [
                            "Training samples : 400\n",
                            "Test samples     : 100\n",
                            "Features         : 3\n"
                        ]
                    }
                ],
                "metadata": {}
            },
            {
                "cell_type": "markdown",
                "source": [
                    "## 4. Model Training\n\n",
                    "We fit an **Ordinary Least Squares** (OLS) Linear Regression model as our baseline.\n"
                ],
                "outputs": [],
                "metadata": {}
            },
            {
                "cell_type": "code",
                "source": [
                    "model = LinearRegression()\n",
                    "model.fit(X_train_scaled, y_train)\n\n",
                    "print('Model Coefficients:')\n",
                    "for feat, coef in zip(X.columns, model.coef_):\n",
                    "    print(f'  {feat:<25} {coef:>10.4f}')\n",
                    "print(f'  {\"Intercept\":<25} {model.intercept_:>10.4f}')"
                ],
                "outputs": [
                    {
                        "output_type": "stream",
                        "name": "stdout",
                        "text": [
                            "Model Coefficients:\n",
                            "  advertising_spend          9128.3412\n",
                            "  region_score                516.7834\n",
                            "  seasonality                 998.1209\n",
                            "  Intercept                 24078.9568\n"
                        ]
                    }
                ],
                "metadata": {}
            },
            {
                "cell_type": "markdown",
                "source": [
                    "## 5. Evaluation & Results\n\n",
                    "We evaluate model performance on the held-out test set using standard regression metrics.\n"
                ],
                "outputs": [],
                "metadata": {}
            },
            {
                "cell_type": "code",
                "source": [
                    "y_pred = model.predict(X_test_scaled)\n\n",
                    "mse  = mean_squared_error(y_test, y_pred)\n",
                    "rmse = np.sqrt(mse)\n",
                    "r2   = r2_score(y_test, y_pred)\n\n",
                    "print('=' * 40)\n",
                    "print('       MODEL EVALUATION REPORT')\n",
                    "print('=' * 40)\n",
                    "print(f'  MSE  : {mse:>15,.2f}')\n",
                    "print(f'  RMSE : {rmse:>15,.2f}')\n",
                    "print(f'  R2   : {r2:>15.6f}')\n",
                    "print('=' * 40)\n",
                    "print(f'  Model explains {r2*100:.1f}% of variance')"
                ],
                "outputs": [
                    {
                        "output_type": "stream",
                        "name": "stdout",
                        "text": [
                            "========================================\n",
                            "       MODEL EVALUATION REPORT\n",
                            "========================================\n",
                            "  MSE  :      248,941.23\n",
                            "  RMSE :         499.94\n",
                            "  R2   :       0.998513\n",
                            "========================================\n",
                            "  Model explains 99.9% of variance\n"
                        ]
                    }
                ],
                "metadata": {}
            },
            {
                "cell_type": "markdown",
                "source": [
                    "## 6. Conclusions\n\n",
                    "The linear regression model achieved an exceptional **R² of 0.9985**, confirming that the three engineered features capture nearly all variance in sales.\n\n",
                    "### Key Findings\n\n",
                    "- `advertising_spend` has the strongest absolute impact on sales\n",
                    "- `seasonality` contributes a significant cyclical component\n",
                    "- `region_score` moderates baseline performance\n\n",
                    "### Next Steps\n\n",
                    "1. Collect real-world data and validate assumptions\n",
                    "2. Explore non-linear models (GBM, Random Forest)\n",
                    "3. Implement cross-validation for more robust evaluation\n\n",
                    "> **Note:** This analysis uses synthetic data. Results should not be used for production forecasting without further validation.\n"
                ],
                "outputs": [],
                "metadata": {}
            }
        ]
    }
    with open(path, "w") as f:
        json.dump(nb, f, indent=2)
    print(f"📓 Sample notebook created: {path}")


# ─────────────────────────────────────────────
# CLI ENTRY POINT
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="nb2pdf — Convert Jupyter Notebooks to professional PDFs",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python nb2pdf_agent.py notebook.ipynb
  python nb2pdf_agent.py notebook.ipynb -o report.pdf
  python nb2pdf_agent.py --demo
        """
    )
    parser.add_argument("notebook", nargs="?", help="Path to .ipynb file")
    parser.add_argument("-o", "--output", help="Output PDF path (default: <notebook>.pdf)")
    parser.add_argument("--demo", action="store_true",
                        help="Generate a sample notebook and convert it")
    args = parser.parse_args()

    if args.demo:
        nb_path = "sample_notebook.ipynb"
        create_sample_notebook(nb_path)
        out_path = "sample_output.pdf"
    elif args.notebook:
        nb_path = args.notebook
        if not os.path.exists(nb_path):
            print(f"❌  File not found: {nb_path}")
            sys.exit(1)
        out_path = args.output or os.path.splitext(nb_path)[0] + ".pdf"
    else:
        parser.print_help()
        sys.exit(0)

    print(f"📖 Parsing notebook: {nb_path}")
    nb_parser = NotebookParser(nb_path)
    title = nb_parser.get_title()
    kernel = nb_parser.kernel
    print(f"   Title  : {title}")
    print(f"   Kernel : {kernel}")
    print(f"   Cells  : {len(nb_parser.cells())}")
    print(f"🔨 Building PDF ...")

    builder = PDFBuilder(out_path, title, kernel, os.path.basename(nb_path))
    builder.build(nb_parser)


if __name__ == "__main__":
    main()
