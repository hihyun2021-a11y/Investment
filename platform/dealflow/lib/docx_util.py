# -*- coding: utf-8 -*-
"""Word 문서 공통 서식(안양 RFP·기안문 양식: 맑은 고딕 10~10.5pt, A4)."""
from __future__ import annotations

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT = "맑은 고딕"
NAVY = RGBColor(0x1F, 0x3A, 0x5F)
GREY = RGBColor(0x59, 0x59, 0x59)
CENTER = WD_ALIGN_PARAGRAPH.CENTER


class Doc:
    def __init__(self, size=10, margin=2.2):
        self.d = Document()
        sec = self.d.sections[0]
        sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
        sec.left_margin = sec.right_margin = Cm(margin)
        sec.top_margin = sec.bottom_margin = Cm(2.0)
        st = self.d.styles["Normal"]
        st.font.name = FONT
        st.font.size = Pt(size)
        st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
        st.paragraph_format.space_after = Pt(3)
        st.paragraph_format.line_spacing = 1.25

    def p(self, text="", bold=False, size=None, color=None, align=None, before=None, after=None, indent=0.0, hanging=0.0):
        para = self.d.add_paragraph()
        f = para.paragraph_format
        if indent or hanging:
            f.left_indent = Cm(indent + hanging)
        if hanging:
            f.first_line_indent = Cm(-hanging)
        if before is not None:
            f.space_before = Pt(before)
        if after is not None:
            f.space_after = Pt(after)
        if align is not None:
            para.alignment = align
        for i, t in enumerate(str(text).split("**")):
            if not t:
                continue
            r = para.add_run(t)
            r.bold = bold or (i % 2 == 1)
            r.font.name = FONT
            r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
            if size:
                r.font.size = Pt(size)
            if color is not None:
                r.font.color.rgb = color
        return para

    def h(self, text):
        return self.p(text, bold=True, size=11, color=NAVY, before=10, after=4)

    def bullets(self, items, indent=0.3, mark="•"):
        for t in items:
            self.p(f"{mark} {t}", indent=indent, hanging=0.4)

    def table(self, rows, widths=None, head=True, size=9.5):
        t = self.d.add_table(rows=len(rows), cols=len(rows[0]))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                c = t.cell(i, j)
                c.text = ""
                para = c.paragraphs[0]
                r = para.add_run(str(v if v is not None else ""))
                r.font.size = Pt(size)
                r.font.name = FONT
                r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
                if (head and i == 0) or j == 0:
                    r.bold = head and i == 0
                    shade = "D9D9D9" if head and i == 0 else "F2F2F2"
                    tcPr = c._tc.get_or_add_tcPr()
                    sh = OxmlElement("w:shd")
                    sh.set(qn("w:val"), "clear")
                    sh.set(qn("w:fill"), shade)
                    tcPr.append(sh)
                if head and i == 0:
                    para.alignment = CENTER
        if widths:
            for j, w in enumerate(widths):
                for row in t.rows:
                    row.cells[j].width = Cm(w)
        self.d.add_paragraph()
        return t

    def save(self, path, author="코람코자산신탁"):
        self.d.core_properties.author = author
        self.d.save(str(path))
        return path
