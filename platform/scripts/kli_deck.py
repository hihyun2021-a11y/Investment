# -*- coding: utf-8 -*-
"""KLI 투심 최종본 v5.8 서식 기반 PPT 생성기 (2026-11-03 운용역 지정 — 모든 PPT의 기준 서식).

기준 파일: platform/templates/kli_deck/KLI_투심최종본_v5.8_서식기준_원본.pptx
빈 템플릿: platform/templates/kli_deck/KLI_빈템플릿_v5.8.pptx (슬라이드 마스터·레이아웃·테마 그대로, 슬라이드 0장)
규격 문서: platform/standards/ppt_style_kli_v58.md

v5.8에서 실측한 규격
- 판형 A4 가로 11.69 × 8.27", 서체 KoPubWorld돋움체 (Light 본문 / Medium 강조·첫 열 / Bold 표 머리·밴드)
- 머리말: 레이아웃 섹션명(Medium 18.86pt) + 바로 뒤 소제목 "> …" (Light Bold 11pt, 제목 상자 lIns로 위치)
- 헤드메시지: "✓ " + Medium 14pt, L0.31 T1.14 W10.95 H0.77 (9_흰배경의 회색 구분 바 y=2.05 위)
- 섹션 밴드: 345B86 탭(우측 사선, 흰 글씨 Bold 10pt) + 탭 끝 회색 삼각 + 345B86 1pt 가로선
- 표: 머리행 D9D9D9 · Bold 10pt, 첫 열 F2F2F2 · Medium 10pt, 본문 흰 바탕 · Light 10pt,
      머리행 윗선 404040 0.5pt, 나머지 선 BFBFBF 0.25pt, 바깥 좌우선 없음, 셀 세로 가운데
- 본문 글자 10pt가 기본(전체 run의 약 60%), **최소 10pt**(운용역 지시: 각주 포함 10pt 미만 금지)
- 색: 밴드 345B86, 진네이비 002060, 먹색 1F1F1F/000000, 보조 646464, 강조 빨강 C00000
"""
from __future__ import annotations

import copy
import pathlib

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = pathlib.Path(__file__).resolve().parents[2]
TEMPLATE = ROOT / "platform/templates/kli_deck/KLI_빈템플릿_v5.8.pptx"

BAND = RGBColor(0x34, 0x5B, 0x86)
NAVY = RGBColor(0x00, 0x20, 0x60)
INK = RGBColor(0x1F, 0x1F, 0x1F)
BLACK = RGBColor(0, 0, 0)
GRAY = RGBColor(0x64, 0x64, 0x64)
RED = RGBColor(0xC0, 0x00, 0x00)
GREEN = RGBColor(0x00, 0x6D, 0x52)
ORANGE = RGBColor(0xF7, 0x8E, 0x3F)
TH = RGBColor(0xD9, 0xD9, 0xD9)       # 표 머리행
COL1 = RGBColor(0xF2, 0xF2, 0xF2)     # 표 첫 열
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE_C = RGBColor(0xBF, 0xBF, 0xBF)
LINE_TOP = RGBColor(0x40, 0x40, 0x40)
TAB_DARK = RGBColor(0x7F, 0x7F, 0x7F)

F_LIGHT = "KoPubWorld돋움체 Light"
F_MEDIUM = "KoPubWorld돋움체 Medium"
F_BOLD = "KoPubWorld돋움체 Bold"
MIN_PT = 10.0

HEAD_TOP, HEAD_LEFT, HEAD_W, HEAD_H = 1.14, 0.31, 10.95, 0.77
BAND_LEFT, BAND_W, BAND_H = 0.50, 10.70, 0.27
BODY_LEFT, BODY_W = 0.56, 10.63
BODY_BOTTOM = 7.38
NOTE_TOP = 7.46


def _font(run, name, size, bold=None, color=INK):
    size = max(float(size), MIN_PT)
    run.font.name = name
    run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", name)


def add_run(par, text, *, font=F_LIGHT, size=10, bold=None, color=INK):
    r = par.add_run()
    r.text = text
    _font(r, font, size, bold, color)
    return r


class KLIDeck:
    """v5.8 서식 덱. header: 레이아웃 머리말 섹션명(예: 'VAT 대출 진행경과'), footer: 꼬리말 앞부분."""

    def __init__(self, header: str, footer: str | None = None, template=TEMPLATE,
                 layouts=("9_흰배경", "8_흰배경")):
        self.prs = Presentation(str(template))
        self.L = {l.name: l for l in self.prs.slide_layouts}
        self.header = header
        for name in layouts:          # 출력본 레이아웃의 섹션명만 교체(템플릿 원본 불변)
            for sh in self.L[name].shapes:
                if sh.has_text_frame and sh.name == "TextBox 10":
                    p = sh.text_frame.paragraphs[0]
                    p.runs[0].text = header
                    for r in p.runs[1:]:
                        r.text = ""
        if footer:   # 꼬리말 "안양물류센터 매입 프로젝트 I 본 투자심의위원회 설명자료 | 팀명 쪽번호"의 앞부분 교체
            for m in self.prs.slide_masters:
                for sh in m.shapes:
                    if not (sh.has_text_frame and "설명자료" in sh.text_frame.text):
                        continue
                    for p in sh.text_frame.paragraphs:
                        runs = [r for r in p.runs]
                        full = "".join(r.text for r in runs)
                        if "설명자료" not in full:
                            continue
                        k = next(i for i, r in enumerate(runs) if "|" in r.text)
                        head = "".join(r.text for r in runs[:k]) + runs[k].text.split("|")[0]
                        runs[0].text = footer + " " + runs[k].text[runs[k].text.index("|"):] if k == 0 else footer + " "
                        for r in runs[1:k]:
                            r.text = ""
                        if k > 0:
                            runs[k].text = runs[k].text[runs[k].text.index("|"):]

    # ── 기본 도형 ──────────────────────────────────────────────────────
    @staticmethod
    def box(slide, x, y, w, h, *, anchor=MSO_ANCHOR.MIDDLE, wrap=True):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = anchor
        return tb, tf

    def text(self, slide, x, y, w, h, lines, *, align=PP_ALIGN.LEFT, size=10, font=F_LIGHT, color=INK):
        """lines: str 또는 (str, {font,size,bold,color}) 목록."""
        tb, tf = self.box(slide, x, y, w, h)
        for i, ln in enumerate(lines):
            s, kw = (ln, {}) if isinstance(ln, str) else ln
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            add_run(p, s, font=kw.get("font", font), size=kw.get("size", size),
                    bold=kw.get("bold"), color=kw.get("color", color))
        return tb

    @staticmethod
    def rect(slide, x, y, w, h, *, fill=None, line=None, lw=0.75, dash=False, shape=MSO_SHAPE.RECTANGLE):
        s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        s.shadow.inherit = False
        if fill is None:
            s.fill.background()
        else:
            s.fill.solid()
            s.fill.fore_color.rgb = fill
        if line is None:
            s.line.fill.background()
        else:
            s.line.color.rgb = line
            s.line.width = Pt(lw)
            if dash:
                s.line.dash_style = 4
        tf = s.text_frame
        tf.margin_left = tf.margin_right = Inches(0.04)
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.word_wrap = True
        return s

    @staticmethod
    def hline(slide, x1, y, x2, *, color=LINE_C, lw=0.75, dash=False):
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y), Inches(x2), Inches(y))
        c.line.color.rgb = color
        c.line.width = Pt(lw)
        if dash:
            c.line.dash_style = 4
        return c

    @staticmethod
    def vline(slide, x, y1, y2, *, color=LINE_C, lw=0.5, dash=False):
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y1), Inches(x), Inches(y2))
        c.line.color.rgb = color
        c.line.width = Pt(lw)
        if dash:
            c.line.dash_style = 4
        return c

    # ── 슬라이드 ──────────────────────────────────────────────────────
    def _title_lins(self):
        """머리말(18.86pt) 길이로 소제목 시작 위치 추정 — v5.8 '투자심의위원회 개요' ≈ 2.8\"."""
        w = 0.20
        for ch in self.header:
            w += 0.10 if ch == " " else (0.30 if ord(ch) > 0x3000 else 0.26)
        return w

    def slide(self, subtitle: str, head: str | None = None, *, title_lins: float | None = None):
        """head가 있으면 9_흰배경(헤드메시지+회색 구분 바), 없으면 8_흰배경."""
        title_lins = title_lins or self._title_lins()
        s = self.prs.slides.add_slide(self.L["9_흰배경" if head else "8_흰배경"])
        tb, tf = self.box(s, 0.0, 0.235, 8.25, 0.58)
        tf.margin_left = Emu(int(title_lins * 914400))
        tf.margin_top = Emu(int(0.39 * 914400))
        add_run(tf.paragraphs[0], f"> {subtitle}", font=F_LIGHT, size=11, bold=True, color=BLACK)
        if head:
            _, tf = self.box(s, HEAD_LEFT, HEAD_TOP, HEAD_W, HEAD_H)
            p = tf.paragraphs[0]
            add_run(p, "✓ ", font=F_MEDIUM, size=14, color=BLACK)
            add_run(p, head, font=F_MEDIUM, size=14, color=BLACK)
        return s

    def band(self, slide, label: str, *, top: float, left: float = BAND_LEFT, width: float = BAND_W,
             tab_w: float | None = None):
        """345B86 탭(우측 사선) + 탭 끝 회색 삼각 + 1pt 가로선."""
        tab_w = tab_w or max(1.9, 0.16 * len(label) + 0.55)
        tab_w = min(tab_w, width)
        h = BAND_H
        ff = slide.shapes.build_freeform(Inches(left), Inches(top))
        ff.add_line_segments([(Inches(left + tab_w), Inches(top)), (Inches(left + tab_w - 0.18), Inches(top + h)),
                              (Inches(left), Inches(top + h))], close=True)
        tab = ff.convert_to_shape()
        tab.fill.solid()
        tab.fill.fore_color.rgb = BAND
        tab.line.fill.background()
        tab.shadow.inherit = False
        tri = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(left + tab_w - 0.07), Inches(top),
                                     Inches(0.07), Inches(0.09))
        tri.rotation = 180
        tri.fill.solid()
        tri.fill.fore_color.rgb = TAB_DARK
        tri.line.fill.background()
        tri.shadow.inherit = False
        if width > tab_w:
            self.hline(slide, left + tab_w - 0.02, top + h * 0.48, left + width, color=BAND, lw=1.0)
        self.text(slide, left + 0.10, top, tab_w - 0.30, h, [(label, {"font": F_BOLD, "color": WHITE})])
        return tab

    def table(self, slide, rows, *, top, left=BODY_LEFT, width=BODY_W, col_widths=None, row_h=0.32,
              size=10, first_col=True, align=None, bold_cells=(), color_cells=None, fills=None):
        """v5.8 표 서식. align: 열별 'l'|'c'|'r' 목록(기본 첫 열 c, 나머지 l). fills: {(r,c): RGB}."""
        nr, nc = len(rows), max(len(r) for r in rows)
        gf = slide.shapes.add_table(nr, nc, Inches(left), Inches(top), Inches(width), Inches(row_h * nr))
        tbl = gf.table
        tblPr = gf._element.graphic.graphicData.tbl.tblPr
        for a in ("firstRow", "bandRow"):
            tblPr.set(a, "0")
        sid = tblPr.find(qn("a:tableStyleId"))
        if sid is not None:
            sid.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"   # No Style, Table Grid → 직접 서식
        if col_widths:
            for i, w in enumerate(col_widths):
                tbl.columns[i].width = Inches(w)
        align = align or (["c"] + ["l"] * (nc - 1))
        amap = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
        color_cells = color_cells or {}
        fills = fills or {}
        for ri, row in enumerate(rows):
            tbl.rows[ri].height = Inches(row_h)
            for ci in range(nc):
                cell = tbl.cell(ri, ci)
                cell.margin_left = cell.margin_right = Inches(0.08)
                cell.margin_top = cell.margin_bottom = Inches(0.02)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                head = ri == 0
                fc = fills.get((ri, ci), TH if head else (COL1 if (first_col and ci == 0) else WHITE))
                cell.fill.solid()
                cell.fill.fore_color.rgb = fc
                txt = row[ci] if ci < len(row) else ""
                tf = cell.text_frame
                tf.word_wrap = True
                lines = str(txt).split("\n")
                for li, ln in enumerate(lines):
                    p = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
                    p.alignment = PP_ALIGN.CENTER if head else amap[align[ci]]
                    if head:
                        fnt = F_BOLD
                    elif first_col and ci == 0:
                        fnt = F_MEDIUM
                    else:
                        fnt = F_BOLD if (ri, ci) in bold_cells else F_LIGHT
                    add_run(p, ln, font=fnt, size=size, color=color_cells.get((ri, ci), BLACK))
                self._borders(cell, top_dark=head, left=ci > 0, right=ci < nc - 1)
        return tbl

    @staticmethod
    def _borders(cell, *, top_dark, left, right):
        tcPr = cell._tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            old = tcPr.find(qn(tag))
            if old is not None:
                tcPr.remove(old)

        def ln(tag, w, color):
            el = tcPr.makeelement(qn(tag), {"w": str(w), "cap": "flat", "cmpd": "sng", "algn": "ctr"})
            if color is None:
                el.append(el.makeelement(qn("a:noFill"), {}))
            else:
                sf = el.makeelement(qn("a:solidFill"), {})
                c = sf.makeelement(qn("a:srgbClr"), {"val": str(color)})
                sf.append(c)
                el.append(sf)
            tcPr.append(el)

        ln("a:lnL", 3175, LINE_C if left else None)
        ln("a:lnR", 3175, LINE_C if right else None)
        ln("a:lnT", 6350 if top_dark else 3175, LINE_TOP if top_dark else LINE_C)
        ln("a:lnB", 3175, LINE_C)
        # 스키마 순서: lnL, lnR, lnT, lnB, …, solidFill → 채우기를 뒤로 이동
        fill = tcPr.find(qn("a:solidFill"))
        if fill is not None:
            tcPr.remove(fill)
            tcPr.append(fill)

    def footnote(self, slide, text: str, *, top: float = NOTE_TOP):
        """각주 — 최소 10pt."""
        return self.text(slide, 0.31, top, 10.95, 0.26, [(text, {"color": GRAY})])

    def save(self, path):
        path = pathlib.Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(path))
        return path

    # ── 점검 ──────────────────────────────────────────────────────────
    def check_min_font(self):
        """MIN_PT 미만 run 목록 반환(빈 목록이어야 정상)."""
        bad = []
        for i, s in enumerate(self.prs.slides, 1):
            for sh in s.shapes:
                frames = []
                if sh.has_text_frame:
                    frames.append(sh.text_frame)
                if getattr(sh, "has_table", False) and sh.has_table:
                    frames += [c.text_frame for r in sh.table.rows for c in r.cells]
                for tf in frames:
                    for p in tf.paragraphs:
                        for r in p.runs:
                            if r.text.strip() and r.font.size is not None and r.font.size.pt < MIN_PT:
                                bad.append((i, r.text[:20], r.font.size.pt))
        return bad
