#!/usr/bin/env python3
"""안양물류센터 본 투자심의 자료 v05 — 표 서식 통일.

- 전 슬라이드 표 글자 10pt 통일, 표 셀 글머리기호(•) 제거
- 10pt 기준으로 열 폭·행 높이·비교자산 수 재조정

[v04 이력]

0. 표 생성 방식 교체: graphicFrame 통째 복제 → python-pptx add_table + 원본 서식 이식
   (v03의 복제 표가 PowerPoint에서 빈 칸으로 렌더되던 문제 해결)
1. 심의안건 3p: v02의 추진일정 타임라인 장표 복원
2. 안건 Ⅱ·Ⅲ(주주간계약서·매입확약서) 주요내용 작성, 매도확약서 내용 추가

[v03 이력]

1. 개요(3p): 양해각서 타임라인 삭제 → 안건 3건 주요내용 요약표
2. 캡처 이미지 → 실제 표(원본 표 서식 복제) + 도형·지도·차트는 원본 좌표 300dpi 클립 렌더, 헤드메시지 2줄 이내
3. 기존 Peer 임대·매매 표 → 시장조사(Savills) 임대사례·거래사례 실제 표(사진 포함)로 대체 (중복 제거)
4. 투자수익률 2p 실제 표 복원(IM 수치), Sell-down 2p를 동일 표 양식(FY1~20)으로 작성

실행: python3 projects/ANYANG-LOGIS/03_IM_투자심의/투심자료/build_deck_v3.py
"""
import copy
import json
import pathlib
import re
import shutil
import subprocess
import sys
import zipfile

import openpyxl
import pymupdf
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Emu, Pt

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from build_deck_v2 import (NS, I, iter_shapes, by_text, by_y, set_tf, set_cell, replace_runs,  # noqa: E402
                           delete, strip_body, band_label, title, lead, clone_table, src_table, fn_add,
                           BASE, SKILL, IM_PDF, MKT_PDF, MODEL, PROJ, HERE)

WORK = pathlib.Path("/tmp/claude-0/-home-user-Investment/38052696-05de-5d1d-9b55-e347b5278ed0/scratchpad/v3")
IMG = WORK / "img"
OUT = HERE / "20260916_투심자료_ANYANG-LOGIS_v05.pptx"
LEAD_MAX, SUB_MAX = 58, 78


# ═════════════════════════════════════════════════════════════════════════
# 0. 이미지: PDF 원본 좌표 클립 렌더 (300dpi)
# ═════════════════════════════════════════════════════════════════════════
def clip(pdf, page, rect, name, dpi=300):
    IMG.mkdir(parents=True, exist_ok=True)
    out = IMG / f"{name}.png"
    if not out.exists():
        pg = pymupdf.open(str(pdf))[page - 1]
        pg.get_pixmap(clip=pymupdf.Rect(*rect), dpi=dpi).save(str(out))
    return out


def add_pic(slide, png, left, top, width=None, height=None):
    im = Image.open(png)
    ar = im.width / im.height
    if width is None:
        width = height * ar
    elif height is None:
        height = width / ar
    return slide.shapes.add_picture(str(png), Inches(left), Inches(top), Inches(width), Inches(height))


def photo_cells(pdf, page):
    """페이지 내 사진 래스터의 사각형을 x 순으로 반환."""
    pg = pymupdf.open(str(pdf))[page - 1]
    rects = []
    for im in pg.get_images(full=True):
        if im[2] < 150 or im[3] < 100:
            continue
        for r in pg.get_image_rects(im[0]):
            if r.width > 60 and r.height > 50:
                rects.append(r)
    rects.sort(key=lambda r: r.x0)
    return rects


# ═════════════════════════════════════════════════════════════════════════
# 1. 구조
# ═════════════════════════════════════════════════════════════════════════
def structure():
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)
    un = WORK / "unpacked"
    zipfile.ZipFile(BASE).extractall(un)

    def dup(src, after):
        r = subprocess.run([sys.executable, str(SKILL / "add_slide.py"), str(un), src, "--after", after],
                           capture_output=True, text=True, check=True)
        return re.search(r"slides/(slide\d+\.xml)", r.stdout).group(1)

    a1 = dup("slide4.xml", "slide5.xml"); a2 = dup("slide4.xml", a1); dup("slide4.xml", a2)      # 안건Ⅱ×2, Ⅲ
    l1 = dup("slide16.xml", "slide16.xml"); l2 = dup("slide16.xml", l1); dup("slide16.xml", l2)  # 입지 보완 3
    m1 = dup("slide17.xml", "slide20.xml"); dup("slide17.xml", m1)                                # 시장 보완 2
    dup("slide22.xml", "slide22.xml")                                                            # 적정임대료 결론
    dup("slide24.xml", "slide24.xml")                                                            # 사업구조도
    d1 = dup("slide31.xml", "slide32.xml"); dup("slide32.xml", d1)                                # Sell-down 2
    dup("slide28.xml", "slide28.xml")                                                            # 운영가정 분할
    pres = un / "ppt/presentation.xml"
    x = pres.read_text(encoding="utf-8")
    rels = (un / "ppt/_rels/presentation.xml.rels").read_text(encoding="utf-8")
    for sl in ("slide6.xml", "slide7.xml"):
        rid = (re.search(r'Id="(rId\d+)"[^>]*Target="slides/%s"' % sl, rels)
               or re.search(r'Target="slides/%s"[^>]*Id="(rId\d+)"' % sl, rels)).group(1)
        x = re.sub(r'<p:sldId [^>]*r:id="%s"[^>]*/>' % rid, "", x)
    pres.write_text(x, encoding="utf-8")
    subprocess.run([sys.executable, str(SKILL / "clean.py"), str(un)], check=True, capture_output=True)
    base = WORK / "base.pptx"
    subprocess.run(["zip", "-qXr", str(base), "."], cwd=un, check=True)
    return base


# ═════════════════════════════════════════════════════════════════════════
# 2. 도우미
# ═════════════════════════════════════════════════════════════════════════
def lead2(slide, main, sub=None):
    assert len(main) <= LEAD_MAX, f"헤드메시지 길이 초과({len(main)}): {main}"
    if sub:
        assert len(sub) <= SUB_MAX, f"보조문장 길이 초과({len(sub)}): {sub}"
    lead(slide, main, [sub] if sub else [])


def band_group(slide):
    g = [sh for sh in slide.shapes if sh.shape_type == 6 and
         (abs((sh.top or 0) / I - 2.17) < 0.06 or abs((sh.top or 0) / I - 1.18) < 0.06)]
    return g[0] if g else None


def _band_textbox(g):
    """밴드 그룹에서 실제 라벨 텍스트박스(TEXT_BOX=17, 가장 넓은 것)를 고른다. 장식 사각형은 제외."""
    tbs = [x for x in iter_shapes(g.shapes) if x.has_text_frame and x.shape_type == 17]
    if not tbs:
        tbs = [x for x in iter_shapes(g.shapes) if x.has_text_frame and x.text_frame.text.strip()]
    return max(tbs, key=lambda x: x.width or 0) if tbs else None


def band_label(slide, text):  # v2 판을 덮어씀: 두 위치 모두 지원 + 라벨 박스만 수정
    g = band_group(slide)
    if g:
        tb = _band_textbox(g)
        if tb is not None:
            set_tf(tb.text_frame, text)


def add_band(slide, top, text):
    src = band_group(slide)
    el = copy.deepcopy(src._element)
    slide.shapes._spTree.append(el)
    g = slide.shapes[-1]
    g.top = Inches(top)
    tb = _band_textbox(g)
    if tb is not None:
        set_tf(tb.text_frame, text)
    return g



# ═════════════════════════════════════════════════════════════════════════
# 2-b. 표 생성 (네이티브 add_table + 원본 서식 이식)
# ═════════════════════════════════════════════════════════════════════════
def _src_cells(tmpl_gf):
    """템플릿 표의 (헤더행 tc 목록, 본문행 tc 목록)."""
    tbl = tmpl_gf._element.find(".//a:tbl", NS)
    trs = tbl.findall("a:tr", NS)
    return tbl, trs[0].findall("a:tc", NS), trs[min(1, len(trs) - 1)].findall("a:tc", NS)


def _fill_cell(cell, src_tc, text, size=None):
    """원본 셀의 tcPr(테두리·배경)과 문단·런 서식을 이식하고 text를 넣는다."""
    tc = cell._tc
    src_pr = src_tc.find("a:tcPr", NS)
    if src_pr is not None:
        new_pr = copy.deepcopy(src_pr)
        for k in ("gridSpan", "rowSpan", "hMerge", "vMerge"):
            new_pr.attrib.pop(k, None)
        for ext in new_pr.findall("a:extLst", NS):
            new_pr.remove(ext)
        old = tc.find("a:tcPr", NS)
        if old is not None:
            tc.replace(old, new_pr)
        else:
            tc.append(new_pr)
    # 원본 문단/런 서식
    sp = src_tc.find(".//a:p", NS)
    pPr = copy.deepcopy(sp.find("a:pPr", NS)) if sp is not None and sp.find("a:pPr", NS) is not None else None
    sr = sp.find("a:r", NS) if sp is not None else None
    rPr = copy.deepcopy(sr.find("a:rPr", NS)) if sr is not None and sr.find("a:rPr", NS) is not None else None
    if rPr is not None and size:
        rPr.set("sz", str(int(size * 100)))
    txBody = cell.text_frame._txBody
    for old_p in txBody.findall("a:p", NS):
        txBody.remove(old_p)
    lines = text.split("\n") if isinstance(text, str) else list(text)
    for line in (lines or [""]):
        el = txBody.makeelement("{%s}p" % NS["a"], {})
        if pPr is not None:
            el.append(copy.deepcopy(pPr))
        if line:
            r = el.makeelement("{%s}r" % NS["a"], {})
            if rPr is not None:
                r.append(copy.deepcopy(rPr))
            t = r.makeelement("{%s}t" % NS["a"], {})
            t.text = line
            r.append(t)
            el.append(r)
        txBody.append(el)


TABLE_PT = 10.0  # 전 표 공통 글자 크기


def clone_table(slide, tmpl_gf, rows, col_widths, row_h=0.30, header_h=None,
                size=TABLE_PT, top=2.58, left=0.50, **_):
    """원본 표 서식을 이식한 새 표를 만든다 (v3의 프레임 복제 방식을 대체)."""
    nr, nc = len(rows), len(col_widths)
    gf = slide.shapes.add_table(nr, nc, Inches(left), Inches(top),
                                Inches(sum(col_widths)), Inches(row_h * nr))
    tb = gf.table
    src_tbl, hdr_tcs, body_tcs = _src_cells(tmpl_gf)
    # 표 속성(스타일 ID·firstRow/bandRow 플래그) 이식
    src_pr = src_tbl.find("a:tblPr", NS)
    new_tbl = gf._element.find(".//a:tbl", NS)
    if src_pr is not None:
        cur = new_tbl.find("a:tblPr", NS)
        new_pr = copy.deepcopy(src_pr)
        for ext in new_pr.findall("a:extLst", NS):
            new_pr.remove(ext)
        if cur is not None:
            new_tbl.replace(cur, new_pr)
        else:
            new_tbl.insert(0, new_pr)
    for j, w in enumerate(col_widths):
        tb.columns[j].width = Inches(w)
    for i in range(nr):
        tb.rows[i].height = Inches(header_h if (i == 0 and header_h) else row_h)
        src_tcs = hdr_tcs if i == 0 else body_tcs
        for j in range(nc):
            val = rows[i][j] if j < len(rows[i]) else ""
            _fill_cell(tb.cell(i, j), src_tcs[min(j, len(src_tcs) - 1)], val, size)
    gf.name = "신규표"
    return gf



def _no_bullet(pPr):
    """문단 속성에서 글머리기호를 제거하고 buNone을 넣는다."""
    for tag in ("buChar", "buAutoNum", "buBlip"):
        for e in pPr.findall("a:" + tag, NS):
            pPr.remove(e)
    if pPr.find("a:buNone", NS) is None:
        bn = pPr.makeelement("{%s}buNone" % NS["a"], {})
        anchor = None
        for tag in ("tabLst", "defRPr", "extLst"):
            f = pPr.find("a:" + tag, NS)
            if f is not None:
                anchor = f
                break
        if anchor is not None:
            anchor.addprevious(bn)
        else:
            pPr.append(bn)


def normalize_tables(prs, pt=TABLE_PT):
    """전 슬라이드 표: 글자 크기 통일 + 글머리기호 제거 + 선행 '• ' 문자 제거."""
    n_sz = n_bu = 0
    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_table:
                continue
            for row in sh.table.rows:
                for cell in row.cells:
                    for para in cell.text_frame.paragraphs:
                        pPr = para._p.find("a:pPr", NS)
                        if pPr is None:
                            pPr = para._p.makeelement("{%s}pPr" % NS["a"], {})
                            para._p.insert(0, pPr)
                        _no_bullet(pPr)
                        n_bu += 1
                        runs = para.runs
                        for k, r in enumerate(runs):
                            if k == 0 and r.text.startswith("• "):
                                r.text = r.text[2:]
                            elif k == 0 and r.text.startswith("•"):
                                r.text = r.text[1:].lstrip()
                            r.font.size = Pt(pt)
                            n_sz += 1
    print(f"  표 정규화: 글자 {n_sz}런 {pt}pt, 문단 {n_bu}개 글머리기호 제거")



def _em_w(ch):
    o = ord(ch)
    if 0xAC00 <= o <= 0xD7A3 or 0x3131 <= o <= 0x318E or 0x4E00 <= o <= 0x9FFF or ch in "·—–…※①②③④":
        return 1.0
    if ch in " .,:/()[]|-":
        return 0.32
    return 0.55


def _nat_w(txt, pt):
    """줄바꿈 단위 최장 자연폭(inch)."""
    em = pt / 72.0
    return max((sum(_em_w(c) for c in seg) * em for seg in txt.split("\n")), default=0.0)


def _lines(txt, cw, pt):
    em = pt / 72.0
    out = 0
    for seg in txt.split("\n"):
        if not seg:
            out += 1
            continue
        cur, ln = 0.0, 1
        for ch in seg:
            a = _em_w(ch) * em
            if cur + a > cw:
                ln += 1
                cur = a
            else:
                cur += a
        out += ln
    return out


def autofit_tables(prs, pt=TABLE_PT, bottom=7.42, min_row=0.175, pad=0.12):
    """10pt 기준으로 열 폭을 내용에 맞게 재배분하고 행 높이를 최소 필요치로 맞춘다."""
    fixed = []
    for si, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if not sh.has_table:
                continue
            tb = sh.table
            top = (sh.top or 0) / I
            total_w = sum(c.width for c in tb.columns) / I
            nr, nc = len(tb.rows), len(tb.columns)

            def layout(widths, keep_declared=False):
                hs = []
                for row in tb.rows:
                    need = (row.height / I) if keep_declared else min_row
                    for ci, cell in enumerate(row.cells):
                        t = cell.text
                        if not t.strip():
                            continue
                        ln = _lines(t, max(0.22, widths[ci] - pad), pt)
                        need = max(need, ln * (pt * 1.12 / 72) + 0.015)
                    hs.append(need)
                return hs

            cur = [c.width / I for c in tb.columns]
            if top + sum(layout(cur, keep_declared=True)) <= bottom:
                continue
            if top + sum(layout(cur)) <= bottom:       # 현재 열 폭으로 행 높이만 압축하면 해결
                hs = layout(cur)
                for ri in range(nr):
                    tb.rows[ri].height = Emu(int(hs[ri] * I))
                fixed.append((si, sh.name, round(top + sum(hs), 2)))
                continue
            # 열별 자연폭 기준 재배분 (최소 0.5", 헤더 제외 본문 기준)
            want = []
            for ci in range(nc):
                mx = 0.0
                for ri, row in enumerate(tb.rows):
                    t = row.cells[ci].text
                    if t.strip():
                        mx = max(mx, _nat_w(t, pt))
                want.append(max(0.5, min(mx + pad, total_w * 0.42)))
            scale = total_w / sum(want)
            widths = [w * scale for w in want]
            hs = layout(widths)
            if top + sum(hs) > bottom:
                # 여전히 넘치면 가장 높이 요구가 큰 열을 더 넓힘
                for _ in range(6):
                    ri = max(range(nr), key=lambda r: hs[r])
                    ci = max(range(nc), key=lambda c: _lines(tb.rows[ri].cells[c].text,
                                                             max(0.22, widths[c] - pad), pt))
                    donor = min(range(nc), key=lambda c: want[c] if c != ci else 9e9)
                    if widths[donor] - 0.25 < 0.5:
                        break
                    widths[donor] -= 0.25
                    widths[ci] += 0.25
                    hs = layout(widths)
                    if top + sum(hs) <= bottom:
                        break
            for ci in range(nc):
                tb.columns[ci].width = Emu(int(widths[ci] * I))
            for ri in range(nr):
                tb.rows[ri].height = Emu(int(hs[ri] * I))
            fixed.append((si, sh.name, round(top + sum(hs), 2)))
    if fixed:
        print("  표 자동보정:", ", ".join(f"[{a}]{b}→{c}" for a, b, c in fixed))



def set_table_width(gf, width_in, left_in=None):
    """표 전체 폭을 바꾸되 열 폭을 비례 확대/축소한다 (프레임만 바꾸면 열은 그대로)."""
    tb = gf.table
    cur = sum(c.width for c in tb.columns)
    k = (width_in * I) / cur if cur else 1.0
    for c in tb.columns:
        c.width = Emu(int(c.width * k))
    gf.width = Emu(int(width_in * I))
    if left_in is not None:
        gf.left = Emu(int(left_in * I))


def set_row_height(table_shape, row_idx, h):
    trs = table_shape._element.findall(".//a:tr", NS)
    trs[row_idx].set("h", str(int(h * I)))


def table_rows_from(rows_json, key):
    return rows_json[key]


def cells_text(line, xmin=0, xmax=9999):
    return [t for x, t in line if xmin <= x < xmax]


# ═════════════════════════════════════════════════════════════════════════
# 3. 데이터
# ═════════════════════════════════════════════════════════════════════════
def im_yield_tables(rows_json):
    """IM p.20/21 → {'1종':[[FY,잔액,배당,CG,합계,Yield]...,합계,투자금액,CoC,IRR,EM], ...}"""
    out = {}
    for key, names in (("im20", ("1종", "2종")), ("im21", ("3종", "보통주"))):
        L = {n: [] for n in names}
        for line in rows_json[key]:
            left = [t for x, t in line if x < 440]
            right = [t for x, t in line if x >= 440]
            for n, part in zip(names, (left, right)):
                if not part:
                    continue
                h = part[0]
                if re.match(r"FY \d+", h) or h in ("합계", "투자금액", "CoC_C.G제외", "IRR_C.G포함", "E.Multiple"):
                    L[n].append(part)
        out.update(L)
    return out


def selldown_periods():
    wb = openpyxl.load_workbook(MODEL, data_only=True)
    ws = wb["Sell-down"]
    c0 = 12
    per = []
    for r in range(68, 88):
        per.append([ws.cell(r, c0 + 2 + j).value or 0 for j in range(16)])  # 4그룹 × (잔액,배당,CG,yield)
    summ = {ws.cell(r, c0 + 1).value: [ws.cell(r, c0 + 5 + 4 * g).value for g in range(4)] for r in (90, 91, 92, 93)}
    return per, summ


def selldown_table(per, summ, g):
    """그룹 g(0:1종,1:2종,2:보통주,3:통합) → 26행 표 (구분,투자잔액,운영배당,C.G,합계,Yield)."""
    m = lambda v: v / 1000.0
    rows = [["구분", "투자잔액", "운영배당", "C.G", "합계", "Yield(%)"]]
    tot_d = tot_c = 0.0
    for k, p in enumerate(per, 1):
        bal, div, cg, yld = m(p[4 * g]), m(p[4 * g + 1]), m(p[4 * g + 2]), p[4 * g + 3]
        tot_d += div; tot_c += cg
        f = lambda v: f"{v:,.0f}" if abs(v) >= 0.5 else "-"
        rows.append([f"FY {k}", f(bal), f(div), f(cg), f(div + cg), (f"{yld:.1%}" if bal else "-")])
    f = lambda v: f"{v:,.0f}" if abs(v) >= 0.5 else "-"
    rows.append(["합계", "", f(tot_d), f(tot_c), f(tot_d + tot_c), ""])
    last_bal = m(per[-1][4 * g])
    rows.append(["투자금액", f"{last_bal:,.0f}", "", "", "", ""])
    rows.append(["CoC_C.G제외", f"{summ['CoC(%)_C.G제외'][g]:.2%}", "", "", "", ""])
    rows.append(["IRR_C.G포함", f"{summ['IRR(%)_C.G포함'][g]:.2%}", "", "", "", ""])
    rows.append(["E.Multiple", f"x {summ['E. Multiple'][g]:.2f}", "", "", "", ""])
    return rows


def fill_yield_table(tbl_shape, rows):
    """26x6 원본 표에 rows(26개) 채우기. 요약행은 원본처럼 값이 마지막 열 근처(병합)에 있으므로 값 셀 위치 유지."""
    tb = tbl_shape.table
    for i, r in enumerate(rows):
        if i < len(tb.rows):
            for j in range(6):
                if j < len(r):
                    set_cell(tb.cell(i, j), r[j])


def fill_im_table(tbl_shape, lines):
    """IM 파싱 행(가변 열)을 원본 26x6 표에 맞춰 채움."""
    tb = tbl_shape.table
    fy = [l for l in lines if re.match(r"FY \d+", l[0])]
    rest = {l[0]: l for l in lines if not re.match(r"FY \d+", l[0])}
    for i, l in enumerate(fy, 1):
        vals = l[1:]
        # 셀 수가 5 미만이면 왼쪽부터 채우고 나머지 '-'
        vals = (vals + ["-"] * 5)[:5]
        for j, v in enumerate(vals, 1):
            set_cell(tb.cell(i, j), v)
    # 합계 (원본: 운영배당·C.G·합계 위치 = col 2,3,4)
    tot = rest.get("합계", ["합계"])[1:]
    if len(tot) == 3:
        set_cell(tb.cell(21, 2), tot[0]); set_cell(tb.cell(21, 3), tot[1]); set_cell(tb.cell(21, 4), tot[2])
    elif len(tot) == 2:  # (배당, 합계) with C.G '-'
        set_cell(tb.cell(21, 2), tot[0]); set_cell(tb.cell(21, 3), "-"); set_cell(tb.cell(21, 4), tot[1])
    else:
        set_cell(tb.cell(21, 2), "-"); set_cell(tb.cell(21, 3), "-"); set_cell(tb.cell(21, 4), "-")
    for i, key in ((22, "투자금액"), (23, "CoC_C.G제외"), (24, "IRR_C.G포함"), (25, "E.Multiple")):
        v = rest.get(key, [key, ""])[-1]
        set_cell(tb.cell(i, 1), v)


# ═════════════════════════════════════════════════════════════════════════
# 4. 본문
# ═════════════════════════════════════════════════════════════════════════
def build():
    rows_json = json.load(open(WORK.parent / "v3/rows.json", encoding="utf-8")) if (WORK.parent / "v3/rows.json").exists() else None
    rows_path = pathlib.Path("/tmp/claude-0/-home-user-Investment/38052696-05de-5d1d-9b55-e347b5278ed0/scratchpad/rows_v3.json")
    if rows_json is None:
        rows_json = json.load(open(rows_path, encoding="utf-8"))
    base = structure()
    prs = Presentation(str(base))
    S = list(prs.slides)
    assert len(S) == 55, len(S)
    tmpl = src_table(S[3])
    per, summ = selldown_periods()
    imt = im_yield_tables(rows_json)

    # ── 표지·목차 ──────────────────────────────────────────────────────
    replace_runs(S[0], "예비투자심의위원회", "본 투자심의위원회"); replace_runs(S[0], "August. 2026", "September. 2026")
    replace_runs(S[1], "예비투자심의위원회", "본 투자심의위원회")
    toc = by_text(S[1], "투자심의위원회 심의 개요")
    set_tf(toc.text_frame, ["투자심의위원회 심의 개요", "투자심의위원회 심의 안건", "안건 Ⅰ. 매매계약서의 체결",
                            "안건 Ⅱ. 주주간계약서의 체결", "안건 Ⅲ. 코람코라이프인프라리츠 매입확약서(LOC) 날인",
                            "참고. 관계사 인수확약서 및 당사 예상수익", "[별첨] 투자 대상 물건 설명자료"])

    # ── 심의안건 및 추진일정 (v02 장표 유지: 타임라인 보존) ──────────────
    s = S[2]
    t = src_table(s)
    tbl = t._element.find(".//a:tbl", NS)          # 3x3 → 4x3
    trs = tbl.findall("a:tr", NS)
    tbl.append(copy.deepcopy(trs[-1]))
    tb = t.table
    vals = [["안건 Ⅰ", "매매계약서의 체결", "코크렙안양㈜와 매매대금 2,700억원의 부동산매매계약 체결 (거래종결 10/22 예정)"],
            ["안건 Ⅱ", "주주간계약서의 체결", "우리투자증권·삼성증권·㈜엘에프·KLI리츠 간 주주간계약 체결 (배당·매입확약·지배구조)"],
            ["안건 Ⅲ", "KLI리츠 매입확약서(LOC) 날인", "제1종(24개월)·제2종(12개월) 종류주식 매입확약서를 주주간계약 체결일에 제출"]]
    for i, v in enumerate(vals, 1):
        for j in range(3):
            set_cell(tb.cell(i, j), v[j])
    for i, tr in enumerate(tbl.findall("a:tr", NS)):   # 타임라인(4.56") 침범 방지
        tr.set("h", str(int((0.26 if i == 0 else 0.55) * I)))
    for old, new_ in [("8월 12일", "8월 18일"), ("8월 [27]일", "8월 27일"), ("9월 2일", "8월 28일"),
                      ("9월 10일", "9월 3일"), ("9월 21일", "9월 17일"), ("10월 7일", "10월 2~13일"),
                      ("10월 29일", "10월 26일"), ("[8/24]", "[9/30]")]:
        replace_runs(s, old, new_)
    # 타임라인 라벨만 교체 (리드 문장이 덮이지 않도록 y>4.0 도형으로 한정)
    def _tl(sub):
        c = [x for x in iter_shapes(s.shapes)
             if x.has_text_frame and sub in x.text_frame.text and (x.top or 0) / I > 4.0]
        return c[0] if c else None
    for sub, lines in [("실사완료 및", ["영업인가 접수(9/11)", "실사 최종본(9/17)"]),
                       ("영업인가 완료", ["매매계약 체결(10/2)", "영업인가 완료(10/12)", "주주간·대출약정(10/13)"]),
                       ("투자심의위원회 개최\n• 매매계약서", ["투자심의위원회 개최", "• 매매계약·주주간계약 체결의 건", "• 매입확약서(LOC) 날인의 건"]),
                       ("• 양해각서 체결의 건", ["예비투자심의위원회", "• 양해각서 체결의 건", "• 실사기관 선정의 건"]),
                       ("이행보증금", ["양해각서 체결"]),
                       ("회사설립", ["발기설립 완료"]),
                       ("사업일정은 협의에 따라", ["사업일정은 협의에 따라 변동될 수 있음.",
                                            "기준 일정: KLI 신규투자(안양) 추진일정(2026-08-11)",
                                            "코람코라이프인프라리츠 설립자본금 3억원 출자 및 발기설립 완료(8/28)"])]:
        sh = _tl(sub)
        if sh:
            set_tf(sh.text_frame, lines)
    g15 = [sh for sh in s.shapes if sh.shape_type == 6
           and any("당 투자심의위원회" in x.text_frame.text for x in iter_shapes(sh.shapes) if x.has_text_frame)]
    if g15:
        g15[0].left = Inches(6.90)
    lead2(s, "본건은 매매계약·주주간계약 체결 및 매입확약서(LOC) 날인 3개 안건임",
          "매매계약 10/2 → 영업인가 완료 10/12 → 주주간계약·대출약정 10/13 → 거래종결 10/22 (기준 일정)")

    # ── 4~8. 안건 슬라이드 (v02 동일, 헤드 2줄) ──────────────────────────
    s = S[3]
    title(s, "> 안건 Ⅰ. 매매계약서 체결의 건 (1/2)"); band_label(s, "매매계약 주요 조건")
    tb = src_table(s).table
    rows = [["구분", "주요내용"],
            ["체결 예정일", "2026년 10월 2일 (기준 일정) — 계약서상 체결일·거래종결예정일은 공란 [협의 중]"],
            ["체결당사자", "매도인: 코크렙안양 주식회사 (PFV) / 매수인: 신설 위탁관리부동산투자회사 (영업인가 진행 중)"],
            ["매매목적물", "경기도 안양시 동안구 관양동 934 외 2필지 토지 및 지상 건물, 부속 동산·시설·설비 일체"],
            ["면적개요", "대지 15,287.5㎡ (4,624.47평) / 연면적 95,474.59㎡ (28,881.06평) / 임대면적 69,058.12㎡ (20,890.08평)"],
            ["매매대금", "금 2,700억원 (건물·동산 부가가치세 별도) — 담보 감정평가액 3,345억원 대비 −19.3%, 연면적 평당 935만원"],
            ["매매 방식", "거래종결일 현재의 법률상·사실상 현황 그대로(as-is, where-is) 매수"],
            ["거래종결시 지급금액", "매매대금 − 승계 임대차보증금 ± 정산금(세금·부담금·유틸리티·임대료 일할 정산) + 부가가치세"],
            ["임대차 승계", "매수인이 임대차계약 전부 승계, 보증금 반환의무 인수 (승계동의서 미징구가 거래종결에 영향 없음)"],
            ["매도인의 진술 및 보장", "설립·존속, 권한, 제3자 동의, 상충 없음, 완전한 소유권(허용된 부담 제외), 소송 부존재 — 거래종결일로부터 2개월간 유효"],
            ["손해배상책임 한도", "매매대금의 5% 상당액 (135억원), 청구기한 거래종결일로부터 2개월"],
            ["거래신고 / 준거법", "계약일 30일 이내 부동산 거래신고 / 대한민국 법, 서울중앙지방법원 전속관할"]]
    tblel = src_table(s)._element.find(".//a:tbl", NS)
    for tr in tblel.findall("a:tr", NS)[len(rows):]:
        tblel.remove(tr)
    for i, v in enumerate(rows):
        set_cell(tb.cell(i, 0), v[0]); set_cell(tb.cell(i, 1), v[1])
    for i, tr in enumerate(tblel.findall("a:tr", NS)):
        tr.set("h", str(int((0.30 if i == 0 else 0.46) * I)))
    fn = by_y(s, 7.45)
    set_tf(fn[0].text_frame, "※ 출처: 부동산매매계약서(clean, 2026-09-07) 제1~5조, 제9조 / 감정평가서(미래새한, 기준시점 2026-08-31). 상기 내용 중 일부는 협의 과정에서 변동될 수 있음.")

    s = S[4]
    title(s, "> 안건 Ⅰ. 매매계약서 체결의 건 (2/2)"); band_label(s, "주요 조항 및 검토 의견")
    delete(src_table(s))
    rows = [["조항", "주요 내용", "검토 의견"],
            ["제4조 진술 및 보장", "매도인: 설립·존속, 권한, 제3자 동의, 상충 없음, 완전한 소유권, 소송 부존재\n매수인: 권한, 실사 완료, 담보책임 배제 (as-is 매입 확인)", "실사 완료 진술로 미발견 하자는 매수인 부담 → 실사 결과의 계약 반영 필요"],
            ["제4.4조 존속기간", "진술·보장은 거래종결일로부터 2개월간 유효, 기간 내 서면청구 없으면 면책", "통상 대비 짧은 기간 — 잔여 리스크 인지"],
            ["제5조 확약", "거래종결 노력, 특정사항 통지, 거래신고, 임대차 승계, 매도인 운영 유지·협상 금지·세금 완납", "승계동의서 확보 현황 별도 관리"],
            ["제6조 선행조건", "확약·의무 이행, 진술·보장의 진실성, 중대한 부정적 법규·행정절차·소송 부존재", "영업인가·자금조달은 선행조건 아님 → 기준 일정 관리로 대응"],
            ["제7조 손해배상", "한도 매매대금의 5%, 청구기한 거래종결일로부터 2개월, 간접·특별·징벌적 손해 배제", "한도 135억원 수준"],
            ["제8조 해제", "서면 합의 / 불가항력 30일 지속 / 중요 위반 후 10영업일 미시정 / 도산절차 / 선행조건 미충족\n위약금 금액 [협의 중]", "자금조달 곤란은 불가항력에서 명시적으로 제외 — 조달 확정성 확보 필수"],
            ["제9조 일반조항", "양도 제한, 비밀유지(종결 후 1년), 비용 각자 부담, 준거법·전속관할", "—"]]
    clone_table(s, tmpl, rows, [1.7, 6.2, 2.9], row_h=0.70, header_h=0.30, top=1.59)
    fn = by_y(s, 7.45)
    set_tf(fn[0].text_frame, "※ 출처: 부동산매매계약서(clean, 2026-09-07) 제4~9조. 본 검토는 내부 검토이며 최종 법률자문은 외부 법무법인 확인이 필요함.")

    s = S[5]
    title(s, "> 안건 Ⅱ. 주주간계약서 체결의 건 (1/2)"); band_label(s, "주주 구성 및 배당 구조")
    delete(src_table(s))
    rows = [["구분", "금액(백만원)", "지분율", "배당", "매각차익 배분", "투자자"],
            ["제1종 종류주식", "24,000", "24.3%", "누적확정 연 7.0%", "없음", "우리투자증권 (LOC 완료)"],
            ["제2종 종류주식", "50,000", "50.6%", "누적확정 연 7.5%", "20%", "삼성증권 (LOC 완료)"],
            ["제3종 종류주식", "10,000", "10.1%", "무배당·무의결권", "없음", "㈜엘에프 (매도인 재투자)"],
            ["보통주식", "14,800", "15.0%", "잔여배당 (초기 1년 연 7.5%)", "80%", "코람코라이프인프라리츠"],
            ["합계", "98,800", "100.0%", "—", "100%", "—"]]
    clone_table(s, tmpl, rows, [1.9, 1.4, 1.1, 2.4, 1.5, 2.5], row_h=0.36, header_h=0.30, top=1.59)
    add_band(s, 3.95, "자본거래 구조 (유상증자·유상감자)")
    rows2 = [["구분", "주요 내용"],
             ["배당 순위 (제1.2조)", "제1종 연 7.0%·제2종 연 7.5% 누적배당 → 제3종 무배당 → 보통주 잔여배당. 보통주는 초기 1년간 발행가액의 연 7.5% 배당에 협조(자본준비금 감액 포함)"],
             ["유상증자 (제2.1~2.2조)", "매매계약상 거래종결일 2영업일 전까지 이사회 결의, 주주별 납입예정일 15:00까지 출자금 전액 납입"],
             ["유상감자 (제2.3조)", "증자 완료 후 발기인(코람코라이프인프라리츠) 보유 보통주식 300,000주 전부를 감자비율 100%(1주당 1,000원)로 유상감자"],
             ["제3종 유상감자 (제5조)", "발행일로부터 24개월 내 유상감자로 소각, ㈜엘에프에 발행가액 기준 감자대금 지급 (채권자보호절차 지연 시 최대 3개월 연장)"]]
    clone_table(s, tmpl, rows2, [2.3, 8.5], row_h=0.56, header_h=0.30, top=4.36)
    fn = by_y(s, 7.45)
    set_tf(fn[0].text_frame, "※ 출처: 주주간계약서 v22(clean) 제1.2조·제2조·제5조 / IM(2026.09) p.15. 종류주식 셀다운 예정으로 계약서상 인수 주식수 표는 지분 구성 확정 후 보충 예정.")

    s = S[6]
    title(s, "> 안건 Ⅱ. 주주간계약서 체결의 건 (2/2)"); band_label(s, "매입확약 · 지배구조 · 양도 제한")
    delete(src_table(s))
    rows = [["구분", "주요 내용", "조항"],
            ["제1종 매입확약", "발행일로부터 24개월 응당일까지 제1종 전부를 발행가액으로 매입 (분할 매입 시 보유 주식수 비례)", "제3.1~3.2조"],
            ["제2종 매입확약", "발행일로부터 12개월 응당일까지 제2종 전부를 발행가액으로 매입", "제4.1~4.2조"],
            ["매입 방식", "직접 양수 또는 유상감자·유상증자 참여 등 동일한 경제적 효과를 달성하는 거래 포함", "제3.2조·제4.2조"],
            ["배당금 정산", "매입완료일 직전일까지 누적배당률 기준 최대한도액에서 기수령 배당을 공제한 금액을 매입금액과 함께 정산", "제3.3~3.4조·제4.3~4.4조"],
            ["확약서 제출", "코람코라이프인프라리츠는 계약 체결일에 매입확약서를, 종류주주는 동일 조건의 매도확약서를 제출", "제3.5조·제4.5조"],
            ["이사회·대표이사", "이사 [3]인 중 [2]인 및 대표이사를 코람코라이프인프라리츠가 지명, 나머지 1인과 감사는 주주총회 선임", "제6.1~6.3조"],
            ["양도 제한", "주주 전원 서면동의 없이 양도·담보제공 금지, 정관에 주식양도 시 이사회 승인 명시", "제7.1조"],
            ["셀다운·후속양도", "인수인(우리투자증권·삼성증권)의 최초 셀다운은 KLI 동의 면제, 그 외 셀다운·후속양도는 KLI 사전 서면동의 필요", "제7.6조"]]
    clone_table(s, tmpl, rows, [1.9, 6.7, 2.2], row_h=0.58, header_h=0.30, top=1.59)
    fn = by_y(s, 7.45)
    set_tf(fn[0].text_frame, "※ 출처: 주주간계약서 v22(clean) 제3~7조. 대괄호 항목은 협의 중. 본 검토는 내부 검토이며 최종 법률자문은 외부 법무법인 확인이 필요함.")

    s = S[7]
    title(s, "> 안건 Ⅲ. 코람코라이프인프라리츠 매입확약서(LOC) 날인의 건"); band_label(s, "매입확약서 주요 내용")
    delete(src_table(s))
    rows = [["구분", "제1종 종류주식", "제2종 종류주식"],
            ["확약 주체", "㈜코람코라이프인프라위탁관리부동산투자회사", "(좌 동)"],
            ["발행회사 / 기초자산", "신설 위탁관리부동산투자회사 / 안양물류센터 (관양동 934 외 2필지)", "(좌 동)"],
            ["우선배당률", "연 7.0%", "연 7.5%"],
            ["매입기한", "발행일로부터 24개월이 되는 달의 응당일", "발행일로부터 12개월이 되는 달의 응당일"],
            ["매입금액", "1주당 발행가액 25,000원 × 매입대상 주식수 (240억원 예정)", "1주당 발행가액 25,000원 × 매입대상 주식수 (500억원 예정)"],
            ["제출 시점", "주주간계약 체결일 (기준 일정 2026-10-13)", "(좌 동)"]]
    clone_table(s, tmpl, rows, [2.2, 4.3, 4.3], row_h=0.34, header_h=0.30, top=1.59)
    add_band(s, 4.09, "매입확약서 · 매도확약서 대응 조건")
    rows2 = [["구분", "매입확약서 (KLI리츠 → 종류주주)", "매도확약서 (종류주주 → KLI리츠)"],
             ["제출 의무", "주주간계약 체결일에 KLI리츠가 각 종류주주에게 제출 (제3.5조·제4.5조)",
              "같은 날 각 종류주주가 KLI리츠에 제출. 셀다운 양수인은 거래 완결 전 제출"],
             ["핵심 내용", "매입기한 내 대상주식 전부를 발행가액(25,000원/주)으로 매입 확약",
              "KLI리츠의 매도 요청일로부터 [*]일 내 동일 조건으로 매도 확약"],
             ["선행조건", "① 영업인가 취득·소유권 확보 ② 종류주주의 완전한 소유권 ③ 부담 부존재 ④ 매도확약서 제출",
              "매입확약서와 동일 — 매도확약서 제출이 매입확약 효력 발생 조건"],
             ["불이행 시", "미매입잔액에 연 5.0% 지연손해금, 초과 손해 별도 배상, 매입의무 존속",
              "미매도잔액에 연 5.0% 지연이자 포함 모든 손해 배상"],
             ["배당금 정산\n· 효력", "매입일 전일까지 종류주주, 매입일부터 KLI리츠 귀속 (5영업일 내 정산)\n본건 매입 목적 외 사용 불가, 주주간계약과 상충 시 주주간계약 우선", "(좌 동)"]]
    clone_table(s, tmpl, rows2, [1.5, 4.7, 4.6], row_h=0.52, header_h=0.30, top=4.50, size=8)
    fn = by_y(s, 7.45)
    set_tf(fn[0].text_frame, "※ 출처: 코람코 제1종·제2종 우선주 매입확약 공문(안)(투자자용) / 주주간계약서 v22 제3.5조·제4.5조. 발행회사명·주식수·문서번호는 확정 후 기재.")

    # ── 11. 당사 예상수익 / 14. Highlights ───────────────────────────────
    for old, new in [("2,985억원", "2,989억원"), ("66.6억원", "66.7억원"), ("80.1억원", "80.2억원"), ("225.0억원", "225.1억원"), ("4.88%", "4.89%")]:
        replace_runs(S[10], old, new)
    replace_runs(S[13], "9.21%", "9.24%")

    # ── 17~19. 입지분석 보완 (실제 표 + 원본 지도 클립) ──────────────────
    s = S[17]
    strip_body(s)
    title(s, "> 2. 입지분석 — 광역 접근성 (시장조사)")
    lead2(s, "안양시는 수도권 중부권역 물류벨트의 중심축으로 광역도로망 접근성이 우수함",
          "서울시청 19km, 평촌IC·북의왕IC 1.9km, 양재IC 10km — 대형 화물차량의 Last-mile 배송 효율 확보")
    band_label(s, "광역입지 분석 및 안양시 개요")
    add_pic(s, clip(MKT_PDF, 6, (44, 191, 475, 505), "mkt06_map"), 0.50, 2.58, height=4.55)
    rows = [["구분", "내용"],
            ["소재지", "경기도 안양시"],
            ["총 인구수", "562,341명 (2026년 기준)"],
            ["총 사업체수 / 종사자수", "66,906개 / 284,815명 (2024년 기준)"],
            ["고속도로", "수도권제1순환·제2경인(인천~안양~성남)·서해안(목포~서울)·평택파주·경부고속도로, 강남순환도시고속도로"],
            ["전철·지하철", "지하철 1호선·4호선, (안양역) ITX-마음·무궁화호, 동탄인덕원선(예정), 월곶판교선(예정)"],
            ["교통 접근성", "서울시청 약 19km / 평촌IC 약 1.9km (수도권제1순환) / 북의왕IC 약 1.9km (제2경인) / 광명역IC 약 8.2km / 양재IC 약 10km (경부)"]]
    clone_table(s, tmpl, rows, [1.7, 2.7], row_h=0.66, header_h=0.30, top=2.58, left=6.85, size=9)
    fn_add(s, "Source: ㈜코람코자산신탁 안양물류센터 매입 시장조사보고서(2026-08-10) p.6 [KOSIS 국가통계포털]")

    s = S[18]
    strip_body(s)
    title(s, "> 2. 입지분석 — 협업입지 및 중부권역 (시장조사)")
    lead2(s, "중부권역은 서울 강남권·핵심 주거권역의 최인접 거점으로 Last-mile 기능을 수행함",
          "안양은 권역 내 유일한 프라임급 물류센터 입지로 상온 임대료 최고 수준(평당 6만원)을 기록 중")
    band_label(s, "협업입지 분석 · 중부권역 지역별 상세분석")
    add_pic(s, clip(MKT_PDF, 7, (44, 191, 917, 505), "mkt07_map"), 0.50, 2.58, width=6.30)
    add_pic(s, clip(MKT_PDF, 9, (44, 191, 470, 505), "mkt09_map"), 7.00, 2.58, height=2.27)
    rows = [["구분", "상온", "저온", "특징"],
            ["안양시", "4.0~6.0", "-", "풍부한 배후인구와 우수한 서울 접근성을 갖춘 라스트마일 핵심 요충지 / 신규개발 가용토지 극도로 제한, 권역 내 프라임급 물류센터인 안양물류센터가 최고 임대료(평당 6만원) 기록"],
            ["과천시", "-", "-", "개발제한구역 비중이 높고 주거·행정 기능에 초점, 대규모 물류센터 개발 사실상 불가(지역 내 물류센터 부재)"],
            ["의왕시", "4.3", "7.1", "수도권 물류거점 의왕ICD 소재로 화물 연계성 우수, 강남권 접근성 기반의 안정적 임대료·수요"],
            ["군포시", "3.3", "-", "군포복합물류터미널이 자리잡은 전통 물류허브, 노후 시설 비중이 높아 임대료는 낮게 형성"],
            ["수원시", "4.5", "-", "거주인구 약 118만명의 자체 소비시장 보유, 서울 도심 접근시간이 길어 경기남부 타겟 물류센터 활성화"]]
    clone_table(s, tmpl, rows, [1.0, 0.9, 0.9, 7.9], row_h=0.40, header_h=0.28, top=5.05, size=8)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.7, p.9 [Savills Korea Research & Consultancy]. 임대료 단위: 만원/평/월")

    s = S[19]
    strip_body(s)
    title(s, "> 2. 입지분석 — 중부권역 물류센터 현황 (시장조사)")
    lead2(s, "중부권역 내 대형 물류센터는 본건을 포함해 9개소로 공급이 극히 제한적임",
          "본건(28,869평, 2023년 준공)은 권역 내 최대 규모·최신축 자산으로 임대료 60,300원/평 수준")
    band_label(s, "중부권역 물류센터 현황")
    add_pic(s, clip(MKT_PDF, 12, (44, 191, 280, 505), "mkt12_map"), 0.50, 2.58, height=4.55)
    rows = [["구분", "자산명", "Type", "준공", "연면적(평)", "상온(원/평)", "저온(원/평)", "주임차인"]]
    for line in rows_json["mkt12"]:
        c = [t for x, t in line]
        if c and (c[0] == "SITE" or re.fullmatch(r"\d", c[0])):
            # SITE|안양시|안양물류센터|Dry|2023|28,869|60,300|-|주임차인
            if len(c) >= 9:
                rows.append([f"{c[0]} {c[1]}", c[2], c[3], c[4], c[5], c[6], c[7], c[8]])
            else:
                rows.append([f"{c[0]} {c[1]}", c[2], c[3], c[4], c[5], c[6], "-", c[7]])
    clone_table(s, tmpl, rows, [0.95, 1.75, 0.5, 0.5, 0.8, 0.8, 0.8, 1.05], row_h=0.33, header_h=0.30, top=2.58, left=4.05, size=8)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.12. Note) 과거 임대마케팅 가격 참고 (롯데제과·당정동)")

    # ── 24~25. 시장현황 보완 (원본 지도·차트 300dpi 클립) ────────────────
    s = S[24]
    strip_body(s)
    title(s, "> 3. 시장현황 — 권역별 임대시장 (시장조사)")
    lead2(s, "수도권 물류센터 임대시장은 권역별 차별화가 뚜렷하며 상온 임대료는 상승세를 지속함",
          "중부권역은 상온 임대료 수준과 임차 수요 모두 수도권 상위권에 위치")
    band_label(s, "권역별 임대료 · 공실률 현황 및 상온 임대료 추이")
    add_pic(s, clip(MKT_PDF, 31, (44, 191, 917, 505), "mkt31"), 0.50, 2.58, width=5.60)
    add_pic(s, clip(MKT_PDF, 32, (44, 191, 917, 505), "mkt32"), 6.25, 2.58, width=4.95)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.31~32 [Savills Korea Research & Consultancy]")

    s = S[25]
    strip_body(s)
    title(s, "> 3. 시장현황 — 공실률 추이 및 Last-mile 공급 (시장조사)")
    lead2(s, "Last-mile 권역 신규 공급은 2025년 이후 급감하고 상온센터 공실률은 낮게 유지됨",
          "안양시 ’26~’28 공급예정량 0㎡ — 공급 제한과 초과 수요로 임대인 우위 시장 지속 예상")
    band_label(s, "권역별 공실률 추이 · 수도권 Last-mile 신규공급 현황")
    add_pic(s, clip(MKT_PDF, 33, (44, 191, 917, 505), "mkt33"), 0.50, 2.58, width=5.20)
    add_pic(s, clip(MKT_PDF, 43, (44, 191, 917, 505), "mkt43"), 5.85, 2.58, width=5.35)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.33, p.43 [Savills RE Strategy & Solutions]. ’26.2월 기준 연면적 5천평 이상 상온창고")

    # ── 26. 임대사례 (실제 표 + 사진) ─────────────────────────────────────
    s = S[26]
    strip_body(s)
    title(s, "> 4. Peer그룹 임대사례 (시장조사)")
    lead2(s, "수도권 Last-mile 임대사례 Eff.Rent 4.0~6.0만원/평 중 본건 5.8만원은 상단 수준",
          "서울복합·부천 삼정동(6.0만원/평) 등 신축 Core 자산과 유사 수준, 준공연도·규모 감안 시 상승 여력")
    band_label(s, "수도권 Last-mile 유사자산 임대사례 (단위: 원/평/월)")
    hdr = ["구분", "본건 (안양)", "서울", "인천·부천", "", "", "김포", "", "고양", ""]
    rows = [hdr]
    order = ["자산명", "소재지", "연면적", "규모", "Type", "임대료(상온)", "관리비(상온)", "R/F+TI(M)", "Eff. Rent", "준공연도", "주임차인"]
    got = {c[0][1]: [t for x, t in c[1:]] for c in rows_json["mkt45"] if c and c[0][1] in order}
    rows.append(["사진"] + [""] * 9)
    for k in order:
        v = got.get(k, [])
        rows.append([k] + (v + [""] * 9)[:9])
    t = clone_table(s, tmpl, rows, [1.0] + [1.078] * 9, row_h=0.27, header_h=0.30, top=2.58, size=8)
    set_row_height(t, 1, 0.95)
    rects = photo_cells(MKT_PDF, 45)
    for j, r in enumerate(rects[:9]):
        add_pic(s, clip(MKT_PDF, 45, (r.x0, r.y0, r.x1, r.y1), f"mkt45_ph{j}", dpi=200),
                0.50 + 1.0 + j * 1.078 + 0.04, 2.58 + 0.30 + 0.03, width=1.0, height=0.89)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.45 [Savills RE Strategy & Solutions]")

    # ── 27. 거래사례 (실제 표 + 사진) ─────────────────────────────────────
    s = S[27]
    strip_body(s)
    title(s, "> 4. Peer그룹 거래사례 (시장조사)")
    lead2(s, "본건 매입가(평당 935만원)는 유사 거래사례 736~1,248만원/평 중 중하단에 위치함",
          "신축 Core 자산인 하남미사(1,248만원)·부천내동(1,220만원) 대비 낮은 수준으로 적정 매입가 범위(891~1,014만원) 내")
    band_label(s, "수도권 Last-mile 유사자산 거래사례")
    hdr = ["구분", "중부권 (본건)", "동남권", "동북권", "", "서남권", "", "", "", ""]
    rows = [hdr]
    order = ["자산명", "소재지", "연면적", "규모", "Type", "거래시점", "매매가격(억원)", "평당가(만원/평)", "Cap. Rate", "매도자", "매수자", "준공연도"]
    got = {c[0][1]: [t for x, t in c[1:]] for c in rows_json["mkt46"] if c and c[0][1] in order}
    rows.append(["사진"] + [""] * 9)
    for k in order:
        v = got.get(k, [])
        rows.append([k] + (v + [""] * 9)[:9])
    t = clone_table(s, tmpl, rows, [1.0] + [1.078] * 9, row_h=0.26, header_h=0.30, top=2.58, size=8)
    set_row_height(t, 1, 0.80)
    rects = photo_cells(MKT_PDF, 46)
    for j, r in enumerate(rects[:9]):
        add_pic(s, clip(MKT_PDF, 46, (r.x0, r.y0, r.x1, r.y1), f"mkt46_ph{j}", dpi=200),
                0.50 + 1.0 + j * 1.078 + 0.04, 2.58 + 0.30 + 0.03, width=1.0, height=0.74)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.46 [Savills RE Strategy & Solutions]. 본건 Cap. Rate 5.16%는 시장조사 산정 기준")

    # ── 28. 적정임대료 결론 (실제 표) ─────────────────────────────────────
    s = S[28]
    strip_body(s)
    title(s, "> 5. 시장조사 결론 — 적정 임대료 및 매매가")
    lead2(s, "본건 적정 Eff.Rent는 2026년 기준 6.14만원/평으로 현 5.8만원 대비 약 6% 상승 여력",
          "적정 매입가 범위 891~1,014만원/평 (본건 935만원), 투자적정성 검토 Exit Cap 5.20%·매각대금 3,667억원")
    band_label(s, "적정 임대가 산정 (비교사례 보정 및 시장 상승률 2.5% 반영)")
    rows = [["구분", "서울복합물류센터", "부천내동물류센터", "CBRE IM 물류센터", "고양 삼송 로지스힐"],
            ["① 보정 전 Eff.Rent (원/평)", "60,000", "55,000", "53,500", "52,300"],
            ["② 보정치 합계", "-4%", "-17%", "-19%", "-6%"],
            ["③ 적용 보정치 값 [1-②]", "104%", "117%", "119%", "106%"],
            ["④ 보정 후 Eff.Rent (원/평) [①×③]", "62,400", "64,350", "63,665", "55,438"],
            ["본 자산 적정 Eff.Rent ('26)", "61,463 (비교사례 평균)", "", "", ""]]
    clone_table(s, tmpl, rows, [2.9, 1.95, 1.95, 1.95, 1.95], row_h=0.30, header_h=0.30, top=2.58, size=9)
    rows = [["구분", "2027년", "2028년", "2029년", "2030년"],
            ["본 건 예상 E.NOC (원/평)", "63,000", "64,575", "66,189", "67,844"]]
    clone_table(s, tmpl, rows, [2.9, 1.95, 1.95, 1.95, 1.95], row_h=0.30, header_h=0.30, top=4.55, size=9)
    add_pic(s, clip(MKT_PDF, 52, (41, 398, 923, 515), "mkt52_range"), 0.50, 5.30, width=10.70)
    rows = [["적정 매입가 (평당)", "891 ~ 1,014만원 (본건 935만원)", "투자적정성 (Exit)", "Exit Cap 5.20% · 매각대금 3,667억원 (NOI N+1 18,959백만원) — 사업계획 4.89%·3,759억원 대비 보수적"]]
    clone_table(s, tmpl, rows, [1.7, 2.6, 1.6, 4.8], row_h=0.34, header_h=0.34, top=6.75, size=8)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.52~53, p.55 종합의견 [Savills RE Strategy & Solutions]. 임대료 상승은 예상치이며 확정된 것이 아님")

    # ── 29~30. 리츠개요·투자구조 수치 갱신 ───────────────────────────────
    for old, new in [("2,985억원 (Equity : 984억원", "2,989억원 (Equity : 988억원"), ("9.21%", "9.24%"), ("16.66%", "16.44%"), ("10.35%", "10.34%")]:
        replace_runs(S[29], old, new)
    fn = by_y(S[29], 7.45)
    set_tf(fn[0].text_frame, "총 자산가액: 부동산 장부가액 2,889억원 및 여유현금 100억원 예정 / 수익률은 재무분석보고서(우리회계법인)·IM 기준, 재무모델(09-15)은 2종 9.15%·보통주 16.18% — 확정값 확인 중")
    for old, new in [("[2,985]", "[2,989]"), ("[144]", "[148]"), ("[185]", "[189]")]:
        replace_runs(S[30], old, new)

    # ── 31. 사업구조도 (IM 도식 300dpi 클립) ─────────────────────────────
    s = S[31]
    strip_body(s)
    title(s, "> 2. 투자구조 — 사업구조도 (IM 기준)")
    lead2(s, "공모예외기관(KLI리츠)이 발기인으로 자리츠를 설립한 뒤 증자로 Equity를 모집하는 구조임",
          "매도인(코크렙안양PFV)으로부터 소유권 이전, AMC 코람코자산신탁·자산보관·사무수탁회사와 위탁계약")
    band_label(s, "사업 구조도")
    add_pic(s, clip(IM_PDF, 6, (35, 185, 805, 565), "im06_diagram"), 0.50, 2.58, width=9.60)
    fn_add(s, "Source: 안양물류센터 Information Memorandum(2026.09) p.5 사업 구조도")

    # ── 32~33. Milestone / 34. 재원조달 / 37. 운영가정 ───────────────────
    for idx in (32, 33):
        for old, new in [("[2,985]", "[2,989]"), ("[144]", "[148]")]:
            replace_runs(S[idx], old, new)
    s = S[34]
    lead_sh = by_text(s, "총 사업비 2,985억원")
    set_tf(lead_sh.text_frame, ["총 사업비 2,989억원으로 매입금액(90.3%), 매입부대비용(6.3%), 예비현금(3.3%)로 구성",
                                "Carry기간 1종 우선주(7.0%) 및 2종 우선주(7.5%) 등 확정배당 지급을 위한 Overfunding금액 100억원 편성"])
    for t in [sh for sh in s.shapes if sh.has_table]:
        tb = t.table
        if len(tb.columns) == 4:
            for ri, (a, b, c) in {1: ("270,000", "90.3%", "매입금액 2,700억원, 감정가 대비 −19.3%"), 2: ("1,350", "0.5%", None), 3: ("12,498", "4.2%", None),
                                  4: ("662", "0.2%", None), 5: ("4,390", "1.5%", None), 6: ("10,000", "3.3%", None), 7: ("298,900", "100.0%", None)}.items():
                set_cell(tb.cell(ri, 1), a); set_cell(tb.cell(ri, 2), b)
                if c:
                    set_cell(tb.cell(ri, 3), c)
        elif len(tb.columns) == 6:
            for ri, (a, b, c, d) in {1: ("2,100", "0.7%", None, None),
                                     2: ("160,000", "53.5%", "금리 4.78%, 수수료 1.45% (All-in 5.50%)", "우리은행 심사 완료(9/3), 대출약정 협의 중"),
                                     3: ("38,000", "12.7%", "금리 6.00%, 수수료 1.00% (All-in 6.50%)", "우리은행 등, 심사 완료"),
                                     4: ("24,000", "8.0%", None, "우리투자증권 인수확약서(LOC) 확보"),
                                     5: ("50,000", "16.7%", None, "삼성증권 인수확약서(LOC) 확보"),
                                     6: ("10,000", "3.3%", None, "㈜엘에프 재투자 (주주간계약 협의 중)"),
                                     7: ("14,800", "5.0%", None, "KLI리츠 이사회 승인 예정(10/1)"),
                                     8: ("298,900", "100.0%", None, None)}.items():
                set_cell(tb.cell(ri, 1), a); set_cell(tb.cell(ri, 2), b)
                if c:
                    set_cell(tb.cell(ri, 3), c)
                if d:
                    set_cell(tb.cell(ri, 5), d)
    replace_runs(S[38], "4.88%", "4.89%")


    # ── 36~37. 운영가정 분할 (렌트롤 / 층별 배치) ─────────────────────────
    sa, sb = S[35], S[36]
    t47a = [x for x in sa.shapes if x.has_table and x.name == "표 47"]
    t6a = [x for x in sa.shapes if x.has_table and x.name == "Table 6"]
    if t47a:
        delete(t47a[0])
    if t6a:
        set_table_width(t6a[0], 10.70, 0.50)
    title(sa, "> 5. 운영가정 — 임대차 현황")
    band_label(sa, "임차인별 임대차 조건")
    t6b = [x for x in sb.shapes if x.has_table and x.name == "Table 6"]
    t47b = [x for x in sb.shapes if x.has_table and x.name == "표 47"]
    if t6b:
        delete(t6b[0])
    if t47b:
        set_table_width(t47b[0], 10.70, 0.50)
        tb = t47b[0].table
        for row in tb.rows:                      # "쿠팡 (2,185.56평)" → "쿠팡 / 2,186평"
            for cell in row.cells:
                m = re.match(r"^(.+?)\s*\(([\d,\.]+)\s*평?\)$", cell.text.strip())
                if m:
                    try:
                        area = f"{float(m.group(2).replace(',', '')):,.0f}평"
                    except ValueError:
                        area = m.group(2) + "평"
                    set_cell(cell, f"{m.group(1).strip()}\n{area}")
    title(sb, "> 5. 운영가정 — 층별 임차인 배치")
    band_label(sb, "층별 임차인 배치 현황")
    lead2(sb, "쿠팡이 지하 2층~지상 8층에 걸쳐 임대면적의 77.2%를 사용 중임",
          "농심(3F)·지오영·용마로지스·아이팜(1F)·키즈밀(8F)이 잔여 면적을 임차, 공실은 8층 141평")
    for sh in list(sb.shapes):
        if sh.has_text_frame and 6.5 < (sh.top or 0) / I < 7.40:
            delete(sh)
    fn_add(sb, "Source: 매도인 제공 렌트롤 (2026-07 기준). 면적은 평 단위 반올림 표기")

    # ── 38~39. 투자수익률 (원본 표 유지, IM 수치로 갱신) ─────────────────
    s = S[40]
    tabs = sorted([sh for sh in s.shapes if sh.has_table], key=lambda x: x.left)
    fill_im_table(tabs[0], imt["1종"]); fill_im_table(tabs[1], imt["2종"])
    lead2(s, "사업기간 10년 가정 투자자별 수익률은 1종 IRR 7.12%, 2종 IRR 9.24% 예상",
          "제1종·제2종 우선주는 누적적 확정 배당수익률 7.0%, 7.5%를 각각 지급 예정")
    s = S[39]
    tabs = sorted([sh for sh in s.shapes if sh.has_table], key=lambda x: x.left)
    fill_im_table(tabs[0], imt["3종"]); fill_im_table(tabs[1], imt["보통주"])
    lead2(s, "3종우선주는 무배당·2년 후 원본감자, 보통주는 IRR 16.44%·E.Multiple 4.35x 예상",
          "보통주는 초기 1년 7.5% 배당 후 잔여배당, 매각차익의 80% 배분 (재무모델 09-15: 16.18%, 확정값 확인 중)")
    for sh in by_y(s, 7.48, tol=0.08):
        if "Source" in sh.text_frame.text:
            set_tf(sh.text_frame, "Source: 안양물류센터 IM(2026.09) p.19~20 / 재무분석보고서(우리회계법인)")

    # ── 40~41. Sell-down (동일 표 양식) ───────────────────────────────────
    s = S[42]
    title(s, "> 6. 투자수익률 — KLI 단계별 매입(Sell-down) 배당스케줄 (1/2)")
    lead2(s, "KLI리츠가 2종(12개월 후)·1종(24개월 후)을 순차 매입할 경우의 배당스케줄 및 수익률",
          "매입가에 간주취득세(1종 2.2억·2종 4.5억)를 포함하여 Yield는 1종 6.9%, 2종 7.4% 수준 (반기 기준 FY)")
    tabs = sorted([sh for sh in s.shapes if sh.has_table], key=lambda x: x.left)
    fill_yield_table(tabs[0], selldown_table(per, summ, 0)); fill_yield_table(tabs[1], selldown_table(per, summ, 1))
    lab = sorted([sh for sh in iter_shapes(s.shapes) if sh.has_text_frame and sh.text_frame.text.strip() in ("1종 우선주", "2종 우선주")],
                 key=lambda x: x.text_frame.text.strip())
    for g, txt in ((0, "1종 우선주 (KLI 매입, FY5~)"), (1, "2종 우선주 (KLI 매입, FY3~)")):
        if len(lab) > g:
            set_tf(lab[g].text_frame, txt)
    fn_add(s, "Source: 재무모델 v01(2026-09-15) Sell-down 시트 ‘Sell-down 인수 시 수익률’ 블록. FY는 운용 개시(2026-10-31) 기준 반기, 단위 백만원")

    s = S[41]
    title(s, "> 6. 투자수익률 — KLI 단계별 매입(Sell-down) 배당스케줄 (2/2)")
    lead2(s, "KLI 통합 기준 투자잔액 895억원, IRR 10.83%, E.Multiple 2.19x 예상",
          "보통주(148억원)는 초기 출자, 2종·1종 매입 완료 후 FY5부터 통합 Yield 약 6.1% 수준")
    tabs = sorted([sh for sh in s.shapes if sh.has_table], key=lambda x: x.left)
    fill_yield_table(tabs[0], selldown_table(per, summ, 2)); fill_yield_table(tabs[1], selldown_table(per, summ, 3))
    lab = sorted([sh for sh in iter_shapes(s.shapes) if sh.has_text_frame and sh.text_frame.text.strip() in ("3종 우선주", "보통주")],
                 key=lambda x: x.text_frame.text.strip())  # '3종 우선주' < '보통주' (좌/우 순)
    if len(lab) == 2:
        set_tf(lab[0].text_frame, "보통주 (KLI 초기 출자)"); set_tf(lab[1].text_frame, "통합 (1종+2종+보통주)")
    for sh in by_y(s, 7.48, tol=0.08):
        if sh.has_text_frame:
            txt = sh.text_frame.text
            if "Source" in txt:
                set_tf(sh.text_frame, "Source: 재무모델 v01(2026-09-15) Sell-down 시트. 통합 C.G 58,168백만원(1종 −215·2종 11,313·보통주 47,070)")
            elif "Carry" in txt:
                set_tf(sh.text_frame, "* 1종 C.G −215백만원은 간주취득세(발행가 초과 매입원가) 미회수분")

    # ── 46. 리스크 / 47. 향후조치 ────────────────────────────────────────
    replace_runs(S[47], "4.88%", "4.89%")
    s = S[48]
    lead_sh = by_text(s, "본 건 (예비)투자심의위원회 이후")
    set_tf(lead_sh.text_frame, ["본 투자심의위원회 이후 보완 예정사항은 다음과 같으며, 10/2 매매계약·10/13 주주간계약·대출약정 체결 예정",
                                "• 추가 심의위원 의견 청취 후 계약 체결 전 주요 계약서 내용 반영 및 보완할 예정임"])
    tb = src_table(s).table
    vals = [["구분", "본 투자심의 후 Checklist", "계약 체결 前 보완방향"],
            ["계약 문안 확정", "매매계약 체결일·거래종결예정일·위약금·지연손해금률 공란 / 주주간계약 인수 주식수 표(셀다운 반영)·이사 수·지명권 조항 대괄호", "법무팀 협업 하에 9/22 매매계약서·주주간계약 합의 시 확정, 매입확약서 발행회사명·주식수 기재 후 날인"],
            ["수치 확정", "재무모델(9/15) 2종 IRR 9.15%·보통주 16.18% vs IM·재무분석보고서 9.24%·16.44% 차이, WALE 산정기준(6.8년 vs 6.66년)", "재무팀 검토로 원인 규명 및 기준값 확정 후 투자자 Q&A·IM 갱신"],
            ["실사 반영", "대지면적 차이(15,287.5㎡ vs 모델 15,019.5㎡, 필지 수) / 법률실사 Recommendation 41건 중 미해소 항목 / 실사보고서 최종본(9/17) 반영", "등기부 확인 및 SPA 선행조건·특약 반영, 물리실사 수익적 지출(5년 12.6억원) 사업계획 반영"],
            ["거래관계 투명성", "매도인 주주(㈜엘에프 95%·코람코 5%)의 3종 재투자 및 매도인·매수인 AMC 동일(코람코자산신탁)에 따른 이해관계인 거래 이슈", "리스크관리팀·법무팀 협업 하에 영업인가 및 거래가격 공정성 증빙(감정평가·시장조사) 구비, 외부 법률의견 반영"]]
    for i, v in enumerate(vals):
        for j in range(3):
            set_cell(tb.cell(i, j), v[j])

    normalize_tables(prs)
    autofit_tables(prs)
    prs.save(str(OUT))
    print("생성 완료:", OUT)
    return OUT


if __name__ == "__main__":
    build()
