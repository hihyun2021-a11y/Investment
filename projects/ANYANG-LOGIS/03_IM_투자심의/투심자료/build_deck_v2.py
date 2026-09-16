#!/usr/bin/env python3
"""안양물류센터 본 투자심의 자료 v02 — 예비투심 원본 덱(KLI 260824 v3)을 개작.

원본의 표지·안건부를 본 투심용으로 교체하고, 별첨의 지도·차트·구조도는 그대로 유지한 뒤
시장조사보고서·IM 도식 페이지와 재무모델 Sell-down 표를 새 슬라이드로 삽입한다.
표는 원본 표 XML을 복제해 서식을 동일하게 유지한다.

실행: python3 projects/ANYANG-LOGIS/03_IM_투자심의/투심자료/build_deck_v2.py
"""
import copy
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
from pptx.util import Inches, Emu

ROOT = pathlib.Path(__file__).resolve().parents[4]
HERE = pathlib.Path(__file__).parent
PROJ = ROOT / "projects/ANYANG-LOGIS"
BASE = ROOT / "platform/templates/ic_deck/투자심의자료_양식_KLI_260824_v3.pptx"
SKILL = pathlib.Path("/root/.claude/skills/synced/3286b16b-1706-4484-a678-751bd9178857_2febf011-5f5f-40c7-acee-e3b1eb5f699f/pptx/scripts")
WORK = pathlib.Path("/tmp/claude-0/-home-user-Investment/38052696-05de-5d1d-9b55-e347b5278ed0/scratchpad/v2")
IMG = WORK / "img"
OUT = HERE / "20260916_투심자료_ANYANG-LOGIS_v02.pptx"
IM_PDF = PROJ / "03_IM_투자심의/IM/20260914_IM_코크렙안양물류센터I_코람코_v01.pdf"
MKT_PDF = PROJ / "02_실사보고서/시장/(2026.08.10) 코람코자산신탁_안양물류센터 매입 시장보고서(Final) (1) (1).pdf"
MODEL = PROJ / "04_Equity투자자/models/20260915_재무모델_Financial_Modeling_v01.xlsm"

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
I = 914400


# ═════════════════════════════════════════════════════════════════════════
# 0. 이미지 준비 (PDF 페이지 → 본문 영역 크롭 PNG)
# ═════════════════════════════════════════════════════════════════════════
def render(pdf, page, box, name, dpi=170):
    IMG.mkdir(parents=True, exist_ok=True)
    out = IMG / f"{name}.png"
    if out.exists():
        return out
    pix = pymupdf.open(str(pdf))[page - 1].get_pixmap(dpi=dpi)
    tmp = IMG / f"_{name}_full.png"
    pix.save(str(tmp))
    im = Image.open(tmp)
    im.crop(box).save(out)
    return out


IM_BOX = (60, 440, 1930, 1260)        # 1988x1406 기준: 섹션 밴드 아래 ~ 꼬리말 위
IM_TBL_BOX = (60, 430, 1930, 1305)    # 수익률 표 페이지 (표 하단까지)
MKT_BOX = (80, 380, 2190, 1205)       # 2267x1275 기준: 리드 문장 아래 ~ 출처 위


# ═════════════════════════════════════════════════════════════════════════
# 1. 구조 편집 (복제·삭제) — unpack → add_slide.py → sldIdLst → clean → zip
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
        m = re.search(r"slides/(slide\d+\.xml)", r.stdout)
        return m.group(1)

    # 안건 슬라이드: slide4 를 3장 복제 (slide5 뒤) → 안건Ⅱ(1), 안건Ⅱ(2), 안건Ⅲ
    a1 = dup("slide4.xml", "slide5.xml")
    a2 = dup("slide4.xml", a1)
    a3 = dup("slide4.xml", a2)
    # 입지 보완 2장 (slide16 뒤)
    l1 = dup("slide16.xml", "slide16.xml")
    l2 = dup("slide16.xml", l1)
    # 시장 보완 2장 (slide20 뒤)
    m1 = dup("slide17.xml", "slide20.xml")
    m2 = dup("slide17.xml", m1)
    # 임대사례 보완 (slide21 뒤), 거래사례 보완 + 적정임대료 결론 (slide22 뒤)
    r1 = dup("slide21.xml", "slide21.xml")
    t1 = dup("slide22.xml", "slide22.xml")
    t2 = dup("slide22.xml", t1)
    # 사업구조도 (slide24 뒤)
    s1 = dup("slide24.xml", "slide24.xml")
    # Sell-down 표 (slide32 뒤)
    d1 = dup("slide4.xml", "slide32.xml")

    # slide6, slide7 (실사업체 선정) 삭제: sldIdLst 에서 제거
    pres = un / "ppt/presentation.xml"
    x = pres.read_text(encoding="utf-8")
    rels = (un / "ppt/_rels/presentation.xml.rels").read_text(encoding="utf-8")
    for sl in ("slide6.xml", "slide7.xml"):
        rid = re.search(r'Id="(rId\d+)"[^>]*Target="slides/%s"' % sl, rels) or \
              re.search(r'Target="slides/%s"[^>]*Id="(rId\d+)"' % sl, rels)
        rid = rid.group(1)
        x = re.sub(r'<p:sldId [^>]*r:id="%s"[^>]*/>' % rid, "", x)
    pres.write_text(x, encoding="utf-8")
    subprocess.run([sys.executable, str(SKILL / "clean.py"), str(un)], check=True, capture_output=True)

    base = WORK / "base.pptx"
    subprocess.run(["zip", "-qXr", str(base), "."], cwd=un, check=True)
    return base


# ═════════════════════════════════════════════════════════════════════════
# 2. 내용 편집 도우미 (서식 보존)
# ═════════════════════════════════════════════════════════════════════════
def iter_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:
            yield from iter_shapes(sh.shapes)


def find(slide, pred):
    return [sh for sh in iter_shapes(slide.shapes) if pred(sh)]


def by_text(slide, sub):
    r = [sh for sh in iter_shapes(slide.shapes) if sh.has_text_frame and sub in sh.text_frame.text]
    return r[0] if r else None


def by_y(slide, y, tol=0.06, has_text=True):
    r = [sh for sh in iter_shapes(slide.shapes)
         if abs((sh.top or 0) / I - y) < tol and (not has_text or sh.has_text_frame)]
    return r


def set_tf(tf, lines, keep_bullets=True):
    """텍스트 프레임의 첫 문단·첫 런 서식을 유지한 채 lines 로 교체. 문단 수만큼 복제."""
    if isinstance(lines, str):
        lines = [lines]
    txBody = tf._txBody
    paras = txBody.findall("a:p", NS)
    tmpl = paras[0]
    # 첫 문단에서 첫 런 서식 확보
    r0 = tmpl.find("a:r", NS)
    rPr = copy.deepcopy(r0.find("a:rPr", NS)) if r0 is not None and r0.find("a:rPr", NS) is not None else None
    pPr = copy.deepcopy(tmpl.find("a:pPr", NS)) if tmpl.find("a:pPr", NS) is not None else None
    endPr = tmpl.find("a:endParaRPr", NS)
    # 두 번째 문단 서식(불릿 등)이 있으면 2행부터 그것을 쓴다
    pPr2 = copy.deepcopy(paras[1].find("a:pPr", NS)) if keep_bullets and len(paras) > 1 and paras[1].find("a:pPr", NS) is not None else pPr
    r1 = paras[1].find("a:r", NS) if len(paras) > 1 else None
    rPr2 = copy.deepcopy(r1.find("a:rPr", NS)) if r1 is not None and r1.find("a:rPr", NS) is not None else rPr
    for p in paras:
        txBody.remove(p)
    for i, line in enumerate(lines):
        p = txBody.makeelement("{%s}p" % NS["a"], {})
        pp = pPr if i == 0 else pPr2
        if pp is not None:
            p.append(copy.deepcopy(pp))
        r = p.makeelement("{%s}r" % NS["a"], {})
        rp = rPr if i == 0 else rPr2
        if rp is not None:
            r.append(copy.deepcopy(rp))
        t = r.makeelement("{%s}t" % NS["a"], {})
        t.text = line
        r.append(t)
        p.append(r)
        if endPr is not None and i == len(lines) - 1:
            p.append(copy.deepcopy(endPr))
        txBody.append(p)


def set_cell(cell, text):
    set_tf(cell.text_frame, text.split("\n") if isinstance(text, str) else text, keep_bullets=False)


def replace_runs(slide, old, new):
    """모든 런에서 old→new 치환 (서식 유지). 치환 건수 반환."""
    n = 0
    for sh in iter_shapes(slide.shapes):
        frames = []
        if sh.has_text_frame:
            frames.append(sh.text_frame)
        if sh.has_table:
            frames += [c.text_frame for r in sh.table.rows for c in r.cells]
        for tf in frames:
            for p in tf.paragraphs:
                hit = False
                for r in p.runs:
                    if old in r.text:
                        r.text = r.text.replace(old, new)
                        n += 1
                        hit = True
                if not hit and p.runs and old in "".join(r.text for r in p.runs):
                    # 런이 쪼개진 경우: 첫 런 서식으로 병합 후 치환
                    joined = "".join(r.text for r in p.runs)
                    p.runs[0].text = joined.replace(old, new)
                    for r in p.runs[1:]:
                        r._r.getparent().remove(r._r)
                    n += 1
    return n


def delete(sh):
    el = sh._element
    el.getparent().remove(el)


def strip_body(slide, keep_footnote=False):
    """제목(y≈0.24)·리드(y≈1.14)·섹션 밴드 그룹(y≈2.17)만 남기고 본문 도형 삭제."""
    for sh in list(slide.shapes):
        y = (sh.top or 0) / I
        keep = (sh.has_text_frame and abs(y - 0.24) < 0.06) or (sh.has_text_frame and abs(y - 1.14) < 0.06) \
               or (sh.shape_type == 6 and abs(y - 2.17) < 0.06) \
               or (keep_footnote and sh.has_text_frame and y >= 7.40)
        if not keep:
            delete(sh)


def band_label(slide, text):
    g = [sh for sh in slide.shapes if sh.shape_type == 6 and abs((sh.top or 0) / I - 2.17) < 0.06]
    if g:
        tb = [s for s in iter_shapes(g[0].shapes) if s.has_text_frame]
        if tb:
            set_tf(tb[0].text_frame, text)


def title(slide, text):
    t = by_y(slide, 0.24)
    if t:
        set_tf(t[0].text_frame, text)


def lead(slide, main, subs=()):
    t = by_y(slide, 1.14)
    if t:
        set_tf(t[0].text_frame, [main, *subs])


def add_pic(slide, png, left, top, width=None, height=None, max_w=10.70, max_h=4.72):
    im = Image.open(png)
    ar = im.width / im.height
    if width is None and height is None:
        width = max_w
        height = width / ar
        if height > max_h:
            height = max_h
            width = height * ar
    elif width is None:
        width = height * ar
    elif height is None:
        height = width / ar
    return slide.shapes.add_picture(str(png), Inches(left), Inches(top), Inches(width), Inches(height))


def clone_table(slide, src_gf, rows, col_widths, row_h=0.30, header_h=None, size=None, top=2.58, left=0.50):
    """src 그래픽프레임(표)의 서식을 복제해 rows(list[list[str]]) 표를 만든다. 열 폭 inch."""
    gf = copy.deepcopy(src_gf._element)
    tbl = gf.find(".//a:tbl", NS)
    grid = tbl.find("a:tblGrid", NS)
    gcs = grid.findall("a:gridCol", NS)
    trs = tbl.findall("a:tr", NS)
    hdr_tr, body_tr = trs[0], trs[1]
    ncol = len(col_widths)
    # gridCol
    for gc in gcs:
        grid.remove(gc)
    for w in col_widths:
        gc = copy.deepcopy(gcs[0])
        gc.set("w", str(int(w * I)))
        # extLst 제거(선택)
        grid.append(gc)
    # 행 제거
    for tr in trs:
        tbl.remove(tr)

    def make_row(tmpl_tr, vals, h):
        tr = copy.deepcopy(tmpl_tr)
        tcs = tr.findall("a:tc", NS)
        for tc in tcs:
            tr.remove(tc)
        for j in range(ncol):
            src_tc = tcs[min(j, len(tcs) - 1)]
            tc = copy.deepcopy(src_tc)
            # gridSpan/hMerge 등 병합 속성 제거
            for k in ("gridSpan", "rowSpan", "hMerge", "vMerge"):
                tc.attrib.pop(k, None)
            tr.append(tc)
        tr.set("h", str(int(h * I)))
        # extLst 안의 rowId 등은 무시
        return tr

    for i, vals in enumerate(rows):
        h = header_h if (i == 0 and header_h) else row_h
        tbl.append(make_row(hdr_tr if i == 0 else body_tr, vals, h))
    # 위치·크기
    xfrm = gf.find("p:xfrm", NS)
    xfrm.find("a:off", NS).set("x", str(int(left * I)))
    xfrm.find("a:off", NS).set("y", str(int(top * I)))
    xfrm.find("a:ext", NS).set("cx", str(int(sum(col_widths) * I)))
    xfrm.find("a:ext", NS).set("cy", str(int((row_h * (len(rows) - 1) + (header_h or row_h)) * I)))
    # id 충돌 방지
    cNvPr = gf.find(".//p:cNvPr", NS)
    cNvPr.set("id", str(9000 + len(slide.shapes) * 7 + 1))
    cNvPr.set("name", "표(복제)")
    slide.shapes._spTree.append(gf)
    # python-pptx 객체로 다시 얻어 텍스트 세팅
    new = slide.shapes[-1]
    tb = new.table
    for i, vals in enumerate(rows):
        for j in range(ncol):
            set_cell(tb.cell(i, j), vals[j] if j < len(vals) else "")
            if size:
                for p in tb.cell(i, j).text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = Emu(size * 12700)
    return new


def src_table(slide):
    return [sh for sh in slide.shapes if sh.has_table][0]


# ═════════════════════════════════════════════════════════════════════════
# 3. Sell-down 연차별 집계 (재무모델)
# ═════════════════════════════════════════════════════════════════════════
def selldown_rows():
    wb = openpyxl.load_workbook(MODEL, data_only=True)
    ws = wb["Sell-down"]
    # 블록: r66 헤더, r67 기간0, r68~r87 기간1~20, r88 합계, r90~93 요약 / 열 12부터
    c0 = 12
    per = {}
    for r in range(68, 88):
        k = int(ws.cell(r, c0 + 1).value)
        v = [ws.cell(r, c0 + 2 + j).value or 0 for j in range(16)]
        per[k] = v  # [1종 잔액,배당,CG,yield, 2종 ..., 보통주 ..., 통합 ...]
    summ = {ws.cell(r, c0 + 1).value: [ws.cell(r, c0 + 5 + 4 * g).value for g in range(4)] for r in (90, 91, 92, 93)}

    def m(v):  # 천원 → 백만원
        return v / 1000.0

    rows = [["연차", "1종 잔액", "1종 배당", "Yield", "2종 잔액", "2종 배당", "Yield",
             "보통주 잔액", "보통주 배당", "Yield", "통합 잔액", "통합 배당+C.G", "Yield"]]
    tot = [0.0] * 4
    cg_tot = [0.0] * 4
    for y in range(1, 11):
        a, b = per[2 * y - 1], per[2 * y]
        line = [f"FY {y}"]
        for g in range(4):
            bal = m(b[4 * g])
            div = m(a[4 * g + 1] + b[4 * g + 1])
            cg = m(a[4 * g + 2] + b[4 * g + 2])
            tot[g] += div
            cg_tot[g] += cg
            yld = (div + (cg if g == 3 else 0)) / bal if bal else 0
            if g == 3:
                line += [f"{bal:,.0f}" if bal else "-", f"{div + cg:,.0f}" if (div + cg) else "-", f"{yld:.1%}" if bal else "-"]
            else:
                line += [f"{bal:,.0f}" if bal else "-", f"{div:,.0f}" if div else "-", f"{yld:.1%}" if bal else "-"]
        rows.append(line)
    rows.append(["합계(배당)", "", f"{tot[0]:,.0f}", "", "", f"{tot[1]:,.0f}", "", "", f"{tot[2]:,.0f}", "", "", f"{tot[3]:,.0f}", ""])
    rows.append(["매각차익(C.G)", "", f"{cg_tot[0]:,.0f}", "", "", f"{cg_tot[1]:,.0f}", "", "", f"{cg_tot[2]:,.0f}", "", "", f"{cg_tot[3]:,.0f}", ""])
    lab = {"CoC(%)_C.G제외": "CoC (C.G 제외)", "IRR(%)_C.G제외": "IRR (C.G 제외)", "IRR(%)_C.G포함": "IRR (C.G 포함)", "E. Multiple": "E.Multiple"}
    for k, name in lab.items():
        v = summ[k]
        fmt = (lambda x: f"x {x:.2f}") if k == "E. Multiple" else (lambda x: f"{x:.2%}")
        rows.append([name, "", fmt(v[0]), "", "", fmt(v[1]), "", "", fmt(v[2]), "", "", fmt(v[3]), ""])
    return rows


# ═════════════════════════════════════════════════════════════════════════
# 4. 본문 편집
# ═════════════════════════════════════════════════════════════════════════
def build():
    base = structure()
    prs = Presentation(str(base))
    S = list(prs.slides)
    print("슬라이드 수:", len(S))
    # 최종 순서 (0-index): 0 표지, 1 목차, 2 심의안건, 3 안건Ⅰ(1), 4 안건Ⅰ(2), 5 안건Ⅱ(1), 6 안건Ⅱ(2), 7 안건Ⅲ,
    # 8,9 참고 인수확약서, 10 당사 예상수익, 11 별첨표지, 12 별첨목차, 13,14 Highlights, 15 자산개요,
    # 16 입지, 17 입지보완1, 18 입지보완2, 19~22 시장현황, 23 시장보완1, 24 시장보완2, 25 Peer임대, 26 임대사례보완,
    # 27 Peer매매, 28 거래사례보완, 29 적정임대료 결론, 30 리츠개요, 31 투자구조, 32 사업구조도(IM), 33,34 Milestone,
    # 35 재원조달, 36~38 운영가정, 39,40 투자수익률, 41 Sell-down, 42 DSCR, 43 민감도, 44 CPM, 45,46 리스크,
    # 47 향후조치, 48 Appendix 간지, 49~52 Appendix, 53 종료
    assert len(S) == 54, len(S)
    tbl_tmpl = src_table(S[3])  # slide4 원본 표(2열) — 복제 원천

    # ── 표지·목차 ──────────────────────────────────────────────────────
    replace_runs(S[0], "예비투자심의위원회", "본 투자심의위원회")
    replace_runs(S[0], "August. 2026", "September. 2026")
    replace_runs(S[1], "예비투자심의위원회", "본 투자심의위원회")
    toc = by_text(S[1], "투자심의위원회 심의 개요")
    if toc:
        set_tf(toc.text_frame, ["투자심의위원회 심의 개요",
                                "투자심의위원회 심의 안건",
                                "안건 Ⅰ. 매매계약서의 체결",
                                "안건 Ⅱ. 주주간계약서의 체결",
                                "안건 Ⅲ. 코람코라이프인프라리츠 매입확약서(LOC) 날인",
                                "참고. 관계사 인수확약서 및 당사 예상수익",
                                "[별첨] 투자 대상 물건 설명자료"])

    # ── 심의안건 및 추진일정 ──────────────────────────────────────────
    s = S[2]
    lead(s, "본건은 안양물류센터 매입을 위한 매매계약·주주간계약 체결 및 코람코라이프인프라리츠 매입확약서(LOC) 날인의 건임",
         ["기준 일정상 매매계약 체결 10/2, 영업인가 완료 10/12, 주주간계약·대출약정 체결 10/13, 거래종결 10/22 예정"])
    t = src_table(s)
    # 3x3 → 4x3: 마지막 행 복제
    tbl = t._element.find(".//a:tbl", NS)
    trs = tbl.findall("a:tr", NS)
    tbl.append(copy.deepcopy(trs[-1]))
    tb = t.table
    vals = [["안건 Ⅰ", "매매계약서의 체결", "코크렙안양㈜와 매매대금 2,700억원의 부동산매매계약 체결 (거래종결 2026-10-22 예정)"],
            ["안건 Ⅱ", "주주간계약서의 체결", "우리투자증권·삼성증권·㈜엘에프·코람코라이프인프라리츠 간 주주간계약 체결 (배당·매입확약·지배구조)"],
            ["안건 Ⅲ", "코람코라이프인프라리츠 매입확약서(LOC) 날인", "제1종(24개월)·제2종(12개월) 종류주식 매입확약서를 주주간계약 체결일에 제출"]]
    for i, v in enumerate(vals, 1):
        for j in range(3):
            set_cell(tb.cell(i, j), v[j])
    for old, new in [("8월 12일", "8월 18일"), ("8월 [27]일", "8월 27일"), ("9월 2일", "8월 28일"),
                     ("9월 10일", "9월 3일"), ("9월 21일", "9월 17일"), ("10월 7일", "10월 2~13일"),
                     ("10월 29일", "10월 26일"), ("[8/24]", "[9/30]")]:
        replace_runs(s, old, new)
    for sub, lines in [("실사완료 및", ["영업인가 접수(9/11)", "실사 최종본(9/17)"]),
                       ("영업인가 완료", ["매매계약 체결(10/2)", "영업인가 완료(10/12)", "주주간·대출약정(10/13)"]),
                       ("투자심의위원회 개최\n• 매매계약서", ["투자심의위원회 개최", "• 매매계약·주주간계약 체결의 건", "• 매입확약서(LOC) 날인의 건"]),
                       ("• 양해각서 체결의 건", ["예비투자심의위원회", "• 양해각서 체결의 건", "• 실사기관 선정의 건"]),
                       ("회사설립", ["발기설립 완료"]),
                       ("사업일정은 협의에 따라", ["사업일정은 협의에 따라 변동될 수 있음.",
                                            "기준 일정: KLI 신규투자(안양) 추진일정(2026-08-11)",
                                            "코람코라이프인프라리츠 설립자본금 3억원 출자 및 발기설립 완료(8/28)"])]:
        sh = by_text(s, sub)
        if sh:
            set_tf(sh.text_frame, lines)
    g15 = [sh for sh in s.shapes if sh.shape_type == 6 and any("당 투자심의위원회" in x.text_frame.text for x in iter_shapes(sh.shapes) if x.has_text_frame)]
    if g15:
        g15[0].left = Inches(6.90)

    # ── 안건 Ⅰ (1) 거래 개요: slide4 표(14x2) 재사용 ────────────────────
    s = S[3]
    title(s, "> 안건 Ⅰ. 매매계약서 체결의 건 (1/2)")
    band_label(s, "매매계약 주요 조건")
    tb = src_table(s).table
    rows = [["구분", "주요내용"],
            ["체결 예정일", "2026년 10월 2일 (기준 일정) — 계약서상 체결일·거래종결예정일은 공란 [협의 중]"],
            ["체결당사자", "매도인: 코크렙안양 주식회사 (PFV)\n매수인: 신설 위탁관리부동산투자회사 (코람코라이프로지스위탁관리부동산투자회사, 영업인가 진행 중)"],
            ["매매목적물", "경기도 안양시 동안구 관양동 934 외 2필지 토지 및 지상 건물, 부속 동산·시설·설비 일체"],
            ["면적개요", "대지면적 15,287.5㎡ (4,624.47평) / 연면적 95,474.59㎡ (28,881.06평) / 임대면적 69,058.12㎡ (20,890.08평)"],
            ["매매대금", "금 2,700억원 (건물·동산 부가가치세 별도) — 담보 감정평가액 3,345억원 대비 −19.3%, 연면적 평당 935만원"],
            ["매매 방식", "거래종결일 현재의 법률상·사실상 현황 그대로(as-is, where-is) 매수"],
            ["거래종결시 지급금액", "매매대금 − 승계 임대차보증금 ± 정산금(세금·부담금·유틸리티·임대료 일할 정산) + 부가가치세"],
            ["임대차 승계", "매수인이 별지 임대차계약 전부 승계, 보증금 반환의무 인수 (승계동의서 미징구가 거래종결에 영향 없음)"],
            ["매도인의 진술 및 보장", "설립·존속, 권한, 제3자 동의, 상충 없음, 완전한 소유권(허용된 부담 제외), 소송 부존재 — 거래종결일로부터 2개월간 유효"],
            ["손해배상책임 한도", "매매대금의 5% 상당액 (135억원), 청구기한 거래종결일로부터 2개월"],
            ["거래신고 / 준거법", "계약일 30일 이내 부동산 거래신고 / 대한민국 법, 서울중앙지방법원 전속관할"]]
    tbl = src_table(s)._element.find(".//a:tbl", NS)
    trs = tbl.findall("a:tr", NS)
    for tr in trs[len(rows):]:
        tbl.remove(tr)
    for i, v in enumerate(rows):
        set_cell(tb.cell(i, 0), v[0]); set_cell(tb.cell(i, 1), v[1])
    for i, tr in enumerate(tbl.findall("a:tr", NS)):
        tr.set("h", str(int((0.30 if i == 0 else 0.44) * I)))
    fn = by_y(s, 7.45)
    if fn:
        set_tf(fn[0].text_frame, "※ 출처: 부동산매매계약서(clean, 2026-09-07) 제1~5조, 제9조 / 감정평가서(미래새한, 기준시점 2026-08-31). 상기 내용 중 일부는 협의 과정에서 변동될 수 있음.")

    # ── 안건 Ⅰ (2) 주요 조항·검토 의견: slide5 표 삭제 후 3열 표 복제 ──
    s = S[4]
    title(s, "> 안건 Ⅰ. 매매계약서 체결의 건 (2/2)")
    band_label(s, "주요 조항 및 검토 의견")
    delete(src_table(s))
    rows = [["조항", "주요 내용", "검토 의견"],
            ["제4조 진술 및 보장", "매도인: 설립·존속, 권한, 제3자 동의, 상충 없음, 완전한 소유권, 소송 부존재\n매수인: 권한, 실사 완료, 담보책임 배제 (as-is 매입 확인)", "실사 완료 진술로 미발견 하자는 매수인 부담 → 실사 결과의 계약 반영 필요"],
            ["제4.4조 존속기간", "진술·보장은 거래종결일로부터 2개월간 유효, 기간 내 서면청구 없으면 면책", "통상 대비 짧은 기간 — 잔여 리스크 인지"],
            ["제5조 확약", "거래종결 노력, 특정사항 통지, 거래신고, 임대차 승계, 매도인 운영 유지·협상 금지·세금 완납", "승계동의서 확보 현황 별도 관리"],
            ["제6조 선행조건", "확약·의무 이행, 진술·보장의 진실성, 중대한 부정적 법규·행정절차·소송 부존재", "영업인가·자금조달은 선행조건 아님 → 기준 일정 관리로 대응"],
            ["제7조 손해배상", "한도 매매대금의 5%, 청구기한 거래종결일로부터 2개월, 간접·특별·징벌적 손해 배제", "한도 135억원 수준"],
            ["제8조 해제", "서면 합의 / 불가항력 30일 지속 / 중요 위반 후 10영업일 미시정 / 도산절차 / 선행조건 미충족\n위약금 금액 [협의 중]", "자금조달 곤란은 불가항력에서 명시적으로 제외 — 조달 확정성 확보 필수"],
            ["제9조 일반조항", "양도 제한, 비밀유지(종결 후 1년), 비용 각자 부담, 준거법·전속관할", "—"]]
    clone_table(s, tbl_tmpl, rows, [1.7, 6.2, 2.9], row_h=0.63, header_h=0.30, top=2.58)
    fn = by_y(s, 7.45)
    if fn:
        set_tf(fn[0].text_frame, "※ 출처: 부동산매매계약서(clean, 2026-09-07) 제4~9조. 본 검토는 내부 검토이며 최종 법률자문은 외부 법무법인 확인이 필요함.")

    # ── 안건 Ⅱ (1) 주주 구성·배당 ──────────────────────────────────────
    s = S[5]
    title(s, "> 안건 Ⅱ. 주주간계약서 체결의 건 (1/2)")
    band_label(s, "주주 구성 및 배당 구조")
    delete(src_table(s))
    rows = [["구분", "금액(백만원)", "지분율", "배당", "매각차익 배분", "투자자"],
            ["제1종 종류주식", "24,000", "24.3%", "누적확정 연 7.0%", "없음", "우리투자증권 (LOC 완료)"],
            ["제2종 종류주식", "50,000", "50.6%", "누적확정 연 7.5%", "20%", "삼성증권 (LOC 완료)"],
            ["제3종 종류주식", "10,000", "10.1%", "무배당·무의결권", "없음", "㈜엘에프 (매도인 재투자)"],
            ["보통주식", "14,800", "15.0%", "잔여배당 (초기 1년 연 7.5%)", "80%", "코람코라이프인프라리츠"],
            ["합계", "98,800", "100.0%", "—", "100%", "—"]]
    clone_table(s, tbl_tmpl, rows, [1.9, 1.4, 1.1, 2.4, 1.5, 2.5], row_h=0.34, header_h=0.30, top=2.58)
    rows2 = [["구분", "주요 내용"],
             ["배당 순위 (제1.2조)", "제1종 연 7.0%·제2종 연 7.5% 누적배당 → 제3종 무배당 → 보통주 잔여배당. 보통주는 초기 1년간 발행가액의 연 7.5% 배당에 협조(자본준비금 감액 포함)"],
             ["유상증자 (제2.1~2.2조)", "매매계약상 거래종결일 2영업일 전까지 이사회 결의, 주주별 납입예정일 15:00까지 출자금 전액 납입"],
             ["유상감자 (제2.3조)", "증자 완료 후 발기인(코람코라이프인프라리츠) 보유 보통주식 300,000주 전부를 감자비율 100%(1주당 1,000원)로 유상감자"],
             ["제3종 유상감자 (제5조)", "발행일로부터 24개월 내 유상감자로 소각, ㈜엘에프에 발행가액 기준 감자대금 지급 (채권자보호절차 지연 시 최대 3개월 연장)"]]
    clone_table(s, tbl_tmpl, rows2, [2.3, 8.5], row_h=0.50, header_h=0.30, top=4.95)
    fn = by_y(s, 7.45)
    if fn:
        set_tf(fn[0].text_frame, "※ 출처: 주주간계약서 v22(clean) 제1.2조·제2조·제5조 / IM(2026.09) p.15. 종류주식 셀다운 예정으로 계약서상 인수 주식수 표는 지분 구성 확정 후 보충 예정.")

    # ── 안건 Ⅱ (2) 매입확약·지배구조 ───────────────────────────────────
    s = S[6]
    title(s, "> 안건 Ⅱ. 주주간계약서 체결의 건 (2/2)")
    band_label(s, "매입확약 · 지배구조 · 양도 제한")
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
    clone_table(s, tbl_tmpl, rows, [1.9, 6.7, 2.2], row_h=0.50, header_h=0.30, top=2.58)
    fn = by_y(s, 7.45)
    if fn:
        set_tf(fn[0].text_frame, "※ 출처: 주주간계약서 v22(clean) 제3~7조. 대괄호 항목은 협의 중. 본 검토는 내부 검토이며 최종 법률자문은 외부 법무법인 확인이 필요함.")

    # ── 안건 Ⅲ LOC ─────────────────────────────────────────────────────
    s = S[7]
    title(s, "> 안건 Ⅲ. 코람코라이프인프라리츠 매입확약서(LOC) 날인의 건")
    band_label(s, "매입확약서 주요 내용")
    delete(src_table(s))
    rows = [["구분", "제1종 종류주식", "제2종 종류주식"],
            ["확약 주체", "㈜코람코라이프인프라위탁관리부동산투자회사", "(좌 동)"],
            ["발행회사 / 기초자산", "신설 위탁관리부동산투자회사 / 안양물류센터 (관양동 934 외 2필지)", "(좌 동)"],
            ["우선배당률", "연 7.0%", "연 7.5%"],
            ["매입기한", "발행일로부터 24개월이 되는 달의 응당일", "발행일로부터 12개월이 되는 달의 응당일"],
            ["매입금액", "1주당 발행가액 25,000원 × 매입대상 주식수 (240억원 예정)", "1주당 발행가액 25,000원 × 매입대상 주식수 (500억원 예정)"],
            ["제출 시점", "주주간계약 체결일 (기준 일정 2026-10-13)", "(좌 동)"]]
    clone_table(s, tbl_tmpl, rows, [2.2, 4.3, 4.3], row_h=0.34, header_h=0.30, top=2.58)
    rows2 = [["구분", "내용"],
             ["선행조건", "① 발행회사의 적법한 영업인가 취득 및 안양물류센터 소유권 취득 ② 종류주주의 완전한 소유권 확보 ③ 질권·가압류 등 부담 부존재 ④ 동일 조건의 매도확약서 제출"],
             ["배당금 정산", "매입일이 포함된 배당기간의 배당금은 매입일 전일까지 종류주주, 매입일부터 KLI 귀속 — 배당 지급일로부터 5영업일 내 정산"],
             ["불이행 시 효과", "미매입잔액에 대해 매입기한 종료 익일부터 연 5.0% 지연손해금, 초과 손해는 별도 배상, 매입의무는 존속"],
             ["효력·우선순위", "본건 매입 목적 외 사용 불가, 제3자 권리 주장 불가, 허용된 양도의 양수인에게 동일 적용. 주주간계약과 상충 시 주주간계약 우선"]]
    clone_table(s, tbl_tmpl, rows2, [2.2, 8.6], row_h=0.50, header_h=0.30, top=5.10)
    fn = by_y(s, 7.45)
    if fn:
        set_tf(fn[0].text_frame, "※ 출처: 코람코 제1종·제2종 우선주 매입확약 공문(안)(투자자용) / 주주간계약서 v22 제3.5조·제4.5조. 발행회사명·주식수·문서번호는 확정 후 기재.")

    # ── 참고. 당사 예상수익 (수치 갱신) ───────────────────────────────
    s = S[10]
    for old, new in [("2,985억원", "2,989억원"), ("66.6억원", "66.7억원"), ("80.1억원", "80.2억원"),
                     ("225.0억원", "225.1억원"), ("4.88%", "4.89%")]:
        replace_runs(s, old, new)

    # ── Highlights 9.21 → 9.24 ─────────────────────────────────────────
    replace_runs(S[13], "9.21%", "9.24%")

    # ── 입지 보완 1: 시장조사 p.6(지리적 이점) + p.7(협업입지) ─────────
    s = S[17]
    strip_body(s)
    title(s, "> 2. 입지분석 (시장조사 — 광역 접근성)")
    lead(s, "안양시는 수도권 중부권역 물류벨트의 중심축으로 대형 화물차량의 광역도로망 접근성이 뛰어나 Last-mile 배송에 유리함",
         ["평촌·인덕원 일대 고밀주거와 지식산업 집적으로 협업 입지 및 인력 수급 여건 양호"])
    band_label(s, "광역 입지 및 협업입지 분석")
    p6 = render(MKT_PDF, 6, MKT_BOX, "mkt06"); p7 = render(MKT_PDF, 7, MKT_BOX, "mkt07")
    add_pic(s, p6, 0.50, 2.58, width=5.30)
    add_pic(s, p7, 5.90, 2.58, width=5.30)
    fn_add(s, "Source: ㈜코람코자산신탁 안양물류센터 매입 시장조사보고서(2026-08-10) p.6~7")

    # ── 입지 보완 2: p.9(중부권역) + p.12(중부권역 물류센터 현황) ───────
    s = S[18]
    strip_body(s)
    title(s, "> 2. 입지분석 (시장조사 — 중부권역)")
    lead(s, "중부권역은 서울 강남권 및 수도권 핵심 주거권역의 최인접 거점으로 라스트마일 기능을 수행하며 신규 공급이 극히 제한적임",
         ["김포·부천·서울동남권 등 핵심 물류단지는 대형 E-Commerce·택배사 장기임차로 공실률이 매우 낮음"])
    band_label(s, "중부권역 분석 및 권역 내 물류센터 현황")
    p9 = render(MKT_PDF, 9, MKT_BOX, "mkt09"); p12 = render(MKT_PDF, 12, MKT_BOX, "mkt12")
    add_pic(s, p9, 0.50, 2.58, width=5.30)
    add_pic(s, p12, 5.90, 2.58, width=5.30)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.9, p.12")

    # ── 시장 보완 1: p.31(권역별 임대시장) + p.32(임대료 추이) ──────────
    s = S[23]
    strip_body(s)
    title(s, "> 3. 시장현황 (시장조사 — 임대시장)")
    lead(s, "수도권 물류센터 임대시장은 권역별 차별화가 뚜렷하며, Last-mile 권역의 상온 임대료는 상승 추세를 지속함",
         ["본건이 속한 중부권역은 임대료 수준과 임차 수요 모두 수도권 상위권"])
    band_label(s, "권역별 임대시장 현황 및 임대료 추이")
    add_pic(s, render(MKT_PDF, 31, MKT_BOX, "mkt31"), 0.50, 2.58, width=5.30)
    add_pic(s, render(MKT_PDF, 32, MKT_BOX, "mkt32"), 5.90, 2.58, width=5.30)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.31~32")

    # ── 시장 보완 2: p.33(공실률 추이) + p.43(Last-mile 공급동향) ───────
    s = S[24]
    strip_body(s)
    title(s, "> 3. 시장현황 (시장조사 — 공실률·공급)")
    lead(s, "수도권 Last-mile 권역의 신규 공급은 2025년 이후 급감하고 있으며 상온센터 공실률은 낮은 수준을 유지함",
         ["공급 제한과 초과 수요가 맞물려 임대인 우위 시장이 지속될 것으로 예상"])
    band_label(s, "공실률 추이 및 Last-mile 권역 공급 동향")
    add_pic(s, render(MKT_PDF, 33, MKT_BOX, "mkt33"), 0.50, 2.58, width=5.30)
    add_pic(s, render(MKT_PDF, 43, MKT_BOX, "mkt43"), 5.90, 2.58, width=5.30)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.33, p.43")

    # ── 임대사례 보완 (p.45) ────────────────────────────────────────────
    s = S[26]
    strip_body(s)
    title(s, "> 4. Peer그룹 매매 및 임대현황 (시장조사 — 임대사례)")
    lead(s, "수도권 Last-mile 물류센터 임대사례의 Eff. Rent는 4.0~6.0만원/평 범위(평균 5.2만원/평)로, 본건 5.8만원/평은 상단에 위치함",
         ["서울복합·부천 삼정동(6.0만원/평) 등 신축 Core 자산과 유사한 수준이며 준공연도·규모를 감안 시 상승 여력 보유"])
    band_label(s, "수도권 Last-mile 유사자산 임대사례 (Savills)")
    add_pic(s, render(MKT_PDF, 45, MKT_BOX, "mkt45"), 0.50, 2.58)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.45 [Savills RE Strategy & Solutions]")

    # ── 거래사례 보완 (p.46) ────────────────────────────────────────────
    s = S[28]
    strip_body(s)
    title(s, "> 4. Peer그룹 매매 및 임대현황 (시장조사 — 거래사례)")
    lead(s, "수도권 Last-mile 유사자산 거래사례 대비 본건 매입가(평당 935만원)는 신축 Core 자산으로서 낮은 가격 수준임",
         ["시장조사상 적정 매입가 범위 891~1,014만원/평의 중하단에 위치"])
    band_label(s, "수도권 Last-mile 유사자산 거래사례 (Savills)")
    add_pic(s, render(MKT_PDF, 46, MKT_BOX, "mkt46"), 0.50, 2.58)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.46 [Savills RE Strategy & Solutions]")

    # ── 적정임대료 결론 (p.52 + 요약 표) ───────────────────────────────
    s = S[29]
    strip_body(s)
    title(s, "> 5. 시장조사 결론 — 적정 임대료 및 매매가")
    lead(s, "비교사례 보정 및 시장 상승률 2.5% 반영 결과, 본건 적정 Eff. Rent는 2026년 기준 6.14만원/평으로 현 수준(5.8만원/평) 대비 약 6% 상승 여력",
         ["적정 매입가 범위는 평당 891~1,014만원이며 본건 매입가 935만원/평은 범위 내에 위치"])
    band_label(s, "적정 임대가 산정 (시장조사 최종 의견)")
    add_pic(s, render(MKT_PDF, 52, MKT_BOX, "mkt52"), 0.50, 2.58, width=6.55)
    rows = [["구분", "시장조사 결론", "본건"],
            ["적정 Eff. Rent ('26)", "6.14만원/평 (범위 5.54~6.43만원)", "약 5.8만원/평"],
            ["예상 E.NOC (원/평)", "'27 63,000 → '28 64,575 → '29 66,189 → '30 67,844", "인상률 2.5% 가정"],
            ["적정 매입가", "평당 891~1,014만원", "평당 935만원"],
            ["투자 적정성 (Exit)", "Exit Cap 5.20%, 매각대금 3,667억원 (NOI N+1 18,959백만원)", "사업계획 Exit Cap 4.89%, 3,759억원"],
            ["시사점", "2029~2030년 만기 도래 임차인 재계약 시 시장 수준 반영으로 NOI 개선 가능", "—"]]
    clone_table(s, tbl_tmpl, rows, [1.35, 2.25, 0.95], row_h=0.62, header_h=0.30, top=2.58, left=7.25)
    fn_add(s, "Source: 시장조사보고서(2026-08-10) p.52~53, p.55 종합의견. 임대료 상승은 시장 상황에 따른 예상이며 확정된 것이 아님.")

    # ── 리츠개요 (수치 갱신) ──────────────────────────────────────────
    s = S[30]
    for old, new in [("2,985억원 (Equity : 984억원", "2,989억원 (Equity : 988억원"), ("9.21%", "9.24%"),
                     ("16.66%", "16.44%"), ("10.35%", "10.34%")]:
        replace_runs(s, old, new)
    fn = by_y(s, 7.45)
    if fn:
        set_tf(fn[0].text_frame, "총 자산가액: 부동산 장부가액 2,889억원 및 여유현금 100억원 예정 / 수익률은 재무분석보고서(우리회계법인) 기준이며, 재무모델(2026-09-15)은 2종 9.15%·보통주 16.18%로 산출 — 확정값 확인 중")

    # ── 투자구조 (수치 갱신) ──────────────────────────────────────────
    s = S[31]
    for old, new in [("[2,985]", "[2,989]"), ("[144]", "[148]"), ("[185]", "[189]")]:
        replace_runs(s, old, new)

    # ── 사업구조도 (IM p.6) ─────────────────────────────────────────────
    s = S[32]
    strip_body(s)
    title(s, "> 2. 투자구조 — 사업구조도 (IM 기준)")
    lead(s, "신설 리츠를 설립한 후 아래 구조로 매입·운용하며, 공모예외기관이 발기인으로 참여하고 종류주주·보통주주 증자로 Equity를 모집함",
         ["코람코라이프인프라리츠가 자리츠를 설립하고 자산관리회사(코람코자산신탁)·자산보관회사·사무수탁회사와 위탁계약 체결"])
    band_label(s, "사업 구조도")
    add_pic(s, render(IM_PDF, 6, IM_BOX, "im06"), 0.50, 2.58)
    fn_add(s, "Source: 안양물류센터 Information Memorandum(2026.09) p.5 사업 구조도")

    # ── Milestone 수치 ────────────────────────────────────────────────
    for idx in (33, 34):
        for old, new in [("[2,985]", "[2,989]"), ("[144]", "[148]")]:
            replace_runs(S[idx], old, new)

    # ── 재원조달 표 갱신 (IM p.15 S&U) ────────────────────────────────
    s = S[35]
    lead_sh = by_text(s, "총 사업비 2,985억원")
    if lead_sh:
        set_tf(lead_sh.text_frame, ["총 사업비 2,989억원으로 매입금액(90.3%), 매입부대비용(6.3%), 예비현금(3.3%)로 구성",
                                    "Carry기간 1종 우선주(7.0%) 및 2종 우선주(7.5%) 등 확정배당 지급을 위한 Overfunding금액 100억원 편성"])
    tabs = [sh for sh in s.shapes if sh.has_table]
    for t in tabs:
        tb = t.table
        if len(tb.columns) == 4:   # Uses
            upd = {1: ("270,000", "90.3%", "매입금액 2,700억원, 감정가 대비 −19.3%"),
                   2: ("1,350", "0.5%", None), 3: ("12,498", "4.2%", None),
                   4: ("662", "0.2%", None), 5: ("4,390", "1.5%", None),
                   6: ("10,000", "3.3%", None), 7: ("298,900", "100.0%", None)}
            for ri, (a, b, c) in upd.items():
                set_cell(tb.cell(ri, 1), a); set_cell(tb.cell(ri, 2), b)
                if c:
                    set_cell(tb.cell(ri, 3), c)
        elif len(tb.columns) == 6:  # Sources
            upd = {1: ("2,100", "0.7%", None, None),
                   2: ("160,000", "53.5%", "금리 4.78%, 수수료 1.45% (All-in 5.50%)", "우리은행 심사 완료(9/3), 대출약정 협의 중"),
                   3: ("38,000", "12.7%", "금리 6.00%, 수수료 1.00% (All-in 6.50%)", "우리은행 등, 심사 완료"),
                   4: ("24,000", "8.0%", None, "우리투자증권 인수확약서(LOC) 확보"),
                   5: ("50,000", "16.7%", None, "삼성증권 인수확약서(LOC) 확보"),
                   6: ("10,000", "3.3%", None, "㈜엘에프 재투자 (주주간계약 협의 중)"),
                   7: ("14,800", "5.0%", None, "KLI리츠 이사회 승인 예정(10/1)"),
                   8: ("298,900", "100.0%", None, None)}
            for ri, (a, b, c, d) in upd.items():
                set_cell(tb.cell(ri, 1), a); set_cell(tb.cell(ri, 2), b)
                if c:
                    set_cell(tb.cell(ri, 3), c)
                if d:
                    set_cell(tb.cell(ri, 5), d)

    # ── 운영가정 4.88 → 4.89 ──────────────────────────────────────────
    replace_runs(S[38], "4.88%", "4.89%")

    # ── 투자수익률: 원본 표 → IM p.20/21 표 이미지 ───────────────────
    for idx, page, name, subs in ((39, 20, "im20", ["제1종 우선주 및 제2종 우선주의 경우 누적적 확정 배당수익률 7.0%, 7.5%를 각각 지급 예정 (IRR 7.12% / 9.24%)"]),
                                   (40, 21, "im21", ["3종우선주는 배당 미지급·2년 후 원본감자 예정이며, 보통주(148억원)는 초기 1년 7.5% 배당 후 잔여배당·매각차익 80% 배분 (IRR 16.44%)"])):
        s = S[idx]
        strip_body(s, keep_footnote=True)
        lead(s, "사업기간 10년을 가정한 투자자별 수익률은 다음과 같음 (IM 2026.09 기준)", subs)
        add_pic(s, render(IM_PDF, page, IM_TBL_BOX, name), 0.50, 2.55, max_h=4.80)
    fn = by_y(S[40], 7.45)
    for sh in by_y(S[40], 7.48, tol=0.08):
        if "Source" in sh.text_frame.text:
            set_tf(sh.text_frame, "Source: 안양물류센터 IM(2026.09) p.19~20 / 재무분석보고서(우리회계법인). 재무모델(2026-09-15)은 2종 9.15%·보통주 16.18%로 산출 — 확정값 확인 중")

    # ── Sell-down 표 (KLI 단계별 매입) ────────────────────────────────
    s = S[41]
    title(s, "> 6. 투자수익률 — KLI 단계별 매입(Sell-down) 배당스케줄 및 수익률")
    band_label(s, "코람코라이프인프라리츠 입장의 단계별 매입 현금흐름 (단위: 백만원)")
    delete(src_table(s))
    lead_sh = by_y(s, 1.14)
    if not lead_sh:
        # slide4 계열에는 리드 박스가 없으므로 slide 2의 리드 박스를 복제
        src = by_y(S[2], 1.14)[0]
        s.shapes._spTree.append(copy.deepcopy(src._element))
    lead(s, "보통주(148억원) 초기 출자 후 12개월 시점 2종(505억원), 24개월 시점 1종(242억원)을 순차 매입하여 총 895억원을 투자하는 구조임",
         ["KLI 통합 기준 CoC 5.40%(C.G 제외), IRR 10.83%(C.G 포함), E.Multiple 2.19x — 매입가에 간주취득세(1종 2.2억·2종 4.5억)가 포함되어 1종 Yield는 6.94%"])
    rows = selldown_rows()
    clone_table(s, tbl_tmpl, rows, [0.75] + [0.83] * 12, row_h=0.255, header_h=0.30, top=2.58, size=8)
    fn = by_y(s, 7.45)
    if fn:
        set_tf(fn[0].text_frame, "※ 출처: 재무모델 v01(2026-09-15) Sell-down 시트. 연차는 운용 개시(2026-10-31) 기준 반기 2개 합산, 잔액은 연말 기준. 통합 배당+C.G는 운영배당과 매각차익 배분액의 합. 수익률은 사업계획 가정에 따른 예상치임.")

    # ── 리스크 4.88 → 4.89 ────────────────────────────────────────────
    replace_runs(S[46], "4.88%", "4.89%")

    # ── 향후조치 (본 투심 이후) ────────────────────────────────────────
    s = S[47]
    lead_sh = by_text(s, "본 건 (예비)투자심의위원회 이후")
    if lead_sh:
        set_tf(lead_sh.text_frame, ["본 투자심의위원회 이후 향후 보완 예정사항은 다음과 같으며, 관계자와의 협의를 통해 10월 2일 매매계약, 10월 13일 주주간계약·대출약정 체결 예정",
                                    "• 추가 심의위원 의견 청취 후 계약 체결 전 주요 계약서 내용 반영 및 보완할 예정임"])
    tb = src_table(s).table
    vals = [["구분", "본 투자심의 후 Checklist", "계약 체결 前 보완방향"],
            ["계약 문안 확정", "매매계약 체결일·거래종결예정일·위약금·지연손해금률 공란 / 주주간계약 인수 주식수 표(셀다운 반영)·이사 수·지명권 조항 대괄호",
             "법무팀 협업 하에 9/22 매매계약서·주주간계약 합의 시 확정, 매입확약서 발행회사명·주식수 기재 후 날인"],
            ["수치 확정", "재무모델(9/15) 2종 IRR 9.15%·보통주 16.18% vs IM·재무분석보고서 9.24%·16.44% 차이, WALE 산정기준(6.8년 vs 6.66년)",
             "재무팀 검토로 원인 규명 및 기준값 확정 후 투자자 Q&A·IM 갱신"],
            ["실사 반영", "대지면적 차이(15,287.5㎡ vs 모델 15,019.5㎡, 필지 수) / 법률실사 Recommendation 41건 중 미해소 항목 / 실사보고서 최종본(9/17) 반영",
             "등기부 확인 및 SPA 선행조건·특약 반영, 물리실사 수익적 지출(5년 12.6억원) 사업계획 반영"],
            ["거래관계 투명성", "매도인 주주(㈜엘에프 95%·코람코 5%)의 3종 재투자 및 매도인·매수인 AMC 동일(코람코자산신탁)에 따른 이해관계인 거래 이슈",
             "리스크관리팀·법무팀 협업 하에 영업인가 및 거래가격 공정성 증빙(감정평가·시장조사) 구비, 외부 법률의견 반영"]]
    for i, v in enumerate(vals):
        for j in range(3):
            set_cell(tb.cell(i, j), v[j])

    prs.save(str(OUT))
    print("생성 완료:", OUT)
    return OUT


def fn_add(slide, text):
    """각주 텍스트박스: 기존 y≥7.45 텍스트가 있으면 갱신, 없으면 slide 3 의 각주 서식을 복제."""
    ex = [sh for sh in slide.shapes if sh.has_text_frame and (sh.top or 0) / I >= 7.40]
    if ex:
        set_tf(ex[0].text_frame, text)
        return
    prs_slide = slide
    # 같은 프레젠테이션의 3번째 슬라이드 각주(TextBox 108)를 복제
    src = None
    for sl in prs_slide.part.package.presentation_part.presentation.slides:
        for sh in sl.shapes:
            if sh.has_text_frame and abs((sh.top or 0) / I - 7.45) < 0.03 and sh.width / I > 8:
                src = sh
                break
        if src:
            break
    el = copy.deepcopy(src._element)
    slide.shapes._spTree.append(el)
    set_tf(slide.shapes[-1].text_frame, text)


if __name__ == "__main__":
    build()
