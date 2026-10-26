# -*- coding: utf-8 -*-
"""대주별 인출요청서(Tranche A-1·A-2·B) 생성.

양식(원본 불변): 약정서/20261026_대출약정_인출요청서_양식_운용역제공.docx (대출약정서 최종본 별첨 가 발췌, 운용역 업로드)
대주·금액: 대출약정서 최종본 별지(대주목록 및 대주별 대출약정금)
참고 서식: KLI 2025-05-29 대주별 인출요청서(Tranche별 부제, 대주별 귀중, '아래 대출약정에 기재된 Tranche ○ 대주로서')
일자(가정, 기준 일정): 대출약정서 2026-10-27, 인출요청서 2026-10-28(제4조 제1항: 인출예정일 1영업일 전), 인출일 2026-10-29
"""
import copy
import os

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(HERE, "20261026_대출약정_인출요청서_양식_운용역제공.docx")
REQ_DATE, LOAN_DATE, DRAW_DATE = "2026년 10월 28일", "2026년 10월 27일", "2026년 10월 29일"

# 대출약정서 최종본 별지 기재 그대로
LENDERS = {
    "A-1": [
        ("롯데캐피탈 주식회사", "금 삼백억원(￦30,000,000,000)"),
        ("주식회사 우리카드", "금 일백억원(￦10,000,000,000)"),
        ("신용협동조합중앙회", "금 일억원(￦100,000,000)"),
        ("서로신용협동조합", "금 사십억원(￦4,000,000,000)"),
        ("수지신용협동조합", "금 이십팔억원(￦2,800,000,000)"),
        ("정선신용협동조합", "금 이십사억원(￦2,400,000,000)"),
        ("연희신용협동조합", "금 십사억원(￦1,400,000,000)"),
        ("금모래신용협동조합", "금 십팔억원(￦1,800,000,000)"),
        ("남인천신용협동조합", "금 십억원(￦1,000,000,000)"),
        ("미호신용협동조합", "금 십삼억원(￦1,300,000,000)"),
        ("대구대서신용협동조합", "금 십삼억원(￦1,300,000,000)"),
        ("한빛신용협동조합", "금 십칠억원(￦1,700,000,000)"),
        ("양동행복 신용협동조합", "금 십구억원(￦1,900,000,000)"),
        ("서귀포동부신용협동조합", "금 십이억원(￦1,200,000,000)"),
        ("광주어룡신용협동조합", "금 이십칠억원(￦2,700,000,000)"),
        ("광주미래신용협동조합", "금 십사억원(￦1,400,000,000)"),
        ("새안양 신용협동조합", "금 사십억원(￦4,000,000,000)"),
        ("남수원 신용협동조합", "금 삼억원(￦300,000,000)"),
        ("북인천신용협동조합", "금 칠억원(￦700,000,000)"),
        ("수협은행", "금 일백억원(￦10,000,000,000)"),
        ("주식회사 에스비아이저축은행", "금 일백억원(￦10,000,000,000)"),
    ],
    "A-2": [("주식회사 우리은행", "금 칠백억원(￦70,000,000,000)")],
    "B": [
        ("키움캐피탈 주식회사", "금 이백억원(￦20,000,000,000)"),
        ("엠지캐피탈 주식회사", "금 일백억원(￦10,000,000,000)"),
        ("주식회사 디비저축은행", "금 팔십억원(￦8,000,000,000)"),
    ],
}
OUTS = {"A-1": "20261026_대출약정_인출요청서_TrA-1_대주별21개사_v02.docx",
        "A-2": "20261026_대출약정_인출요청서_TrA-2_우리은행_v02.docx",
        "B": "20261026_대출약정_인출요청서_TrB_대주별3개사_v02.docx"}


def ptext(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def set_text(p, new):
    """문단의 첫 글자 런 서식을 유지하고 나머지 글자 런은 지운다."""
    runs = [r for r in p.findall(qn("w:r")) if r.find(qn("w:t")) is not None]
    first = runs[0]
    for r in runs[1:]:
        p.remove(r)
    for t in first.findall(qn("w:t"))[1:]:
        first.remove(t)
    t = first.find(qn("w:t"))
    t.text = new
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return first


def new_par_like(p_src, text, align=None, size_half_pt=None, bold=False):
    p = copy.deepcopy(p_src)
    for r in p.findall(qn("w:r")):
        p.remove(r)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        fonts.set(qn(a), "Times New Roman" if a != "w:eastAsia" else "바탕")
    rpr.append(fonts)
    if bold:
        rpr.append(OxmlElement("w:b"))
    if size_half_pt:
        for tag in ("w:sz", "w:szCs"):
            e = OxmlElement(tag)
            e.set(qn("w:val"), str(size_half_pt))
            rpr.append(e)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    p.append(r)
    if align:
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            p.insert(0, ppr)
        jc = ppr.find(qn("w:jc"))
        if jc is None:
            jc = OxmlElement("w:jc")
            ppr.append(jc)
        jc.set(qn("w:val"), align)
    return p


def page_break_before(p):
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    ppr.insert(0, OxmlElement("w:pageBreakBefore"))


def cell_set(tc, text):
    p = tc.find(qn("w:p"))
    runs = [r for r in p.findall(qn("w:r"))]
    if runs:
        set_text(p, text)
    else:
        # 빈 칸: 머리 행 글자 서식을 빌려 런 생성
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        fonts = OxmlElement("w:rFonts")
        for a in ("w:ascii", "w:hAnsi"):
            fonts.set(qn(a), "Times New Roman")
        fonts.set(qn("w:eastAsia"), "바탕")
        rpr.append(fonts)
        for tag in ("w:sz", "w:szCs"):
            e = OxmlElement(tag)
            e.set(qn("w:val"), "22")
            rpr.append(e)
        r.append(rpr)
        t = OxmlElement("w:t")
        t.text = text
        r.append(t)
        p.append(r)



def restart_numbering(doc, els):
    """복제한 확인사항 목록 번호가 장마다 (1)부터 시작하도록 새 번호 정의를 붙인다."""
    numbering = doc.part.numbering_part.element
    nums = {n.get(qn("w:numId")): n for n in numbering.findall(qn("w:num"))}
    new_ids = {}
    for e in els:
        for np_ in e.iter(qn("w:numPr")):
            nid = np_.find(qn("w:numId"))
            old = nid.get(qn("w:val"))
            if old not in new_ids:
                abs_id = nums[old].find(qn("w:abstractNumId")).get(qn("w:val"))
                nxt = str(max(int(k) for k in nums) + 1)
                num = OxmlElement("w:num")
                num.set(qn("w:numId"), nxt)
                a = OxmlElement("w:abstractNumId")
                a.set(qn("w:val"), abs_id)
                num.append(a)
                ov = OxmlElement("w:lvlOverride")
                ov.set(qn("w:ilvl"), "0")
                so = OxmlElement("w:startOverride")
                so.set(qn("w:val"), "1")
                ov.append(so)
                num.append(ov)
                numbering.append(num)
                nums[nxt] = num
                new_ids[old] = nxt
            nid.set(qn("w:val"), new_ids[old])


def build(tranche):
    doc = Document(TPL)
    body = doc.element.body
    sect = body.find(qn("w:sectPr"))
    tpl = [copy.deepcopy(e) for e in body if e.tag != qn("w:sectPr")]
    for e in [e for e in body if e.tag != qn("w:sectPr")]:
        body.remove(e)
    label = f"Tranche {tranche}"
    for idx, (name, amount) in enumerate(LENDERS[tranche]):
        els = [copy.deepcopy(e) for e in tpl]
        paras = [e for e in els if e.tag == qn("w:p")]
        title = paras[0]
        assert ptext(title) == "인출요청서"
        # 부제 (v02: "별첨 가" 표시 삭제 — 운용역 요청 2026-10-26)
        sub = new_par_like(paras[2], f"-{label} 대출 인출용-", align="center")
        if idx:
            page_break_before(title)
        # 일자·수신
        set_text(paras[2], REQ_DATE)
        set_text(paras[3], f"{name} 귀중")
        set_text(paras[4], f"(아래 대출약정에 기재된 {label} 대주로서)")
        body_p = next(p for p in paras if ptext(p).startswith("주식회사 코람코라이프로지스위탁관리부동산투자회사(이하"))
        s = ptext(body_p)
        for a, b in [("2026년 월  일자", f"{LOAN_DATE}자"),
                     ("인출 의사를 통지하며", f"{label} 대출의 인출 의사를 통지하며"),
                     ("______년 ___월 ___일(이하 “인출일”)", f"{DRAW_DATE}(이하 “인출일”)")]:
            assert s.count(a) == 1, (a, s)
            s = s.replace(a, b)
        set_text(body_p, s)
        tbl = next(e for e in els if e.tag == qn("w:tbl"))
        rows = tbl.findall(qn("w:tr"))
        assert len(rows) == 2
        cells = rows[1].findall(qn("w:tc"))
        cell_set(cells[0], f"{label} 대출")
        cell_set(cells[1], amount)
        # 조립: 제목 → 부제 → 나머지
        # 한 쪽 유지: 빈 제목줄(paras[1])과 확인사항 머리말 다음 빈 줄 1개를 뺀다
        confirm = next(p for p in paras if ptext(p).startswith("차주는 이 인출요청서 일자 현재"))
        gap = els[els.index(confirm) + 1]
        assert gap.tag == qn("w:p") and not ptext(gap).strip()
        out = [title, sub] + [e for e in els if e is not title and e is not paras[1] and e is not gap]
        if idx:
            restart_numbering(doc, out)
        for e in out:
            sect.addprevious(e)
    path = os.path.join(HERE, OUTS[tranche])
    doc.save(path)
    return path


if __name__ == "__main__":
    for t in ("A-1", "A-2", "B"):
        print("saved", build(t), len(LENDERS[t]))
