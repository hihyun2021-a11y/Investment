# -*- coding: utf-8 -*-
"""대출금 입금(인출) 요청 공문 v02 — KLI 공문 서식(리파이낸싱 동의 공문 Draft 261023)을 틀로 사용.
틀(원본 불변): 20261028_참고_KLI공문서식_리파이낸싱동의공문Draft_261023.docx
내용 근거: 대출약정서 최종 체결본(2026-10-27) 별지, 날인 인출요청서 25부(2026-10-28, 인출일 10-29)
출력: 20261028_대출약정_대출금입금요청공문_KLL명의_v02.docx
"""
import copy
import os

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(HERE, "20261028_참고_KLI공문서식_리파이낸싱동의공문Draft_261023.docx")
OUT = os.path.join(HERE, "20261028_대출약정_대출금입금요청공문_KLL명의_v02.docx")
BORROWER = "주식회사 코람코라이프로지스위탁관리부동산투자회사"
FONT = "HY신명조"

ROWS = [  # (구분, 대주, 약정액 백만원)
    ("Tr.A-1", "롯데캐피탈(주)", 30000), ("Tr.A-1", "(주)우리카드", 10000), ("Tr.A-1", "신용협동조합중앙회", 100),
    ("Tr.A-1", "서로신용협동조합", 4000), ("Tr.A-1", "수지신용협동조합", 2800), ("Tr.A-1", "정선신용협동조합", 2400),
    ("Tr.A-1", "연희신용협동조합", 1400), ("Tr.A-1", "금모래신용협동조합", 1800), ("Tr.A-1", "남인천신용협동조합", 1000),
    ("Tr.A-1", "미호신용협동조합", 1300), ("Tr.A-1", "대구대서신용협동조합", 1300), ("Tr.A-1", "한빛신용협동조합", 1700),
    ("Tr.A-1", "양동행복 신용협동조합", 1900), ("Tr.A-1", "서귀포동부신용협동조합", 1200), ("Tr.A-1", "광주어룡신용협동조합", 2700),
    ("Tr.A-1", "광주미래신용협동조합", 1400), ("Tr.A-1", "새안양 신용협동조합", 4000), ("Tr.A-1", "남수원 신용협동조합", 300),
    ("Tr.A-1", "북인천신용협동조합", 700), ("Tr.A-1", "수협은행", 10000), ("Tr.A-1", "(주)에스비아이저축은행", 10000),
    ("Tr.A-2", "(주)우리은행", 70000),
    ("Tr.B", "키움캐피탈(주)", 20000), ("Tr.B", "엠지캐피탈(주)", 10000), ("Tr.B", "(주)디비저축은행", 8000),
]
assert sum(r[2] for r in ROWS) == 198000


def ptext(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def set_text(p, new):
    runs = [r for r in p.findall(qn("w:r")) if r.find(qn("w:t")) is not None]
    first = runs[0]
    for r in runs[1:]:
        p.remove(r)
    ts = first.findall(qn("w:t"))
    for t in ts[1:]:
        first.remove(t)
    ts[0].text = new
    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def run(text, size=None, bold=False):
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    f = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:eastAsia", "w:hAnsi"):
        f.set(qn(a), FONT)
    rpr.append(f)
    if bold:
        rpr.append(OxmlElement("w:b"))
    if size:
        for tag in ("w:sz", "w:szCs"):
            e = OxmlElement(tag)
            e.set(qn("w:val"), str(int(size * 2)))
            rpr.append(e)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


def table():
    """KLI 공문 글꼴(HY신명조)에 맞춘 단순 격자 표."""
    hdr = ["구분", "대    주", "약정액", "인출 전 잔액", "금번 인출액", "인출 후 미인출한도"]
    widths = [1000, 2900, 1250, 1250, 1250, 1500]  # twips 합계 9150(본문 폭 이내)
    tbl = OxmlElement("w:tbl")
    pr = OxmlElement("w:tblPr")
    for tag, attrs in (("w:tblW", {"w:w": str(sum(widths)), "w:type": "dxa"}), ("w:jc", {"w:val": "center"}),
                       ("w:tblLayout", {"w:type": "fixed"})):
        e = OxmlElement(tag)
        for k, v in attrs.items():
            e.set(qn(k), v)
        pr.append(e)
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{side}")
        for k, v in (("w:val", "single"), ("w:sz", "4"), ("w:space", "0"), ("w:color", "000000")):
            e.set(qn(k), v)
        b.append(e)
    pr.append(b)
    mar = OxmlElement("w:tblCellMar")
    for side in ("left", "right"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:w"), "70")
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    pr.append(mar)
    tbl.append(pr)
    grid = OxmlElement("w:tblGrid")
    for w in widths:
        g = OxmlElement("w:gridCol")
        g.set(qn("w:w"), str(w))
        grid.append(g)
    tbl.append(grid)

    def tr(cells, fill=None, bold=False, merge=None, header=False):
        row = OxmlElement("w:tr")
        trpr = OxmlElement("w:trPr")
        trpr.append(OxmlElement("w:cantSplit"))
        if header:  # 다음 쪽으로 넘어가면 머리행 반복
            trpr.append(OxmlElement("w:tblHeader"))
        h = OxmlElement("w:trHeight")
        h.set(qn("w:val"), "250")
        trpr.append(h)
        row.append(trpr)
        j = 0
        for k, (txt, align) in enumerate(cells):
            tc = OxmlElement("w:tc")
            tcpr = OxmlElement("w:tcPr")
            span = 2 if (merge == "first2" and k == 0) else 1
            w = sum(widths[j:j + span])
            tw = OxmlElement("w:tcW")
            tw.set(qn("w:w"), str(w))
            tw.set(qn("w:type"), "dxa")
            tcpr.append(tw)
            if span > 1:
                gs = OxmlElement("w:gridSpan")
                gs.set(qn("w:val"), str(span))
                tcpr.append(gs)
            if isinstance(txt, tuple):  # 세로 병합 표시 (text, 'restart'|'continue')
                vm = OxmlElement("w:vMerge")
                if txt[1] == "restart":
                    vm.set(qn("w:val"), "restart")
                tcpr.append(vm)
                txt = txt[0]
            if fill:
                sh = OxmlElement("w:shd")
                sh.set(qn("w:val"), "clear")
                sh.set(qn("w:color"), "auto")
                sh.set(qn("w:fill"), fill)
                tcpr.append(sh)
            va = OxmlElement("w:vAlign")
            va.set(qn("w:val"), "center")
            tcpr.append(va)
            tc.append(tcpr)
            p = OxmlElement("w:p")
            ppr = OxmlElement("w:pPr")
            sp = OxmlElement("w:spacing")
            sp.set(qn("w:before"), "0")
            sp.set(qn("w:after"), "0")
            sp.set(qn("w:line"), "240")
            sp.set(qn("w:lineRule"), "auto")
            ppr.append(sp)
            jc = OxmlElement("w:jc")
            jc.set(qn("w:val"), align)
            ppr.append(jc)
            p.append(ppr)
            if txt:
                p.append(run(txt, 9, bold))
            tc.append(p)
            row.append(tc)
            j += span
        tbl.append(row)

    tr([(h, "center") for h in hdr], fill="D9D9D9", bold=True, header=True)
    for g, label in (("Tr.A-1", "Tr.A-1"), ("Tr.A-2", "Tr.A-2"), ("Tr.B", "Tr.B")):
        items = [x for x in ROWS if x[0] == g]
        for i, (_, name, amt) in enumerate(items):
            first = (label, "restart") if i == 0 else ("", "continue")
            tr([(first, "center"), (name, "center"), (f"{amt:,}", "right"), ("0", "right"), (f"{amt:,}", "right"), ("0", "right")])
        tot = sum(x[2] for x in items)
        tr([("소    계", "center"), (f"{tot:,}", "right"), ("-", "center"), (f"{tot:,}", "right"), ("0", "right")],
           fill="F2F2F2", bold=True, merge="first2")
    tr([("합    계", "center"), ("198,000", "right"), ("-", "center"), ("198,000", "right"), ("0", "right")],
       fill="E7E6E6", bold=True, merge="first2")
    return tbl


doc = Document(TPL)
body = doc.element.body
P = [e for e in body if e.tag == qn("w:p")]
assert ptext(P[0]).startswith("주식회사 코람코라이프인프라")

# 머리글·문서정보
set_text(P[0], BORROWER)
set_text(P[1], "06158 서울시 강남구 삼성로 511(삼성동) 4층  안용현 차장 TEL : 02-787-0049")
t3 = ptext(P[3])
set_text(P[3], t3.replace("KLI 2026-10-[  ]", "KLL 2026-10-[  ]").replace("2026. 10. [  ].", "2026. 10. 28."))
set_text(P[4], "수   신: 주식회사 우리은행 등 대주기관(수신처 참조)")
set_text(P[5], "참   조: 업무 담당자")
set_text(P[6], "제   목: 코람코라이프로지스리츠 담보대출 인출(대출금 입금) 요청의 건")

# 본문
set_text(P[10], ("2. 2026년 10월 27일 체결한 『코람코라이프로지스리츠 담보대출』 대출약정서(이하 “대출약정서”)와 관련하여, "
                 f"당사({BORROWER}, 이하 “차주”)는 대출약정서 제4조 제1항에 따라 2026년 10월 28일 각 대주에게 인출요청서를 "
                 "제출하였습니다. 이에 귀사가 약정하신 대출금을 아래와 같이 입금하여 주시기 바랍니다."))
items = [
    "1) 입금일시: 2026년 10월 29일(목)",
    f"2) 입금계좌: 우리은행 1005-504-949583 (예금주: {BORROWER})",
    "3) 대출약정금 합계: 금 일천구백팔십억원(￦198,000,000,000) - Tranche A-1 금 900억원, Tranche A-2 금 700억원, Tranche B 금 380억원",
    "4) 대주별 입금금액: 아래 표와 같음 (단위: 백만원)",
]
for p, txt in zip(P[13:17], items):
    set_text(p, txt)
    ppr = p.find(qn("w:pPr"))
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        ppr.append(jc)
    jc.set(qn("w:val"), "left")  # 양쪽맞춤 시 계좌번호 줄 간격 벌어짐 방지
body.remove(P[17])           # 기존 5) 항목
P[16].addnext(table())
set_text(P[19], "3. 대출금 입금과 관련하여 문의사항이 있으시면 당사 담당자(안용현 차장 02-787-0049, 염진경 과장 02-787-0139)에게 연락하여 주시기 바랍니다.")
for i in (20, 21, 22, 23, 24, 25, 26, 27, 28):  # ①②③, 4. 항목과 사이 빈 줄 제거
    body.remove(P[i])
set_text(P[29], "별첨 1. 차주 인출요청서 25부(대주별 각 1부).  끝.")
for i in (31, 32, 33):     # 서명 위 빈 줄 일부 축소
    body.remove(P[i])
set_text(P[35], BORROWER)
set_text(P[36], "대표이사 원 재 이")

# 수신처
recv = OxmlElement("w:p")
ppr = OxmlElement("w:pPr")
bd = OxmlElement("w:pBdr")
top = OxmlElement("w:top")
for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "4"), ("w:color", "000000")):
    top.set(qn(k), v)
bd.append(top)
ppr.append(bd)
sp = OxmlElement("w:spacing")
sp.set(qn("w:before"), "240")
ppr.append(sp)
recv.append(ppr)
recv.append(run("수신처: " + ", ".join(["(주)우리은행"] + [r[1] for r in ROWS if r[0] != "Tr.A-2"]), 8))
P[36].addnext(recv)

doc.core_properties.author = "코람코 상장리츠팀"
doc.save(OUT)
print("saved", OUT)
