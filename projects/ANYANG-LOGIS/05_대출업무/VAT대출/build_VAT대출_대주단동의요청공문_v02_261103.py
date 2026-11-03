# -*- coding: utf-8 -*-
"""VAT 대출 대주단 동의요청 공문 v02 — 운용역 수정본(v01_운용역수정, 2026-11-02)을 틀로
'계좌 신규 개설 + 해당 계좌 예금근질권 + 부수행위'까지 포괄 동의하도록 문구 보완.
근거: 담보대출약정서(BKL v38) 제8조 제1항(프로젝트계좌 외 계좌 추가 개설 시 대주 사전 서면동의),
      제8조 제3항(본건 부동산 관련 수입금의 수입금계좌 입금), 제11조 제7항 제(1)·(2)호, 제13조
틀(원본 불변): 20261102_공문_VAT대출_대주단동의요청_KLL명의_v01_운용역수정.docx
출력: 20261103_공문_VAT대출_대주단동의요청_KLL명의_v02.docx
"""
import copy
import os

from docx import Document
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "20261102_공문_VAT대출_대주단동의요청_KLL명의_v01_운용역수정.docx")
OUT = os.path.join(HERE, "20261103_공문_VAT대출_대주단동의요청_KLL명의_v02.docx")


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


def add_after(anchor, like, text):
    p = copy.deepcopy(like)
    set_text(p, text)
    anchor.addnext(p)
    return p


doc = Document(SRC)
body = doc.element.body
P = [e for e in body if e.tag == qn("w:p")]
assert ptext(P[20]).startswith("6) 담보") and ptext(P[37]).startswith("동  의  서")
ITEM, SUB = P[19], P[21]          # 항목 문단(1) 2) …), 하위 문단(① ②) 서식

# ── 제목·요청 취지
set_text(P[6], "제   목: 코람코라이프로지스리츠 부가가치세 대출 차입, 계좌 개설 및 담보제공 동의 요청의 건")
t = ptext(P[10])
old = "이에 대출약정서 제11조 제7항 제(1)호 (추가차입의 금지) 및 제(2)호(담보제공의 불가)에 따른 귀사의 사전 서면동의를 요청드립니다."
assert old in t
set_text(P[10], t.replace(old, (
    "이에 VAT 대출의 차입, VAT 대출 관련 계좌의 신규 개설, 담보의 제공 및 이에 부수하는 행위 일체에 대하여 "
    "대출약정서 제8조 제1항(계좌의 개설), 제11조 제7항 제(1)호(추가차입의 금지) 및 제(2)호(담보제공의 불가)에 따른 "
    "귀사의 사전 서면동의를 요청드립니다.")))

# ── 담보 ②: 신설 계좌임을 명확히
set_text(P[22], "② 위 7)에 따라 신규 개설하는 부가세환급계좌에 대한 제1순위")

# ── 7)~9) 항목 추가 (6) 담보 ② 다음)
a = add_after(P[23], ITEM, "7) 계좌 개설: VAT 대출 대주에 차주 명의의 부가세환급계좌 등 VAT 대출의 실행·상환에 "
                           "필요한 예금계좌 신규 개설(VAT 대출금 수령 및 부가가치세 환급금 수령·예치 용도)")
a = add_after(a, ITEM, "8) 환급금 관리: 부가가치세 환급금은 수입금계좌가 아닌 부가세환급계좌로 수령하여 "
                       "VAT 대출 원리금 상환에 우선 사용하고, 상환 후 잔액은 지체 없이 운영계좌로 이체")
a = add_after(a, ITEM, "9) 부수 행위: 위 1)~8)의 이행에 필요한 담보계약의 체결, 국세환급금양도요구서의 작성, "
                       "관할 세무서 환급계좌 신고, 근질권 설정 통지 및 승낙 등 일체의 행위와 VAT 대출 상환 후 "
                       "관련 담보 및 계좌의 해지")

# ── 확인 문단 보완
set_text(P[25], ("VAT 대출의 담보는 부가가치세환급청구권과 신규 개설하는 부가세환급계좌에 한정되며, 본건 부동산, "
                 "보험금청구권 및 대출약정서상 프로젝트계좌 등 기존 담보목적물에는 어떠한 담보도 설정하지 않습니다. "
                 "VAT 대출은 부가가치세 환급금으로 전액 상환되며, 상환 즉시 관련 담보를 해지하고 부가세환급계좌는 "
                 "해지할 예정입니다."))

# ── 별첨 동의서
set_text(P[39], "제   목: 부가가치세 대출 차입, 계좌 개설 및 담보제공에 대한 동의")
t = ptext(P[40])
old = "차주가 아래 부가가치세 대출을 차입하고 그 담보를 제공하는 것에 대하여 대출약정서 제11조 제7항 제(1)호 및 제(2)호에 따라 동의합니다."
assert old in t
set_text(P[40], t.replace(old, (
    "차주가 아래 부가가치세 대출을 차입하고, 이를 위한 예금계좌를 신규 개설하며, 그 담보를 제공하는 것 및 이에 "
    "부수하는 행위 일체에 대하여 대출약정서 제8조 제1항, 제11조 제7항 제(1)호 및 제(2)호에 따라 동의합니다.")))
set_text(P[45], "3) 담보: 부가가치세환급청구권에 대한 제1순위 채권양도담보 및 아래 4)의 부가세환급계좌에 대한 "
                "제1순위 예금채권근질권(각 대출금의 120% 한도)")
a = add_after(P[45], P[45], "4) 계좌 개설: 부가세환급계좌 등 VAT 대출의 실행·상환에 필요한 차주 명의 예금계좌의 신규 개설")
a = add_after(a, P[45], "5) 환급금 관리: 부가가치세 환급금의 부가세환급계좌 수령 및 VAT 대출 원리금 상환에의 우선 사용")
a = add_after(a, P[45], "6) 부수 행위: 위 1)~5)의 이행에 필요한 일체의 행위 및 VAT 대출 상환 후 관련 담보·계좌의 해지")
# 동의서 항목 문단 간격 축소(1쪽 유지)
from docx.oxml import OxmlElement
items = [P[43], P[44], P[45]] + [x for x in [P[45].getnext(), P[45].getnext().getnext(), P[45].getnext().getnext().getnext()]]
for p in items:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        ppr.append(sp)
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "60")
# 단서: 계좌 개설 포함, 끝 마침표가 별도 문단(47)으로 떨어진 것 정리
set_text(P[46], "다만, 본 동의는 위에 기재된 차입, 계좌 개설, 담보제공 및 그 부수 행위에 한하며, 대출약정서상 그 밖의 "
                "권리를 포기하거나 조건을 변경하는 것으로 해석되지 아니합니다.")
body.remove(P[47])

# 동의서 1쪽 유지: 제목·단서·일자 간격 축소, '- 아 래 -' 뒤 빈 줄 삭제
start = next(i for i, p in enumerate(body) if ptext(p).startswith("[별첨 1]"))
tail = list(body)[start:]
for p in tail:
    sp = p.find(qn("w:pPr") + "/" + qn("w:spacing")) if p.find(qn("w:pPr")) is not None else None
    if sp is None:
        continue
    txt = ptext(p)
    if txt.startswith("동  의  서"):
        sp.set(qn("w:before"), "120"); sp.set(qn("w:after"), "360")
    elif txt.startswith("다만, 본 동의는") or txt.startswith("2026년"):
        sp.set(qn("w:after"), "360")
for p in tail:
    if p.tag == qn("w:p") and not ptext(p).strip() and ptext(p.getprevious()).startswith("- 아 래"):
        body.remove(p)
        break

doc.save(OUT)
print("saved", OUT)
