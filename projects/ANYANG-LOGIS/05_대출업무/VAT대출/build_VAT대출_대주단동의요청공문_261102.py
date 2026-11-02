# -*- coding: utf-8 -*-
"""VAT 대출 차입·담보제공에 대한 기존 대주단 동의 요청 공문 v01 (KLL 명의).
틀(원본 불변): 05_대출업무/약정서/20261028_참고_KLI공문서식_리파이낸싱동의공문Draft_261023.docx
  — 기존 대출금 입금요청 공문 v02(build_대출금입금요청공문_v02_KLI서식_261028.py)와 같은 방식으로 문단만 교체
근거:
  · 담보대출약정서 제11조 제7항 제(1)호(추가차입 금지)·제(2)호(담보제공 불가), 제13조(다수대주 결의) — BKL v38 텍스트 기준
  · 부동산매매계약서 제1.5조(부가가치세 대리납부) — BKL v18
  · VAT 대출 조건: 20261102_IM_부가가치세대출_수협은행_v01.pptx (운용역 지시 T/C)
  · 담보 구조: 참고계약서(코크렙제72호 VAT 대출, 2026-04-27) 제6조
출력: 20261102_공문_VAT대출_대주단동의요청_KLL명의_v01.docx
"""
import copy
import os

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(HERE, "../약정서/20261028_참고_KLI공문서식_리파이낸싱동의공문Draft_261023.docx")
OUT = os.path.join(HERE, "20261102_공문_VAT대출_대주단동의요청_KLL명의_v01.docx")
BORROWER = "주식회사 코람코라이프로지스위탁관리부동산투자회사"
FONT = "HY신명조"
LENDERS = ["(주)우리은행", "롯데캐피탈(주)", "(주)우리카드", "신용협동조합중앙회", "서로신용협동조합", "수지신용협동조합",
           "정선신용협동조합", "연희신용협동조합", "금모래신용협동조합", "남인천신용협동조합", "미호신용협동조합",
           "대구대서신용협동조합", "한빛신용협동조합", "양동행복 신용협동조합", "서귀포동부신용협동조합",
           "광주어룡신용협동조합", "광주미래신용협동조합", "새안양 신용협동조합", "남수원 신용협동조합",
           "북인천신용협동조합", "수협은행", "(주)에스비아이저축은행", "키움캐피탈(주)", "엠지캐피탈(주)", "(주)디비저축은행"]
assert len(LENDERS) == 25


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


def left(p):
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        p.insert(0, ppr)
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        ppr.append(jc)
    jc.set(qn("w:val"), "left")


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


def para(text="", size=11, bold=False, align="both", before=0, after=0, page_break=False):
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    if page_break:
        ppr.append(OxmlElement("w:pageBreakBefore"))
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), str(before))
    sp.set(qn("w:after"), str(after))
    sp.set(qn("w:line"), "360")
    sp.set(qn("w:lineRule"), "auto")
    ppr.append(sp)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), align)
    ppr.append(jc)
    p.append(ppr)
    if text:
        p.append(run(text, size, bold))
    return p


doc = Document(TPL)
body = doc.element.body
P = [e for e in body if e.tag == qn("w:p")]
assert ptext(P[0]).startswith("주식회사 코람코라이프인프라")
assert ptext(P[19]).startswith("3. 당사는")

# ── 머리글·문서정보
set_text(P[0], BORROWER)
set_text(P[1], "06158 서울시 강남구 삼성로 511(삼성동) 4층  안용현 차장 TEL : 02-787-0049")
set_text(P[3], ptext(P[3]).replace("KLI 2026-10-[  ]", "KLL 2026-11-[  ]").replace("2026. 10. [  ].", "2026. 11. [  ]."))
set_text(P[4], "수   신: 주식회사 우리은행 등 대주기관(수신처 참조)")
set_text(P[5], "참   조: 여신 담당자")
set_text(P[6], "제   목: 코람코라이프로지스리츠 담보대출 관련 부가가치세 대출 차입 및 담보제공 동의 요청의 건")

# ── 2. 요청 취지
set_text(P[10], ("2. 2026년 10월 27일 체결한 『코람코라이프로지스리츠 담보대출』 대출약정서(이하 “대출약정서”)와 관련하여, "
                 f"당사({BORROWER}, 이하 “차주”)는 2026년 10월 29일 경기도 안양시 동안구 관양동 934 외 2필지 소재 "
                 "안양물류센터(이하 “본건 부동산”)의 매입 거래를 종결하였습니다. 차주는 본건 부동산 매매계약서 제1.5조 및 "
                 "부가가치세법 제52조 제4항에 따라 본건 건물 및 동산 등에 대한 부가가치세를 대리납부하여야 하며, 그 재원 마련을 위하여 "
                 "아래와 같이 단기 부가가치세 대출(이하 “VAT 대출”)을 차입하고자 합니다. 이에 대출약정서 제11조 제7항 제(1)호 "
                 "(추가차입의 금지) 및 제(2)호(담보제공의 불가)에 따른 귀사의 사전 서면동의를 요청드립니다."))

items = [
    "1) VAT 대출 대주: 수협은행 [협의 중]",
    "2) 대출금액: 금 일백이십칠억원(￦12,700,000,000) 한도",
    "3) 자금용도: 본건 건물 및 본건 동산 등에 대한 부가가치세 대리납부(납부기한 2026년 11월 25일)",
    "4) 대출기간: 인출일로부터 3개월(인출 예정일 2026년 11월 25일, 만기 예정일 2027년 2월 25일)",
    "5) 상환재원: 부가가치세 환급금(조기환급 신고, 2026년 12월 [10]일 환급 예정)",
    "6) 담보: ① 부가가치세환급청구권에 대한 제1순위 채권양도담보(담보한도 대출금의 120%) "
    "② VAT 대출 대주에 개설하는 부가세환급계좌에 대한 제1순위 예금채권근질권(채권최고액 대출금의 120%)",
]
src = P[17]
for p, txt in zip(P[13:18], items[:5]):
    set_text(p, txt)
    left(p)
extra = copy.deepcopy(src)
set_text(extra, items[5])
left(extra)
src.addnext(extra)

# ── 3. 기존 대주 권리에 미치는 영향
set_text(P[19], "3. 차주는 VAT 대출과 관련하여 다음 사항을 확인하여 드립니다.")
set_text(P[21], ("① 부가가치세 대리납부는 본건 매매계약에 따라 예정된 차주의 의무이며, 대주 심의 시 제공된 투자설명자료에도 "
                 "VAT 대출이 재원조달 계획으로 반영되어 있었습니다.[IM 판본·페이지 확인 필요]"))
set_text(P[22], ("② VAT 대출의 담보는 부가가치세환급청구권과 부가세환급계좌에 한정되며, 본건 부동산, 보험금청구권 및 대출약정서상 "
                 "각 계좌 등 기존 담보목적물에는 어떠한 담보도 설정하지 않습니다."))
set_text(P[23], ("③ VAT 대출은 부가가치세 환급금으로 전액 상환되며, 상환 즉시 관련 담보를 해지합니다. 환급이 지연되는 경우에도 "
                 "일반환급 기한(2027년 2월 24일)이 VAT 대출 만기 이내에 도래합니다."))

# ── 4. 회신 요청
set_text(P[26], ("4. 위 사항에 동의하시는 경우 별첨 동의서에 날인하시어 2026년 11월 [  ]일까지 회신하여 주시기 바랍니다. "
                 "본 동의는 대출약정서 제13조에 따라 다수대주의 찬성으로 결정되며, 문의사항이 있으시면 당사 담당자"
                 "(안용현 차장 02-787-0049, 염진경 과장 02-787-0139)에게 연락하여 주시기 바랍니다."))
set_text(P[29], "별첨 1. 동의서 양식 1부.  끝.")
for i in (25, 28, 31, 32, 33):  # 빈 줄 일부 축소
    body.remove(P[i])
set_text(P[35], BORROWER)
set_text(P[36], "대표이사 원 재 이")

# ── 수신처
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
recv.append(run("수신처: " + ", ".join(LENDERS), 8))
P[36].addnext(recv)

# ── 별첨: 동의서 양식
anchor = recv
blocks = [
    para("[별첨 1]", 10, align="left", page_break=True),
    para("동  의  서", 18, bold=True, align="center", before=360, after=480),
    para(f"수   신: {BORROWER}", 11, align="left"),
    para("제   목: 부가가치세 대출 차입 및 담보제공에 대한 동의", 11, align="left", after=360),
    para(("당사는 2026년 10월 27일 체결된 『코람코라이프로지스리츠 담보대출』 대출약정서(이하 “대출약정서”)의 대주로서, "
          f"{BORROWER}(이하 “차주”)의 2026년 11월 [  ]일자 공문(문서번호 KLL 2026-11-[  ])에 기재된 바에 따라 "
          "차주가 아래 부가가치세 대출을 차입하고 그 담보를 제공하는 것에 대하여 대출약정서 제11조 제7항 제(1)호 및 "
          "제(2)호에 따라 동의합니다."), 11, after=240),
    para("- 아 래 -", 11, align="center", after=120),
    para("1) 대주: 수협은행 [협의 중]", 11, align="left"),
    para("2) 대출금액: 금 일백이십칠억원(￦12,700,000,000) 한도, 대출기간 인출일로부터 3개월", 11, align="left"),
    para("3) 자금용도: 본건 건물 및 본건 동산 등에 대한 부가가치세 대리납부", 11, align="left"),
    para(("4) 담보: 부가가치세환급청구권에 대한 제1순위 채권양도담보 및 부가세환급계좌에 대한 제1순위 예금채권근질권"
          "(각 대출금의 120% 한도)"), 11, align="left", after=240),
    para("다만, 본 동의는 위에 기재된 차입 및 담보제공에 한하며, 대출약정서상 그 밖의 권리를 포기하거나 조건을 변경하는 것으로 "
         "해석되지 아니합니다.", 11, after=600),
    para("2026년      월      일", 11, align="center", after=600),
    para("대 주 :                                         (인)", 11, align="right", after=120),
    para("대표자(또는 대리인) :                       ", 11, align="right"),
]
for b in blocks:
    anchor.addnext(b)
    anchor = b

doc.core_properties.author = "코람코 상장리츠팀"
doc.save(OUT)
print("saved", OUT)
