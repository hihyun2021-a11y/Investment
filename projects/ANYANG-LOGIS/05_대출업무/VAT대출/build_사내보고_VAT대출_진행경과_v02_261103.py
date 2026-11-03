# -*- coding: utf-8 -*-
"""VAT 대출 진행경과 사내 보고자료 v02 (2장) — KLI 투심 최종본 v5.8 서식(platform/scripts/kli_deck.py) 적용.

- 2026-11-03 운용역 지시: 모든 PPT는 v5.8 서식(폰트·글씨크기·헤드메시지·슬라이드 마스터·색 톤) 기준, 글씨 최소 10pt
- 내용: v01(build_사내보고_VAT대출_진행경과_261103.py)과 동일 — 1p 우선 참여기관 및 T/C, 2p 진행일정(IM 3p 시계열 막대)
- 법령: 부가가치세법 제52조 제4항·시행령 제95조(대리납부 기한), 제59조 제2항·시행령 제107조(조기환급 15일)
실행: python3 build_사내보고_VAT대출_진행경과_v02_261103.py
"""
import datetime as dt
import os
import sys

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "../../../.."))
sys.path.insert(0, os.path.join(ROOT, "platform/scripts"))
from kli_deck import (BAND, BLACK, F_BOLD, F_MEDIUM, GRAY, GREEN, NAVY, ORANGE, RED,  # noqa: E402
                      WHITE, KLIDeck)

OUT = os.path.join(HERE, "20261103_사내보고_VAT대출_진행경과_v02.pptx")
LGRAY = RGBColor(0xBF, 0xBF, 0xBF)
ROWBG = RGBColor(0xF7, 0xF7, 0xF7)

AMT = 127.0
D_DRAW, D_REFUND, D_MAT = dt.date(2026, 11, 25), dt.date(2026, 12, 10), dt.date(2027, 2, 25)
DAYS_MAT = (D_MAT - D_DRAW).days           # 92일
HC_RATE, HC_MIN = 0.065, 45
D_HC_MIN = D_DRAW + dt.timedelta(days=HC_MIN)  # 2027-01-09
hc_min = AMT * HC_RATE * HC_MIN / 365
hc_mat = AMT * HC_RATE * DAYS_MAT / 365

d = KLIDeck(header="VAT 대출 진행경과", footer="안양물류센터 VAT 대출 I 진행경과 보고")

# =====================================================================
# 1p. 우선 참여기관 및 T/C
# =====================================================================
s = d.slide("우선 참여기관 및 T/C",
            head="VAT 대출(127억원)은 우리은행·수협은행·현대커머셜 3개 기관으로 압축하여 11월 중순 승인을 추진 중이며, "
                 "기존 대주 4개 기관은 Drop함")
d.band(s, "기관별 진행 현황 및 주요 조건", top=2.17)
rows = [
    ["구분", "기관", "승인 일정 / 절차", "금리", "만기", "부대조건", "진행 상태"],
    ["우선 ①", "우리은행", "~11/18 승인\n셀다운 기관 승인 시 지점전결", "[TBD]%", "3개월", "-", "협의 중"],
    ["우선 ②", "수협은행", "~11/13 신규 승인 심의", "[TBD]%", "3개월", "지점 여유자금 예치 요건\n(환급금 예치)", "심의 준비"],
    ["우선 ③", "현대커머셜", "신규 승인 필요\n(심의 일정 [TBD])", "[6.5]%", "3개월\n(조기상환 가능)", "최소 45일 잔액 유지", "신규 검토"],
    ["Drop", "SBI저축은행 · 우리카드\n신협중앙회 · 롯데캐피탈", "-", "-", "-", "-", "Drop"],
]
d.table(s, rows, top=2.56, col_widths=[0.95, 2.20, 2.45, 0.90, 1.35, 1.88, 0.90], row_h=0.48,
        align=["c", "c", "c", "c", "c", "c", "c"],
        bold_cells={(1, 1), (2, 1), (3, 1)},
        color_cells={(1, 1): NAVY, (2, 1): NAVY, (3, 1): NAVY, (3, 3): RED, (2, 5): RED, (3, 5): RED,
                     **{(4, c): GRAY for c in range(7)}},
        fills={(0, 0): RGBColor(0xD9, 0xD9, 0xD9), **{(4, c): RGBColor(0xF2, 0xF2, 0xF2) for c in range(7)}})

Y2 = 5.12
d.band(s, "공통 조건", top=Y2, width=5.05)
d.table(s, [["항목", "내용"],
            ["한도 · 용도", "127억원 / 건물분 부가가치세 대리납부"],
            ["담보", "환급청구권 양도담보 · 환급계좌 근질권(각 120%)"],
            ["선행조건", "기존 대주단 사전 서면동의(공문 v02)"],
            ["상환재원", "조기환급금('26.12.10 예정)"]],
        top=Y2 + 0.38, width=5.05, col_widths=[1.25, 3.80], row_h=0.36)
XR = 5.82
d.band(s, "이자비용 비교 (127억원 기준, 억원)", top=Y2, left=XR, width=5.38)
d.table(s, [["기관", "12/10 상환 시", "만기(92일) 유지", "비고"],
            ["우리은행", "127×[r]×15/365", "127×[r]×92/365", "조기상환 조건 [TBD]"],
            ["수협은행", "-", "127×[r]×92/365", "예치이자로 일부 상계"],
            ["현대커머셜", f"{hc_min:.2f}(최소 45일)", f"{hc_mat:.2f}", "'27.01.09까지 유지"]],
        top=Y2 + 0.38, left=XR + 0.06, width=5.32, col_widths=[1.05, 1.42, 1.40, 1.45], row_h=0.45,
        align=["c", "c", "c", "c"], color_cells={(3, 1): RED, (3, 2): RED}, bold_cells={(3, 1), (3, 2)})
d.footnote(s, "* '26.11.03 운용역 협의 현황 기준. 이자비용은 단리·실일수/365 [추정], 수수료·예치이자 미반영, [r]은 확정 전 금리.")

# =====================================================================
# 2p. 진행 일정 (IM 3p 시계열 막대 도식)
# =====================================================================
s = d.slide("진행 일정",
            head="11/13~11/18 기관 승인 후 11/25 VAT 납부(대출 인출)·조기환급 신고, 12/10 조기환급 수령 일정으로 진행함")
d.band(s, "VAT 대출 진행 일정 (시계열)", top=2.17)

D0, D1 = dt.date(2026, 10, 26), dt.date(2027, 3, 5)
LBL_X, LBL_W = 0.56, 3.10
CX0, CX1 = 3.75, 11.15
TOPC, ROWH = 2.50, 0.33


def xpos(day):
    return CX0 + (CX1 - CX0) * (day - D0).days / (D1 - D0).days


tasks = [
    ("거래종결 · 대리납부 의무 발생", dt.date(2026, 10, 29), None, NAVY, "10.29"),
    ("기존 대주단 동의 (다수대주 2/3)", dt.date(2026, 11, 4), dt.date(2026, 11, 20), BAND, "~11.20 [TBD]"),
    ("수협은행 신규 승인 심의", dt.date(2026, 11, 4), dt.date(2026, 11, 13), BAND, "심의 ~11.13"),
    ("우리은행 승인 (지점전결)", dt.date(2026, 11, 4), dt.date(2026, 11, 18), BAND, "승인 ~11.18"),
    ("현대커머셜 신규 승인", dt.date(2026, 11, 4), dt.date(2026, 11, 20), LGRAY, "[TBD]"),
    ("기관 확정 · 약정 · 담보 설정", dt.date(2026, 11, 19), dt.date(2026, 11, 24), ORANGE, "11.19~11.24 [TBD]"),
    ("VAT 납부 (대출 인출 · 대리납부)", D_DRAW, None, RED, "11.25 납부기한"),
    ("조기환급 심사 (15일)", D_DRAW, D_REFUND, ORANGE, "~12.10"),
    ("조기환급일 · 상환 / 예치", D_REFUND, None, GREEN, "12.10 조기환급"),
    ("VAT 대출 기간 (3개월)", D_DRAW, D_MAT, NAVY, "11.25 ~ '27.02.25"),
]
NR = len(tasks)
y_end = TOPC + 0.30 + NR * ROWH
for m in [dt.date(2026, 11, 1), dt.date(2026, 12, 1), dt.date(2027, 1, 1), dt.date(2027, 2, 1), dt.date(2027, 3, 1)]:
    x = xpos(m)
    d.vline(s, x, TOPC + 0.27, y_end)
    d.text(s, x - 0.45, TOPC, 0.9, 0.25, [(f"'{m:%y}.{m.month:02d}", {"font": F_MEDIUM, "color": NAVY})],
           align=PP_ALIGN.CENTER)
for i, (name, a, b_, col, tag) in enumerate(tasks):
    y = TOPC + 0.30 + i * ROWH
    if i % 2 == 0:
        d.rect(s, LBL_X, y, CX1 - LBL_X, ROWH, fill=ROWBG)
    d.text(s, LBL_X + 0.06, y, LBL_W, ROWH, [(name, {"font": F_MEDIUM, "color": BLACK})])
    yc = y + ROWH / 2
    if b_ is None:
        x = xpos(a)
        d.rect(s, x - 0.09, yc - 0.09, 0.18, 0.18, fill=col, shape=MSO_SHAPE.DIAMOND)
        d.text(s, x + 0.13, yc - 0.13, 1.9, 0.26, [(tag, {"font": F_BOLD, "color": col})])
    else:
        x1, x2 = xpos(a), xpos(b_)
        gray = col == LGRAY
        bar = d.rect(s, x1, yc - 0.09, x2 - x1, 0.18, fill=None if gray else col,
                     line=GRAY if gray else None, dash=gray)
        if x2 - x1 > 1.7:
            d.text(s, x1, yc - 0.13, x2 - x1, 0.26, [(tag, {"font": F_BOLD, "color": WHITE})], align=PP_ALIGN.CENTER)
        else:
            d.text(s, x2 + 0.06, yc - 0.13, 2.0, 0.26, [(tag, {"font": F_BOLD, "color": GRAY if gray else col})])
for day, col in ((D_DRAW, RED), (D_HC_MIN, GRAY), (D_MAT, RED)):
    d.vline(s, xpos(day), TOPC + 0.27, y_end, color=col, lw=0.75, dash=True)
d.text(s, xpos(D_HC_MIN) + 0.05, TOPC + 0.30 + 8 * ROWH, 1.9, ROWH,
       [("현대커머셜 45일('27.01.09)", {"color": GRAY})])

YB = y_end + 0.10
d.band(s, "핵심 일정 요약", top=YB)
d.table(s, [["대출 심의", "대출 승인", "VAT 납부", "조기환급", "만기"],
            ["수협 ~11/13 · 우리 ~11/18 · 현대 [TBD]", "11/13~11/18(기관 확정 11/19 [TBD])",
             "'26.11.25(수)", "'26.12.10(목) 예정", "'27.02.25(목)"]],
        top=YB + 0.36, col_widths=[3.05, 2.85, 1.50, 1.78, 1.45], row_h=0.30, first_col=False,
        align=["c"] * 5, color_cells={(1, 2): RED, (1, 3): RED}, bold_cells={(1, 2), (1, 3)})
d.footnote(s, "* 대리납부: 지급일 다음 달 25일까지(부가가치세법 제52조 제4항), 조기환급: 신고기한 후 15일 이내(제59조 제2항). 세무 판단은 외부 확인 필요.")

bad = d.check_min_font()
assert not bad, f"10pt 미만 글자: {bad}"
p = d.save(OUT)
print("생성:", p, len(d.prs.slides._sldIdLst), "장 / 10pt 미만 없음")
