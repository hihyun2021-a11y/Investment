# -*- coding: utf-8 -*-
"""이자적립금(예금근질권 대상)·3개월 이자 스케줄·기표일 기관별 수수료·통지처 엑셀 생성.

근거(원본 불변):
- 대출약정서 최종 `약정서/대출 약정서 최종/코람코라이프로지스리츠_담보대출약정서(25730072.v.final)_clean.docx`
  제5조(이자율·이자기간·이자지급일·이자의 계산), 제8조(이자적립금), 별지(대주별 약정금), 부록5(통지처)
- 수수료: 같은 폴더 개별 수수료약정서 10건, `대주별 통지처 및 수수료 내역_261012.xlsx_vf - 복사본.xlsx`
출력: 05_대출업무/20261025_금융_이자적립금_이자스케줄_기표일수수료_v01.xlsx
실행: python3 build_이자적립금_이자스케줄_기표일수수료_261025.py
"""
import datetime as dt
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "20261025_금융_이자적립금_이자스케줄_기표일수수료_v01.xlsx")
if len(sys.argv) > 1:
    OUT = sys.argv[1]

# ------------------------------------------------------------------ 기초 데이터
LENDERS = [  # (Tranche, 대주명(약정서 별지), 약정금, 수수료구분, 요율 또는 None, 고정액 또는 None, 근거 약정서)
    ("A-1", "롯데캐피탈 주식회사", 30_000_000_000),
    ("A-1", "주식회사 우리카드", 10_000_000_000),
    ("A-1", "신용협동조합중앙회", 100_000_000),
    ("A-1", "서로신용협동조합", 4_000_000_000),
    ("A-1", "수지신용협동조합", 2_800_000_000),
    ("A-1", "정선신용협동조합", 2_400_000_000),
    ("A-1", "연희신용협동조합", 1_400_000_000),
    ("A-1", "금모래신용협동조합", 1_800_000_000),
    ("A-1", "남인천신용협동조합", 1_000_000_000),
    ("A-1", "미호신용협동조합", 1_300_000_000),
    ("A-1", "대구대서신용협동조합", 1_300_000_000),
    ("A-1", "한빛신용협동조합", 1_700_000_000),
    ("A-1", "양동행복 신용협동조합", 1_900_000_000),
    ("A-1", "서귀포동부신용협동조합", 1_200_000_000),
    ("A-1", "광주어룡신용협동조합", 2_700_000_000),
    ("A-1", "광주미래신용협동조합", 1_400_000_000),
    ("A-1", "새안양 신용협동조합", 4_000_000_000),
    ("A-1", "남수원 신용협동조합", 300_000_000),
    ("A-1", "북인천신용협동조합", 700_000_000),
    ("A-1", "수협은행", 10_000_000_000),
    ("A-1", "주식회사 에스비아이저축은행", 10_000_000_000),
    ("A-2", "주식회사 우리은행", 70_000_000_000),
    ("B", "키움캐피탈 주식회사", 20_000_000_000),
    ("B", "엠지캐피탈 주식회사", 10_000_000_000),
    ("B", "주식회사 디비저축은행", 8_000_000_000),
]
assert sum(x[2] for x in LENDERS) == 198_000_000_000

SHINHYUP = [n for t, n, a in LENDERS if "신용협동조합" in n and n != "신용협동조합중앙회"]
FEES = []  # (Tranche, 대주명, 구분, 요율, 고정액, 근거)
for t, n, a in LENDERS:
    if n == "주식회사 우리은행":
        FEES.append((t, n, "대출주선수수료", None, 830_000_000, "대출주선수수료약정서(Tr A-2 우리은행)(25906867.v4)"))
        FEES.append((t, n, "대출참여수수료", 0.008, None, "대출참여수수료약정서(Tr A-2 우리은행)(25905690.v3)"))
    elif n == "주식회사 에스비아이저축은행":
        FEES.append((t, n, "대출주선수수료", None, 100_000_000, "대출주선수수료약정서(Tr A-1 에스비아이저축은행)(25905688.v3)"))
    elif n == "롯데캐피탈 주식회사":
        FEES.append((t, n, "대출참여수수료", 0.018, None, "대출참여수수료약정서(Tr A-1 롯데캐피탈)(25905685.v3)"))
    elif n == "주식회사 우리카드":
        FEES.append((t, n, "대출참여수수료", 0.012, None, "대출참여수수료약정서(Tr A-1 우리카드)(25905689.v3)"))
    elif n == "수협은행":
        FEES.append((t, n, "대출참여수수료", 0.010, None, "대출참여수수료약정서(Tr A-1 수협은행)(25905686.v3)"))
    elif "신용협동조합" in n:
        FEES.append((t, n, "대출참여수수료", 0.010, None, "대출참여수수료약정서(Tr A-1 신협대주단)(25905687.v4) 제2조 표"))
    elif n == "키움캐피탈 주식회사":
        FEES.append((t, n, "대출취급수수료", 0.010, None, "대출취급수수료약정서(Tr B 키움캐피탈)(25845759.v3)"))
    elif n == "엠지캐피탈 주식회사":
        FEES.append((t, n, "대출취급수수료", 0.0064, None, "대출취급수수료약정서(Tr B 엠지캐피탈)(25845758.v3)"))
    elif n == "주식회사 디비저축은행":
        FEES.append((t, n, "대출참여수수료", 0.010, None, "대출참여수수료약정서(Tr B 디비저축은행)(25845757.v3)"))
# 기대 합계(약정서 금액 기준) 검증용
FEE_EXPECT = {"A-1": 100_000_000 + 540_000_000 + 120_000_000 + 100_000_000 + 300_000_000,
              "A-2": 830_000_000 + 560_000_000, "B": 200_000_000 + 64_000_000 + 80_000_000}

# 부록5 통지처(최종 대출약정서) — 원문 문자열
NOTICE_RAW = {
 "주식회사 우리은행": "주소: 서울특별시 영등포구 여의대로 108, 34층(여의도동, 파크원) / 전화번호: 02-2002-5921 / 02-2002-4567 / 담당자: 이인환 과장 / 홍진경 과장 / 전자우편: gsk3150@wooribank.com / jkhong@wooribank.com",
 "롯데캐피탈 주식회사": "주소: 서울특별시 강남구 테헤란로 142 ARC PLACE 9층 Corporate본부 Corporate영업3팀 / 전화번호: 02-3451-0168 / 010-9872-2219 / 담당자: 곽경호 책임 / 전자우편: kyungho.kwak@lotte.net/ lottecap_corp3@lotte.net",
 "주식회사 우리카드": "주소: 서울특별시 종로구 종로1길 50(중학동) 케이트윈타워 A동 / 전화번호: 02-6968-3531 / 010-9764-8118 / 팩스번호: 0505-001-8544 / 담당자: 강민구 과장 / 전자우편: kangmg@wooricard.com",
 "신용협동조합중앙회": "주소: 서울특별시 중구 소월로 10 단암빌딩 14층 / 전화번호: 02-590-5849 / 010-2241-1820 / 담당자: 안소진 대리 / 전자우편: sjahn@cu.co.kr/ cuf1803@cu.co.kr",
 "서로신용협동조합": "주소: 인천광역시 서해구 가정로 288(석남동) / 전화번호: 032-574-0166 / 010-7152-5264 / 담당자: 장성진 과장 / 전자우편: SJin87@cu.co.kr",
 "수지신용협동조합": "주소: 경기도 용인시 수지구 문정로 7번길 8 1층 / 전화번호: 031-262-6801 / 010-7143-7369 / 담당자: 권영욱 대리 / 전자우편: suji_youngwook@naver.com",
 "정선신용협동조합": "주소: 강원특별자치도 정선군 정선읍 봉양6길 31 / 전화번호: 033-563-0760 / 010-6679-3464 / 담당자: 장용현 주임 / 전자우편: cu04008@cu.co.kr",
 "연희신용협동조합": "주소: 서울특별시 서대문구 연희로 64-22 / 전화번호: 02-323-2636 / 010-6519-0830 / 010-8467-2620 / 담당자: 양주성 과장 / 박준희 대리 / 전자우편: cu01221@cu.co.kr",
 "금모래신용협동조합": "주소: 서울시 서대문구 가재울미래로 11 (남가좌동) / 전화번호: 02-376-2036 / 010-2527-0576 / 담당자: 정덕호 주임 / 전자우편: jeong0576@cu.co.kr",
 "남인천신용협동조합": "주소: 인천시 부평구 배곶로 64 / 전화번호: 032-437-9563 / 010-2006-7819 / 담당자: 임인희 차장 / 전자우편: hee4977@cu.co.kr",
 "미호신용협동조합": "주소: 충청도 청주시 흥덕구 강내면 태성탑연로 448-10 / 전화번호: 043-232-2652 / 010-7735-5225 / 담당자: 이주희 차장 / 전자우편: cu05053@cu.co.kr",
 "대구대서신용협동조합": "주소: 대구 서구 서대구로305, 1층 / 전화번호: 053-351-9500 / 010-8740-2848 / 담당자: 곽경열 대리 / 전자우편: kklcome@cu.co.kr",
 "한빛신용협동조합": "주소: 대전 대덕구 계족산로81번길 87 / 전화번호: 042-627-4501 / 010-6608-1804 / 담당자: 구진숙 차장 / 전자우편: cjlover1804@cu.co.kr",
 "양동행복 신용협동조합": "주소: 전남광주통합특별시 서구 천변좌로 238(양동) / 전화번호: 062-454-6514 / 010-5694-2003 / 담당자: 안혜진 과장 / 전자우편: 080762002002@cu.co.kr",
 "서귀포동부신용협동조합": "주소: 제주특별자치도 서귀포시 효돈로 136 / 전화번호: 064-767-2820 / 010-7564-0888 / 담당자: 홍영기 대리 / 전자우편: cu11008@cu.co.kr",
 "광주어룡신용협동조합": "주소: 전남광주통합특별시 광산구 송정로 15 / 전화번호: 062-720-8811 / 010-3318-5799 / 담당자: 박태훈 대리 / 전자우편: hoon91@cu.co.kr",
 "광주미래신용협동조합": "주소: 전남광주통합특별시 북구 북문대로 153 (운암동) / 전화번호: 062-515-0070 / 010-3075-0887 / 담당자: 박민성 사원 / 전자우편: pmsok101999@cu.co.kr",
 "새안양 신용협동조합": "주소: 경기도 안양시 만안구 장내로 120번길 22(안양동) / 전화번호: 031-446-1101 / 010-8019-3350 / 담당자: 이필승 대리 / 전자우편: cu03112@cu.co.kr / 09protos@cu.co.kr",
 "남수원 신용협동조합": "주소: 경기도 수원시 팔달구 효원로 177(인계동) 1층 / 전화번호: 031-238-0107 / 010-4487-6879 / 010-9411-0227 / 담당자: 이지운 과장 / 김지선 대리 / 전자우편: krlejw@naver.com / jisunya86@naver.com",
 "북인천신용협동조합": "주소: 인천광역시 부평구 후정로 8, 상가2동 B04호 / 전화번호: 032-516-8101 / 010-9664-0702 / 담당자: 김희승 과장 / 전자우편: kkhs101@nate.com / cu03236@cu.co.kr",
 "수협은행": "주소: 서울특별시 동작구 동작대로 43, 1층 수협은행 사당역지점 / 전화번호: 02-874-8781(내선: 4518) / 010-6828-8205 / 팩스번호: 02-585-7745 / 담당자: 이진용 대리 / 전자우편: rinyong2@suhyup-bank.com",
 "주식회사 에스비아이저축은행": "주소: 서울특별시 중구 을지로5길 26, 9층(수하동, 미래에셋센터원빌딩 동관) / 전화번호: 02-3416-4828 / 010-9158-6842 / 팩스번호: 02-2136-0293 / 담당자: 이정민 매니저 / 전자우편: ljm88@sbisb.co.kr",
 "키움캐피탈 주식회사": "주소: 서울특별시 영등포구 여의대로 66, 17층 // 전화번호: 02-780-4295 / 010-4339-3878 // 팩스번호: 02-780-1517 // 담당자: 이승근 주임 // 전자우편: tmd283@kiwoomcapital.com",
 "엠지캐피탈 주식회사": "주소: 서울특별시 중구 남대문로 55, MG파이낸스센터 12층 // 전화번호: 02-2200-7423 / 010-5241-8897 // 팩스번호: 02-511-3021 // 담당자: 박소현 과장 // 전자우편: sopark@mgcap.co.kr",
 "주식회사 디비저축은행": "주소: 서울특별시 중구 남대문로 113, 다동 // 전화번호: 02-3705-1759 / 010-4440-4730 // 팩스번호: 02-3705-1850 // 담당자: 정성범 선임 // 전자우편: seongbeom4730j@dbgroup.co.kr",
}
# 통지처 엑셀(261012) '신협통지처(최종)' 시트의 원리금·수수료 수취계좌 (통장사본 확인 전 참고값)
SH_ACCT_REF = {
 "신용협동조합중앙회": "농협 081-01-168860 / 신협중앙회", "서로신용협동조합": "신협 131-000-941666 / 서로신용협동조합",
 "수지신용협동조합": "신협 131-014-493393 / 수지신협", "정선신용협동조합": "신협 131-019-058167 / 정선신용협동조합",
 "연희신용협동조합": "신협 131-000-189579 / 연희신협", "금모래신용협동조합": "신협 131-000-001170 / 금모래신협",
 "남인천신용협동조합": "신협 131-000-942952 / 남인천신협", "미호신용협동조합": "신협 131-002-821894 / 미호신협",
 "대구대서신용협동조합": "신협 131-005-005290 / 대구대서신용협동조합", "한빛신용협동조합": "신협 131-003-657209 / 한빛신용협동조합",
 "양동행복 신용협동조합": "신협 131-021-819738 / 양동행복신용협동조합", "서귀포동부신용협동조합": "신협 131-005-410020 / 서귀포동부신협",
 "광주어룡신용협동조합": "신협 131-004-571106 / 광주어룡신협", "광주미래신용협동조합": "신협 131-004-612864 / 광주미래신용협동조합",
 "새안양 신용협동조합": "신협 131-001-041701 / 새안양신용협동조합", "남수원 신용협동조합": "신협 131-001-156258 / 남수원신용협동조합",
 "북인천신용협동조합": "신협 131-002-051950 / 북인천신용협동조합",
}

# 은행 휴무일(공휴일·대체공휴일·근로자의 날·연말 휴무). 정부 발표로 확인·추가 필요
HOLIDAYS = [
 ("2026-12-25", "성탄절"), ("2026-12-31", "은행 연말 휴무"),
 ("2027-01-01", "신정"), ("2027-02-06", "설날 연휴"), ("2027-02-07", "설날"), ("2027-02-08", "설날 연휴"),
 ("2027-02-09", "대체공휴일(설날)"), ("2027-03-01", "삼일절"), ("2027-05-01", "근로자의 날(은행 휴무)"),
 ("2027-05-05", "어린이날"), ("2027-05-13", "부처님오신날"), ("2027-06-06", "현충일"), ("2027-08-15", "광복절"),
 ("2027-08-16", "대체공휴일(광복절)"), ("2027-09-14", "추석 연휴"), ("2027-09-15", "추석"), ("2027-09-16", "추석 연휴"),
 ("2027-10-03", "개천절"), ("2027-10-04", "대체공휴일(개천절)"), ("2027-10-09", "한글날"),
 ("2027-10-11", "대체공휴일(한글날)"), ("2027-12-25", "성탄절"), ("2027-12-27", "대체공휴일(성탄절)"),
 ("2027-12-31", "은행 연말 휴무"),
 ("2028-01-01", "신정"), ("2028-01-25", "설날 연휴"), ("2028-01-26", "설날"), ("2028-01-27", "설날 연휴"),
 ("2028-03-01", "삼일절"), ("2028-04-12", "국회의원 선거일"), ("2028-05-01", "근로자의 날(은행 휴무)"),
 ("2028-05-02", "부처님오신날"), ("2028-05-05", "어린이날"), ("2028-06-06", "현충일"), ("2028-08-15", "광복절"),
 ("2028-10-02", "추석 연휴"), ("2028-10-03", "추석·개천절"), ("2028-10-04", "추석 연휴"),
 ("2028-10-05", "대체공휴일(개천절·추석 겹침)"), ("2028-10-09", "한글날"), ("2028-12-25", "성탄절"),
]

N_PERIODS = 8  # 24개월 / 3개월

# ------------------------------------------------------------------ 서식
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill("solid", fgColor="1F3864")
SUB = PatternFill("solid", fgColor="D9E1F2")
INPUT = PatternFill("solid", fgColor="FFF2CC")
TOTAL = PatternFill("solid", fgColor="E2EFDA")
WHITE_B = Font(bold=True, color="FFFFFF", name="맑은 고딕", size=10)
BOLD = Font(bold=True, name="맑은 고딕", size=10)
BASE = Font(name="맑은 고딕", size=10)
NOTE = Font(name="맑은 고딕", size=9, color="595959")
TITLE = Font(bold=True, name="맑은 고딕", size=13)
WON = '#,##0;[Red]-#,##0'
PCT = '0.00%'
PCT4 = '0.0000%'
DATE = 'yyyy-mm-dd(aaa)'
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")


def style_range(ws, ref, font=BASE, fill=None, fmt=None, align=None, border=True):
    for row in ws[ref]:
        for c in row:
            c.font = font
            if fill:
                c.fill = fill
            if fmt:
                c.number_format = fmt
            if align:
                c.alignment = align
            if border:
                c.border = BOX


def header(ws, row, labels, col=1):
    for i, h in enumerate(labels):
        c = ws.cell(row, col + i, h)
        c.font, c.fill, c.alignment, c.border = WHITE_B, HEAD, CENTER, BOX


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


wb = Workbook()

# ================================================================== 1. 기본가정
ws = wb.active
ws.title = "기본가정"
ws["A1"] = "안양물류센터 담보대출 — 이자·이자적립금·기표일 수수료 계산 기본가정"
ws["A1"].font = TITLE
ws["A2"] = "노란 칸은 입력값입니다. 나머지 시트는 이 값과 '휴일' 시트를 참조해 자동 계산됩니다."
ws["A2"].font = NOTE
header(ws, 4, ["항목", "값", "비고·근거"])
rows = [
    ("대출약정 체결일", dt.date(2026, 10, 27), "가정 (기준 일정 TIMELINE.md)", DATE, True, None),
    ("인출일(기표일)", dt.date(2026, 10, 29), "가정 (기준 일정). 최초 이자기간 시작일·이자율결정일", DATE, True, "DrawDate"),
    ("대출기간(개월)", 24, "대출약정서 제1조 '대출만기일': 인출일로부터 24개월 응당일", "0", True, "TermMonths"),
    ("이자기간(개월)", 3, "대출약정서 제5조 제2항", "0", True, "PeriodMonths"),
    ("대출만기일(영업일 조정)", "=WORKDAY(EDATE(DrawDate,TermMonths)-1,1,Holidays)", "응당일이 영업일이 아니면 직후 영업일", DATE, False, "MaturityDate"),
    ("Tranche A-1 금리(고정)", 0.05, "대출약정서 제5조 제1항: 연 5.00%", PCT, True, "RateA1"),
    ("Tranche A-2 가산금리", 0.0143, "대출약정서 제5조 제1항: 기준금리 + 1.43%", PCT, True, "SpreadA2"),
    ("Tranche A-2 최초 기준금리(CD 91일)", 0.0324, "[추정] 10/12 대주 엑셀 쿠폰 4.67% − 1.43%. 실제는 인출일 직전 3영업일(10/26~10/28) CD수익률 단순평균(소수점 셋째 자리 반올림)으로 교체", PCT, True, "CD1"),
    ("Tranche B 금리(고정)", 0.06, "대출약정서 제5조 제1항: 연 6.00%", PCT, True, "RateB"),
    ("평년 일수", 365, "대출약정서 제5조 제5항", "0", True, "Days365"),
    ("윤년 일수", 366, "같은 항 '다만 윤년인 경우 366일'", "0", True, "Days366"),
    ("윤년 시작일", dt.date(2028, 1, 1), "대출기간 중 윤년(2028년). 윤년에 속한 일수는 366일, 나머지는 365일로 나눠 계산 [대리은행 계산방식 확인 필요]", DATE, True, "LeapStart"),
    ("윤년 종료일", dt.date(2028, 12, 31), "", DATE, True, "LeapEnd"),
]
r = 5
for label, val, note, fmt, is_input, nm in rows:
    ws.cell(r, 1, label)
    ws.cell(r, 2, val)
    ws.cell(r, 3, note)
    style_range(ws, f"A{r}:C{r}", align=LEFT)
    ws.cell(r, 2).number_format = fmt
    ws.cell(r, 2).alignment = RIGHT
    if is_input:
        ws.cell(r, 2).fill = INPUT
    if nm:
        name(wb, nm, f"기본가정!$B${r}")
    r += 1
r += 1
notes = [
    "계산 규칙 (대출약정서 최종본 제5조·제8조)",
    "· 이자기간: 최초는 인출일~3개월 응당일, 이후는 직전 말일~다음 3개월 응당일, 마지막은 대출만기일까지. 말일이 영업일이 아니면 다음 영업일에 종료하되, 다음 기간 응당일은 원래 일정대로 계산",
    "· 이자지급일: 각 이자기간 말일(영업일 조정). 후불",
    "· 이자: 원금 × 연이율 × 일수 ÷ 365(윤년 366). 초일 산입·말일 불산입. 원 미만 절사(대주별·회차별)",
    "· Tranche A-2 기준금리: 각 이자기간 초일(이자율결정일) 직전 3영업일 CD 91일물 단순평균 → '이자기간' 시트 노란 칸에 회차별로 입력",
    "· 이자적립금: 인출일 당일 Tranche별 '3개월분 이자 상당액'을 각 이자유보계좌에 적립(제8조). 본 파일은 최초 이자기간(인출일~첫 이자지급일) 실제 일수 기준으로 산정하고, 연이자×3/12 방식도 참고로 표시 [대리은행 확인 필요]",
    "· 휴일: '휴일' 시트 목록(공휴일·대체공휴일·근로자의 날·연말 은행 휴무). 임시공휴일 지정 시 추가 입력",
    "· 수익 확약이 아닌 계산 도구이며, 실제 청구액은 대리은행(우리은행) 이자계산서를 기준으로 확인",
]
for i, t in enumerate(notes):
    c = ws.cell(r + i, 1, t)
    c.font = BOLD if i == 0 else NOTE
    ws.merge_cells(start_row=r + i, start_column=1, end_row=r + i, end_column=3)
    c.alignment = LEFT
widths(ws, [34, 18, 95])

# ================================================================== 2. 휴일
wh = wb.create_sheet("휴일")
wh["A1"] = "은행 휴무일 목록 (영업일 조정에 사용, 행 추가 가능)"
wh["A1"].font = TITLE
header(wh, 3, ["일자", "명칭", "요일"])
for i, (d, nm) in enumerate(HOLIDAYS):
    rr = 4 + i
    wh.cell(rr, 1, dt.date.fromisoformat(d)).number_format = 'yyyy-mm-dd'
    wh.cell(rr, 2, nm)
    wh.cell(rr, 3, f'=IF(A{rr}="","",TEXT(A{rr},"aaa"))')
last_h = 4 + len(HOLIDAYS) + 20
style_range(wh, f"A4:C{last_h}", align=CENTER)
style_range(wh, f"A4:A{last_h}", fill=INPUT, fmt='yyyy-mm-dd', align=CENTER)
name(wb, "Holidays", f"휴일!$A$4:$A${last_h}")
wh.cell(last_h + 2, 1, "※ 2027~2028년 공휴일(설·추석·부처님오신날 음력 환산, 대체공휴일)은 작성자가 추정한 값입니다. 정부 발표(관보)로 확인하고, 임시공휴일·선거일이 지정되면 빈 행에 추가하세요.").font = NOTE
widths(wh, [16, 30, 8])

# ================================================================== 3. 대주목록
wl = wb.create_sheet("대주목록")
wl["A1"] = "대주별 대출약정금 (대출약정서 최종본 별지)"
wl["A1"].font = TITLE
header(wl, 3, ["순번", "Tranche", "대주명", "대출약정금(원)", "금리 구분", "최초 이자기간 적용금리"])
for i, (t, n, a) in enumerate(LENDERS):
    rr = 4 + i
    wl.cell(rr, 1, i + 1)
    wl.cell(rr, 2, t)
    wl.cell(rr, 3, n)
    wl.cell(rr, 4, a)
    wl.cell(rr, 5, {"A-1": "고정 연 5.00%", "A-2": "변동 CD+1.43%", "B": "고정 연 6.00%"}[t])
    wl.cell(rr, 6, f'=IF(B{rr}="A-1",RateA1,IF(B{rr}="A-2",CD1+SpreadA2,RateB))')
L0, L1 = 4, 3 + len(LENDERS)
style_range(wl, f"A{L0}:F{L1}", align=CENTER)
style_range(wl, f"C{L0}:C{L1}", align=LEFT)
style_range(wl, f"D{L0}:D{L1}", fmt=WON, align=RIGHT, fill=INPUT)
style_range(wl, f"F{L0}:F{L1}", fmt=PCT, align=RIGHT)
tr = L1 + 1
for j, t in enumerate(["A-1", "A-2", "B"]):
    rr = tr + j
    wl.cell(rr, 3, f"Tranche {t} 소계")
    wl.cell(rr, 4, f'=SUMIF($B${L0}:$B${L1},"{t}",$D${L0}:$D${L1})')
wl.cell(tr + 3, 3, "합계")
wl.cell(tr + 3, 4, f"=SUM(D{L0}:D{L1})")
style_range(wl, f"A{tr}:F{tr+3}", font=BOLD, fill=TOTAL)
style_range(wl, f"D{tr}:D{tr+3}", font=BOLD, fill=TOTAL, fmt=WON, align=RIGHT)
widths(wl, [6, 9, 30, 20, 16, 18])
wl.freeze_panes = "A4"

# ================================================================== 4. 이자기간 (스케줄)
wp = wb.create_sheet("이자기간")
wp["A1"] = "이자기간별 이자지급 스케줄 (인출일 기준 3개월 단위, 후불)"
wp["A1"].font = TITLE
wp["A2"] = "노란 칸: Tranche A-2 회차별 기준금리(CD 91일, 이자율결정일 직전 3영업일 평균). 확정되면 덮어쓰세요. 미입력 회차는 직전 회차 값을 그대로 씁니다."
wp["A2"].font = NOTE
cols = ["회차", "이자기간 시작일\n(이자율결정일)", "응당일(명목)", "이자기간 말일\n= 이자지급일", "일수", "윤년 일수",
        "평년 일수", "A-2 기준금리\n(CD 91일)", "A-2 적용금리", "Tranche A-1 이자", "Tranche A-2 이자",
        "Tranche A 소계", "Tranche B 이자", "합계"]
header(wp, 4, cols)
P0 = 5
for k in range(1, N_PERIODS + 1):
    rr = P0 + k - 1
    wp.cell(rr, 1, k)
    wp.cell(rr, 2, "=DrawDate" if k == 1 else f"=D{rr-1}")
    wp.cell(rr, 3, f"=EDATE(DrawDate,PeriodMonths*A{rr})")
    wp.cell(rr, 4, "=MaturityDate" if k == N_PERIODS else f"=WORKDAY(C{rr}-1,1,Holidays)")
    wp.cell(rr, 5, f"=D{rr}-B{rr}")
    wp.cell(rr, 6, f"=MAX(0,MIN(D{rr},LeapEnd+1)-MAX(B{rr},LeapStart))")
    wp.cell(rr, 7, f"=E{rr}-F{rr}")
    wp.cell(rr, 8, "=CD1" if k == 1 else f"=H{rr-1}")
    wp.cell(rr, 9, f"=H{rr}+SpreadA2")
    colL = get_column_letter(4 + k)  # 기관별이자 시트의 회차 열 (E부터)
    wp.cell(rr, 10, f'=SUMIF(기관별이자!$B$5:$B$29,"A-1",기관별이자!{colL}$5:{colL}$29)')
    wp.cell(rr, 11, f'=SUMIF(기관별이자!$B$5:$B$29,"A-2",기관별이자!{colL}$5:{colL}$29)')
    wp.cell(rr, 12, f"=J{rr}+K{rr}")
    wp.cell(rr, 13, f'=SUMIF(기관별이자!$B$5:$B$29,"B",기관별이자!{colL}$5:{colL}$29)')
    wp.cell(rr, 14, f"=L{rr}+M{rr}")
P1 = P0 + N_PERIODS - 1
style_range(wp, f"A{P0}:N{P1}", align=CENTER)
style_range(wp, f"B{P0}:D{P1}", fmt=DATE, align=CENTER)
style_range(wp, f"E{P0}:G{P1}", fmt="0", align=CENTER)
style_range(wp, f"H{P0}:H{P1}", fmt=PCT, align=RIGHT, fill=INPUT)
style_range(wp, f"I{P0}:I{P1}", fmt=PCT, align=RIGHT)
style_range(wp, f"J{P0}:N{P1}", fmt=WON, align=RIGHT)
tt = P1 + 1
wp.cell(tt, 1, "합계")
wp.cell(tt, 5, f"=SUM(E{P0}:E{P1})")
for c in range(10, 15):
    L = get_column_letter(c)
    wp.cell(tt, c, f"=SUM({L}{P0}:{L}{P1})")
style_range(wp, f"A{tt}:N{tt}", font=BOLD, fill=TOTAL, align=CENTER)
style_range(wp, f"J{tt}:N{tt}", font=BOLD, fill=TOTAL, fmt=WON, align=RIGHT)
wp.cell(tt + 2, 1, "※ 이자지급일이 휴일이면 다음 영업일로 미뤄지고, 다음 회차 이자기간은 그 날부터 시작합니다(일수에 반영). 원 미만 절사는 대주별·회차별로 적용한 뒤 합산했습니다.").font = NOTE
wp.cell(tt + 3, 1, "※ 2028-04-29(토) → 5/1 근로자의 날(은행 휴무)·5/2 부처님오신날 → 5/3 지급. 5/1 은행 휴무 여부는 대리은행과 확인 필요.").font = NOTE
widths(wp, [6, 16, 15, 16, 7, 8, 8, 12, 11, 16, 16, 17, 16, 17])
wp.row_dimensions[4].height = 34
wp.freeze_panes = "B5"

# ================================================================== 5. 기관별이자
wi = wb.create_sheet("기관별이자")
wi["A1"] = "대주별 회차별 이자 (원 미만 절사)"
wi["A1"].font = TITLE
wi["A2"] = "각 칸 = ROUNDDOWN(약정금 × 적용금리 × (평년일수/365 + 윤년일수/366), 0)"
wi["A2"].font = NOTE
hdr = ["순번", "Tranche", "대주명", "대출약정금"] + [f"{k}회차" for k in range(1, N_PERIODS + 1)] + ["합계"]
header(wi, 3, hdr)
for k in range(1, N_PERIODS + 1):  # 4행: 지급일 표시
    c = wi.cell(4, 4 + k, f"=이자기간!D{P0+k-1}")
    c.number_format, c.font, c.fill, c.alignment, c.border = DATE, BOLD, SUB, CENTER, BOX
c = wi.cell(4, 4, "이자지급일 →")
c.font, c.fill, c.alignment, c.border = BOLD, SUB, CENTER, BOX
I0 = 5
for i, (t, n, a) in enumerate(LENDERS):
    rr = I0 + i
    src = L0 + i
    wi.cell(rr, 1, i + 1)
    wi.cell(rr, 2, f"=대주목록!B{src}")
    wi.cell(rr, 3, f"=대주목록!C{src}")
    wi.cell(rr, 4, f"=대주목록!D{src}")
    for k in range(1, N_PERIODS + 1):
        pr = P0 + k - 1
        rate = f'IF($B{rr}="A-1",RateA1,IF($B{rr}="A-2",이자기간!$I${pr},RateB))'
        wi.cell(rr, 4 + k, f"=ROUNDDOWN($D{rr}*{rate}*(이자기간!$G${pr}/Days365+이자기간!$F${pr}/Days366),0)")
    wi.cell(rr, 5 + N_PERIODS, f"=SUM(E{rr}:{get_column_letter(4+N_PERIODS)}{rr})")
I1 = I0 + len(LENDERS) - 1
assert I0 == 5 and I1 == 29
style_range(wi, f"A{I0}:{get_column_letter(5+N_PERIODS)}{I1}", align=CENTER)
style_range(wi, f"C{I0}:C{I1}", align=LEFT)
style_range(wi, f"D{I0}:{get_column_letter(5+N_PERIODS)}{I1}", fmt=WON, align=RIGHT)
tr = I1 + 1
for j, t in enumerate(["A-1", "A-2", "B"]):
    rr = tr + j
    wi.cell(rr, 3, f"Tranche {t} 소계")
    for c in range(4, 6 + N_PERIODS):
        L = get_column_letter(c)
        wi.cell(rr, c, f'=SUMIF($B${I0}:$B${I1},"{t}",{L}${I0}:{L}${I1})')
wi.cell(tr + 3, 3, "합계")
for c in range(4, 6 + N_PERIODS):
    L = get_column_letter(c)
    wi.cell(tr + 3, c, f"=SUM({L}{I0}:{L}{I1})")
style_range(wi, f"A{tr}:{get_column_letter(5+N_PERIODS)}{tr+3}", font=BOLD, fill=TOTAL, align=CENTER)
style_range(wi, f"D{tr}:{get_column_letter(5+N_PERIODS)}{tr+3}", font=BOLD, fill=TOTAL, fmt=WON, align=RIGHT)
widths(wi, [6, 9, 28, 17] + [15] * N_PERIODS + [17])
wi.freeze_panes = "E5"

# ================================================================== 6. 이자적립금(예금근질권)
wr = wb.create_sheet("이자적립금")
wr["A1"] = "인출일 이자적립금 — Tranche별 이자유보계좌 (예금채권근질권 대상)"
wr["A1"].font = TITLE
wr["A2"] = "대출약정서 제8조: 인출일 당일 Tranche A·B 대출금의 각 3개월분 이자 상당액을 각 이자유보계좌에 적립. 예금채권근질권설정계약 별지2 근질권대상계좌."
wr["A2"].font = NOTE
header(wr, 4, ["구분", "이자유보계좌(우리은행)", "대출약정금", "적용금리", "산정기간", "일수",
               "이자적립금\n(최초 이자기간 실일수, 적립 기준)", "참고: 연이자 × 3/12", "차이"])
p1 = P0  # 1회차 행
items = [
    ("Tranche A-1", "", "A-1", "=RateA1"),
    ("Tranche A-2", "", "A-2", "=이자기간!I5"),
]
rr = 5
for lab, acct, t, rate in items:
    wr.cell(rr, 1, f"  {lab}")
    wr.cell(rr, 3, f'=SUMIF(대주목록!$B${L0}:$B${L1},"{t}",대주목록!$D${L0}:$D${L1})')
    wr.cell(rr, 4, rate)
    wr.cell(rr, 5, f'=TEXT(이자기간!B{p1},"yyyy-mm-dd")&" ~ "&TEXT(이자기간!D{p1},"yyyy-mm-dd")')
    wr.cell(rr, 6, f"=이자기간!E{p1}")
    wr.cell(rr, 7, f'=SUMIF(기관별이자!$B${I0}:$B${I1},"{t}",기관별이자!$E${I0}:$E${I1})')
    wr.cell(rr, 8, f"=ROUNDDOWN(C{rr}*D{rr}*3/12,0)")
    wr.cell(rr, 9, f"=G{rr}-H{rr}")
    rr += 1
# Tranche A 합계
wr.cell(rr, 1, "Tranche A 이자적립금")
wr.cell(rr, 2, "1005-504-965088")
wr.cell(rr, 3, "=C5+C6")
wr.cell(rr, 7, "=G5+G6")
wr.cell(rr, 8, "=H5+H6")
wr.cell(rr, 9, f"=G{rr}-H{rr}")
ra = rr
rr += 1
wr.cell(rr, 1, "Tranche B 이자적립금")
wr.cell(rr, 2, "1005-504-970089")
wr.cell(rr, 3, f'=SUMIF(대주목록!$B${L0}:$B${L1},"B",대주목록!$D${L0}:$D${L1})')
wr.cell(rr, 4, "=RateB")
wr.cell(rr, 5, f"=E5")
wr.cell(rr, 6, f"=F5")
wr.cell(rr, 7, f'=SUMIF(기관별이자!$B${I0}:$B${I1},"B",기관별이자!$E${I0}:$E${I1})')
wr.cell(rr, 8, f"=ROUNDDOWN(C{rr}*D{rr}*3/12,0)")
wr.cell(rr, 9, f"=G{rr}-H{rr}")
rb = rr
rr += 1
wr.cell(rr, 1, "합계 (인출일 적립 필요액)")
wr.cell(rr, 3, f"=C{ra}+C{rb}")
wr.cell(rr, 7, f"=G{ra}+G{rb}")
wr.cell(rr, 8, f"=H{ra}+H{rb}")
wr.cell(rr, 9, f"=G{rr}-H{rr}")
rt = rr
style_range(wr, f"A5:I{rt}", align=CENTER)
style_range(wr, f"A5:A{rt}", align=LEFT)
style_range(wr, f"C5:C{rt}", fmt=WON, align=RIGHT)
style_range(wr, f"D5:D{rt}", fmt=PCT, align=RIGHT)
style_range(wr, f"G5:I{rt}", fmt=WON, align=RIGHT)
for r_ in (ra, rb):
    style_range(wr, f"A{r_}:I{r_}", font=BOLD, fill=SUB, align=CENTER)
    style_range(wr, f"C{r_}:C{r_}", font=BOLD, fill=SUB, fmt=WON, align=RIGHT)
    style_range(wr, f"D{r_}:D{r_}", font=BOLD, fill=SUB, fmt=PCT, align=RIGHT)
    style_range(wr, f"G{r_}:I{r_}", font=BOLD, fill=SUB, fmt=WON, align=RIGHT)
    wr.cell(r_, 1).alignment = LEFT
style_range(wr, f"A{rt}:I{rt}", font=BOLD, fill=TOTAL, align=CENTER)
style_range(wr, f"C{rt}:C{rt}", font=BOLD, fill=TOTAL, fmt=WON, align=RIGHT)
style_range(wr, f"G{rt}:I{rt}", font=BOLD, fill=TOTAL, fmt=WON, align=RIGHT)
wr.cell(rt, 1).alignment = LEFT
nn = rt + 2
for t in [
    "※ 적립 기준(G열)은 최초 이자기간(인출일~첫 이자지급일)의 실제 일수로 대주별 이자를 계산해 원 미만을 절사한 뒤 합산한 금액입니다. 첫 이자지급일에 실제로 지급할 이자와 같습니다.",
    "※ '3개월분 이자 상당액'의 산정 방식(실일수 / 연이자×3/12)은 약정서에 정해져 있지 않으므로, 적립 전 대리은행(우리은행)과 확인하세요 [확인 필요].",
    "※ Tranche A-2는 인출일 기준금리(10/26~10/28 CD 평균)가 확정되면 '기본가정' 시트에 입력해야 확정됩니다. 현재 값은 [추정]입니다.",
    "※ 이자적립금은 대출금 중에서 적립할 수 있습니다(대출약정서 제2조 자금용도). 예금채권근질권 채권최고액(1순위 1,920억·2순위 456억)은 적립액과 무관합니다.",
]:
    wr.cell(nn, 1, t).font = NOTE
    nn += 1
widths(wr, [26, 18, 18, 10, 26, 7, 22, 20, 14])
wr.row_dimensions[4].height = 34

# ================================================================== 7. 통지처·계좌
wn = wb.create_sheet("통지처_계좌")
wn["A1"] = "대주 통지처 및 수취계좌 (통지처: 대출약정서 최종본 부록5)"
wn["A1"].font = TITLE
wn["A2"] = "노란 칸(수취은행·계좌번호·예금주·통장사본 수령일)은 통장사본을 받은 뒤 입력하세요. 입력하면 '기표일수수료' 시트에 자동 표시됩니다."
wn["A2"].font = NOTE
ncols = ["순번", "Tranche", "대주명", "주소", "담당자", "전화", "휴대폰", "팩스", "전자우편",
         "수취은행", "계좌번호", "예금주", "통장사본 수령일", "비고",
         "참고: 통지처 엑셀(261012) 기재 수취계좌\n(통장사본 확인 전)"]
header(wn, 4, ncols)


def parse_notice(s):
    toks = [t.strip() for t in re.split(r"\s*//\s*|\s+/\s+", s) if t.strip()]
    d = {"주소": [], "전화번호": [], "팩스번호": [], "담당자": [], "전자우편": []}
    key = None
    for t in toks:
        m = re.match(r"(주소|전화번호|팩스번호|담당자|전자우편)\s*:\s*(.*)", t)
        if m:
            key = m.group(1)
            t = m.group(2)
        if key and t:
            d[key].extend([x.strip() for x in re.split(r"/\s*", t) if x.strip()] if key == "전자우편" else [t])
    phones = [p for p in d["전화번호"] if not p.startswith("010")]
    mobiles = [p for p in d["전화번호"] if p.startswith("010")]
    return (", ".join(d["주소"]), ", ".join(d["담당자"]), ", ".join(phones), ", ".join(mobiles),
            ", ".join(d["팩스번호"]), ", ".join(d["전자우편"]))


N0 = 5
# 차주 행
borrower = ("차주", "주식회사 코람코라이프로지스위탁관리부동산투자회사", "서울 강남구 삼성로 511, 4층",
            "안용현 차장, 염진경 과장", "02-787-0049, 02-787-0139", "", "",
            "yonghyeonahn@koramco.com, jk.yeom@koramco.com")
wn.cell(N0, 1, "-")
for j, v in enumerate(borrower):
    wn.cell(N0, 2 + j, v)
for i, (t, n, a) in enumerate(LENDERS):
    rr = N0 + 1 + i
    addr, person, tel, mob, fax, mail = parse_notice(NOTICE_RAW[n])
    vals = [i + 1, t, n, addr, person, tel, mob, fax, mail, None, None, None, None, None, SH_ACCT_REF.get(n, "")]
    for j, v in enumerate(vals):
        wn.cell(rr, 1 + j, v)
N1 = N0 + len(LENDERS)
style_range(wn, f"A{N0}:O{N1}", align=LEFT)
style_range(wn, f"A{N0}:B{N1}", align=CENTER)
style_range(wn, f"J{N0+1}:N{N1}", fill=INPUT, align=LEFT)
style_range(wn, f"M{N0+1}:M{N1}", fill=INPUT, fmt='yyyy-mm-dd', align=CENTER)
style_range(wn, f"A{N0}:O{N0}", fill=SUB, align=LEFT)
wn.cell(N1 + 2, 1, "※ 통지처 원문의 '서울시·충청도' 등 약식 표기는 원문 그대로 두었습니다(미호신협 주소는 충청북도). 수협은행 통지처는 사당역지점입니다.").font = NOTE
wn.cell(N1 + 3, 1, "※ O열은 대주가 통지처 엑셀에 적어 낸 계좌로, 통장사본으로 확인한 값이 아닙니다. 통장사본과 대조한 뒤 J~L열에 입력하세요.").font = NOTE
widths(wn, [6, 8, 30, 40, 20, 24, 24, 14, 34, 12, 20, 22, 14, 18, 34])
wn.row_dimensions[4].height = 34
wn.freeze_panes = "D5"

# ================================================================== 8. 기표일수수료
wf = wb.create_sheet("기표일수수료")
wf["A1"] = "기표일(인출일) 지급 기관별 수수료"
wf["A1"].font = TITLE
wf["A2"] = "근거: 개별 수수료약정서(모두 '대출금의 인출일에 지급', '부가가치세 없음' 기재). 수취계좌는 '통지처_계좌' 시트 입력값을 자동 표시합니다."
wf["A2"].font = NOTE
fcols = ["순번", "Tranche", "대주명", "수수료 구분", "대출약정금", "요율", "수수료 금액(원)", "부가가치세",
         "지급일", "수취은행", "계좌번호", "예금주", "통장사본", "담당자", "전화", "전자우편", "근거 약정서"]
header(wf, 4, fcols)
F0 = 5
for i, (t, n, kind, rate, fixed, src) in enumerate(FEES):
    rr = F0 + i
    lrow = L0 + [x[1] for x in LENDERS].index(n)
    nrow = N0 + 1 + [x[1] for x in LENDERS].index(n)
    wf.cell(rr, 1, i + 1)
    wf.cell(rr, 2, t)
    wf.cell(rr, 3, n)
    wf.cell(rr, 4, kind)
    wf.cell(rr, 5, f"=대주목록!D{lrow}")
    if rate is not None:
        wf.cell(rr, 6, rate)
        wf.cell(rr, 7, f"=ROUND(E{rr}*F{rr},0)")
    else:
        wf.cell(rr, 6, f"=G{rr}/E{rr}")
        wf.cell(rr, 7, fixed)
    wf.cell(rr, 8, "없음")
    wf.cell(rr, 9, "=DrawDate")
    for j, colN in enumerate(["J", "K", "L"]):
        wf.cell(rr, 10 + j, f'=IF(통지처_계좌!{colN}{nrow}="","",통지처_계좌!{colN}{nrow})')
    wf.cell(rr, 13, f'=IF(통지처_계좌!M{nrow}="","미수령","수령")')
    wf.cell(rr, 14, f"=통지처_계좌!E{nrow}")
    wf.cell(rr, 15, f'=통지처_계좌!F{nrow}&IF(통지처_계좌!G{nrow}="","",", "&통지처_계좌!G{nrow})')
    wf.cell(rr, 16, f"=통지처_계좌!I{nrow}")
    wf.cell(rr, 17, src)
F1 = F0 + len(FEES) - 1
style_range(wf, f"A{F0}:Q{F1}", align=CENTER)
style_range(wf, f"C{F0}:C{F1}", align=LEFT)
style_range(wf, f"N{F0}:Q{F1}", align=LEFT)
style_range(wf, f"E{F0}:E{F1}", fmt=WON, align=RIGHT)
style_range(wf, f"F{F0}:F{F1}", fmt='0.00%;;"-"', align=RIGHT)
for i, f in enumerate(FEES):
    if f[3] is None:
        wf.cell(F0 + i, 6).number_format = '0.0000%'
style_range(wf, f"G{F0}:G{F1}", fmt=WON, align=RIGHT, font=BOLD)
style_range(wf, f"I{F0}:I{F1}", fmt=DATE, align=CENTER)
style_range(wf, f"J{F0}:L{F1}", fill=INPUT, align=LEFT)
tr = F1 + 1
for j, t in enumerate(["A-1", "A-2", "B"]):
    rr = tr + j
    wf.cell(rr, 3, f"Tranche {t} 소계")
    wf.cell(rr, 7, f'=SUMIF($B${F0}:$B${F1},"{t}",$G${F0}:$G${F1})')
wf.cell(tr + 3, 3, "선순위(Tranche A) 합계")
wf.cell(tr + 3, 7, f"=G{tr}+G{tr+1}")
wf.cell(tr + 4, 3, "중순위(Tranche B) 합계")
wf.cell(tr + 4, 7, f"=G{tr+2}")
wf.cell(tr + 5, 3, "총 합계")
wf.cell(tr + 5, 7, f"=SUM(G{F0}:G{F1})")
style_range(wf, f"A{tr}:Q{tr+5}", font=BOLD, fill=TOTAL, align=LEFT)
style_range(wf, f"G{tr}:G{tr+5}", font=BOLD, fill=TOTAL, fmt=WON, align=RIGHT)
nn = tr + 7
for t in [
    "※ 우리은행은 주선수수료(8.3억)와 참여수수료(5.6억) 두 건이며, 계좌가 다르면 해당 행 J~L열에 직접 입력하세요(수식을 덮어씀).",
    "※ 주선수수료(우리은행·에스비아이저축은행)는 약정서상 '부가가치세 없음'이나, 금융주선 용역의 과세 여부는 세무 확인이 필요합니다 [확인 필요].",
    "※ 엠지캐피탈 1종 종류주식 인수수수료(30억 × 1.2% = 3,600만원)는 대출 수수료가 아니어서 이 표에 넣지 않았습니다.",
]:
    wf.cell(nn, 1, t).font = NOTE
    nn += 1
widths(wf, [6, 8, 28, 14, 17, 9, 16, 9, 15, 12, 20, 22, 9, 22, 26, 32, 52])
wf.freeze_panes = "D5"

# 시트 순서
wb._sheets = [wb[s] for s in ["기본가정", "이자적립금", "이자기간", "기관별이자", "기표일수수료", "통지처_계좌", "대주목록", "휴일"]]
for s in wb.worksheets:
    s.sheet_view.showGridLines = False
wb.active = 0
wb.save(OUT)
print("saved", OUT)
print("fee expect", FEE_EXPECT, sum(FEE_EXPECT.values()))
