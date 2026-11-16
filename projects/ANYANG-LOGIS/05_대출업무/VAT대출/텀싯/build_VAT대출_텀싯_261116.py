# -*- coding: utf-8 -*-
"""부가가치세 대출 Term Sheet (약정서 v03·근질권 v02·양도담보 v02 기준) Word 정리."""
import pathlib
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = pathlib.Path(__file__).resolve().parent
OUT = HERE / "20261116_텀싯_VAT대출_부가가치세대출TermSheet_v01.docx"
FONT = "맑은 고딕"
NAVY = RGBColor(0x1F, 0x3A, 0x5F)


def font(run, size=10, bold=False, color=None):
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color


def para(doc, text, size=10, bold=False, color=None, align=None, after=4):
    p = doc.add_paragraph()
    font(p.add_run(text), size, bold, color)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(0)
    if align:
        p.alignment = align
    return p


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear")
    s.set(qn("w:color"), "auto")
    s.set(qn("w:fill"), hexcolor)
    tcPr.append(s)


def cell_text(cell, text, bold=False, size=10, color=None):
    cell.text = ""
    lines = text.split("\n")
    p = cell.paragraphs[0]
    for k, ln in enumerate(lines):
        if k:
            p = cell.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        font(p.add_run(ln), size, bold, color)


def table(doc, rows, widths, header=None):
    n = len(rows) + (1 if header else 0)
    t = doc.add_table(rows=n, cols=len(widths))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    r0 = 0
    if header:
        for j, h in enumerate(header):
            c = t.rows[0].cells[j]
            cell_text(c, h, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
            shade(c, "1F3A5F")
        r0 = 1
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = t.rows[r0 + i].cells[j]
            cell_text(c, v, bold=(j == 0))
            if j == 0:
                shade(c, "EAF0F7")
    t.autofit = False
    for j, w in enumerate(widths):
        t.columns[j].width = Cm(w)
        for row in t.rows:
            row.cells[j].width = Cm(w)
    grid = t._tbl.tblGrid
    for j, gc in enumerate(grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(widths[j] * 567)))
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def heading(doc, text):
    p = para(doc, text, 12, True, NAVY, after=3)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.keep_with_next = True


def main():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.font.size = Pt(10)

    para(doc, "코람코라이프로지스리츠 부가가치세 대출", 16, True, NAVY, WD_ALIGN_PARAGRAPH.CENTER, 2)
    para(doc, "Term Sheet (약정서 기준 정리)", 13, True, NAVY, WD_ALIGN_PARAGRAPH.CENTER, 6)
    para(doc, "작성 2026-11-16 · 코람코자산신탁 상장리츠팀", 10, align=WD_ALIGN_PARAGRAPH.RIGHT, after=2)
    para(doc, "기준 문서: 부가가치세 대출약정서(BKL 26376115.v6 → 대주정보 반영 v03), 예금채권근질권설정계약서(26376116.v4 → v02), "
              "채권양도담보권설정계약서(26376117.v4 → v02), 부동산매매계약서(BKL v18). 체결 전 초안 기준이며 [*]는 미확정 항목입니다.",
         10, after=6)

    heading(doc, "1. 거래 개요")
    table(doc, [
        ("차주", "주식회사 코람코라이프로지스위탁관리부동산투자회사 (대표이사 원재이)\n자산관리회사: 코람코자산신탁"),
        ("대주", "주식회사 우리은행 단독 (은행장 정진완, 위 지배인 배동호)\n서울특별시 서초구 방배천로 91, 1층(방배동, 구산타워)"),
        ("대출 목적", "안양물류센터 매입에 따른 건물분 부가가치세 납부 — 부가가치세법 제52조 제4항 양수인 대리납부"),
        ("부가세 근거", "매매계약 별첨 2: 건물 공급가액 1,270억원 × 10% = 부가가치세 127억원 (토지 1,430억원은 면세)\n세금계산서는 매도인(코크렙안양)이 거래종결일(2026-10-29) 발급"),
        ("기존 PF와의 관계", "기존 담보대출(2026-10-27, 1,980억원)과 별개 계약. 기존 PF의 담보·채무는 본 약정상 '허용된 부담·허용된 채무'\n기존 대주단에 차입·계좌 개설·담보 제공 동의 요청(2026-11-03 공문)"),
    ], [3.6, 13.4])

    heading(doc, "2. 대출 조건")
    table(doc, [
        ("대출약정금", "금 127억원 (₩12,700,000,000), 1회 일시 인출, 상환 금액 재차입 불가"),
        ("체결일", "2026년 11월 [*]일"),
        ("인출예정일", "2026-11-25 (인출요청서 양식 기준, 부가세 대리납부 기한과 일치)"),
        ("인출가능기간", "체결일부터 [*]개월 — 미확정"),
        ("대출만기", "2027-02-24 (영업일이 아니면 다음 영업일), 만기 일시상환"),
        ("금리", "CD(91일) + 1.90%, 변동금리"),
        ("기준금리", "이자율결정일 직전 3영업일 CD 수익률 단순평균(금투협 고시)\nCD 산출 중단 시 채권평가 4사 3개월 CD(AAA) 평균으로 대체"),
        ("이자기간·지급", "1개월 단위, 각 이자기간 말일 후취 (최종 이자기간은 만기일까지)"),
        ("이자 계산", "실제 경과일수 / 365 (윤년 366), 한편넣기"),
        ("연체이자", "적용금리 + 연 3.0%p"),
        ("임의조기상환", "1억원 이상·1억원 단위, 3영업일 전 서면 통지, 수수료 없음"),
        ("의무조기상환", "본건 매매 관련 부가세 환급금 수령일부터 3영업일 이내 환급금 전액으로 상환, 수수료 없음"),
        ("수수료", "취급·약정 수수료 조항 없음"),
        ("비용", "법률자문비, 인지세, 담보 설정·해지 비용 등 차주 부담"),
    ], [3.6, 13.4])

    heading(doc, "3. 담보")
    table(doc, [
        ("예금채권 근질권", "부가세환급계좌 예금지급청구권에 대주 1순위 근질권\n채권최고액 152.4억원(대출약정금의 120%)\n근질권자 동의가 있는 경우에만 출금 가능 (계좌개설은행 업무상 서면동의 필요 시 근질권자가 적시 협조)\n※ '인출 금원에 한한 근질권 일부 해지 간주' 문구 없음 (우리은행 요청 반영 확인)"),
        ("부가세 환급청구권 양도담보", "본건 건물분 부가가치세 환급청구권에 대주 1순위 양도담보\n인출일까지 양도담보설정통지서·국세환급금양도요구서(차주·대주 공동 작성)에 확정일자 → 원본을 대주에 교부\n대주는 만기 미상환·기한이익상실 시 관할 세무서에 제출·통지"),
        ("원본 보관", "약정서·담보계약 각 1부 작성, 원본은 대주(근질권자·담보권자) 보관, 차주는 사본 보관"),
        ("간인", "법무법인 또는 당사자가 정한 자의 날인·철인으로 갈음 가능"),
    ], [3.6, 13.4])

    heading(doc, "4. 계좌 및 자금 흐름")
    table(doc, [
        ("부가세대출금입금계좌", "차주 명의 신규 개설(개설은행·계좌번호 [*])\n대출금 입금 → 인출일 당일 건물분 부가세 대리납부 → 잔액은 부가세환급계좌로 이체"),
        ("부가세환급계좌", "차주 명의 신규 개설(개설은행·계좌번호 [*]), 대주 동의 없이 변경 불가\n과세관청 환급금 입금 → 3영업일 이내 대출 상환\n관할 세무서에 환급계좌 신고 필요"),
        ("부가세 신고", "예정·확정신고 전 신고서를 대주에 제출해 확인받고, 신고 후 2영업일 이내 신고서 사본·접수증 제출"),
    ], [3.6, 13.4])

    heading(doc, "5. 인출선행·동시이행·후행조건")
    table(doc, [
        ("인출선행조건\n(인출 1영업일 전)", "인출요청서(별첨 가)\n법인등기부, 법인인감증명서, 정관·사업자등록증 사본(원본대조필)\n약정 체결·담보 제공 승인 이사회의사록 사본(원본대조필)\n위임장(대표이사 외 서명 시), 두 계좌 통장 사본\n부동산매매계약서 사본, 예금근질권·채권양도담보 계약서\n본 약정 적법 체결·선행조건 충족 법무법인 법률의견서"),
        ("인출일 당일\n(동시이행)", "부가세 대리납부 납부서·영수증 사본\n확정일자부 예금근질권설정통지서 사본\n확정일자부 채권양도담보권설정통지서\n확정일자부 국세환급금양도요구서"),
        ("인출 후 1영업일", "본건 매입 부가세액이 기재된 세금계산서 사본"),
        ("인출 후 10영업일", "계좌개설은행의 확정일자부 예금근질권설정 승낙서"),
    ], [3.6, 13.4])

    heading(doc, "6. 준수사항 및 기한이익상실")
    table(doc, [
        ("주요 준수사항", "대출금은 부가세 납부 용도로만 사용\n허용된 채무·부담 외 추가 차입·담보 제공·신용공여 금지(대주 사전 서면동의 시 예외)\n대주 권리에 영향을 주는 계약 변경 금지\n영업양수도·재산 처분·합병·분할·해산 제한(본건 부동산 매각 후 청산은 예외)\n중요 변동·EOD 사유 즉시 서면 통지"),
        ("자동 기한이익상실", "만기 미상환, 지급정지·회생·파산 등, 법령상 이행 위법"),
        ("선언에 의한 상실", "기타 미지급(5영업일 치유), 준수사항 위반(치유 요청 후 30영업일), 후행조건 미충족(10영업일), 담보권 효력 상실, 본건 부동산 멸실 등으로 상환능력 중대 훼손"),
    ], [3.6, 13.4])

    heading(doc, "7. 기타")
    table(doc, [
        ("양도", "대주: 차주에 사전 서면통지로 권리·지위 양도 가능(차주는 약정 체결로 사전 동의)\n차주: 대주의 사전 서면동의 없이 양도 불가"),
        ("적용 약관", "약정에서 정하지 않은 사항은 우리은행 기업용 여신거래기본약관(별첨 나)"),
        ("대주 통지처", "서울특별시 서초구 방배천로 91, 1층(방배동, 구산타워) / 02-523-2171\n이택주 부지점장 / 송준민 계장 / TJ.LEE@wooribank.com, dalppong0219@wooribank.com"),
        ("차주 통지처", "서울특별시 강남구 삼성로 511, 4층 / 02-787-0049, 02-787-0139\n안용현 차장 / 염진경 과장 / yonghyeonahn@koramco.com, jk.yeom@koramco.com"),
    ], [3.6, 13.4])

    heading(doc, "8. 미확정 사항 및 유의점")
    for s in [
        "미확정 공란: 체결일, 인출가능기간, 두 계좌 개설은행·계좌번호, 근질권 계좌개설은행, 양도담보 통지서의 관할 세무서·환급 입금계좌",
        "상환 재원은 부가세 환급금뿐이며 만기(2027-02-24) 연장 조항이 없음 — 환급 예상 시기와 만기 사이 여유를 대주와 사전 확인 필요",
        "국세환급금 양도 요구는 세무서의 환급금 통지서 발급 전에 해야 효력이 있으므로, 실질적인 담보는 근질권이 설정된 부가세환급계좌로 환급금을 받는 구조",
        "세금계산서 발급(공급가액 1,270억원·세액 127억원·작성일자 2026-10-29) 여부를 홈택스로 확인",
        "본 자료는 약정서 초안을 요약한 내부 참고자료입니다. 법률·세무 사항은 내부 검토이며 외부 법무법인·세무대리인 확인이 필요합니다.",
    ]:
        p = doc.add_paragraph(style="List Bullet")
        font(p.add_run(s), 10)
        p.paragraph_format.space_after = Pt(2)

    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
