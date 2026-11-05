# -*- coding: utf-8 -*-
"""안양물류센터(ANYANG-LOGIS) 실제 진행 자료로 예시 딜 폴더 만들기 — 1단계 검증용.

python platform/dealflow/samples/build_sample_anyang.py
출처: projects/ANYANG-LOGIS/PROJECT.md, 07_일정_회의록/TIMELINE.md, 00_원본자료/processed/20260916_렌트롤_IM기준_v01.csv,
      05_대출업무/20261013_금융_대주통지처및수수료_vf2_통합통지처.xlsx, 03_IM_투자심의/IM/…v02.pptx(사진), 투심 추가장표 v02(실사 수수료)
"""
import csv
import datetime as dt
import pathlib
import re
import shutil
import sys

import openpyxl
from pptx import Presentation

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parents[2]
sys.path.insert(0, str(HERE.parent))
from lib import deal as D  # noqa: E402

P = ROOT / "projects/ANYANG-LOGIS"
OUT = HERE / "안양물류센터"
d = dt.date.fromisoformat

ASSET = {
    "code": "ANYANG-LOGIS", "name": "안양물류센터", "sector": "물류센터",
    "addr": "경기도 안양시 동안구 관양동 934 외 2필지", "jibun": "경기도 안양시 동안구 관양동 934 외 2필지",
    "gfa_m2": 95474.59, "floors": "B2/8F", "use": "창고시설(물류센터)", "approval": "2023-11",
    "owner": "코크렙안양㈜ (㈜엘에프 95%·코람코 5%)", "leasable_py": 20890, "seller_advisor": "JLL",
    "price": 2700, "appraisal": 3345, "vehicle": "신설 사모 위탁관리리츠(AMC 코람코자산신탁)", "term": 10,
    "deal_type": "공개입찰(우선협상대상자 선정)",
    "highlight1": "서울·경기남부 30분 커버 Last-mile 도심 신축 상온 물류센터, 임대율 99.3%",
    "highlight2": "쿠팡 77.2% 장기 임차('34.04), WALE 6.8년의 안정적 현금흐름",
    "highlight3": "'29~'30 만기 갱신 시 시장임대료(권역 56~63천원/평) 수준으로 임대료 상승 여지",
    "location": "경기도 안양 도심(Last-mile Zone), 서울 및 경기남부 소비자를 30분 내 커버하는 입지",
    "manager": "안용현", "manager_title": "차장", "team": "리츠투자부문/상장리츠본부/상장리츠팀",
}
SCHED = {"bid_notice": "2026-07-16", "loc_equity": "2026-07-31", "bid_due": "2026-08-05", "interview": "2026-08-10",
         "dd_vendor": "2026-08-13", "preferred": "2026-08-18", "pre_ic": "2026-08-24", "mou": "2026-08-27",
         "dd_start": "2026-08-27", "incorp": "2026-09-11", "dd_draft": "2026-09-11", "license_file": "2026-09-23",
         "ic": "2026-10-14", "license_ok": "2026-10-23", "board": "2026-10-26", "loan_sign": "2026-10-27",
         "drawdown_req": "2026-10-28", "closing": "2026-10-29"}
VENDORS = [  # 투심 추가장표 v02(기준 재무모델 v03 실사/자문비용) — 연락처는 운용역 입력
    ["감정", "미래새한감정평가법인", "", "", "", 85.3, 3, "담보 목적 감정평가", "O", "대주 지정 감정평가법인"],
    ["법률", "법무법인 태평양", "", "", "", 80, 4, "법률실사·매매·주주간·대출 계약 자문", "O", "매매·주주간·대출 계약 일괄 자문 가능"],
    ["재무", "우리회계법인", "", "", "", 40, 4, "재무·세무실사", "O", ""],
    ["시장", "Savills", "", "", "", 40, 3, "수도권 Last-mile 물류 시장조사", "O", "권역 임대·매매 사례 DB 보유"],
    ["물리", "CBRE", "", "", "", 25, 3, "건축·설비·CAPEX 실사", "O", ""],
]
OTHER_WGL = [  # 연락처 미기재(운용역 입력)
    ["매수인(당사)", "코람코자산신탁", "AMC·딜 총괄", "안용현", "차장", "상장리츠팀", "", "", "", ""],
    ["매도인", "코크렙안양㈜", "매도인", "", "", "", "", "", "", "㈜엘에프 95%·코람코 5%"],
    ["매각자문사", "JLL", "매각자문", "", "", "", "", "", "", ""],
    ["Equity 투자자", "우리투자증권", "1종 종류주 잔액인수", "", "", "", "", "", "", "LOC 240억"],
    ["Equity 투자자", "삼성증권", "2종 종류주", "", "", "", "", "", "", "LOC 500억"],
    ["Equity 투자자", "이지스자산운용", "1종·2종", "", "", "", "", "", "", "LOC 각 100억"],
    ["법무법인", "법무법인 태평양", "법률실사·계약", "", "", "", "", "", "", ""],
    ["회계법인", "우리회계법인", "재무실사", "", "", "", "", "", "", ""],
    ["감정평가법인", "미래새한감정평가법인", "감정평가", "", "", "", "", "", "", ""],
    ["시장실사", "Savills", "시장실사", "", "", "", "", "", "", ""],
    ["물리실사", "CBRE", "물리실사", "", "", "", "", "", "", ""],
    ["자산보관·사무수탁", "우리은행", "자산보관(부동산)", "", "", "", "", "", "", ""],
    ["임차인", "쿠팡", "주 임차인(77.2%)", "", "", "", "", "", "", "'29.04 중도해지 선택권"],
]
ASSUME = {"price": 2700, "senior_amt": 1600, "senior_rate": 5.5, "mezz_amt": 380, "mezz_rate": 6.5, "hold": 10,
          "exit_cap": 4.89, "rent_growth": 2.5, "downtime": 3}


def lenders():
    wb = openpyxl.load_workbook(P / "05_대출업무/20261013_금융_대주통지처및수수료_vf2_통합통지처.xlsx", data_only=True)
    ws = wb["통합 통지처"]
    out = []
    for r in ws.iter_rows(min_row=4, values_only=True):
        org, info, tr, amt = r[2], r[6], r[3], r[4]
        if not org or not info:
            continue
        t = str(info)
        mails = re.findall(r"[\w.\-]+@[\w.\-]+", t)
        phones = re.findall(r"0\d{1,2}[.\-]\d{3,4}[.\-]\d{4}", t)
        mob = next((p for p in phones if p.startswith("010")), "")
        tel = next((p for p in phones if not p.startswith("010")), "")
        m = re.search(r"담당자\s*[:/]?\s*([^\n]+)", t)
        who = (m.group(1) if m else "").strip(" :/")
        first = re.split(r"[,/]|\n", who)[0].strip().split()
        name, title = (first[0], " ".join(first[1:])) if first else ("", "")
        out.append(["대주", org, f"{str(tr).replace(chr(10), ' ')} {amt}억원", name, title, "", "; ".join(mails), tel, mob,
                    who if ("," in who or "/" in who) else ""])
    return out


def rentroll():
    rows = []
    with open(P / "00_원본자료/processed/20260916_렌트롤_IM기준_v01.csv", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if r["층"] == "합계" or not r["임차인"]:
                continue
            rows.append([r["층"], r["임차인"], r["용도"], float(r["임대면적_평"]), float(r["보증금_원per평"] or 0),
                         float(r["월임대료_원per평"] or 0), float(r["관리비_원per평"] or 0), None,
                         d(r["임대차만료"]) if r["임대차만료"] else None, float(r["인상률_pct"]) if r["인상률_pct"] else None,
                         (r["비고"] + " " + r["출처"]).strip()])
    return rows


def photos(folder):
    im = next((P / "03_IM_투자심의/IM").glob("*InformationMemorandum_v02.pptx"))
    prs = Presentation(str(im))
    picks = {5: "01_자산전경", 6: "02_입지", 32: "03_도면"}
    for idx, nm in picks.items():
        s = prs.slides[idx - 1]
        pic = next((sh for sh in s.shapes if sh.shape_type == 13), None)
        if pic:
            (folder / f"{nm}.{pic.image.ext}").write_bytes(pic.image.blob)


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    sample = {"자산개요": ASSET, "사업일정": {k: d(v) for k, v in SCHED.items()}, "WGL": OTHER_WGL + lenders(),
              "실사업체": VENDORS, "재무가정": ASSUME, "렌트롤": rentroll()}
    D.scaffold(OUT, sample)
    photos(OUT / D.IN / "사진")
    for src in [P / "02_실사보고서/법률/1. 코크렙안양 토지대장(관양동 934외 2필지)_지목변경완료(창고용지).pdf",
                P / "02_실사보고서/법률/안양물류센터 일부멸실 완료등기부 (1).pdf"]:
        shutil.copy2(src, OUT / D.IN / "대장_등기" / src.name)
    shutil.copy2(P / "04_Equity투자자/투자자별/20260921_메일_LF_3종종류주식_재투자협의_광장의견_v01.eml", OUT / D.IN / "메일")
    (OUT / "settings.txt").unlink(missing_ok=True)
    print("예시 딜 폴더:", OUT)


if __name__ == "__main__":
    main()
