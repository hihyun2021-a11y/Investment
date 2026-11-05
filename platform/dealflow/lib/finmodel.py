# -*- coding: utf-8 -*-
"""협의용 재무모델(1단계) — 렌트롤 기반 연간 현금흐름, Equity IRR, ICR, LTV, Cap Rate.

같은 계산을 ① 파이썬(요약·비교용)과 ② 엑셀 수식(협의 중 직접 수정용)으로 만든다.
- 가정 시트: 협의안 | 실사반영 | 적용(=실사반영이 있으면 실사반영) — 실사 결과를 받으면 딜입력.xlsx '재무가정'의 실사반영 열을 채우고 다시 실행.
- 단순화(3단계 섹터별 모델에서 고도화): 대출은 만기일시·이자만 지급, 임차인별 최초 만기 1회만 시장임대료로 갱신(공실기간 적용),
  고정비 정액, 법인세·감가상각·배당가능이익 제약 미반영, Overfunding 미반영.
"""
from __future__ import annotations

import datetime as dt
import math

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName

from .common import TODAY, num
from .inputbook import ASSUME

FONT = "맑은 고딕"
HEAD = PatternFill("solid", fgColor="D9D9D9")
COL1 = PatternFill("solid", fgColor="F2F2F2")
INP = PatternFill("solid", fgColor="FFF2CC")
DDF = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
PCT = {"acq_tax", "other_cost", "amc_acq", "senior_rate", "senior_fee", "mezz_rate", "mezz_fee", "rent_growth",
       "market_growth", "vacancy", "opex_ratio", "amc_fee", "exit_cap", "sale_cost", "amc_sale"}


def irr(flows, lo=-0.99, hi=1.0):
    def npv(r):
        return sum(f / (1 + r) ** i for i, f in enumerate(flows))
    if npv(lo) * npv(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        if npv(lo) * npv(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


def months_to(start: dt.date, end: dt.date | None) -> int:
    if not end:
        return 999
    return max(0, math.floor((end - start).days / 30.4375 + 0.5))


def resolve(d: dict) -> tuple[dict, dict, dict]:
    """(협의안, 실사반영, 적용) — 단위: % 항목은 소수(0.05)."""
    a = dict(d["assume"])
    price = a.get("price") or num(d["asset"].get("price"))
    a["price"] = price
    if a.get("senior_amt") is None and price:
        a["senior_amt"] = round(price * 0.55, 1)
    if a.get("fixed_cost") is None and price:
        a["fixed_cost"] = round(price * 0.003, 2)
    for k in ("mezz_amt", "capex", "reserve"):
        a[k] = a.get(k) or 0.0
    for k, _l, _u, dflt, _n in ASSUME:
        if a.get(k) is None and dflt is not None:
            a[k] = dflt
    dd = dict(d.get("assume_dd", {}))
    applied = {**a, **dd}
    conv = lambda m: {k: (v / 100 if (k in PCT and v is not None) else v) for k, v in m.items()}
    return conv(a), conv(dd), conv(applied)


def compute(d: dict, A: dict | None = None) -> dict:
    if A is None:
        A = resolve(d)[2]
    S = d.get("closing") or TODAY
    H = int(A.get("hold") or 5)
    rr = d["rentroll"]
    years = H + 1
    rent = [0.0] * years
    cam = [0.0] * years
    mkt, mg, DT = A.get("market_rent") or 0, A.get("market_growth") or 0, A.get("downtime") or 0
    for r in rr:
        esc = (r["esc"] / 100) if r["esc"] is not None else A["rent_growth"]
        E = months_to(S, r["expiry"])
        for t in range(1, years + 1):
            g = (1 + esc) ** (t - 1)
            old_m = min(12, max(0, E - 12 * (t - 1)))
            new_m = min(12, max(0, 12 * t - max(E + DT, 12 * (t - 1))))
            m_rent = mkt * (1 + mg) ** (t - 1) if mkt > 0 else r["rent_py"] * g
            rent[t - 1] += r["area_py"] * (r["rent_py"] * g * old_m + m_rent * new_m) / 1e8
            cam[t - 1] += r["area_py"] * r["cam_py"] * g * (old_m + new_m) / 1e8
    price = A["price"] or 0
    noi, cfe, icr = [], [], []
    interest = A["senior_amt"] * A["senior_rate"] + A["mezz_amt"] * A["mezz_rate"]
    for t in range(years):
        n = rent[t] * (1 - A["vacancy"]) + cam[t] - cam[t] * A["opex_ratio"] - A["fixed_cost"]
        noi.append(n)
        if t < H:
            cfe.append(n - price * A["amc_fee"] - A["capex"] - interest)
            icr.append(n / interest if interest else None)
    deposit = sum(r["area_py"] * r["dep_py"] for r in rr) / 1e8
    loans = A["senior_amt"] + A["mezz_amt"]
    acq = price * (A["acq_tax"] + A["other_cost"] + A["amc_acq"])
    fees = A["senior_amt"] * A["senior_fee"] + A["mezz_amt"] * A["mezz_fee"]
    uses = price + acq + fees + A["reserve"]
    equity = uses - loans - deposit
    cap_in = noi[0] / price if price else None
    exit_cap = A.get("exit_cap") or cap_in
    sale = noi[H] / exit_cap if exit_cap else 0
    exit_net = sale * (1 - A["sale_cost"] - A["amc_sale"]) - loans - deposit + A["reserve"]
    flows = [-equity] + cfe[:]
    flows[-1] += exit_net
    pos = sum(f for f in flows[1:])
    appraisal = num(d["asset"].get("appraisal"))
    return {
        "years": H, "start": S.isoformat(), "price": price, "rent": rent, "cam": cam, "noi": noi, "cf_equity": cfe, "icr": icr,
        "deposit": deposit, "loans": loans, "acq_cost": acq, "loan_fees": fees, "uses": uses, "equity": equity,
        "cap_in": cap_in, "exit_cap": exit_cap, "sale": sale, "exit_net": exit_net, "flows": flows,
        "irr": irr(flows), "multiple": (pos / equity) if equity > 0 else None,
        "cash_yield": (sum(cfe) / len(cfe) / equity) if equity > 0 and cfe else None,
        "ltv_price": loans / price if price else None, "ltv_appr": loans / appraisal if appraisal else None,
        "min_icr": min(x for x in icr if x is not None) if any(icr) else None, "interest": interest,
    }


# ── 엑셀 ───────────────────────────────────────────────────────────────
def _hdr(ws, r, c0, vals):
    for i, v in enumerate(vals):
        cell = ws.cell(r, c0 + i, v)
        cell.fill, cell.border = HEAD, BOX
        cell.font = Font(name=FONT, bold=True, size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _c(ws, r, c, v, fmt=None, bold=False, fill=None):
    cell = ws.cell(r, c, v)
    cell.font = Font(name=FONT, size=10, bold=bold)
    cell.border = BOX
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


def write_xlsx(d: dict, path, title: str, version_note: str = ""):
    base, dd, A = resolve(d)
    res = compute(d, A)
    H, Y = res["years"], res["years"] + 1
    wb = openpyxl.Workbook()
    sm = wb.active
    sm.title = "요약"
    ga = wb.create_sheet("가정")
    rs = wb.create_sheet("렌트롤")
    ri = wb.create_sheet("임대수입")
    cf = wb.create_sheet("CF")
    lg = wb.create_sheet("실사반영이력")

    def name(n, ref):
        wb.defined_names[n] = DefinedName(n, attr_text=ref)

    # 가정
    ga.append(["협의용 재무모델 가정 — 노란 칸(협의안)·주황 칸(실사반영)을 바꾸면 전체가 다시 계산됩니다. 적용 = 실사반영이 있으면 실사반영"])
    ga["A1"].font = Font(name=FONT, bold=True, size=11)
    _hdr(ga, 3, 1, ["키", "항목", "단위", "협의안", "실사반영", "적용", "출처·비고"])
    labels = {k: (l, u, n) for k, l, u, _d, n in ASSUME}
    row = 4
    rows = {}
    for k, *_ in ASSUME:
        l, u, n = labels[k]
        pct = k in PCT
        f = "0.00%" if pct else ("#,##0" if u == "원/평/월" else "#,##0.0")
        _c(ga, row, 1, k)
        _c(ga, row, 2, l, fill=COL1)
        _c(ga, row, 3, u)
        if k == "exit_cap" and d["assume"].get("exit_cap") is None:
            _c(ga, row, 4, "=a_capin", f, fill=INP)
            n = "협의안 = Going-in Cap Rate(요약 시트)"
        else:
            _c(ga, row, 4, base.get(k), f, fill=INP)
        _c(ga, row, 5, dd.get(k), f, fill=DDF)
        _c(ga, row, 6, f'=IF(E{row}="",D{row},E{row})', f, bold=True)
        _c(ga, row, 7, d.get("assume_src", {}).get(k) or n)
        name("a_" + k, f"'가정'!$F${row}")
        rows[k] = row
        row += 1
    S = dt.date.fromisoformat(res["start"])
    _c(ga, row + 1, 2, "운용 개시일(거래종결)", fill=COL1)
    _c(ga, row + 1, 6, S, "yyyy-mm-dd", bold=True)
    name("a_start", f"'가정'!$F${row + 1}")
    for c, w in zip("ABCDEFG", (11, 34, 9, 12, 12, 12, 46)):
        ga.column_dimensions[c].width = w
    ga.column_dimensions["A"].hidden = True

    # 렌트롤
    _hdr(rs, 1, 1, ["층", "임차인", "용도", "임대면적(평)", "보증금(원/평)", "월임대료(원/평)", "관리비(원/평)",
                    "임대차만료", "인상률(적용)", "만기까지(개월)", "보증금(억원)", "비고"])
    for i, r in enumerate(d["rentroll"], 2):
        _c(rs, i, 1, r["floor"])
        _c(rs, i, 2, r["tenant"])
        _c(rs, i, 3, r["use"])
        _c(rs, i, 4, r["area_py"], "#,##0.00", fill=INP)
        _c(rs, i, 5, r["dep_py"], "#,##0", fill=INP)
        _c(rs, i, 6, r["rent_py"], "#,##0", fill=INP)
        _c(rs, i, 7, r["cam_py"], "#,##0", fill=INP)
        _c(rs, i, 8, r["expiry"], "yyyy-mm-dd", fill=INP)
        _c(rs, i, 9, (r["esc"] / 100) if r["esc"] is not None else "=a_rent_growth", "0.00%")
        _c(rs, i, 10, f'=IF(H{i}="",999,MAX(0,ROUND((H{i}-a_start)/30.4375,0)))', "0")
        _c(rs, i, 11, f"=D{i}*E{i}/100000000", "#,##0.00")
        _c(rs, i, 12, r["note"])
    n_rr = len(d["rentroll"])
    last = n_rr + 1
    tr = last + 1
    _c(rs, tr, 2, "합계", bold=True, fill=COL1)
    _c(rs, tr, 4, f"=SUM(D2:D{last})", "#,##0.00", bold=True)
    _c(rs, tr, 11, f"=SUM(K2:K{last})", "#,##0.00", bold=True)
    name("a_deposit", f"'렌트롤'!$K${tr}")
    for c, w in zip("ABCDEFGHIJKL", (6, 16, 10, 12, 12, 12, 11, 12, 10, 11, 11, 26)):
        rs.column_dimensions[c].width = w

    # 임대수입: 임대료(억원) 표 + 관리비 표, 열 = 연차 1..H+1
    ri["A1"] = "임차인별 연간 임대료·관리비 수입(억원) — 최초 만기 후 시장임대료(가정 '만기 후 시장임대료', 빈칸이면 현 임대료 추세)·공실기간 반영"
    ri["A1"].font = Font(name=FONT, bold=True, size=10)
    _hdr(ri, 3, 1, ["임차인", "층"] + [t for t in range(1, Y + 1)])
    r0 = 4
    for j, r in enumerate(d["rentroll"]):
        rr_row = j + 2
        rw = r0 + j
        _c(ri, rw, 1, f"='렌트롤'!B{rr_row}", fill=COL1)
        _c(ri, rw, 2, f"='렌트롤'!A{rr_row}")
        for t in range(1, Y + 1):
            col = L(2 + t)
            tt = f"{col}$3"
            R = f"'렌트롤'!$%s{rr_row}"
            g = f"(1+{R % 'I'})^({tt}-1)"
            old = f"MIN(12,MAX(0,{R % 'J'}-12*({tt}-1)))"
            new = f"MIN(12,MAX(0,12*{tt}-MAX({R % 'J'}+a_downtime,12*({tt}-1))))"
            mk = f"IF(a_market_rent>0,a_market_rent*(1+a_market_growth)^({tt}-1),{R % 'F'}*{g})"
            _c(ri, rw, 2 + t, f"={R % 'D'}*({R % 'F'}*{g}*{old}+{mk}*{new})/100000000", "#,##0.00")
    rent_tot = r0 + n_rr
    _c(ri, rent_tot, 1, "임대료 합계", bold=True, fill=COL1)
    for t in range(1, Y + 1):
        col = L(2 + t)
        _c(ri, rent_tot, 2 + t, f"=SUM({col}{r0}:{col}{rent_tot - 1})", "#,##0.00", bold=True)
    c0 = rent_tot + 2
    _hdr(ri, c0, 1, ["관리비 수입", "층"] + [t for t in range(1, Y + 1)])
    for j, r in enumerate(d["rentroll"]):
        rr_row = j + 2
        rw = c0 + 1 + j
        _c(ri, rw, 1, f"='렌트롤'!B{rr_row}", fill=COL1)
        _c(ri, rw, 2, f"='렌트롤'!A{rr_row}")
        for t in range(1, Y + 1):
            col = L(2 + t)
            tt = f"{col}${c0}"
            R = f"'렌트롤'!$%s{rr_row}"
            old = f"MIN(12,MAX(0,{R % 'J'}-12*({tt}-1)))"
            new = f"MIN(12,MAX(0,12*{tt}-MAX({R % 'J'}+a_downtime,12*({tt}-1))))"
            _c(ri, rw, 2 + t, f"={R % 'D'}*{R % 'G'}*(1+{R % 'I'})^({tt}-1)*({old}+{new})/100000000", "#,##0.00")
    cam_tot = c0 + 1 + n_rr
    _c(ri, cam_tot, 1, "관리비 합계", bold=True, fill=COL1)
    for t in range(1, Y + 1):
        col = L(2 + t)
        _c(ri, cam_tot, 2 + t, f"=SUM({col}{c0 + 1}:{col}{cam_tot - 1})", "#,##0.00", bold=True)
    ri.column_dimensions["A"].width = 16
    for t in range(1, Y + 2):
        ri.column_dimensions[L(2 + t)].width = 10

    # CF (열 B = 0년차, C.. = 1..H+1)
    cf["A1"] = "연간 현금흐름(억원) — Equity 기준"
    cf["A1"].font = Font(name=FONT, bold=True, size=11)
    _hdr(cf, 3, 1, ["구분", 0] + [t for t in range(1, Y + 1)])
    items = [("임대료 수입", "rent"), ("공실·대손", "vac"), ("관리비 수입", "cam"), ("운영비", "opex"),
             ("보유세·보험 등", "fixed"), ("NOI", "noi"), ("AMC 운용보수", "amc"), ("CAPEX", "capex"),
             ("이자비용", "int"), ("Equity 현금흐름", "cfe"), ("ICR(NOI/이자)", "icr")]
    R = {k: 4 + i for i, (_n, k) in enumerate(items)}
    for lbl, k in items:
        _c(cf, R[k], 1, lbl, bold=k in ("noi", "cfe"), fill=COL1)
    for t in range(1, Y + 1):
        col = L(2 + t)
        sc = L(2 + t)
        _c(cf, R["rent"], 2 + t, f"='임대수입'!{sc}{rent_tot}", "#,##0.00")
        _c(cf, R["vac"], 2 + t, f"=-{col}{R['rent']}*a_vacancy", "#,##0.00")
        _c(cf, R["cam"], 2 + t, f"='임대수입'!{sc}{cam_tot}", "#,##0.00")
        _c(cf, R["opex"], 2 + t, f"=-{col}{R['cam']}*a_opex_ratio", "#,##0.00")
        _c(cf, R["fixed"], 2 + t, "=-a_fixed_cost", "#,##0.00")
        _c(cf, R["noi"], 2 + t, f"=SUM({col}{R['rent']}:{col}{R['fixed']})", "#,##0.00", bold=True)
        if t <= H:
            _c(cf, R["amc"], 2 + t, "=-a_price*a_amc_fee", "#,##0.00")
            _c(cf, R["capex"], 2 + t, "=-a_capex", "#,##0.00")
            _c(cf, R["int"], 2 + t, "=-(a_senior_amt*a_senior_rate+a_mezz_amt*a_mezz_rate)", "#,##0.00")
            _c(cf, R["cfe"], 2 + t, f"={col}{R['noi']}+{col}{R['amc']}+{col}{R['capex']}+{col}{R['int']}", "#,##0.00", bold=True)
            _c(cf, R["icr"], 2 + t, f"=IF({col}{R['int']}=0,\"-\",{col}{R['noi']}/-{col}{R['int']})", "0.00\"x\"")
    e = R["icr"] + 2
    su = [("매입가", "=a_price"), ("취득세·부대비용·매입보수", "=a_price*(a_acq_tax+a_other_cost+a_amc_acq)"),
          ("대출 취급수수료", "=a_senior_amt*a_senior_fee+a_mezz_amt*a_mezz_fee"), ("운영 예비비", "=a_reserve")]
    _hdr(cf, e, 1, ["재원·투자비(0년차)", "억원"])
    for i, (lbl, f) in enumerate(su, 1):
        _c(cf, e + i, 1, lbl, fill=COL1)
        _c(cf, e + i, 2, f, "#,##0.0")
    uses_r = e + len(su) + 1
    _c(cf, uses_r, 1, "총 투자비", bold=True, fill=COL1)
    _c(cf, uses_r, 2, f"=SUM(B{e + 1}:B{uses_r - 1})", "#,##0.0", bold=True)
    srcs = [("선순위 대출", "=a_senior_amt"), ("중순위 대출", "=a_mezz_amt"), ("임대보증금 승계", "=a_deposit")]
    for i, (lbl, f) in enumerate(srcs, 1):
        _c(cf, uses_r + i, 1, lbl, fill=COL1)
        _c(cf, uses_r + i, 2, f, "#,##0.0")
    eq_r = uses_r + len(srcs) + 1
    _c(cf, eq_r, 1, "Equity(총 투자비 − 대출 − 보증금)", bold=True, fill=COL1)
    _c(cf, eq_r, 2, f"=B{uses_r}-SUM(B{uses_r + 1}:B{eq_r - 1})", "#,##0.0", bold=True)
    name("a_equity", f"'CF'!$B${eq_r}")
    ex = eq_r + 2
    hc, fc = L(2 + H), L(2 + H + 1)
    exits = [("매각가(H+1년 NOI ÷ Exit Cap)", f"={fc}{R['noi']}/a_exit_cap"),
             ("매각비용·매각보수", f"=-B{ex + 1}*(a_sale_cost+a_amc_sale)"),
             ("대출 상환", "=-(a_senior_amt+a_mezz_amt)"), ("보증금 승계(매수인)", "=-a_deposit"),
             ("예비비 회수", "=a_reserve")]
    _hdr(cf, ex, 1, [f"매각({H}년차 말)", "억원"])
    for i, (lbl, f) in enumerate(exits, 1):
        _c(cf, ex + i, 1, lbl, fill=COL1)
        _c(cf, ex + i, 2, f, "#,##0.0")
    exn = ex + len(exits) + 1
    _c(cf, exn, 1, "매각 순현금", bold=True, fill=COL1)
    _c(cf, exn, 2, f"=SUM(B{ex + 1}:B{exn - 1})", "#,##0.0", bold=True)
    fl = exn + 2
    _c(cf, fl, 1, "Equity 순현금흐름", bold=True, fill=COL1)
    _c(cf, fl, 2, f"=-B{eq_r}", "#,##0.0", bold=True)
    for t in range(1, H + 1):
        col = L(2 + t)
        extra = f"+B{exn}" if t == H else ""
        _c(cf, fl, 2 + t, f"={col}{R['cfe']}{extra}", "#,##0.0", bold=True)
    cf.column_dimensions["A"].width = 30
    for t in range(0, Y + 2):
        cf.column_dimensions[L(2 + t)].width = 11
    name("a_capin", f"'CF'!$C${R['noi']}/a_price")
    name("a_flows", f"'CF'!$B${fl}:${hc}${fl}")

    # 요약
    sm["A1"] = title
    sm["A1"].font = Font(name=FONT, bold=True, size=13)
    sm["A2"] = (f"협의용 재무모델 — 작성 {TODAY:%Y-%m-%d}. 모든 수치는 [추정]이며 실사 결과에 따라 바뀝니다. {version_note}").strip()
    sm["A2"].font = Font(name=FONT, size=10, color="646464")
    _hdr(sm, 4, 1, ["지표", "값", "산식·비고"])
    kp = [("매입가(억원)", "=a_price", "#,##0.0", "가정"),
          ("총 투자비(억원)", f"='CF'!B{uses_r}", "#,##0.0", "매입가+취득세·부대비용+수수료+예비비"),
          ("대출(억원)", "=a_senior_amt+a_mezz_amt", "#,##0.0", "선순위+중순위"),
          ("Equity(억원)", "=a_equity", "#,##0.0", "총 투자비 − 대출 − 보증금"),
          ("1년차 NOI(억원)", f"='CF'!C{R['noi']}", "#,##0.00", "임대료·관리비 − 공실·운영비·고정비"),
          ("Going-in Cap Rate", "=a_capin", "0.00%", "1년차 NOI ÷ 매입가"),
          ("Exit Cap Rate", "=a_exit_cap", "0.00%", "가정"),
          ("LTV(매입가)", "=(a_senior_amt+a_mezz_amt)/a_price", "0.0%", ""),
          ("최저 ICR", f"=MIN('CF'!C{R['icr']}:{hc}{R['icr']})", "0.00\"x\"", "NOI ÷ 이자"),
          ("평균 현금배당수익률", f"=AVERAGE('CF'!C{R['cfe']}:{hc}{R['cfe']})/a_equity", "0.00%", "Equity 현금흐름 평균 ÷ Equity(세전·회계 배당 제약 미반영)"),
          ("Equity IRR", "=IRR(a_flows)", "0.00%", f"{H}년 보유 후 매각"),
          ("Equity Multiple", f"=SUM('CF'!C{fl}:{hc}{fl})/a_equity", "0.00\"x\"", "")]
    for i, (lbl, f, fm, note) in enumerate(kp, 5):
        _c(sm, i, 1, lbl, fill=COL1)
        _c(sm, i, 2, f, fm, bold=True)
        _c(sm, i, 3, note)
    sm.column_dimensions["A"].width = 24
    sm.column_dimensions["B"].width = 14
    sm.column_dimensions["C"].width = 60
    k = 5 + len(kp) + 1
    sm.cell(k, 1, "파이썬 검산값(프로그램 계산, 엑셀 수식과 같아야 함)").font = Font(name=FONT, size=9, color="646464")
    sm.cell(k + 1, 1, f"IRR {res['irr'] * 100:.2f}% / Cap {res['cap_in'] * 100:.2f}% / Equity {res['equity']:,.1f}억원"
            if res["irr"] is not None else "IRR 계산 불가(현금흐름 부호 확인)").font = Font(name=FONT, size=9, color="646464")

    # 실사반영이력
    _hdr(lg, 1, 1, ["항목", "협의안", "실사반영", "출처", "반영일"])
    i = 2
    for kk, *_ in ASSUME:
        if kk in dd:
            _c(lg, i, 1, labels[kk][0])
            _c(lg, i, 2, base.get(kk))
            _c(lg, i, 3, dd[kk])
            _c(lg, i, 4, d.get("assume_src", {}).get(kk, ""))
            _c(lg, i, 5, TODAY, "yyyy-mm-dd")
            i += 1
    if i == 2:
        lg.cell(2, 1, "아직 실사 결과 반영 없음(딜입력.xlsx 재무가정 '실사반영' 열을 채우고 다시 실행)")
    for c, w in zip("ABCDE", (34, 12, 12, 40, 12)):
        lg.column_dimensions[c].width = w
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return res
