# -*- coding: utf-8 -*-
"""딜플로우 — 신규 딜 1건을 딜 진행 단계별로 자동 처리하는 프로그램 (현재 1단계).

사용법
  python dealflow.py                       메뉴
  python dealflow.py new  <딜폴더>          새 딜 폴더·딜입력.xlsx 만들기
  python dealflow.py run  <딜폴더>          1단계 전체 실행(아래 전부)
  python dealflow.py teaser|rfp|vendor|schedule|wgl|mail|model <딜폴더>   항목별 실행

단계 로드맵은 ROADMAP.md. 산출물은 <딜폴더>/02_산출물/1단계/, 누적 데이터는 <딜폴더>/deal.json.
"""
from __future__ import annotations

import datetime as dt
import os
import pathlib
import sys
import traceback

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from lib import deal as D  # noqa: E402
from lib import finmodel, maillog, rfp, schedule, teaser, vendor, wgl  # noqa: E402
from lib.common import fmt, kdate, outpath  # noqa: E402

STEPS = ["model", "teaser", "rfp", "vendor", "schedule", "wgl", "mail"]
LABEL = {"model": "협의용 재무모델", "teaser": "Teaser", "rfp": "실사 RFP 5종", "vendor": "실사업체 선정 보고·기안",
         "schedule": "사업일정·체크리스트·캘린더", "wgl": "Working Group List·주소록", "mail": "WGL 메일 이슈로그"}


def _has_price(d):
    return bool(d["assume"].get("price") or d["asset"].get("price"))


def step_model(d, out, prior):
    name = d["asset"].get("name", "딜")
    if not _has_price(d):
        return {"skip": "매입가(자산개요 또는 재무가정)가 없어 건너뜀"}
    has_dd = bool(d.get("assume_dd"))
    p = outpath(out / "재무모델", "재무모델", f"{name}_협의용" + ("_실사반영" if has_dd else ""), "xlsx")
    res = finmodel.write_xlsx(d, p, f"{name} 협의용 재무모델", "실사반영 항목: " + ", ".join(d["assume_dd"]) if has_dd else "")
    d["model"] = res
    info = {"xlsx": str(p), "irr": res["irr"], "cap_in": res["cap_in"], "equity": res["equity"]}
    old = (prior or {}).get("model")
    if old and old.get("irr") is not None and res["irr"] is not None:
        rows = [("Equity IRR", old["irr"], res["irr"], "%"), ("Going-in Cap", old["cap_in"], res["cap_in"], "%"),
                ("Equity(억원)", old["equity"], res["equity"], ""), ("1년차 NOI(억원)", old["noi"][0], res["noi"][0], ""),
                ("최저 ICR", old.get("min_icr"), res.get("min_icr"), "x")]
        md = [f"# {name} 재무모델 변동 비교 ({dt.date.today():%Y-%m-%d})", "",
              f"직전 실행 대비. 실사반영 항목: {', '.join(d['assume_dd']) or '없음'}", "", "| 지표 | 직전 | 이번 | 변동 |", "|---|---|---|---|"]
        for lbl, a, b, u in rows:
            if a is None or b is None:
                continue
            if u == "%":
                md.append(f"| {lbl} | {a * 100:.2f}% | {b * 100:.2f}% | {(b - a) * 1e4:+.0f}bp |")
            else:
                md.append(f"| {lbl} | {a:,.2f} | {b:,.2f} | {b - a:+,.2f} |")
        md += ["", "출처 표기: " + "; ".join(f"{k} ← {v}" for k, v in d.get("assume_src", {}).items()) if d.get("assume_src") else ""]
        cp = outpath(out / "재무모델", "검토메모", f"{name}_재무모델_변동비교", "md")
        cp.write_text("\n".join(md) + "\n", encoding="utf-8")
        info["compare"] = str(cp)
    return info


def step_teaser(d, out, prior):
    name = d["asset"].get("name", "딜")
    p = teaser.build(d, d.get("model"), outpath(out / "Teaser", "Teaser", name, "pptx"))
    return {"pptx": str(p)}


def step_rfp(d, out, prior):
    return rfp.build_all(d, d.get("model"), out / "RFP")


def step_vendor(d, out, prior):
    r = vendor.build(d, out / "실사업체선정")
    return r or {"skip": "실사업체 시트에 '선정(O)' 표시가 없어 건너뜀"}


def step_schedule(d, out, prior):
    if not any(m["date"] for m in d["schedule"]):
        return {"skip": "사업일정 시트에 날짜가 없어 건너뜀(거래종결일만 넣어도 전체 일정을 추정)"}
    return schedule.build(d, out / "일정")


def step_wgl(d, out, prior):
    return wgl.build(d, out / "WGL")


def step_mail(d, out, prior):
    folder = pathlib.Path(d["folder"]) / D.IN / "메일"
    log, notes, new = maillog.collect(d, folder, (prior or {}).get("mail_log", []))
    d["mail_log"] = log
    d["notes"] += notes
    if not log:
        return {"skip": "01_입력자료/메일 폴더에 메일이 없어 건너뜀"}
    r = maillog.write(d, log, out / "메일로그")
    r["new"] = new
    return r


FUNCS = {k: globals()[f"step_{k}"] for k in STEPS}


def run(folder, only=None):
    folder = pathlib.Path(folder).resolve()
    prior = D.prior_state(folder)
    d = D.load(folder)
    out = D.out_dir(d)
    steps = only or STEPS
    if "model" not in steps and _has_price(d):
        d["model"] = finmodel.compute(d)
    produced, errors = {}, {}
    print(f"\n■ {d['asset'].get('name', folder.name)} — 딜플로우 1단계")
    for n in d["notes"]:
        print("  · " + n)
    for k in steps:
        try:
            produced[k] = FUNCS[k](d, out, prior)
            r = produced[k]
            print(f"  – {LABEL[k]} — {r['skip']}" if "skip" in r else f"  ✔ {LABEL[k]}")
        except Exception as e:
            errors[k] = f"{type(e).__name__}: {e}"
            print(f"  ✘ {LABEL[k]} 실패: {errors[k]}")
            if os.environ.get("DEALFLOW_DEBUG"):
                traceback.print_exc()
    D.save_state(d, produced)
    rp = report(d, produced, errors, out)
    print(f"\n  결과 요약: {rp}\n  산출 폴더: {out}")
    return d, produced, errors


def report(d, produced, errors, out):
    a, rr, m = d["asset"], d["rr"], d.get("model")
    L = [f"# {a.get('name', '')} 딜플로우 1단계 실행 결과 ({dt.datetime.now():%Y-%m-%d %H:%M})", ""]
    if m:
        L += ["## 협의용 재무모델 핵심 [추정]", "",
              f"- 매입가 {fmt(m.get('price'), 0)}억원 / 총 투자비 {fmt(m['uses'], 1)}억원 / "
              f"Equity {fmt(m['equity'], 1)}억원 / 대출 {fmt(m['loans'], 1)}억원",
              f"- 1년차 NOI {fmt(m['noi'][0], 1)}억원, Going-in Cap {m['cap_in'] * 100:.2f}%, "
              f"Equity IRR {m['irr'] * 100:.2f}% ({m['years']}년), 최저 ICR {fmt(m['min_icr'], 2)}x"
              if m.get("irr") is not None and m.get("cap_in") else "- IRR 계산 불가 — 가정 확인", ""]
    L += ["## 임대 현황(렌트롤)", "",
          f"- {rr['n']}행, 임대면적 {fmt(rr['area_py'], 0)}평, 임대율 {fmt((rr['occupancy'] or 0) * 100, 1)}%, "
          f"WALE {fmt(rr['wale'], 2)}년, 월 임대료 {fmt(rr['rent_month'] / 1e6, 1)}백만원", ""]
    L += ["## 산출물", ""]
    for k, r in produced.items():
        files = []
        for v in r.values():
            if isinstance(v, str) and ("/" in v or "\\" in v):
                files.append(v)
            elif isinstance(v, list):
                files += [x for x in v if isinstance(x, str) and ("/" in x or "\\" in x)]
        L.append(f"- **{LABEL[k]}**" + (f": {r['skip']}" if "skip" in r else f" ({len(files)}개)"))
        for f in files:
            L.append(f"  - `{pathlib.Path(f).relative_to(out) if str(f).startswith(str(out)) else f}`")
    if errors:
        L += ["", "## 실패", ""] + [f"- {LABEL[k]}: {v}" for k, v in errors.items()]
    nx = produced.get("schedule", {}).get("next_2w")
    if nx:
        L += ["", "## 앞으로 2주 할 일", ""] + [f"- {x}" for x in nx]
    miss = [k for k in ("name", "sector", "price", "manager", "manager_email") if not a.get(k)]
    est = [s["name"] for s in d["schedule"] if s["estimated"]]
    L += ["", "## 확인 필요", ""]
    if miss:
        L.append(f"- 자산개요 미입력: {', '.join(miss)}")
    L.append(f"- 일정 {len(est)}개는 [추정](안양 일정 간격 적용) — 확정되면 사업일정 시트에 입력")
    for n in d["notes"]:
        L.append(f"- {n}")
    L += ["", "법률·세무 관련 문구는 내부 검토 초안이며 외부 법무법인·회계법인 확인이 필요합니다. 수익률은 목표·추정치입니다."]
    p = out / "00_실행결과.md"
    p.write_text("\n".join(L) + "\n", encoding="utf-8")
    return p


def new(folder):
    folder = pathlib.Path(folder).resolve()
    book = D.scaffold(folder)
    print(f"새 딜 폴더를 만들었습니다: {folder}\n  1) {book.name} 을 채우고\n  2) 01_입력자료/사진·대장_등기·메일 에 자료를 넣은 뒤\n  3) [1단계 전체 실행]")
    return book


def _ask_folder(default=None):
    deals = HERE / "deals"
    existing = sorted(p for p in deals.glob("*") if (p / "딜입력.xlsx").exists()) if deals.exists() else []
    samples = sorted(p for p in (HERE / "samples").glob("*") if (p / "딜입력.xlsx").exists())
    opts = existing + samples
    for i, p in enumerate(opts, 1):
        print(f"   {i}. {p.name}" + ("  (예시)" if p in samples else ""))
    s = input("  딜 번호 또는 폴더 경로: ").strip().strip('"')
    if s.isdigit() and 1 <= int(s) <= len(opts):
        return opts[int(s) - 1]
    return pathlib.Path(s) if s else default


def menu():
    while True:
        print("\n==== 딜플로우 (1단계: 소싱·실사 착수) ====")
        print(" 1. 새 딜 만들기            2. 1단계 전체 실행")
        print(" 3. Teaser만               4. 실사 RFP만")
        print(" 5. 실사업체 선정 보고·기안  6. 일정·체크리스트·캘린더")
        print(" 7. WGL·주소록              8. 메일 이슈로그")
        print(" 9. 재무모델(실사반영 포함)  0. 끝내기")
        c = input("번호: ").strip()
        try:
            if c == "0":
                return
            if c == "1":
                nm = input("  딜 이름(폴더명, 예: 20261105_OO물류): ").strip()
                if nm:
                    new(HERE / "deals" / nm)
                continue
            m = {"2": None, "3": ["teaser"], "4": ["rfp"], "5": ["vendor"], "6": ["schedule"], "7": ["wgl"], "8": ["mail"], "9": ["model"]}
            if c in m:
                f = _ask_folder()
                if f:
                    run(f, m[c])
                    if os.name == "nt":
                        os.startfile(str(pathlib.Path(f).resolve() / D.OUT / "1단계"))
        except Exception as e:
            print(f"[오류] {type(e).__name__}: {e}")
            traceback.print_exc()


def main(argv):
    if len(argv) < 2:
        return menu()
    cmd = argv[1]
    folder = argv[2] if len(argv) > 2 else None
    if cmd == "new" and folder:
        return new(folder)
    if cmd == "run" and folder:
        return run(folder)
    if cmd in STEPS and folder:
        return run(folder, [cmd])
    print(__doc__)


if __name__ == "__main__":
    main(sys.argv)
