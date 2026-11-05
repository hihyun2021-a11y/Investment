# -*- coding: utf-8 -*-
"""딜 폴더 → 통합 데이터(deal.json). 다음 단계(2~13)는 이 deal.json을 누적해서 쓴다.

딜 폴더 구조
  <딜폴더>/딜입력.xlsx
  <딜폴더>/settings.txt                (선택: USE_CLAUDE, ANTHROPIC_API_KEY — 깃·메일로 보내지 않음)
  <딜폴더>/01_입력자료/사진/            대표 사진은 파일명 순 첫 장(예: 01_전경.jpg)
  <딜폴더>/01_입력자료/대장_등기/       토지대장·건축물대장·등기부 PDF
  <딜폴더>/01_입력자료/메일/            WGL 메일(.eml/.msg) — Outlook에서 끌어다 놓기
  <딜폴더>/01_입력자료/렌트롤*.xlsx|csv (선택: 딜입력 렌트롤 시트 대신)
  <딜폴더>/02_산출물/1단계/…            산출물
"""
from __future__ import annotations

import datetime as dt
import json
import pathlib

from . import docparse, inputbook
from .common import TODAY, load_json, num, prev_bizday, to_date

IN = "01_입력자료"
OUT = "02_산출물"


def settings(folder: pathlib.Path) -> dict:
    s = {}
    p = folder / "settings.txt"
    if p.exists():
        for ln in p.read_text(encoding="utf-8-sig").splitlines():
            if "=" in ln and not ln.strip().startswith("#"):
                k, v = ln.split("=", 1)
                s[k.strip()] = v.strip()
    return s


def scaffold(folder: pathlib.Path, sample: dict | None = None):
    for sub in ("사진", "대장_등기", "메일"):
        (folder / IN / sub).mkdir(parents=True, exist_ok=True)
    (folder / OUT).mkdir(parents=True, exist_ok=True)
    book = folder / "딜입력.xlsx"
    if not book.exists():
        inputbook.create(book, sample)
    st = folder / "settings.txt"
    if not st.exists():
        st.write_text("# 딜플로우 개인 설정 — 이 파일은 공유하지 마세요\n"
                      "# 스캔본 대장·등기부를 Claude로 읽으려면 1 (자료가 Anthropic API로 전송됨)\n"
                      "USE_CLAUDE=0\nANTHROPIC_API_KEY=\n", encoding="utf-8")
    return book


# ── 일정 ───────────────────────────────────────────────────────────────
def build_schedule(given: dict) -> list[dict]:
    ms = load_json("milestones.json")["milestones"]
    off = {m["key"]: m["offset"] for m in ms}
    closing = given.get("closing")
    if not closing:
        cands = [given[k] - dt.timedelta(days=off[k]) for k in given if k in off]
        closing = sorted(cands)[len(cands) // 2] if cands else None
    out = []
    for m in sorted(ms, key=lambda x: x["offset"]):
        if m["key"] in given:
            d, est = given[m["key"]], False
        elif closing:
            d, est = prev_bizday(closing + dt.timedelta(days=m["offset"])), True
        else:
            d, est = None, True
        out.append({"key": m["key"], "name": m["name"], "group": m["group"], "date": d, "estimated": est,
                    "checklist": m["checklist"]})
    return out


# ── 렌트롤 요약 ────────────────────────────────────────────────────────
def rr_summary(rr: list[dict], leasable_py: float | None, asof: dt.date) -> dict:
    area = sum(r["area_py"] for r in rr)
    rent_m = sum(r["area_py"] * r["rent_py"] for r in rr)
    cam_m = sum(r["area_py"] * r["cam_py"] for r in rr)
    dep = sum(r["area_py"] * r["dep_py"] for r in rr)
    wale_num = 0.0
    for r in rr:
        if r["expiry"]:
            yrs = max(0.0, (r["expiry"] - asof).days / 365.25)
            wale_num += yrs * r["area_py"] * r["rent_py"]
    by_t = {}
    for r in rr:
        t = by_t.setdefault(r["tenant"], {"tenant": r["tenant"], "area_py": 0.0, "rent_m": 0.0, "expiry": r["expiry"], "uses": set()})
        t["area_py"] += r["area_py"]
        t["rent_m"] += r["area_py"] * r["rent_py"]
        t["uses"].add(r["use"])
        if r["expiry"] and (not t["expiry"] or r["expiry"] > t["expiry"]):
            t["expiry"] = r["expiry"]
    tenants = sorted(by_t.values(), key=lambda x: -x["area_py"])
    for t in tenants:
        t["uses"] = "·".join(sorted(u for u in t["uses"] if u))
        t["share"] = t["area_py"] / area if area else 0
    expiry = {}
    for r in rr:
        if r["expiry"]:
            expiry[r["expiry"].year] = expiry.get(r["expiry"].year, 0) + r["area_py"]
    leasable = leasable_py or area
    vacant = [t for t in tenants if t["tenant"] in ("공실", "Vacant", "vacant")]
    leased = area - sum(v["area_py"] for v in vacant)
    return {"n": len(rr), "area_py": area, "leasable_py": leasable, "leased_py": leased,
            "occupancy": (leased / leasable) if leasable else None,
            "rent_month": rent_m, "cam_month": cam_m, "deposit": dep,
            "avg_rent_py": rent_m / leased if leased else None,
            "wale": (wale_num / rent_m) if rent_m else None, "tenants": tenants,
            "expiry_by_year": dict(sorted(expiry.items()))}


# ── 통합 ───────────────────────────────────────────────────────────────
def find_rentroll_file(folder: pathlib.Path):
    for pat in ("렌트롤*.xlsx", "렌트롤*.csv", "*rentroll*.xlsx", "*rent_roll*.csv"):
        hits = sorted((folder / IN).glob(pat))
        if hits:
            return hits[-1]
    return None


def load(folder: pathlib.Path, extract_docs=True) -> dict:
    folder = pathlib.Path(folder).resolve()
    book = folder / "딜입력.xlsx"
    if not book.exists():
        raise FileNotFoundError(f"{book} 이 없습니다. 먼저 [새 딜 만들기]를 실행하세요.")
    st = settings(folder)
    raw = inputbook.load(book)
    notes = []
    auto = {}
    if extract_docs:
        auto, notes = docparse.extract(folder / IN / "대장_등기", st)
        if auto:
            try:
                inputbook.write_auto(book, auto)
            except PermissionError:
                notes.append("딜입력.xlsx 가 열려 있어 '자동추출' 열은 기록하지 못했습니다(산출물에는 반영)")
    asset = {k: v for k, (v, _s) in auto.items()}
    asset.update({k: v for k, v in raw["asset_auto"].items() if k not in asset})
    asset.update(raw["asset"])                      # 입력값 우선
    asset.setdefault("company", "코람코자산신탁")
    asset.setdefault("team", "리츠투자부문/상장리츠본부/상장리츠팀")
    src = {k: ("입력" if k in raw["asset"] else (auto[k][1] if k in auto else "자동추출(이전 실행)")) for k in asset}

    rr = raw["rentroll"]
    rrf = find_rentroll_file(folder)
    if not rr and rrf:
        rr = inputbook.normalize_rr(inputbook.load_rentroll_file(rrf))
        notes.append(f"렌트롤: {rrf.name} 에서 {len(rr)}행")
    sched = build_schedule(raw["schedule_in"])
    closing = next((s["date"] for s in sched if s["key"] == "closing"), None)
    rs = rr_summary(rr, num(asset.get("leasable_py")), closing or TODAY)
    photos = sorted(p for p in (folder / IN / "사진").glob("*") if p.suffix.lower() in (".jpg", ".jpeg", ".png"))
    d = {"folder": str(folder), "asset": asset, "asset_src": src, "schedule": sched, "closing": closing,
         "wgl": raw["wgl"], "vendors": raw["vendors"], "assume": raw["assume"], "assume_dd": raw["assume_dd"],
         "assume_src": raw["assume_src"], "rentroll": rr, "rr": rs, "photos": [str(p) for p in photos],
         "settings": st, "notes": notes}
    return d


def out_dir(d: dict, *parts) -> pathlib.Path:
    p = pathlib.Path(d["folder"]) / OUT / "1단계"
    for x in parts:
        p = p / x
    p.mkdir(parents=True, exist_ok=True)
    return p


def _ser(o):
    if isinstance(o, (dt.date, dt.datetime)):
        return o.isoformat()
    if isinstance(o, set):
        return sorted(o)
    if isinstance(o, pathlib.Path):
        return str(o)
    raise TypeError(type(o))


def save_state(d: dict, produced: dict):
    """deal.json — 단계별 산출물 이력을 누적 저장(다음 단계 입력)."""
    p = pathlib.Path(d["folder"]) / "deal.json"
    old = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
    hist = old.get("history", [])
    hist.append({"at": dt.datetime.now().isoformat(timespec="seconds"), "stage": 1, "produced": produced})
    state = {k: v for k, v in d.items() if k not in ("settings",)}
    state["stage_done"] = sorted(set(old.get("stage_done", [])) | {1})
    state["mail_log"] = d.get("mail_log", old.get("mail_log", []))
    state["model"] = d.get("model", old.get("model"))
    state["history"] = hist[-200:]
    p.write_text(json.dumps(state, ensure_ascii=False, indent=1, default=_ser), encoding="utf-8")
    return p


def prior_state(folder: pathlib.Path) -> dict:
    p = pathlib.Path(folder) / "deal.json"
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
