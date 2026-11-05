# -*- coding: utf-8 -*-
"""RFIS 마케팅 캘린더 자동 입력 — 구글 캘린더("마케팅 캘린더") 일정 → RFIS [계획 등록]·[결과 등록].

흐름: 구글 캘린더(iCal 비공개 주소) 읽기 → 입력안 만들기 → [승인 창]에서 확인·수정 → 승인한 건만 크롬으로 자동 입력
      → (설정 confirm_each) 저장 직전 건마다 한 번 더 확인 → 저장 → 등록 이력 기록(중복 방지)

구글 캘린더 일정 작성 규칙
  제목: [마케팅구분] 고객명 / 마케팅목적
        예) [F.I/LP/대주/증권] 삼성증권(주) / 시장 동향 전달
            [Tenant] 컴인워시(주) / 프로젝트 운용-임대차 관리      (같은 이름 목적이 여러 묶음에 있으면 "묶음-목적")
  메모(설명): 아래 항목을 줄마다 "항목: 내용"으로 (없으면 승인 창에서 입력)
        주요내용: KLI 세차장 용도전환 관련 진행사항 논의
        코멘트: 잠재임차인과 사업 추진 현황 공유
        당사담당자: 안용현, 임영규

실행: run_rfis.bat 더블클릭(메뉴). Python 3.11~3.13 + Playwright(첫 실행 때 자동 설치) + PC에 설치된 크롬 사용.
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import os
import re
import sys
import time
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
CFG = json.loads((HERE / "config.json").read_text(encoding="utf-8"))
DATA = HERE / "data"
STATE = DATA / "state.json"
LOGCSV = DATA / "등록이력.csv"
PROFILE = DATA / "chrome-profile"
KST = dt.timezone(dt.timedelta(hours=9))


# ────────────────────────────────────────────────────────────── 설정·상태
def load_settings():
    for name in ("settings.txt",):
        p = HERE / name
        if p.exists():
            for line in p.read_text(encoding="utf-8-sig").splitlines():
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip())


def load_state() -> dict:
    try:
        return json.loads(STATE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"plan": {}, "result": {}}


def save_state(st: dict):
    DATA.mkdir(parents=True, exist_ok=True)
    STATE.write_text(json.dumps(st, ensure_ascii=False, indent=1), encoding="utf-8")


def log_row(kind: str, item: dict, status: str, msg: str = ""):
    DATA.mkdir(parents=True, exist_ok=True)
    new = not LOGCSV.exists()
    with LOGCSV.open("a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["기록시각", "구분", "상태", "월/완료일", "주차", "마케팅구분", "고객명", "목적", "내용", "메시지"])
        w.writerow([dt.datetime.now(KST).strftime("%Y-%m-%d %H:%M"), kind, status,
                    item.get("month_label") or item.get("date"), item.get("week", ""), item.get("category"),
                    item.get("customer", ""), item.get("purpose", ""),
                    item.get("summary_text") or item.get("main", ""), msg])


# ────────────────────────────────────────────────────────────── 고객명 느슨한 일치
ALIAS = DATA / "고객명_별칭.json"
_CORP = re.compile(r"\(주\)|㈜|주식회사|\(유\)|유한회사|\(재\)|재단법인|\(사\)|사단법인|\(株\)|inc\.?|co\.,?\s*ltd\.?|ltd\.?",
                   re.I)


def norm_name(s: str) -> str:
    """'삼성증권(주)', '(주)삼성증권', '삼성 증권' → '삼성증권'."""
    return re.sub(r"[\s·.,\-()]", "", _CORP.sub("", s or "")).lower()


def load_alias() -> dict:
    try:
        return json.loads(ALIAS.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_alias(short: str, official: str):
    a = load_alias()
    if a.get(norm_name(short)) != official:
        a[norm_name(short)] = official
        DATA.mkdir(parents=True, exist_ok=True)
        ALIAS.write_text(json.dumps(a, ensure_ascii=False, indent=1), encoding="utf-8")


def rank_names(query: str, names: list[str]) -> list[tuple[int, str]]:
    """점수 높은 순: 3=정규화 일치, 2=앞부분 일치, 1=포함."""
    q = norm_name(query)
    out = []
    for n in names:
        k = norm_name(n)
        s = 3 if k == q else 2 if k.startswith(q) else 1 if q and (q in k or k in q) else 0
        if s:
            out.append((s, n))
    return sorted(out, key=lambda x: (-x[0], len(x[1])))


CONTACTS = DATA / "고객담당자.json"
RANKS = ["임원급", "팀장급", "팀원급"]


def load_contacts() -> dict:
    try:
        return json.loads(CONTACTS.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_contact(customer: str, name: str, dept: str, rank: str):
    c = load_contacts()
    c[norm_name(customer)] = {"customer": customer, "name": name, "dept": dept, "rank": rank}
    DATA.mkdir(parents=True, exist_ok=True)
    CONTACTS.write_text(json.dumps(c, ensure_ascii=False, indent=1), encoding="utf-8")


def import_result_excel(path: Path) -> int:
    """RFIS '마케팅 결과 관리' 엑셀 다운로드 파일 → 고객별 최근 담당자·부서·직급 사전."""
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True, read_only=True).active
    head = None
    c = load_contacts()
    n = 0
    rows = list(ws.iter_rows(values_only=True))
    for row in rows:
        vals = [str(x).strip() if x is not None else "" for x in row]
        if head is None:
            if "고객명" in vals and "완료일" in vals:
                head = {k: i for i, k in enumerate(vals)}
            continue
        g = lambda k: vals[head[k]] if k in head and head[k] < len(vals) else ""  # noqa: E731
        cust, dept, rank = g("고객명"), g("고객담당자부서"), g("고객담당자직급")
        if not cust or not dept or not rank or "," in cust:
            continue
        key = norm_name(cust)
        if key in c and c[key].get("date", "") >= g("완료일"):
            continue
        c[key] = {"customer": cust, "name": g("고객담당자명"), "dept": dept, "rank": rank, "date": g("완료일")}
        n += 1
    DATA.mkdir(parents=True, exist_ok=True)
    CONTACTS.write_text(json.dumps(c, ensure_ascii=False, indent=1), encoding="utf-8")
    return n


# ────────────────────────────────────────────────────────────── 주차 계산
def week_of(d: dt.date) -> tuple[int, int, int]:
    """(연, 월, 주차). thursday: 월~일 주를 목요일이 속한 달로 / day7: 1~7일=1주차."""
    if CFG.get("week_rule") == "day7":
        return d.year, d.month, (d.day - 1) // 7 + 1
    thu = d - dt.timedelta(days=d.weekday()) + dt.timedelta(days=3)
    first = thu.replace(day=1)
    first_thu = first + dt.timedelta(days=(3 - first.weekday()) % 7)
    return thu.year, thu.month, (thu - first_thu).days // 7 + 1


# ────────────────────────────────────────────────────────────── 구글 캘린더(iCal)
def fetch_ics(url: str) -> str:
    import urllib.error
    url = url.strip().strip('"').strip("'")
    if "calendar.google.com" in url and "/private-" not in url:
        raise SystemExit("ICS_URL이 'iCal 형식의 비공개 주소'가 아닙니다(공개 주소·embed 주소는 404가 납니다). "
                         "캘린더 설정 → 캘린더 통합 → 'iCal 형식의 비공개 주소'(…/private-…/basic.ics)를 넣어 주세요.")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return r.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        raise SystemExit(f"구글 캘린더 주소 오류(HTTP {e.code}). settings.txt 의 ICS_URL을 다시 복사해 넣어 주세요 — "
                         "'iCal 형식의 비공개 주소'를 통째로(…/basic.ics까지) 붙여야 합니다. "
                         "비공개 주소를 '재설정'했다면 새 주소로 바꿔야 합니다.") from e


def _unescape(v: str) -> str:
    return v.replace("\\n", "\n").replace("\\N", "\n").replace("\\,", ",").replace("\\;", ";").replace("\\\\", "\\")


def _parse_dt(key: str, val: str) -> dt.date | None:
    val = val.strip()
    try:
        if "VALUE=DATE" in key or re.fullmatch(r"\d{8}", val):
            return dt.datetime.strptime(val[:8], "%Y%m%d").date()
        t = dt.datetime.strptime(val[:15], "%Y%m%dT%H%M%S")
        if val.endswith("Z"):
            t = t.replace(tzinfo=dt.timezone.utc).astimezone(KST)
        return t.date()          # TZID=Asia/Seoul 등은 현지 시각 그대로
    except ValueError:
        return None


def parse_ics(text: str) -> list[dict]:
    text = re.sub(r"\r?\n[ \t]", "", text)        # 줄 접힘 해제
    events, cur = [], None
    for line in text.splitlines():
        if line == "BEGIN:VEVENT":
            cur = {}
        elif line == "END:VEVENT" and cur is not None:
            if cur.get("start") and cur.get("status") != "CANCELLED" and not cur.get("recurrence_id"):
                events.append(cur)
            cur = None
        elif cur is not None and ":" in line:
            k, v = line.split(":", 1)
            name = k.split(";")[0].upper()
            if name == "DTSTART":
                cur["start"] = _parse_dt(k, v)
            elif name == "SUMMARY":
                cur["summary"] = _unescape(v).strip()
            elif name == "DESCRIPTION":
                cur["desc"] = _unescape(v)
            elif name == "UID":
                cur["uid"] = v.strip()
            elif name == "STATUS":
                cur["status"] = v.strip().upper()
            elif name == "RRULE":
                cur["rrule"] = v
            elif name == "RECURRENCE-ID":
                cur["recurrence_id"] = v
    return events


# ────────────────────────────────────────────────────────────── 제목·메모 해석
def norm_category(raw: str) -> str | None:
    s = raw.strip()
    for c in CFG["categories"]:
        if s.replace(" ", "").lower() == c.replace(" ", "").lower():
            return c
    return CFG["category_alias"].get(s.upper()) or CFG["category_alias"].get(s)


def resolve_purpose(raw: str) -> tuple[str | None, str | None, str]:
    """→ (묶음, 목적, 경고). '묶음-목적' 또는 '목적'만."""
    s = (raw or "").strip()
    if not s:
        return None, None, "목적 없음"
    groups = CFG["purposes"]
    if "-" in s:
        g, p = s.split("-", 1)
        if g.strip() in groups and p.strip() in groups[g.strip()]:
            return g.strip(), p.strip(), ""
    hits = [(g, p) for g, ps in groups.items() for p in ps if p == s]
    if not hits:
        hits = [(g, p) for g, ps in groups.items() for p in ps if p.replace(" ", "") == s.replace(" ", "")]
    if len(hits) == 1:
        return hits[0][0], hits[0][1], ""
    if len(hits) > 1:
        g, p = hits[0]
        return g, p, f"'{s}'는 여러 묶음에 있어 '{g}'로 가정 — 다르면 수정"
    return None, s, f"목적 '{s}'를 RFIS 목록에서 못 찾음 — 수정 필요"


def parse_event(ev: dict) -> dict | None:
    title = (ev.get("summary", "") or "").translate(str.maketrans({"［": "[", "］": "]", "【": "[", "】": "]", "〔": "[", "〕": "]"}))
    m = re.match(r"^[\s\W]*?\[([^\]]+)\]\s*(.*)$", title)
    if not m:
        return None
    cat = norm_category(m.group(1))
    rest = m.group(2).strip()
    parts = re.split(r"\s*[/|]\s*", rest, maxsplit=1)
    customer = parts[0].strip() if parts else ""
    customer = load_alias().get(norm_name(customer), customer)   # 예전에 고른 정식 이름
    purpose_raw = parts[1].strip() if len(parts) > 1 else ""
    group, purpose, warn = resolve_purpose(purpose_raw)
    memo = {}
    for line in (ev.get("desc") or "").splitlines():
        mm = re.match(r"^\s*([^:：]+)[:：]\s*(.*)$", line)
        if mm:
            memo[mm.group(1).strip().replace(" ", "")] = mm.group(2).strip()
    y, mo, w = week_of(ev["start"])
    warns = [x for x in [warn] if x]
    if not cat:
        warns.append(f"마케팅구분 '{m.group(1)}' 인식 불가")
    if not customer:
        warns.append("고객명 없음")
    return {"uid": ev.get("uid") or f"{ev['start']}_{ev.get('summary')}", "date": ev["start"].isoformat(),
            "year": y, "month": mo, "week": w, "category": cat or m.group(1), "customer": customer,
            "group": group, "purpose": purpose, "main": memo.get("주요내용", ""),
            "comment": memo.get("코멘트", ""),
            "staff": memo.get("당사담당자") or CFG.get("default_staff", ""),
            **_contact_from_memo(memo, customer),
            "warn": " / ".join(warns), "recurring": bool(ev.get("rrule"))}


def _contact_from_memo(memo: dict, customer: str) -> dict:
    """메모 '고객담당자: 이름 / 부서 / 직급' 또는 '담당자명:'·'부서명:'·'직급:' 각각. 없으면 지난 등록 이력."""
    name = dept = rank = ""
    raw = memo.get("고객담당자") or memo.get("고객측담당자") or ""
    if raw:
        parts = [x.strip() for x in re.split(r"[/|,]", raw) if x.strip()]
        for x in parts:
            if x in RANKS or x.endswith("급"):
                rank = x
            elif not name:
                name = x
            elif not dept:
                dept = x
    name = memo.get("담당자명") or memo.get("고객담당자명") or name
    dept = memo.get("부서명") or memo.get("부서") or memo.get("고객담당자부서") or dept
    rank = memo.get("직급") or memo.get("고객담당자직급") or rank
    if not (dept and rank):
        hist = load_contacts().get(norm_name(customer)) or {}
        name, dept, rank = name or hist.get("name", ""), dept or hist.get("dept", ""), rank or hist.get("rank", "")
    return {"c_name": name, "c_dept": dept, "c_rank": rank}


def load_events() -> list[dict]:
    load_settings()
    url = os.environ.get("ICS_URL")
    if not url:
        raise SystemExit("settings.txt 에 ICS_URL(구글 캘린더 'iCal 형식의 비공개 주소')을 넣어 주세요.")
    out, skipped = [], []
    for ev in parse_ics(fetch_ics(url)):
        e = parse_event(ev)
        if e:
            out.append(e)
        else:
            skipped.append(f"{ev['start']} '{ev.get('summary', '')}'")
    if skipped:
        print(f"  (제목이 [마케팅구분]으로 시작하지 않아 건너뛴 일정 {len(skipped)}건)")
        for s in skipped[-5:]:
            print(f"     - {s}")
        print("     → 제목 형식: [F.I/LP/대주/증권] 고객명 / 마케팅목적")
    return sorted(out, key=lambda x: x["date"])


# ────────────────────────────────────────────────────────────── 입력안
def build_plans(events: list[dict], st: dict) -> list[dict]:
    today = dt.date.today()
    lim = today + dt.timedelta(days=CFG.get("plan_lookahead_days", 45))
    groups: dict = {}
    for e in events:
        d = dt.date.fromisoformat(e["date"])
        if not (week_of(today)[:2] <= (e["year"], e["month"]) and d <= lim):
            continue
        key = (e["year"], e["month"], e["category"])
        groups.setdefault(key, []).append(e)
    plans = []
    for (y, mo, cat), evs in sorted(groups.items()):
        weeks = {w: [] for w in range(1, 6)}
        for e in evs:
            if e["customer"] and e["customer"] not in weeks[min(e["week"], 5)]:
                weeks[min(e["week"], 5)].append(e["customer"])
        done = [e["uid"] for e in evs if e["uid"] in st["plan"]]
        warn = [x["warn"] for x in evs if x["warn"]]
        if done:
            warn.insert(0, f"이 월·구분은 이미 {len(done)}건 등록됨 — 다시 등록하면 중복될 수 있음(기존 카드 수정 권장)")
        plans.append({"kind": "계획", "year": y, "month": mo, "month_label": f"{y}-{mo:02d}", "category": cat,
                      "weeks": {w: ", ".join(v) for w, v in weeks.items()}, "uids": [e["uid"] for e in evs],
                      "summary_text": " / ".join(f"{w}주차: {', '.join(v)}" for w, v in weeks.items() if v),
                      "warn": " / ".join(warn), "selected": not done and not any(e["warn"] for e in evs if "인식 불가" in e["warn"])})
    return plans


def build_results(events: list[dict], st: dict) -> list[dict]:
    today = dt.date.today()
    lo = today - dt.timedelta(days=CFG.get("result_lookback_days", 30))
    res = []
    for e in events:
        d = dt.date.fromisoformat(e["date"])
        if not (lo <= d <= today) or e["uid"] in st["result"]:
            continue
        warns = [e["warn"]] if e["warn"] else []
        if not e["main"]:
            warns.append("주요내용 없음")
        if not e["comment"]:
            warns.append("코멘트 없음")
        if not e["c_dept"]:
            warns.append("고객 부서명 없음")
        if not e["c_rank"]:
            warns.append("고객 직급 없음")
        elif e["c_rank"] not in RANKS:
            warns.append(f"직급 '{e['c_rank']}' → 임원급·팀장급·팀원급 중 하나로")
        ready = e["main"] and e["comment"] and e["purpose"] and e["c_dept"] and e["c_rank"] and not e["warn"]
        res.append({**e, "kind": "결과", "warn": " / ".join(warns), "selected": bool(ready)})
    return res


# ────────────────────────────────────────────────────────────── 승인 창
PLAN_COLS = [("category", "마케팅구분", 150), ("month_label", "월", 70)] + \
            [(f"w{i}", f"{i}주차", 150) for i in range(1, 6)] + [("warn", "확인 필요", 260)]
RES_COLS = [("date", "완료일", 85), ("week", "주차", 45), ("category", "마케팅구분", 125), ("customer", "고객명", 120),
            ("c_name", "고객담당자", 75), ("c_dept", "부서명", 100), ("c_rank", "직급", 60),
            ("purpose_full", "마케팅목적", 180), ("staff", "당사담당자", 100), ("main", "주요내용", 200),
            ("comment", "코멘트", 200), ("warn", "확인 필요", 200)]


def _row_val(it, key):
    if key.startswith("w") and key[1:].isdigit():
        return it["weeks"].get(int(key[1:]), "")
    if key == "purpose_full":
        return f"{it.get('group') or '?'}-{it.get('purpose') or '?'}"
    return str(it.get(key, "") or "")


def _set_val(it, key, val):
    if key.startswith("w") and key[1:].isdigit():
        it["weeks"][int(key[1:])] = val
        it["summary_text"] = " / ".join(f"{w}주차: {v}" for w, v in it["weeks"].items() if v)
    elif key == "purpose_full":
        g, p, _ = resolve_purpose(val)
        it["group"], it["purpose"] = g, p
    elif key == "week":
        it["week"] = int(re.sub(r"\D", "", val) or it["week"])
    else:
        it[key] = val


def approve(items: list[dict], kind: str) -> list[dict]:
    """tkinter 승인 창. 체크(✔) 칸 클릭 = 선택 전환, 다른 칸 더블클릭 = 수정."""
    if not items:
        why = ("결과 등록은 오늘까지 지난 일정(최근 30일) 중 미등록 건만 대상입니다" if kind == "결과"
               else "계획 등록은 이번 달부터 45일 안의 일정이 대상입니다")
        print(f"\n[{kind}] 새로 입력할 항목이 없습니다. ({why})")
        return []
    try:
        import tkinter as tk
        from tkinter import ttk
    except ImportError:
        return _approve_console(items, kind)
    cols = PLAN_COLS if kind == "계획" else RES_COLS
    for it in items:
        print(f"   {'✔' if it['selected'] else '·'} {it.get('month_label') or it.get('date')} {it['category']} "
              f"{it.get('customer', '') or it.get('summary_text', '')}" + (f"  ⚠ {it['warn']}" if it["warn"] else ""))
    root = tk.Tk()
    root.title(f"RFIS {kind} 등록 — 입력 전 승인")
    root.geometry("1500x560+40+40")
    root.attributes("-topmost", True)            # 명령 창 뒤에 숨지 않도록 맨 앞에
    root.after(800, lambda: root.attributes("-topmost", False))
    root.lift()
    root.focus_force()
    tk.Label(root, text=f"아래 내용으로 RFIS [{kind} 등록]에 입력합니다.  ✔ 칸 클릭: 선택/해제 · 다른 칸 더블클릭: 수정",
             font=("맑은 고딕", 11)).pack(anchor="w", padx=10, pady=6)
    frame = tk.Frame(root)
    frame.pack(fill="both", expand=True, padx=10)
    tv = ttk.Treeview(frame, columns=["sel"] + [c[0] for c in cols], show="headings", height=18)
    tv.heading("sel", text="✔")
    tv.column("sel", width=36, anchor="center")
    for k, title, w in cols:
        tv.heading(k, text=title)
        tv.column(k, width=w, anchor="w")
    sb = ttk.Scrollbar(frame, orient="horizontal", command=tv.xview)
    tv.configure(xscrollcommand=sb.set)
    tv.pack(fill="both", expand=True)
    sb.pack(fill="x")

    def refresh():
        tv.delete(*tv.get_children())
        for i, it in enumerate(items):
            tv.insert("", "end", iid=str(i), values=["✔" if it["selected"] else ""] + [_row_val(it, c[0]) for c in cols])

    def on_click(ev):
        if tv.identify_column(ev.x) == "#1" and tv.identify_row(ev.y):
            i = int(tv.identify_row(ev.y))
            items[i]["selected"] = not items[i]["selected"]
            refresh()

    def on_double(ev):
        row, col = tv.identify_row(ev.y), tv.identify_column(ev.x)
        if not row or col in ("#1", ""):
            return
        key = cols[int(col[1:]) - 2][0]
        if key == "warn":
            return
        box = tv.bbox(row, col)
        if not box:
            return
        x, y, w, h = box
        var = tk.StringVar(value=_row_val(items[int(row)], key))
        ent = tk.Entry(tv, textvariable=var)
        ent.place(x=x, y=y, width=max(w, 200), height=h)
        ent.focus()

        def done(_=None):
            _set_val(items[int(row)], key, var.get().strip())
            ent.destroy()
            refresh()
        ent.bind("<Return>", done)
        ent.bind("<FocusOut>", done)

    tv.bind("<Button-1>", on_click, add="+")
    tv.bind("<Double-1>", on_double)
    result = {"ok": False}
    bar = tk.Frame(root)
    bar.pack(fill="x", padx=10, pady=8)
    def on_ok():
        if not any(it["selected"] for it in items):
            from tkinter import messagebox
            messagebox.showwarning("선택 없음", "체크(✔)된 항목이 없습니다.\n맨 왼쪽 ✔ 칸을 클릭해 입력할 항목을 선택하세요.\n"
                                   "'확인 필요'에 적힌 내용(예: 주요내용 없음)은 해당 칸을 더블클릭해 채울 수 있습니다.", parent=root)
            return
        result.update(ok=True)
        root.destroy()
    tk.Button(bar, text="  승인 — 선택 항목 입력 시작  ", bg="#C00000", fg="white", font=("맑은 고딕", 11, "bold"),
              command=on_ok).pack(side="right")
    tk.Button(bar, text="  취소  ", font=("맑은 고딕", 11), command=root.destroy).pack(side="right", padx=8)
    refresh()
    root.mainloop()
    return [it for it in items if it["selected"]] if result["ok"] else []


def _approve_console(items, kind):
    for i, it in enumerate(items, 1):
        print(f"{i}. {it.get('month_label') or it.get('date')} {it['category']} {it.get('customer', '')} "
              f"{it.get('summary_text') or it.get('main', '')} {('⚠ ' + it['warn']) if it['warn'] else ''}")
    s = input("입력할 번호(쉼표, 전체=a, 취소=Enter): ").strip()
    if s.lower() == "a":
        return items
    return [items[int(x) - 1] for x in re.findall(r"\d+", s) if 0 < int(x) <= len(items)]


# ────────────────────────────────────────────────────────────── RFIS 자동 입력(Playwright)
class RFIS:
    def __init__(self):
        from playwright.sync_api import sync_playwright
        self.pw = sync_playwright().start()
        PROFILE.mkdir(parents=True, exist_ok=True)
        try:
            self.ctx = self.pw.chromium.launch_persistent_context(str(PROFILE), channel="chrome", headless=False,
                                                                  no_viewport=True, args=["--start-maximized"],
                                                                  chromium_sandbox=True)
        except Exception:  # noqa: BLE001  크롬이 없으면 Edge
            self.ctx = self.pw.chromium.launch_persistent_context(str(PROFILE), channel="msedge", headless=False,
                                                                  no_viewport=True, args=["--start-maximized"],
                                                                  chromium_sandbox=True)
        self.page = self.ctx.pages[0] if self.ctx.pages else self.ctx.new_page()
        self.page.set_default_timeout(8000)

    def open(self):
        p = self.page
        p.goto(CFG["rfis_url"])
        for _ in range(3):
            try:
                p.get_by_text("마케팅 캘린더", exact=True).first.wait_for(timeout=8000)
                if p.get_by_role("button", name="계획 등록").count():
                    return
            except Exception:  # noqa: BLE001
                pass
            input("\n  ▶ 브라우저에서 RFIS에 로그인한 뒤 이 창에서 Enter를 누르세요 (로그인은 다음부터 유지됩니다)...")
            p.goto(CFG["rfis_url"])
        raise RuntimeError("마케팅 캘린더 화면을 열지 못했습니다.")

    def close(self):
        try:
            self.ctx.close()
            self.pw.stop()
        except Exception:  # noqa: BLE001
            pass

    # 도우미 ─────────────────────────────────
    def drawer(self, title: str):
        p = self.page
        d = p.locator('[role="dialog"]').filter(has_text=title).last
        d.wait_for(timeout=8000)
        return d

    @staticmethod
    def row(scope, label: str):
        x = (f"xpath=.//*[normalize-space(text())='{label}']"
             f"/ancestor::*[.//input or .//textarea or .//*[@role='combobox']][1]")
        return scope.locator(x).first

    def _control(self, scope, label):
        r = self.row(scope, label)
        c = r.locator("[role='combobox'], input:not([type='hidden']):not([aria-hidden='true']), textarea")
        return c.first

    def fill(self, scope, label, value, placeholder=None):
        el = scope.get_by_placeholder(placeholder).first if placeholder else self._control(scope, label)
        el.click()
        el.fill(value)

    def pick(self, scope, label, value, typed=False):
        """드롭다운: 열고 → (필요시 검색어 입력) → 옵션 클릭."""
        p = self.page
        self._control(scope, label).click()
        if typed:
            p.keyboard.type(value, delay=40)
            time.sleep(0.8)
        time.sleep(0.3)
        opts = p.get_by_role("option", name=value, exact=True)
        if opts.count() == 0:
            opts = p.get_by_role("option").filter(has_text=value)
        if opts.count() == 0:
            opts = p.locator("[data-part='item'], li").filter(has_text=value)
        if opts.count() == 0:
            p.keyboard.press("Escape")
            raise LookupError(f"'{label}' 목록에 '{value}'이(가) 없습니다")
        exact = [i for i in range(opts.count()) if opts.nth(i).inner_text().strip() == value]
        opts.nth(exact[0] if exact else 0).click()

    def pick_customer(self, scope, name: str) -> str:
        """고객 선택: '삼성증권'만 써도 '삼성증권(주)'를 찾아 고른다. 고른 이름을 돌려준다."""
        p = self.page
        self._control(scope, "고객명").click()
        base = _CORP.sub("", name).strip() or name
        query = next((v for k, v in CFG.get("customer_alias", {}).items() if norm_name(k) == norm_name(base)), base)
        p.keyboard.type(query, delay=40)
        time.sleep(1.2)
        opts = p.get_by_role("option")
        if opts.count() == 0:
            opts = p.locator("[data-part='item'], li, tr").filter(has_text=query[:2])
        texts = [opts.nth(i).inner_text().strip().split("\n")[0] for i in range(min(opts.count(), 50))]
        ranked = rank_names(query, texts) or rank_names(name, texts)
        if not ranked:
            p.keyboard.press("Escape")
            raise LookupError(f"고객정보에서 '{name}'을(를) 찾지 못했습니다 — RFIS 고객정보에 먼저 등록하거나 이름을 확인하세요")
        best = [n for s, n in ranked if s == ranked[0][0]]
        choice = best[0]
        if len(best) > 1 and ranked[0][0] < 3:
            print(f"\n  '{name}'과(와) 비슷한 고객이 여러 개입니다:")
            for i, n in enumerate(best[:9], 1):
                print(f"    {i}. {n}")
            a = input("  번호를 고르세요(Enter=1번): ").strip()
            choice = best[int(a) - 1] if a.isdigit() and 0 < int(a) <= len(best[:9]) else best[0]
        idx = texts.index(choice)
        opts.nth(idx).click()
        save_alias(name, choice)
        if choice != name:
            print(f"  · 고객명 '{name}' → '{choice}'")
        return choice

    def purpose(self, scope, group, purpose):
        p = self.page
        self._control(scope, "마케팅 목적").click()
        chip = p.locator(f"xpath=//*[normalize-space(text())='{group}']/following::*[normalize-space(text())='{purpose}'][1]")
        chip.first.click()
        p.keyboard.press("Escape")

    def save(self, scope, desc: str) -> bool:
        if CFG.get("confirm_each", True):
            a = input(f"\n  ▶ [{desc}] 입력을 마쳤습니다. 화면을 확인하고 Enter=저장 / s=건너뛰기 / q=중단: ").strip().lower()
            if a == "q":
                raise KeyboardInterrupt
            if a == "s":
                self.back(scope)
                return False
        scope.get_by_role("button", name="저장").first.click()
        scope.wait_for(state="hidden", timeout=15000)
        return True

    def back(self, scope):
        try:
            self.page.keyboard.press("Escape")
            if scope.is_visible():
                scope.locator("button").first.click()   # 좌상단 ← 버튼
        except Exception:  # noqa: BLE001
            pass

    def dump(self, name: str):
        DATA.mkdir(parents=True, exist_ok=True)
        stamp = dt.datetime.now().strftime("%m%d_%H%M%S")
        try:
            (DATA / f"diag_{name}_{stamp}.html").write_text(self.page.content(), encoding="utf-8")
            self.page.screenshot(path=str(DATA / f"diag_{name}_{stamp}.png"), full_page=False)
        except Exception:  # noqa: BLE001
            pass

    # 등록 ─────────────────────────────────
    def add_plan(self, it: dict) -> bool:
        p = self.page
        p.get_by_role("button", name="계획 등록").click()
        d = self.drawer("계획 등록")
        self.pick(d, "마케팅구분", it["category"])
        self._month(d, it["year"], it["month"])
        for w in range(1, 6):
            if it["weeks"].get(w):
                self.fill(d, f"{w}주차", it["weeks"][w])
        return self.save(d, f"계획 {it['month_label']} {it['category']}")

    def _month(self, d, y, m):
        """월 선택: 입력칸이면 'YYYY-MM' 입력, 달력 팝업이면 연·월 클릭."""
        p = self.page
        ctl = self._control(d, "월")
        ctl.click()
        try:
            tag = ctl.evaluate("e => e.tagName")
            if tag == "INPUT" and not ctl.get_attribute("readonly"):
                ctl.fill(f"{y}-{m:02d}")
                p.keyboard.press("Enter")
                return
        except Exception:  # noqa: BLE001
            pass
        for cand in (f"{m}월", f"{m:02d}월", f"{m}"):
            loc = p.get_by_text(cand, exact=True)
            if loc.count():
                loc.last.click()
                return
        raise LookupError("월 선택 달력에서 월을 찾지 못했습니다 — 메뉴 4번(화면 구조 저장) 결과를 보내 주세요")

    def modal(self, has_text: str):
        d = self.page.locator('[role="dialog"], [role="alertdialog"]').filter(has_text=has_text).last
        d.wait_for(timeout=8000)
        return d

    def _check_row(self, row):
        for sel in ("label", "[data-part='control']", "input[type='checkbox']"):
            loc = row.locator(sel)
            if loc.count():
                try:
                    loc.first.click(timeout=3000)
                    return
                except Exception:  # noqa: BLE001
                    continue
        row.locator("td").first.click()

    def choose_customer(self, d, name: str) -> str:
        """고객 선택: 팝업 검색(조회) → 행 체크 → 선택 완료. 고른 정식 이름을 돌려준다."""
        p = self.page
        d.get_by_placeholder("고객 선택").first.click()
        m = self.modal("선택 완료")
        base = _CORP.sub("", name).strip() or name
        query = next((v for k, v in CFG.get("customer_alias", {}).items() if norm_name(k) == norm_name(base)), base)
        box = m.get_by_placeholder("고객명으로 검색").first
        box.fill(query)
        m.get_by_role("button", name="조회").first.click()
        time.sleep(1.5)
        rows = m.locator("tbody tr")
        if rows.count() == 0:
            rows = m.locator("[role='row']")
        cands = []
        for i in range(min(rows.count(), 60)):
            tds = rows.nth(i).locator("td")
            cells = [tds.nth(j).inner_text().strip() for j in range(tds.count())]
            cname = cells[3] if len(cells) >= 5 else max(cells, key=lambda c: len(rank_names(query, [c])) and len(c), default="")
            if cname:
                cands.append((i, cname))
        ranked = rank_names(query, [c for _, c in cands]) or rank_names(name, [c for _, c in cands])
        if not ranked:
            m.get_by_role("button", name="닫기").first.click()
            raise LookupError(f"고객정보에서 '{name}'(검색어 '{query}')을(를) 찾지 못했습니다 — RFIS 고객정보에 먼저 등록하세요")
        best = [n for s_, n in ranked if s_ == ranked[0][0]]
        choice = best[0]
        if len(set(best)) > 1:
            print(f"\n  '{name}'과(와) 비슷한 고객이 여러 개입니다:")
            for i, n in enumerate(best[:9], 1):
                print(f"    {i}. {n}")
            a = input("  번호를 고르세요(Enter=1번): ").strip()
            choice = best[int(a) - 1] if a.isdigit() and 0 < int(a) <= len(best[:9]) else best[0]
        idx = next(i for i, c in cands if c == choice)
        self._check_row(rows.nth(idx))
        m.get_by_role("button", name="선택 완료").first.click()
        save_alias(name, choice)
        if choice != name:
            print(f"  · 고객명 '{name}' → '{choice}'")
        return choice

    def fill_contact(self, it: dict):
        """'고객사별 담당자 등록' 팝업: 담당자명·부서명·직급 → 담당자 등록 완료."""
        p = self.page
        m = self.modal("고객사별 담당자 등록")
        if it.get("c_name"):
            m.get_by_placeholder("예: 김철수").first.fill(it["c_name"])
        m.get_by_placeholder("예: 투자운용팀").first.fill(it["c_dept"])
        m.get_by_text("직급 선택").first.click()
        opt = p.get_by_role("option", name=it["c_rank"], exact=True)
        if opt.count() == 0:
            opt = p.locator("[data-part='item']").filter(has_text=it["c_rank"])
        opt.first.click()
        m.get_by_role("button", name="담당자 등록 완료").first.click()
        m.wait_for(state="hidden", timeout=8000)

    def choose_staff(self, d, names: list[str]):
        """당사담당자 선택: 검색 팝업(체크 → 선택 완료) 또는 드롭다운, 화면 형태에 맞춰 처리."""
        p = self.page
        for nm in names:
            d.get_by_placeholder("당사담당자 선택").first.click()
            time.sleep(0.8)
            dlg = p.locator('[role="dialog"]').filter(has=p.get_by_role("button", name="선택 완료")).last
            if dlg.count() and dlg.is_visible():
                box = dlg.locator("input[type='text'], input:not([type])").first
                box.fill(nm)
                btn = dlg.get_by_role("button", name="조회")
                if btn.count():
                    btn.first.click()
                else:
                    p.keyboard.press("Enter")
                time.sleep(1.2)
                row = dlg.locator("tbody tr, [role='row']").filter(has_text=nm)
                if row.count() == 0:
                    dlg.get_by_role("button", name="닫기").first.click()
                    raise LookupError(f"당사담당자 '{nm}'을(를) 찾지 못했습니다")
                self._check_row(row.first)
                dlg.get_by_role("button", name="선택 완료").first.click()
            else:
                p.keyboard.type(nm, delay=40)
                time.sleep(0.8)
                opt = p.get_by_role("option").filter(has_text=nm)
                if opt.count() == 0:
                    p.keyboard.press("Escape")
                    raise LookupError(f"당사담당자 '{nm}'을(를) 찾지 못했습니다")
                opt.first.click()
                p.keyboard.press("Escape")

    def add_result(self, it: dict) -> bool:
        p = self.page
        p.get_by_role("button", name="결과 등록").click()
        d = self.drawer("결과 등록")
        # 1) 고객명 → 담당자 등록
        it["customer"] = self.choose_customer(d, it["customer"])
        self.fill_contact(it)
        # 2) 완료일(키인)
        dt_in = d.get_by_placeholder("YYYY-MM-DD").first
        dt_in.click()
        dt_in.fill(it["date"])
        p.keyboard.press("Tab")
        p.keyboard.press("Escape")
        # 3) 완료주차·4) 마케팅구분 드롭다운
        self.pick(d, "완료주차", f"{it['week']}주차")
        self.pick(d, "마케팅구분", it["category"])
        # 5) 마케팅 목적 버튼
        self.purpose(d, it["group"], it["purpose"])
        # 당사담당자
        names = [x.strip() for x in re.split(r"[,/]", it["staff"]) if x.strip()]
        if names:
            self.choose_staff(d, names)
        # 6) 주요내용·코멘트 직접 입력
        d.get_by_placeholder("주요 내용 입력").first.fill(it["main"])
        d.get_by_placeholder("주요 논의내용", exact=False).first.fill(it["comment"])
        saved = self.save(d, f"결과 {it['date']} {it['customer']}")
        if saved:
            save_contact(it["customer"], it.get("c_name", ""), it["c_dept"], it["c_rank"])
        return saved


def run(kind: str):
    st = load_state()
    print("\n[1/3] 구글 캘린더를 읽는 중...")
    events = load_events()
    items = build_plans(events, st) if kind == "계획" else build_results(events, st)
    print(f"[2/3] 입력안 {len(items)}건 — 승인 창이 열립니다. 안 보이면 작업 표시줄에서 'RFIS {kind} 등록' 창을 누르세요.")
    chosen = approve(items, kind)
    if not chosen:
        if items:
            print("  승인 창을 닫아 입력을 취소했습니다.")
        return
    print(f"[3/3] RFIS에 {len(chosen)}건 입력을 시작합니다. 크롬 창을 건드리지 마세요.")
    r = RFIS()
    ok = 0
    try:
        r.open()
        for it in chosen:
            desc = it.get("month_label") or f"{it['date']} {it['customer']}"
            try:
                saved = r.add_plan(it) if kind == "계획" else r.add_result(it)
                if saved:
                    ok += 1
                    for u in (it.get("uids") or [it["uid"]]):
                        st["plan" if kind == "계획" else "result"][u] = dt.datetime.now(KST).isoformat()
                    save_state(st)
                    log_row(kind, it, "저장")
                    print(f"  ✔ 저장: {desc}")
                else:
                    log_row(kind, it, "건너뜀")
                    print(f"  - 건너뜀: {desc}")
            except KeyboardInterrupt:
                print("  중단했습니다.")
                break
            except Exception as e:  # noqa: BLE001
                r.dump(f"오류_{kind}")
                log_row(kind, it, "오류", str(e)[:200])
                print(f"  ✖ 오류: {desc} — {e}\n    (data 폴더의 diag_오류 파일을 보내 주시면 고치겠습니다)")
                r.back(r.page.locator('[role="dialog"]').last)
                r.page.goto(CFG["rfis_url"])
    finally:
        input(f"\n완료 {ok}/{len(chosen)}건. Enter를 누르면 크롬을 닫습니다...")
        r.close()


def diagnose():
    """등록 창·드롭다운을 연 상태의 화면 구조를 저장(입력은 하지 않음)."""
    r = RFIS()
    try:
        r.open()
        p = r.page
        for btn in ("계획 등록", "결과 등록"):
            p.get_by_role("button", name=btn).click()
            time.sleep(1.5)
            r.dump(btn)
            d = p.locator('[role="dialog"]').last
            for label in (["마케팅구분", "월"] if btn == "계획 등록" else ["고객명", "완료주차", "마케팅 목적", "당사담당자"]):
                try:
                    RFIS._control(r, d, label).click()
                    time.sleep(1.0)
                    r.dump(f"{btn}_{label}")
                    p.keyboard.press("Escape")
                    time.sleep(0.3)
                except Exception as e:  # noqa: BLE001
                    print(f"  {btn}/{label}: {e}")
            r.back(d)
            p.goto(CFG["rfis_url"])
            time.sleep(1)
        print(f"\n화면 구조를 저장했습니다: {DATA}\\diag_*.html / .png  → 이 파일들을 Claude에게 보내 주세요.")
    finally:
        input("Enter를 누르면 크롬을 닫습니다...")
        r.close()


def preview():
    st = load_state()
    events = load_events()
    print(f"\n구글 캘린더 마케팅 일정 {len(events)}건")
    for e in events[-15:]:
        print(f"  {e['date']} {e['year']}-{e['month']:02d} {e['week']}주차 | {e['category']} | {e['customer']} | "
              f"{e.get('group')}-{e.get('purpose')} {('⚠ ' + e['warn']) if e['warn'] else ''}")
    pl, rs = build_plans(events, st), build_results(events, st)
    print(f"\n계획 등록 후보 {len(pl)}건:")
    for x in pl:
        print(f"  {x['month_label']} {x['category']}: {x['summary_text']} {('⚠ ' + x['warn']) if x['warn'] else ''}")
    print(f"결과 등록 후보 {len(rs)}건:")
    for x in rs:
        print(f"  {x['date']} {x['week']}주차 {x['category']} {x['customer']} {('⚠ ' + x['warn']) if x['warn'] else ''}")


def menu():
    if not (HERE / "settings.txt").exists():
        (HERE / "settings.txt").write_text((HERE / "settings.example.txt").read_text(encoding="utf-8"), encoding="utf-8")
    items = {"1": "미리보기 (구글 캘린더 → 입력안 확인, RFIS 접속 안 함)", "2": "계획 등록 (승인 후 RFIS 입력)",
             "3": "결과 등록 (승인 후 RFIS 입력)", "4": "RFIS 화면 구조 저장 (처음 한 번 / 오류 시)",
             "5": "설정 파일(settings.txt) 열기", "6": "등록 이력 열기",
             "7": "고객담당자 사전 만들기 (RFIS 결과관리 엑셀 다운로드 파일 가져오기)", "0": "끝내기"}
    while True:
        print("\n ==========================================\n  RFIS 마케팅 캘린더 자동 입력\n ==========================================")
        for k, v in items.items():
            print(f"  {k}. {v}")
        sel = input("\n번호를 입력하고 Enter: ").strip()
        try:
            if sel == "1":
                preview()
            elif sel == "2":
                run("계획")
            elif sel == "3":
                run("결과")
            elif sel == "4":
                diagnose()
            elif sel == "5":
                os.system(f'notepad "{HERE / "settings.txt"}"')
            elif sel == "6":
                if LOGCSV.exists():
                    os.startfile(str(LOGCSV))  # noqa: S606
                else:
                    print("아직 등록 이력이 없습니다.")
            elif sel == "7":
                xs = sorted(HERE.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True)
                if not xs:
                    print("이 폴더에 RFIS '마케팅 결과 관리' → [엑셀 다운로드] 파일(.xlsx)을 넣고 다시 실행하세요.")
                else:
                    print(f"  {xs[0].name} → 고객 {import_result_excel(xs[0])}곳의 최근 담당자·부서·직급을 저장했습니다.")
            elif sel == "0":
                return
        except SystemExit as e:
            print(f"\n[안내] {e}")
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            print(f"\n[오류] {e} — 이 창을 캡처해 보내 주세요.")


if __name__ == "__main__":
    if "--preview" in sys.argv:
        preview()
    else:
        menu()
