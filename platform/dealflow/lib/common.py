# -*- coding: utf-8 -*-
"""딜 플로우 공통 도구: 경로, 파일명 규칙(YYYYMMDD_<유형>_<제목>_v##), 영업일, 숫자 표기."""
from __future__ import annotations

import datetime as dt
import json
import pathlib
import re
import sys

HERE = pathlib.Path(__file__).resolve().parents[1]          # platform/dealflow
REPO = HERE.parents[1] if (HERE.parents[1] / "platform").exists() else None
TPL = HERE / "templates"

# kli_deck(PPT 표준 서식)은 저장소의 platform/scripts 또는 배포본의 vendor/ 에서 읽는다
for p in (HERE / "vendor", (REPO / "platform/scripts") if REPO else None):
    if p and p.exists() and str(p) not in sys.path:
        sys.path.insert(0, str(p))


def kli_template() -> pathlib.Path:
    for p in (HERE / "vendor/KLI_빈템플릿_v5.8.pptx",
              (REPO / "platform/templates/kli_deck/KLI_빈템플릿_v5.8.pptx") if REPO else None):
        if p and p.exists():
            return p
    raise FileNotFoundError("KLI_빈템플릿_v5.8.pptx 를 찾지 못했습니다 (vendor 폴더 확인)")


def load_json(name: str):
    return json.loads((TPL / name).read_text(encoding="utf-8"))


TODAY = dt.date.today()


def stamp(d: dt.date | None = None) -> str:
    return (d or TODAY).strftime("%Y%m%d")


def outpath(folder: pathlib.Path, kind: str, title: str, ext: str, d: dt.date | None = None) -> pathlib.Path:
    """파일명 규칙 + 같은 이름이 있으면 v## 를 올린다(기존 파일은 절대 덮어쓰지 않음)."""
    folder.mkdir(parents=True, exist_ok=True)
    title = re.sub(r'[\\/:*?"<>|\s]+', "", title)
    v = 1
    while True:
        p = folder / f"{stamp(d)}_{kind}_{title}_v{v:02d}.{ext}"
        if not p.exists():
            return p
        v += 1


# ── 날짜 ───────────────────────────────────────────────────────────────
def holidays() -> set[dt.date]:
    try:
        return {dt.date.fromisoformat(x) for x in load_json("holidays_kr.json")["dates"]}
    except Exception:
        return set()


_HOL = None


def is_bizday(d: dt.date) -> bool:
    global _HOL
    if _HOL is None:
        _HOL = holidays()
    return d.weekday() < 5 and d not in _HOL


def prev_bizday(d: dt.date) -> dt.date:
    while not is_bizday(d):
        d -= dt.timedelta(days=1)
    return d


def to_date(v) -> dt.date | None:
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    m = re.match(r"(\d{4})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})", s)
    if m:
        return dt.date(int(m[1]), int(m[2]), int(m[3]))
    return None


WK = "월화수목금토일"


def kdate(d: dt.date | None, wd=True) -> str:
    if not d:
        return "[TBD]"
    return f"'{d:%y}.{d.month:02d}.{d.day:02d}" + (f"({WK[d.weekday()]})" if wd else "")


# ── 숫자 ───────────────────────────────────────────────────────────────
def num(v, default=None):
    if v is None or v == "":
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s not in ("", "-", ".") else default
    except ValueError:
        return default


def fmt(v, d=0, unit="") -> str:
    if v is None:
        return "[확인 필요]"
    return f"{v:,.{d}f}{unit}"


PYEONG = 3.305785   # 1평 = 3.305785㎡


def to_py(m2):
    return None if m2 is None else m2 / PYEONG


def dd_label(kind: str) -> str:
    """실사구분 표시명: 감정 → 감정평가, 그 외 → OO실사."""
    return "감정평가" if kind == "감정" else f"{kind}실사"
