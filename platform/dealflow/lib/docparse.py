# -*- coding: utf-8 -*-
"""토지대장·건축물대장·등기부등본 자동추출.

1순위: PDF 글자층에서 정규식 추출(정부24·세움터·인터넷등기소 전자발급본은 글자층이 있음).
2순위: 스캔본(글자층 없음)이고 settings.txt 에 USE_CLAUDE=1 + ANTHROPIC_API_KEY 가 있으면 Claude가 PDF를 읽어 JSON으로 답한다.
       (자료가 Anthropic API로 전송되므로 회사 정책 확인 후 켠다.)
3순위: 못 읽은 항목은 [확인 필요] — 딜입력.xlsx '입력값'에 직접 적는다.
"""
from __future__ import annotations

import base64
import json
import os
import pathlib
import re

FIELDS = {  # key: 설명 (Claude 추출 스키마에도 사용)
    "jibun": "대지 소재지와 지번 전체(예: 경기도 안양시 동안구 관양동 934 외 2필지)",
    "land_m2": "대지면적 합계(㎡, 숫자)",
    "gfa_m2": "연면적(㎡, 숫자)",
    "floors": "규모 '지하n/지상n' → 'B2/8F' 형식",
    "structure": "주구조",
    "zoning": "용도지역",
    "use": "주용도",
    "approval": "사용승인일 YYYY-MM-DD",
    "owner": "현재 소유자(등기부 갑구 최종 소유자)",
    "encumbrance": "을구에 남아 있는 근저당권 등 제한물권 요약(권리자·채권최고액). 말소된 것은 제외",
}


def pdf_text(p: pathlib.Path) -> str:
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    with fitz.open(str(p)) as d:
        return "\n".join(pg.get_text() for pg in d)


def kind_of(name: str) -> str | None:
    n = name.replace(" ", "")
    if "건축물대장" in n or "표제부" in n:
        return "building"
    if "토지대장" in n or "임야대장" in n:
        return "land"
    if "등기" in n:
        return "registry"
    return None


def _n(s):
    return float(s.replace(",", "")) if s else None


def parse_building(t: str) -> dict:
    out = {}
    flat = re.sub(r"[ \t]+", " ", t)
    m = re.search(r"대지위치\s*(.+?)(?=\s*지\s*번|\n|$)", flat)
    jb = re.search(r"지\s*번\s*([0-9\-]+(?:\s*외\s*\d+\s*필지)?)", flat)
    if m:
        out["jibun"] = (m.group(1).strip() + (" " + jb.group(1) if jb else "")).strip()
    for key, pat in [("land_m2", r"대지면적\s*[:：]?\s*([\d,]+\.?\d*)"),
                     ("gfa_m2", r"연\s*면\s*적\s*[:：]?\s*([\d,]+\.?\d*)")]:
        m = re.search(pat, flat)
        if m:
            out[key] = _n(m.group(1))
    m = re.search(r"주\s*구\s*조\s*[:：]?\s*([^\n]+?)(?:\s{2,}|\n|주\s*용\s*도)", flat)
    if m:
        out["structure"] = m.group(1).strip()
    m = re.search(r"주\s*용\s*도\s*[:：]?\s*([^\n]+?)(?:\s{2,}|\n|층\s*수)", flat)
    if m:
        out["use"] = m.group(1).strip()
    b = re.search(r"지하\s*[:：]?\s*(\d+)\s*층", flat)
    a = re.search(r"지상\s*[:：]?\s*(\d+)\s*층", flat)
    if a:
        out["floors"] = (f"B{b.group(1)}/" if b and b.group(1) != "0" else "") + f"{a.group(1)}F"
    m = re.search(r"사용승인일\s*[:：]?\s*(\d{4})[.\-년\s]+(\d{1,2})[.\-월\s]+(\d{1,2})", flat)
    if m:
        out["approval"] = f"{m[1]}-{int(m[2]):02d}-{int(m[3]):02d}"
    m = re.search(r"(제\d종\s*\S*주거지역|\S*상업지역|\S*공업지역|\S*녹지지역|관리지역)", flat)
    if m:
        out["zoning"] = m.group(1)
    return out


def parse_land(t: str) -> dict:
    out = {}
    flat = re.sub(r"[ \t]+", " ", t)
    m = re.search(r"토지\s*소재\s*(.+?)(?=\s*지\s*번|\n|$)", flat)
    jb = re.search(r"지\s*번\s*([0-9\-]+)", flat)
    if m:
        out["jibun"] = (m.group(1).strip() + (" " + jb.group(1) if jb else "")).strip()
    areas = [_n(x) for x in re.findall(r"면적\s*\(?㎡\)?\s*([\d,]+\.?\d*)", flat)]
    if areas:
        out["land_m2"] = areas[-1]           # 최종 이동(변동) 후 면적
    m = re.findall(r"지\s*목\s*[:：]?\s*\(?\d*\)?\s*([가-힣]+)", flat)
    if m:
        out["land_use"] = m[-1]
    return out


def parse_registry(t: str) -> dict:
    out = {}
    owners = re.findall(r"소유자\s+([^\s]+(?:\s?주식회사|\s?\(주\))?)", t)
    if owners:
        out["owner"] = owners[-1].strip()
    liens = []
    for m in re.finditer(r"채권최고액\s*금\s*([\d,]+)\s*원[\s\S]{0,200}?근저당권자\s+([^\s]+)", t):
        liens.append(f"근저당 {m.group(2)} 채권최고액 {int(m.group(1).replace(',', '')) / 1e8:,.1f}억원")
    if liens:
        out["encumbrance"] = "; ".join(liens) + " [말소 여부 확인 필요]"
    return out


PARSERS = {"building": parse_building, "land": parse_land, "registry": parse_registry}


def _claude_extract(files: list[pathlib.Path], settings: dict) -> dict:
    """스캔본을 Claude로 읽기(옵트인). 실패하면 빈 dict."""
    if settings.get("USE_CLAUDE") != "1":
        return {}
    if settings.get("ANTHROPIC_API_KEY"):
        os.environ.setdefault("ANTHROPIC_API_KEY", settings["ANTHROPIC_API_KEY"])
    try:
        import anthropic
    except ImportError:
        return {}
    content = []
    for p in files[:6]:
        data = p.read_bytes()
        if len(data) > 25 * 1024 * 1024:
            continue
        if p.suffix.lower() == ".pdf":
            content.append({"type": "document", "title": p.name,
                            "source": {"type": "base64", "media_type": "application/pdf",
                                       "data": base64.standard_b64encode(data).decode()}})
        elif p.suffix.lower() in (".jpg", ".jpeg", ".png"):
            mt = "image/png" if p.suffix.lower() == ".png" else "image/jpeg"
            content.append({"type": "image", "source": {"type": "base64", "media_type": mt,
                                                         "data": base64.standard_b64encode(data).decode()}})
    if not content:
        return {}
    props = {k: {"type": "string", "description": v} for k, v in FIELDS.items()}
    content.append({"type": "text", "text": "첨부한 한국 토지대장·건축물대장·등기사항전부증명서에서 아래 항목을 읽어 JSON으로 답하세요. "
                                            "문서에 없거나 읽을 수 없는 항목은 빈 문자열로 둡니다. 숫자는 쉼표 없이 씁니다."})
    try:
        client = anthropic.Anthropic()
        resp = client.beta.messages.create(
            model=settings.get("CLAUDE_MODEL", "claude-opus-5-5"), max_tokens=16000,
            betas=["server-side-fallback-2026-07-01"], fallbacks="default",
            output_config={"effort": "medium", "format": {"type": "json_schema", "schema": {
                "type": "object", "properties": props, "required": list(FIELDS), "additionalProperties": False}}},
            messages=[{"role": "user", "content": content}])
        if resp.stop_reason != "end_turn":
            return {}
        text = next(b.text for b in resp.content if b.type == "text")
        data = json.loads(text)
    except Exception as e:  # 네트워크·인증 오류는 자동추출만 건너뛴다
        print(f"  [안내] Claude 자동추출을 건너뜁니다: {type(e).__name__}: {e}")
        return {}
    out = {}
    for k, v in data.items():
        if v not in ("", None):
            out[k] = float(v) if k in ("land_m2", "gfa_m2") and re.fullmatch(r"[\d.]+", str(v)) else v
    return out


def extract(folder: pathlib.Path, settings: dict | None = None) -> tuple[dict, list[str]]:
    """반환: ({key: (값, 출처)}, 안내 메시지 목록)."""
    settings = settings or {}
    auto, notes, scanned = {}, [], []
    if not folder.exists():
        return auto, [f"{folder} 폴더가 없습니다"]
    for p in sorted(folder.iterdir()):
        k = kind_of(p.name)
        if not k:
            continue
        if p.suffix.lower() != ".pdf":
            scanned.append(p)
            continue
        t = pdf_text(p)
        if len(t.strip()) < 30:
            scanned.append(p)
            notes.append(f"{p.name}: 글자층 없는 스캔본")
            continue
        got = PARSERS[k](t)
        for key, v in got.items():
            auto.setdefault(key, (v, p.name))
        notes.append(f"{p.name}: {len(got)}개 항목 추출 ({', '.join(got) or '없음'})")
    if scanned:
        got = _claude_extract(scanned, settings)
        if got:
            for key, v in got.items():
                auto.setdefault(key, (v, "Claude 판독: " + ", ".join(x.name for x in scanned)))
            notes.append(f"Claude 판독 {len(got)}개 항목 [원본 대조 필요]")
        elif settings.get("USE_CLAUDE") != "1":
            notes.append("스캔본은 자동추출하지 못했습니다 → 딜입력.xlsx 자산개요 '입력값'에 직접 적거나, "
                         "settings.txt 에 USE_CLAUDE=1 과 API 키를 넣으면 Claude가 읽습니다")
    return auto, notes
