# -*- coding: utf-8 -*-
"""로컬 SQLite 저장소(data/app.db). 첫 실행 때 안양물류센터 작업 데이터로 예시 행을 넣는다."""
import json
import re
import sqlite3
from contextlib import contextmanager

from .config import DATA, ROOT

DB = DATA / "app.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS deals (id INTEGER PRIMARY KEY, name TEXT, asset_type TEXT, address TEXT, lat REAL, lng REAL,
  size TEXT, price TEXT, advisor TEXT, marketing TEXT, stage TEXT, deadline TEXT, notes TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS events (id INTEGER PRIMARY KEY, project TEXT, title TEXT, date TEXT, kind TEXT, done INTEGER DEFAULT 0, notes TEXT);
CREATE TABLE IF NOT EXISTS qa (id INTEGER PRIMARY KEY, project TEXT, institution TEXT, category TEXT, question TEXT, answer TEXT,
  status TEXT, asked TEXT);
CREATE TABLE IF NOT EXISTS rfps (id INTEGER PRIMARY KEY, project TEXT, category TEXT, title TEXT, scope TEXT, due TEXT, status TEXT);
CREATE TABLE IF NOT EXISTS quotes (id INTEGER PRIMARY KEY, rfp_id INTEGER, vendor TEXT, amount TEXT, period TEXT, notes TEXT, selected INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS templates (id INTEGER PRIMARY KEY, name TEXT, category TEXT, subject TEXT, body TEXT);
CREATE TABLE IF NOT EXISTS generations (id INTEGER PRIMARY KEY, module TEXT, subtype TEXT, title TEXT, inputs TEXT, output TEXT,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP);
"""

TABLES = {  # 화면에서 편집 가능한 표와 열
    "deals": ["name", "asset_type", "address", "lat", "lng", "size", "price", "advisor", "marketing", "stage", "deadline", "notes"],
    "events": ["project", "title", "date", "kind", "done", "notes"],
    "qa": ["project", "institution", "category", "question", "answer", "status", "asked"],
    "rfps": ["project", "category", "title", "scope", "due", "status"],
    "quotes": ["rfp_id", "vendor", "amount", "period", "notes", "selected"],
    "templates": ["name", "category", "subject", "body"],
}


@contextmanager
def conn():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    try:
        yield c
        c.commit()
    finally:
        c.close()


def rows(table: str, where: str = "", args: tuple = ()) -> list[dict]:
    with conn() as c:
        return [dict(r) for r in c.execute(f"SELECT * FROM {table} {where} ORDER BY id", args)]


def upsert(table: str, data: dict) -> int:
    cols = [k for k in TABLES[table] if k in data]
    with conn() as c:
        if data.get("id"):
            c.execute(f"UPDATE {table} SET {', '.join(f'{k}=?' for k in cols)} WHERE id=?", [data[k] for k in cols] + [data["id"]])
            return int(data["id"])
        cur = c.execute(f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})", [data[k] for k in cols])
        return cur.lastrowid


def delete(table: str, rid: int) -> None:
    with conn() as c:
        c.execute(f"DELETE FROM {table} WHERE id=?", (rid,))


def save_generation(module: str, subtype: str, title: str, inputs: dict, output: str) -> int:
    with conn() as c:
        cur = c.execute("INSERT INTO generations (module, subtype, title, inputs, output) VALUES (?,?,?,?,?)",
                        (module, subtype, title, json.dumps(inputs, ensure_ascii=False), output))
        return cur.lastrowid


def init() -> None:
    with conn() as c:
        c.executescript(SCHEMA)
        if c.execute("SELECT COUNT(*) FROM deals").fetchone()[0] == 0:
            _seed(c)


def _timeline_events() -> list[tuple]:
    """projects/ANYANG-LOGIS/07_일정_회의록/TIMELINE.md 표에서 일자·내용을 읽는다(취소선 행 제외)."""
    p = ROOT / "projects/ANYANG-LOGIS/07_일정_회의록/TIMELINE.md"
    out = []
    if not p.exists():
        return out
    section = ""
    for line in p.read_text(encoding="utf-8").splitlines():
        if line.startswith("## "):
            section = line[3:5]
        if section not in ("1.", "2."):  # '1. 완료', '2. 진행 중 및 예정' 표만 사용(변경 이력 등 제외)
            continue
        m = re.match(r"\|\s*(20\d\d-\d\d-\d\d)[^|]*\|(.*)\|\s*$", line)
        if not m or "~~" in line.split("|")[1]:
            continue
        cells = [re.sub(r"\*\*|~~", "", x).strip() for x in m.group(2).split("|")]
        title = " / ".join(x for x in cells if x and x not in ("[CP]",))[:200]
        if title:
            out.append(("ANYANG-LOGIS", title, m.group(1), "마일스톤", 1 if m.group(1) <= "2026-10-29" else 0, "TIMELINE.md에서 가져옴"))
    return out


def _seed(c) -> None:
    c.execute("INSERT INTO deals (name, asset_type, address, lat, lng, size, price, advisor, marketing, stage, deadline, notes) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
              ("안양물류센터 (ANYANG-LOGIS)", "물류(상온·라스트마일)", "경기도 안양시 동안구 흥안대로439번길 30 (관양동 934 외 2필지)",
               37.4003, 126.9668, "연면적 95,475㎡ / B2~8F", "2,700억원 (Cap 4.9%)", "JLL(매각자문)", "입찰공고·TM 배포 2026-07-16 → 우선협상 2026-08-18",
               "거래종결", "2026-10-29", "임대율 99.3%·WALE 6.7년, 쿠팡 등. 좌표는 대략값[추정] — 지도에서 클릭해 수정"))
    c.executemany("INSERT INTO events (project, title, date, kind, done, notes) VALUES (?,?,?,?,?,?)", _timeline_events())
    c.executemany("INSERT INTO qa (project, institution, category, question, answer, status, asked) VALUES (?,?,?,?,?,?,?)", [
        ("ANYANG-LOGIS", "우리은행(대주)", "담보", "신탁계약 특약 제15조②: 신규 임차인 보증금이 선순위 대주보다 먼저 변제되는가?",
         "특약상 보증금 반환채권의 우선 순위 해석과 근거 조문을 정리해 회신 [안양 사례]", "답변완료", "2026-10-20"),
        ("ANYANG-LOGIS", "대주단", "선행조건", "농심 임대차 승계동의서 미징구 — 인출에 문제가 없는가?",
         "농심 임대차계약 제21조(임차인은 임대인 변경에 조건 없이 동의, 요청 시 승계동의서 제공) + 상가임대차법 제3조②. 기표일 기준 제출 예정", "답변완료", "2026-10-28"),
        ("ANYANG-LOGIS", "삼성증권(투자자)", "세무", "인수수수료 부가가치세 과세 여부",
         "증권사 주선·잔액인수 용역 → 금융용역 면세 검토(부가세법 제26조①11호). 계산서 발행 여부 확인 필요", "확인중", "2026-10-28"),
    ])
    c.execute("INSERT INTO rfps (project, category, title, scope, due, status) VALUES (?,?,?,?,?,?)",
              ("ANYANG-LOGIS", "회계", "사업타당성검토보고서 (영업인가용)", "부동산투자회사법상 영업인가 첨부 사업타당성 검토", "2026-10-20", "발주완료"))
    c.executemany("INSERT INTO templates (name, category, subject, body) VALUES (?,?,?,?)", [
        ("대출금 입금 요청(정오까지)", "대주",
         "[{{프로젝트}} 담보대출] {{인출일}} 대출금 입금 요청 – {{대주기관명}}",
         "{{대주기관명}} {{담당자}}님, 안녕하세요.\n코람코자산신탁 {{발신자}}입니다.\n\n{{약정일}} 체결한 「{{대출명}}」 대출약정서와 관련하여, 인출요청서와 대출금 입금 요청 공문을 첨부와 같이 송부드립니다.\n\n인출일 당일 매매대금 지급 등 후속 일정이 있어 {{인출일}} 정오(12:00)까지 아래 계좌로 입금하여 주시기를 부탁드립니다.\n\n- 입금금액: 금 {{입금금액}}원 (Tranche {{트랜치}})\n- 입금계좌: {{입금계좌}}\n- 예금주: {{예금주}}\n\n입금 후 회신 주시면 바로 확인하겠습니다.\n\n감사합니다.\n{{발신자}} 드림"),
        ("동시이행 서류(등기 접수증) 제출 안내", "대주",
         "[{{프로젝트}}] 인출 동시이행 서류 제출 – 등기 접수증",
         "{{대주기관명}} {{담당자}}님,\n\n대출약정서 부록 2(인출일 당일 제공 서류)에 따라 아래 접수증 사본을 송부드립니다.\n\n{{제출서류}}\n\n{{미제출서류}}는 법무법인으로부터 수령하는 대로 추가 제출하겠습니다.\n\n감사합니다.\n{{발신자}} 드림"),
        ("투자자 계약서 최종안 회람", "투자자",
         "[{{프로젝트}}] {{계약서명}} 최종안 송부",
         "{{투자자명}} {{담당자}}님, 안녕하세요.\n\n{{계약서명}} 최종안을 첨부와 같이 송부드립니다. {{회신기한}}까지 최종 의견을 회신 부탁드립니다.\n\n주요 변경사항\n{{변경사항}}\n\n감사합니다.\n{{발신자}} 드림"),
    ])
