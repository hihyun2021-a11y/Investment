# -*- coding: utf-8 -*-
"""신규투자 업무 지원 웹앱(로컬) — FastAPI 서버.
실행: platform/webapp/run.sh (또는 python -m uvicorn app.main:app --port 8765)"""
import base64
import csv
import glob
import io
import json
import re
import secrets
import zipfile
from email.message import EmailMessage
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles

from . import config, db
from .claude import build_content, stream_generate
from .exporter import md_to_docx
from .modules import MODULES, get_subtype

app = FastAPI(title="신규투자 워크벤치")
STATIC = config.APP_DIR / "static"
SAMPLES = config.DATA / "samples"
db.init()


@app.middleware("http")
async def password_gate(request: Request, call_next):
    """APP_PASSWORD 를 설정하면 브라우저 기본 인증(아이디는 아무거나)으로 접속을 막는다 — 외부 접속 단계용."""
    if config.APP_PASSWORD:
        ok = False
        auth = request.headers.get("authorization", "")
        if auth.startswith("Basic "):
            try:
                _, pw = base64.b64decode(auth[6:]).decode().split(":", 1)
                ok = secrets.compare_digest(pw, config.APP_PASSWORD)
            except Exception:
                ok = False
        if not ok:
            return Response(status_code=401, headers={"WWW-Authenticate": 'Basic realm="workbench"'})
    return await call_next(request)


# ---------------------------------------------------------------- 기본 정보
@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")


@app.get("/api/config")
def get_config():
    mods = [{k: v for k, v in m.items()} | {"subtypes": [{k: v for k, v in s.items() if k != "samples"} for s in m["subtypes"]]}
            for m in MODULES]
    return {"modules": mods, "model": config.MODEL, "google_maps": bool(config.GOOGLE_MAPS_API_KEY)}


# ---------------------------------------------------------------- 샘플 라이브러리
def _sample_list(module: str, subtype: str) -> list[dict]:
    _, st = get_subtype(module, subtype)
    out, seen = [], set()
    for pat in st.get("samples", []):
        for p in sorted(glob.glob(str(config.ROOT / pat))):
            pp = Path(p)
            if pp.is_file() and p not in seen:
                seen.add(p)
                out.append({"id": "b:" + str(pp.relative_to(config.ROOT)), "name": pp.name, "source": "기본(안양·표준서식)"})
    folder = SAMPLES / module / subtype
    if folder.exists():
        for pp in sorted(folder.iterdir()):
            out.append({"id": "u:" + str(pp.relative_to(SAMPLES)), "name": pp.name, "source": "내가 올린 샘플"})
    return out


def _resolve(sid: str) -> Path:
    kind, rel = sid.split(":", 1)
    base = config.ROOT if kind == "b" else SAMPLES
    p = (base / rel).resolve()
    if not str(p).startswith(str(base.resolve())) or not p.is_file():
        raise HTTPException(400, f"샘플을 찾을 수 없습니다: {rel}")
    return p


def _safe_name(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", Path(name).name) or "file"


@app.get("/api/samples")
def list_samples(module: str, subtype: str):
    return _sample_list(module, subtype)


@app.post("/api/samples")
async def add_samples(module: str = Form(...), subtype: str = Form(...), files: list[UploadFile] = File(...)):
    get_subtype(module, subtype)
    folder = SAMPLES / module / subtype
    folder.mkdir(parents=True, exist_ok=True)
    for f in files:
        (folder / _safe_name(f.filename)).write_bytes(await f.read())
    return _sample_list(module, subtype)


@app.delete("/api/samples")
def remove_sample(id: str):
    if not id.startswith("u:"):
        raise HTTPException(400, "기본 샘플은 삭제할 수 없습니다.")
    _resolve(id).unlink()
    return {"ok": True}


# ---------------------------------------------------------------- 생성
def _include_text(kind: str, fields: dict) -> str:
    if kind == "qa":
        inst = fields.get("institution", "").strip()
        qs = db.rows("qa", "WHERE institution LIKE ?", (f"%{inst}%",)) if inst else db.rows("qa")
        lines = ["[Q&A DB]", "| # | 기관 | 구분 | 질문 | 답변 | 상태 | 일자 |", "|---|---|---|---|---|---|---|"]
        lines += [f"| {q['id']} | {q['institution']} | {q['category']} | {q['question']} | {q['answer'] or ''} | {q['status']} | {q['asked']} |" for q in qs]
        return "\n".join(lines)
    if kind == "quotes":
        rid = re.sub(r"\D", "", str(fields.get("rfp_id", "")))
        if not rid:
            return "[견적 DB] RFP 번호가 입력되지 않았습니다."
        rfp = db.rows("rfps", "WHERE id=?", (int(rid),))
        qs = db.rows("quotes", "WHERE rfp_id=?", (int(rid),))
        lines = ["[견적 DB]", f"RFP: {json.dumps(rfp[0], ensure_ascii=False) if rfp else '없음'}",
                 "| 업체 | 금액 | 기간 | 비고 |", "|---|---|---|---|"]
        lines += [f"| {q['vendor']} | {q['amount']} | {q['period']} | {q['notes'] or ''} |" for q in qs]
        return "\n".join(lines)
    return ""


@app.post("/api/generate")
async def generate(module: str = Form(...), subtype: str = Form(...), fields: str = Form("{}"), sample_ids: str = Form("[]"),
                   extra: str = Form(""), web_search: str = Form("0"), save_samples: str = Form("0"),
                   files: list[UploadFile] = File(default=[])):
    mod, st = get_subtype(module, subtype)
    fields_d = json.loads(fields or "{}")
    samples = [{"name": _resolve(s).name, "path": _resolve(s)} for s in json.loads(sample_ids or "[]")]
    if files:
        folder = (SAMPLES / module / subtype) if save_samples == "1" else (config.DATA / "outputs" / "_uploads")
        folder.mkdir(parents=True, exist_ok=True)
        for f in files:
            if not f.filename:
                continue
            p = folder / _safe_name(f.filename)
            p.write_bytes(await f.read())
            samples.append({"name": p.name, "path": p})
    content, warnings = build_content(mod, st, fields_d, samples, extra, _include_text(st.get("include", ""), fields_d))
    use_web = web_search == "1" and mod.get("web_search", False)

    def events():
        if warnings:
            yield f"data: {json.dumps({'warn': warnings}, ensure_ascii=False)}\n\n"
        out = []
        for chunk in stream_generate(content, use_web):
            out.append(chunk)
            yield f"data: {json.dumps({'t': chunk}, ensure_ascii=False)}\n\n"
        title = f"{st['name']} — {fields_d.get('project') or fields_d.get('purpose') or ''}".strip(" —")
        gid = db.save_generation(module, subtype, title, {"fields": fields_d, "samples": [s["name"] for s in samples], "extra": extra}, "".join(out))
        yield f"data: {json.dumps({'done': True, 'id': gid})}\n\n"

    return StreamingResponse(events(), media_type="text/event-stream", headers={"Cache-Control": "no-cache"})


@app.get("/api/generations")
def generations():
    with db.conn() as c:
        return [dict(r) for r in c.execute("SELECT id, module, subtype, title, created_at FROM generations ORDER BY id DESC LIMIT 200")]


@app.get("/api/generations/{gid}")
def generation(gid: int):
    r = db.rows("generations", "WHERE id=?", (gid,))
    if not r:
        raise HTTPException(404)
    return r[0]


@app.put("/api/generations/{gid}")
async def update_generation(gid: int, request: Request):
    body = await request.json()
    with db.conn() as c:
        c.execute("UPDATE generations SET output=? WHERE id=?", (body.get("output", ""), gid))
    return {"ok": True}


def _fname(g: dict, ext: str) -> str:
    date = (g["created_at"] or "")[:10].replace("-", "")
    return f"{date}_{_safe_name(g['title'] or 'result')}.{ext}"


@app.get("/api/generations/{gid}/docx")
def generation_docx(gid: int):
    g = generation(gid)
    return Response(md_to_docx(g["output"]), media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    headers={"Content-Disposition": f"attachment; filename*=UTF-8''{_quote(_fname(g, 'docx'))}"})


@app.get("/api/generations/{gid}/md")
def generation_md(gid: int):
    g = generation(gid)
    return Response(g["output"].encode("utf-8"), media_type="text/markdown",
                    headers={"Content-Disposition": f"attachment; filename*=UTF-8''{_quote(_fname(g, 'md'))}"})


def _quote(s: str) -> str:
    from urllib.parse import quote
    return quote(s)


# ---------------------------------------------------------------- 데이터 표(딜·일정·Q&A·발주·견적·양식)
@app.get("/api/t/{table}")
def table_rows(table: str):
    if table not in db.TABLES:
        raise HTTPException(404)
    return db.rows(table)


@app.post("/api/t/{table}")
async def table_upsert(table: str, request: Request):
    if table not in db.TABLES:
        raise HTTPException(404)
    return {"id": db.upsert(table, await request.json())}


@app.delete("/api/t/{table}/{rid}")
def table_delete(table: str, rid: int):
    if table not in db.TABLES:
        raise HTTPException(404)
    db.delete(table, rid)
    return {"ok": True}


@app.get("/api/events.ics")
def events_ics():
    """휴대폰·아웃룩 캘린더로 가져오기용 iCalendar 파일."""
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//workbench//KO", "CALSCALE:GREGORIAN"]
    for e in db.rows("events"):
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", e["date"] or ""):
            continue
        d = e["date"].replace("-", "")
        title = f"[{e['project']}] {e['title']}".replace("\n", " ").replace(",", "\\,")
        lines += ["BEGIN:VEVENT", f"UID:wb-{e['id']}@workbench", f"DTSTART;VALUE=DATE:{d}", f"SUMMARY:{title}",
                  f"CATEGORIES:{e['kind'] or ''}", "END:VEVENT"]
    lines.append("END:VCALENDAR")
    return Response("\r\n".join(lines), media_type="text/calendar",
                    headers={"Content-Disposition": "attachment; filename=workbench_events.ics"})


# ---------------------------------------------------------------- 배포(이메일 양식 채우기)
def _fill(text: str, values: dict) -> str:
    return re.sub(r"\{\{\s*([^}]+?)\s*\}\}", lambda m: str(values.get(m.group(1), m.group(0))), text)


def _fill_rows(template_id: int, rows_: list[dict]) -> list[dict]:
    t = db.rows("templates", "WHERE id=?", (template_id,))
    if not t:
        raise HTTPException(404, "양식이 없습니다.")
    t = t[0]
    return [{"to": r.get("이메일") or r.get("email") or "", "cc": r.get("참조") or "",
             "subject": _fill(t["subject"], r), "body": _fill(t["body"], r)} for r in rows_]


@app.post("/api/fill")
async def fill(request: Request):
    body = await request.json()
    return _fill_rows(int(body["template_id"]), body.get("rows", []))


@app.post("/api/fill_csv")
async def fill_csv(template_id: int = Form(...), file: UploadFile = File(...)):
    raw = (await file.read()).decode("utf-8-sig", errors="ignore")
    return _fill_rows(template_id, list(csv.DictReader(io.StringIO(raw))))


@app.post("/api/eml_zip")
async def eml_zip(request: Request):
    """채워진 메일들을 .eml 파일(아웃룩에서 열어 바로 발송 가능)로 묶어 내려준다."""
    mails = (await request.json()).get("mails", [])
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for i, m in enumerate(mails, 1):
            msg = EmailMessage()
            msg["To"] = m.get("to", "")
            if m.get("cc"):
                msg["Cc"] = m["cc"]
            msg["Subject"] = m.get("subject", "")
            msg["X-Unsent"] = "1"  # 아웃룩에서 '보내기 전' 상태로 열림
            msg.set_content(m.get("body", ""))
            z.writestr(f"{i:02d}_{_safe_name(m.get('subject', 'mail'))[:60]}.eml", msg.as_bytes())
    return Response(buf.getvalue(), media_type="application/zip",
                    headers={"Content-Disposition": "attachment; filename=mails.zip"})


# ---------------------------------------------------------------- 프로젝트 대시보드
@app.get("/api/projects")
def projects():
    out = []
    for p in sorted((config.ROOT / "projects").glob("*/PROJECT.md")):
        text = p.read_text(encoding="utf-8")
        title = next((l.lstrip("# ").strip() for l in text.splitlines() if l.startswith("#")), p.parent.name)
        stage = re.search(r"(단계|Stage)[^\n|]*\|\s*([^|\n]+)", text)
        out.append({"code": p.parent.name, "title": title, "stage": stage.group(2).strip() if stage else ""})
    return out


@app.get("/api/projects/{code}")
def project(code: str):
    if not re.fullmatch(r"[A-Za-z0-9_\-]+", code):
        raise HTTPException(400)
    p = config.ROOT / "projects" / code / "PROJECT.md"
    if not p.exists():
        raise HTTPException(404)
    return {"code": code, "markdown": p.read_text(encoding="utf-8")}


app.mount("/static", StaticFiles(directory=STATIC), name="static")
