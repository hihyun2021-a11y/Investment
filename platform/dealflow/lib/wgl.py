# -*- coding: utf-8 -*-
"""Working Group List: 엑셀(기관별 정리본) + 주소록 HTML(검색·클릭 시 연락처 팝업·선택 주소 복사).

주소록 HTML은 인터넷 연결 없이 열리는 단일 파일이다(데이터가 파일 안에 들어 있으므로 외부로 보내지 말 것).
"""
from __future__ import annotations

import html
import json
import pathlib

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .common import TODAY, outpath
from .inputbook import WGL_COLS, WGL_GROUPS

FONT = "맑은 고딕"
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def ordered(wgl):
    rank = {g: i for i, g in enumerate(WGL_GROUPS)}
    return sorted(wgl, key=lambda r: (rank.get(r["구분"], 99), r["기관명"], r["성명"]))


def write_xlsx(d, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "WGL"
    name = d["asset"].get("name", "")
    ws.append([f"{name} Working Group List ({TODAY:%Y.%m.%d} 기준)"])
    ws["A1"].font = Font(name=FONT, bold=True, size=12)
    ws.append([])
    ws.append(WGL_COLS)
    for c in ws[3]:
        c.fill = PatternFill("solid", fgColor="D9D9D9")
        c.font, c.border = Font(name=FONT, bold=True, size=10), BOX
        c.alignment = Alignment(horizontal="center")
    prev = None
    for r in ordered(d["wgl"]):
        if r["구분"] != prev:
            ws.append([r["구분"] or "기타"])
            row = ws.max_row
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(WGL_COLS))
            ws.cell(row, 1).fill = PatternFill("solid", fgColor="345B86")
            ws.cell(row, 1).font = Font(name=FONT, bold=True, size=10, color="FFFFFF")
            prev = r["구분"]
        ws.append([r[c] for c in WGL_COLS])
        for c in ws[ws.max_row]:
            c.font, c.border = Font(name=FONT, size=10), BOX
    for i, w in enumerate((16, 22, 20, 10, 9, 16, 30, 15, 15, 20), 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    # 메일 주소 묶음(아웃룩 받는사람 칸에 붙여넣기용)
    ws2 = wb.create_sheet("메일주소묶음")
    ws2.append(["구분", "기관", "주소(; 구분)"])
    groups = {}
    for r in d["wgl"]:
        if r["이메일"]:
            groups.setdefault((r["구분"], r["기관명"]), []).append(r["이메일"])
    for (g, org), ms in sorted(groups.items()):
        ws2.append([g, org, "; ".join(ms)])
    for g in WGL_GROUPS:
        ms = [r["이메일"] for r in d["wgl"] if r["구분"] == g and r["이메일"]]
        if ms:
            ws2.append([g, "(구분 전체)", "; ".join(ms)])
    for c, w in zip("ABC", (16, 24, 100)):
        ws2.column_dimensions[c].width = w
    wb.save(path)
    return path


PAGE = """<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITLE__</title><style>
:root{--band:#345B86;--ink:#1F1F1F;--sub:#646464;--line:#BFBFBF;--th:#D9D9D9;--c1:#F2F2F2}
*{box-sizing:border-box}body{margin:0;font-family:"맑은 고딕","Malgun Gothic",sans-serif;color:var(--ink);background:#fff;font-size:14px}
header{background:var(--band);color:#fff;padding:12px 16px}header h1{margin:0;font-size:17px}header p{margin:4px 0 0;font-size:12px;opacity:.85}
.bar{display:flex;flex-wrap:wrap;gap:8px;padding:12px 16px;border-bottom:1px solid var(--line);position:sticky;top:0;background:#fff}
input[type=search]{flex:1 1 260px;padding:8px 10px;border:1px solid var(--line);border-radius:4px;font-size:14px}
button{padding:7px 12px;border:1px solid var(--band);background:#fff;color:var(--band);border-radius:4px;cursor:pointer;font-size:13px}
button.on,button.pri{background:var(--band);color:#fff}.chips{display:flex;flex-wrap:wrap;gap:6px;padding:8px 16px}
table{border-collapse:collapse;width:100%}th,td{border-bottom:1px solid #e3e3e3;padding:7px 8px;text-align:left;vertical-align:middle}
th{background:var(--th);font-size:13px}tr.g td{background:var(--band);color:#fff;font-weight:bold}
td.org{background:var(--c1)}tr.p{cursor:pointer}tr.p:hover td{background:#EAF1F8}.wrap{padding:0 16px 40px;overflow-x:auto}
.cnt{font-size:12px;color:var(--sub);align-self:center}
#m{position:fixed;inset:0;background:rgba(0,0,0,.35);display:none;align-items:center;justify-content:center;padding:16px}
#m .box{background:#fff;border-radius:6px;max-width:440px;width:100%;padding:18px 20px;box-shadow:0 6px 24px rgba(0,0,0,.25)}
#m h2{margin:0 0 4px;font-size:18px}#m .sub{color:var(--sub);font-size:13px;margin-bottom:12px}
#m dl{display:grid;grid-template-columns:72px 1fr;gap:6px 10px;margin:0 0 14px}#m dt{color:var(--sub)}#m dd{margin:0;word-break:break-all}
#m .act{display:flex;flex-wrap:wrap;gap:8px}#toast{position:fixed;bottom:18px;left:50%;transform:translateX(-50%);background:#333;color:#fff;padding:8px 14px;border-radius:4px;display:none}
@media(max-width:700px){.hide-s{display:none}}
</style></head><body>
<header><h1>__TITLE__</h1><p>이름·기관·역할로 검색 → 이름을 누르면 연락처 팝업. 체크한 사람의 메일 주소를 한 번에 복사해 아웃룩 받는사람 칸에 붙여넣으세요. (대외비)</p></header>
<div class="bar"><input id="q" type="search" placeholder="검색: 예) 우리은행, 법무, 김" autofocus>
<button id="all">보이는 사람 모두 선택</button><button id="none">선택 해제</button><button class="pri" id="copy">선택 주소 복사</button>
<button id="mail">선택 대상 메일 쓰기</button><span class="cnt" id="cnt"></span></div>
<div class="chips" id="chips"></div>
<div class="wrap"><table><thead><tr><th style="width:30px"></th><th>기관</th><th>성명</th><th class="hide-s">직급·부서</th><th class="hide-s">역할</th><th>이메일</th><th class="hide-s">휴대폰</th></tr></thead><tbody id="tb"></tbody></table></div>
<div id="m"><div class="box"><h2 id="mn"></h2><div class="sub" id="ms"></div><dl id="md"></dl><div class="act" id="ma"></div></div></div>
<div id="toast"></div>
<script>
const D=__DATA__;const G=__GROUPS__;let sel=new Set(),grp=null;
const $=s=>document.querySelector(s),esc=s=>(s||"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
function toast(t){const e=$("#toast");e.textContent=t;e.style.display="block";setTimeout(()=>e.style.display="none",1600)}
function copy(t){(navigator.clipboard?navigator.clipboard.writeText(t):Promise.reject()).then(()=>toast("복사했습니다")).catch(()=>{const a=document.createElement("textarea");a.value=t;document.body.appendChild(a);a.select();document.execCommand("copy");a.remove();toast("복사했습니다")})}
function vis(){const q=$("#q").value.trim().toLowerCase();return D.map((r,i)=>[r,i]).filter(([r])=>(!grp||r.g===grp)&&(!q||(t=>t.includes(q))((v=>v+" "+v.replace(/신용협동조합/g,"신협").replace(/주식회사|\(주\)|㈜/g,""))(Object.values(r).join(" ").toLowerCase()))))}
function chips(){$("#chips").innerHTML=["전체",...G].map(g=>`<button class="${(g==="전체"&&!grp)||g===grp?"on":""}" data-g="${esc(g)}">${esc(g)}</button>`).join("");
 document.querySelectorAll("#chips button").forEach(b=>b.onclick=()=>{grp=b.dataset.g==="전체"?null:b.dataset.g;chips();draw()})}
function draw(){let h="",prev=null;for(const [r,i] of vis()){if(r.g!==prev){h+=`<tr class="g"><td colspan="7">${esc(r.g||"기타")}</td></tr>`;prev=r.g}
 h+=`<tr class="p" data-i="${i}"><td><input type="checkbox" data-c="${i}" ${sel.has(i)?"checked":""}></td><td class="org">${esc(r.o)}</td><td><b>${esc(r.n)}</b></td><td class="hide-s">${esc([r.t,r.d].filter(Boolean).join(" · "))}</td><td class="hide-s">${esc(r.r)}</td><td>${esc(r.e)}</td><td class="hide-s">${esc(r.m)}</td></tr>`}
 $("#tb").innerHTML=h||'<tr><td colspan="7">검색 결과가 없습니다</td></tr>';$("#cnt").textContent=`선택 ${sel.size}명`;
 document.querySelectorAll("tr.p").forEach(tr=>tr.onclick=e=>{if(e.target.dataset.c!==undefined)return;pop(+tr.dataset.i)});
 document.querySelectorAll("input[data-c]").forEach(c=>c.onchange=()=>{const i=+c.dataset.c;c.checked?sel.add(i):sel.delete(i);$("#cnt").textContent=`선택 ${sel.size}명`})}
function pop(i){const r=D[i];$("#mn").textContent=`${r.n||"(이름 없음)"} ${r.t||""}`;$("#ms").textContent=[r.g,r.o,r.d].filter(Boolean).join(" · ");
 $("#md").innerHTML=[["역할",r.r],["이메일",r.e],["전화",r.p],["휴대폰",r.m],["비고",r.x]].filter(x=>x[1]).map(([k,v])=>`<dt>${k}</dt><dd>${esc(v)}</dd>`).join("");
 const a=[];if(r.e){a.push(`<button class="pri" data-a="ce">이메일 복사</button>`);a.push(`<button data-a="me">메일 쓰기</button>`)}
 if(r.m)a.push(`<button data-a="cm">휴대폰 복사</button><button data-a="tm">전화 걸기</button>`);if(r.p)a.push(`<button data-a="cp">전화 복사</button>`);
 a.push(`<button data-a="x">닫기</button>`);$("#ma").innerHTML=a.join("");$("#m").style.display="flex";
 document.querySelectorAll("#ma button").forEach(b=>b.onclick=()=>{const k=b.dataset.a;if(k==="ce")copy(r.e);if(k==="me")location.href="mailto:"+r.e;
 if(k==="cm")copy(r.m);if(k==="tm")location.href="tel:"+r.m.replace(/[^0-9+]/g,"");if(k==="cp")copy(r.p);if(k==="x")$("#m").style.display="none"})}
$("#m").onclick=e=>{if(e.target.id==="m")$("#m").style.display="none"};document.onkeydown=e=>{if(e.key==="Escape")$("#m").style.display="none"};
$("#q").oninput=draw;$("#all").onclick=()=>{vis().forEach(([r,i])=>r.e&&sel.add(i));draw()};$("#none").onclick=()=>{sel.clear();draw()};
const picked=()=>[...sel].map(i=>D[i].e).filter(Boolean).flatMap(e=>e.split(/[;,/ ]+/)).filter(e=>e.includes("@"));
$("#copy").onclick=()=>{const m=picked();m.length?copy(m.join("; ")):toast("선택한 사람이 없습니다")};
$("#mail").onclick=()=>{const m=picked();m.length?location.href="mailto:"+m.join(";"):toast("선택한 사람이 없습니다")};
chips();draw();
</script></body></html>"""


def write_html(d, path):
    rows = [{"g": r["구분"], "o": r["기관명"], "r": r["역할"], "n": r["성명"], "t": r["직급"], "d": r["부서"], "e": r["이메일"],
             "p": r["전화"], "m": r["휴대폰"], "x": r["비고"]} for r in ordered(d["wgl"])]
    groups = [g for g in WGL_GROUPS if any(r["g"] == g for r in rows)] + sorted({r["g"] for r in rows if r["g"] not in WGL_GROUPS and r["g"]})
    title = f"{d['asset'].get('name', '')} WGL 주소록"
    page = (PAGE.replace("__TITLE__", html.escape(title))
            .replace("__DATA__", json.dumps(rows, ensure_ascii=False).replace("</", "<\\/"))
            .replace("__GROUPS__", json.dumps(groups, ensure_ascii=False)))
    path.write_text(page, encoding="utf-8")
    return path


def build(d, folder: pathlib.Path) -> dict:
    name = d["asset"].get("name", "딜")
    x = write_xlsx(d, outpath(folder, "WGL", f"{name}_WorkingGroupList", "xlsx"))
    h = folder / f"{name}_WGL주소록.html"     # 즐겨찾기용으로 고정 이름(항상 최신)
    write_html(d, h)
    return {"xlsx": str(x), "html": str(h), "n": len(d["wgl"])}
