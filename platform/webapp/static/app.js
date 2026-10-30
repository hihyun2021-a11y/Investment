// 신규투자 워크벤치 — 화면 스크립트 (프레임워크 없이 작성)
const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const api = async (url, opt = {}) => {
  const r = await fetch(url, opt);
  if (!r.ok) throw new Error(`${r.status} ${await r.text()}`);
  return r.headers.get("content-type")?.includes("json") ? r.json() : r;
};
const post = (url, body) => api(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
let CFG = null;

// ---------------------------------------------------------------- Markdown 표시
function md(src) {
  const lines = String(src || "").split("\n");
  let out = "", i = 0, list = null;
  const inline = (t) => esc(t).replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/(https?:\/\/[^\s)<]+)/g, '<a href="$1" target="_blank">$1</a>');
  const close = () => { if (list) { out += `</${list}>`; list = null; } };
  while (i < lines.length) {
    const l = lines[i];
    if (/^\s*\|/.test(l)) {
      close();
      const rows = [];
      while (i < lines.length && /^\s*\|/.test(lines[i])) {
        const cells = lines[i].trim().replace(/^\||\|$/g, "").split("|").map((c) => c.trim());
        if (!cells.every((c) => /^:?-{2,}:?$/.test(c) || c === "")) rows.push(cells);
        i++;
      }
      out += '<div class="scroll"><table>' + rows.map((r, ri) => "<tr>" + r.map((c) => ri === 0 ? `<th>${inline(c)}</th>` : `<td>${inline(c)}</td>`).join("") + "</tr>").join("") + "</table></div>";
      continue;
    }
    let m;
    if ((m = l.match(/^(#{1,4})\s+(.*)/))) { close(); const n = Math.min(m[1].length + 1, 5); out += `<h${n}>${inline(m[2])}</h${n}>`; }
    else if ((m = l.match(/^\s*[-*]\s+(.*)/))) { if (list !== "ul") { close(); out += "<ul>"; list = "ul"; } out += `<li>${inline(m[1])}</li>`; }
    else if ((m = l.match(/^\s*\d+[.)]\s+(.*)/))) { if (list !== "ol") { close(); out += "<ol>"; list = "ol"; } out += `<li>${inline(m[1])}</li>`; }
    else if (l.trim() === "") { close(); }
    else { close(); out += `<p>${inline(l)}</p>`; }
    i++;
  }
  close();
  return out;
}

// ---------------------------------------------------------------- 메뉴
const EXTRA_NAV = [{ id: "home", name: "홈", desc: "프로젝트 대시보드" }];
const TAIL_NAV = [{ id: "archive", name: "결과 보관함", desc: "생성한 결과물 다시 보기" }];

function renderNav(active) {
  const items = [...EXTRA_NAV, ...CFG.modules, ...TAIL_NAV];
  $("#nav").innerHTML = items.map((m, i) =>
    (i === 1 || i === items.length - 1 ? '<div class="sep"></div>' : "") +
    `<a href="#${m.id}" class="${m.id === active ? "active" : ""}">${esc(m.name)}<small>${esc(m.desc)}</small></a>`).join("");
}

async function route() {
  const id = location.hash.slice(1) || "home";
  renderNav(id);
  $("#nav").classList.remove("open");
  const main = $("#main");
  if (id === "home") return renderHome(main);
  if (id === "archive") return renderArchive(main);
  const mod = CFG.modules.find((m) => m.id === id);
  if (!mod) return (main.innerHTML = "<p>메뉴를 선택하세요.</p>");
  main.innerHTML = `<h2>${esc(mod.name)}</h2><p class="desc">${esc(mod.desc)}</p><div id="special"></div><div id="gen"></div>`;
  const views = { sourcing: viewSourcing, schedule: viewSchedule, ic: viewIC, procurement: viewProcurement, distribution: viewDistribution };
  if (mod.view && views[mod.view]) await views[mod.view]($("#special"));
  renderGenerator($("#gen"), mod);
}

// ---------------------------------------------------------------- 홈
async function renderHome(main) {
  const [projects, events, gens] = await Promise.all([api("/api/projects"), api("/api/t/events"), api("/api/generations")]);
  const today = new Date().toISOString().slice(0, 10);
  const upcoming = events.filter((e) => e.date >= today && !e.done).sort((a, b) => a.date.localeCompare(b.date)).slice(0, 10);
  main.innerHTML = `<h2>홈</h2><p class="desc">진행 중인 프로젝트와 다가오는 일정을 한눈에 봅니다.</p>
  <div class="grid2">
    <div class="card"><h3>프로젝트</h3><table class="tbl"><tr><th>코드명</th><th>제목</th></tr>
      ${projects.map((p) => `<tr class="click" data-code="${esc(p.code)}"><td>${esc(p.code)}</td><td>${esc(p.title)}</td></tr>`).join("")}</table>
      <div id="projDetail" class="muted" style="margin-top:8px">행을 누르면 PROJECT.md를 표시합니다.</div></div>
    <div class="card"><h3>다가오는 일정</h3><table class="tbl"><tr><th>일자</th><th>내용</th></tr>
      ${upcoming.map((e) => `<tr><td>${esc(e.date)}</td><td>[${esc(e.project)}] ${esc(e.title)}</td></tr>`).join("") || '<tr><td colspan="2">예정된 일정이 없습니다.</td></tr>'}</table>
      <div class="row"><a class="btn sub" href="#schedule">캘린더 열기</a></div></div>
  </div>
  <div class="card"><h3>최근 생성 결과</h3><table class="tbl"><tr><th>#</th><th>단계</th><th>제목</th><th>일시</th></tr>
    ${gens.slice(0, 8).map((g) => `<tr class="click" data-gid="${g.id}"><td>${g.id}</td><td>${esc(g.module)}</td><td>${esc(g.title)}</td><td>${esc(g.created_at)}</td></tr>`).join("") || '<tr><td colspan="4">아직 없습니다.</td></tr>'}</table></div>`;
  main.querySelectorAll("[data-code]").forEach((tr) => tr.onclick = async () => {
    const p = await api(`/api/projects/${tr.dataset.code}`);
    $("#projDetail").innerHTML = `<div class="output" style="max-height:420px;overflow:auto">${md(p.markdown)}</div>`;
  });
  main.querySelectorAll("[data-gid]").forEach((tr) => tr.onclick = () => { location.hash = "archive"; setTimeout(() => openGeneration(+tr.dataset.gid), 50); });
}

// ---------------------------------------------------------------- 생성 패널(모든 모듈 공통)
function renderGenerator(box, mod) {
  box.innerHTML = `<div class="card"><h3>문서 생성 — 샘플 양식으로 작성</h3>
    <label>만들 문서</label><select id="subtype">${mod.subtypes.map((s) => `<option value="${s.id}">${esc(s.name)}</option>`).join("")}</select>
    <div class="muted" id="stDesc" style="margin-top:6px"></div>
    <div id="fields"></div>
    <label>참고 샘플 (체크한 서류의 양식을 따라 작성)</label><div class="samples" id="samples"></div>
    <label>새 샘플 올리기 (PDF·Word·PPT·Excel·이메일 등)</label><input type="file" id="files" multiple>
    <div class="row"><label style="margin:0"><input type="checkbox" id="saveSamples" checked> 샘플 라이브러리에 저장(다음에도 사용)</label></div>
    <label>추가 요청</label><textarea id="extra" placeholder="예: 2쪽 이내, 표 위주로, 대주 25개사 수신처 포함"></textarea>
    ${mod.web_search ? '<div class="row"><label style="margin:0"><input type="checkbox" id="web" checked> 웹 검색으로 최신 자료 찾기</label></div>' : ""}
    <div class="row"><button id="go">생성하기</button><span id="status" class="muted"></span></div>
  </div>
  <div class="card" id="resultCard" style="display:none"><h3>결과</h3><div id="warns" class="warn"></div>
    <div class="output" id="out"></div><textarea id="edit" style="display:none;min-height:400px"></textarea>
    <div class="row" id="resultBtns"></div></div>`;
  const sel = $("#subtype", box);
  const load = async () => {
    const st = mod.subtypes.find((s) => s.id === sel.value);
    $("#stDesc", box).textContent = st.instructions;
    $("#fields", box).innerHTML = (st.fields || []).map((f) => `<label>${esc(f.label)}</label>` +
      (f.type === "area" ? `<textarea data-k="${f.key}" placeholder="${esc(f.placeholder || "")}"></textarea>` : `<input type="text" data-k="${f.key}" placeholder="${esc(f.placeholder || "")}">`)).join("");
    await loadSamples(mod.id, st.id);
  };
  sel.onchange = load;
  load();
  $("#go", box).onclick = () => generate(box, mod);
}

async function loadSamples(mid, sid) {
  const list = await api(`/api/samples?module=${mid}&subtype=${sid}`);
  $("#samples").innerHTML = list.map((s) => `<div><input type="checkbox" value="${esc(s.id)}" checked>
    <span>${esc(s.name)} <small>(${esc(s.source)})</small></span>
    ${s.id.startsWith("u:") ? `<a href="#" data-del="${esc(s.id)}" class="muted">삭제</a>` : ""}</div>`).join("") || '<span class="muted">기본 샘플이 없습니다. 아래에서 올려 주세요.</span>';
  $("#samples").querySelectorAll("[data-del]").forEach((a) => a.onclick = async (e) => {
    e.preventDefault();
    if (!confirm("이 샘플을 라이브러리에서 지울까요?")) return;
    await api(`/api/samples?id=${encodeURIComponent(a.dataset.del)}`, { method: "DELETE" });
    loadSamples(mid, sid);
  });
}

async function generate(box, mod) {
  const st = $("#subtype", box).value;
  const fields = {};
  box.querySelectorAll("[data-k]").forEach((el) => (fields[el.dataset.k] = el.value));
  const fd = new FormData();
  fd.append("module", mod.id); fd.append("subtype", st); fd.append("fields", JSON.stringify(fields));
  fd.append("sample_ids", JSON.stringify([...box.querySelectorAll("#samples input:checked")].map((c) => c.value)));
  fd.append("extra", $("#extra", box).value);
  fd.append("web_search", $("#web", box)?.checked ? "1" : "0");
  fd.append("save_samples", $("#saveSamples", box).checked ? "1" : "0");
  for (const f of $("#files", box).files) fd.append("files", f);
  const btn = $("#go", box); btn.disabled = true;
  $("#status", box).textContent = "생성 중… (샘플 분량에 따라 1~5분)";
  $("#resultCard", box).style.display = "block";
  const out = $("#out", box); out.innerHTML = ""; $("#warns", box).textContent = ""; $("#resultBtns", box).innerHTML = "";
  let text = "", gid = null;
  try {
    const r = await fetch("/api/generate", { method: "POST", body: fd });
    if (!r.ok) throw new Error(await r.text());
    const reader = r.body.getReader(), dec = new TextDecoder();
    let buf = "";
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      buf += dec.decode(value, { stream: true });
      let k;
      while ((k = buf.indexOf("\n\n")) >= 0) {
        const line = buf.slice(0, k); buf = buf.slice(k + 2);
        if (!line.startsWith("data: ")) continue;
        const ev = JSON.parse(line.slice(6));
        if (ev.warn) $("#warns", box).innerHTML = ev.warn.map(esc).join("<br>");
        if (ev.t) { text += ev.t; out.innerHTML = md(text); }
        if (ev.done) gid = ev.id;
      }
    }
    $("#status", box).innerHTML = '<span class="ok">완료 — 결과 보관함에도 저장했습니다.</span>';
    $("#files", box).value = "";
    if ($("#saveSamples", box).checked) loadSamples(mod.id, st);
    resultButtons(box, gid, () => text, (t) => { text = t; });
  } catch (e) {
    $("#status", box).innerHTML = `<span class="warn">오류: ${esc(e.message)}</span>`;
  } finally { btn.disabled = false; }
}

function resultButtons(box, gid, getText, setText) {
  const b = $("#resultBtns", box);
  b.innerHTML = `<button class="sub" id="copy">복사</button><button class="sub" id="toggle">편집</button>
    <a class="btn" href="/api/generations/${gid}/docx">Word 다운로드</a><a class="btn sub" href="/api/generations/${gid}/md">Markdown 다운로드</a>`;
  $("#copy", b).onclick = () => navigator.clipboard.writeText(getText()).then(() => ($("#copy", b).textContent = "복사됨"));
  $("#toggle", b).onclick = async () => {
    const ed = $("#edit", box), out = $("#out", box);
    if (ed.style.display === "none") { ed.value = getText(); ed.style.display = "block"; out.style.display = "none"; $("#toggle", b).textContent = "편집 저장"; }
    else {
      setText(ed.value); await api(`/api/generations/${gid}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ output: ed.value }) });
      out.innerHTML = md(ed.value); ed.style.display = "none"; out.style.display = "block"; $("#toggle", b).textContent = "편집";
    }
  };
}

// ---------------------------------------------------------------- 결과 보관함
async function renderArchive(main) {
  const gens = await api("/api/generations");
  main.innerHTML = `<h2>결과 보관함</h2><p class="desc">지금까지 생성한 결과물입니다. 눌러서 다시 보고 Word로 내려받을 수 있습니다.</p>
  <div class="card"><table class="tbl"><tr><th>#</th><th>단계</th><th>문서</th><th>제목</th><th>일시</th></tr>
  ${gens.map((g) => `<tr class="click" data-gid="${g.id}"><td>${g.id}</td><td>${esc(g.module)}</td><td>${esc(g.subtype)}</td><td>${esc(g.title)}</td><td>${esc(g.created_at)}</td></tr>`).join("") || '<tr><td colspan="5">아직 없습니다.</td></tr>'}</table></div>
  <div id="gen"><div class="card" id="resultCard" style="display:none"><h3 id="gTitle"></h3><div id="warns"></div><div class="output" id="out"></div><textarea id="edit" style="display:none;min-height:400px"></textarea><div class="row" id="resultBtns"></div></div></div>`;
  main.querySelectorAll("[data-gid]").forEach((tr) => (tr.onclick = () => openGeneration(+tr.dataset.gid)));
}
async function openGeneration(gid) {
  const g = await api(`/api/generations/${gid}`);
  const box = $("#gen"); $("#resultCard").style.display = "block"; $("#gTitle").textContent = `#${g.id} ${g.title}`;
  let text = g.output; $("#out").innerHTML = md(text);
  resultButtons(box, gid, () => text, (t) => (text = t));
  $("#resultCard").scrollIntoView({ behavior: "smooth" });
}

// ---------------------------------------------------------------- 표 편집기(공통)
async function tableEditor(box, table, cols, opts = {}) {
  // cols: [{k, label, type:'text'|'area'|'date'|'num'|'check', show:true}]
  const rows = (await api(`/api/t/${table}`)).filter(opts.filter || (() => true));
  const shown = cols.filter((c) => c.show !== false);
  box.innerHTML = `<div class="scroll"><table class="tbl"><tr><th>#</th>${shown.map((c) => `<th>${esc(c.label)}</th>`).join("")}</tr>
    ${rows.map((r) => `<tr class="click" data-id="${r.id}"><td>${r.id}</td>${shown.map((c) => `<td>${esc(c.type === "check" ? (r[c.k] ? "✔" : "") : r[c.k])}</td>`).join("")}</tr>`).join("") || `<tr><td colspan="${shown.length + 1}">자료가 없습니다.</td></tr>`}</table></div>
    <details ${opts.open ? "open" : ""}><summary class="muted" style="margin-top:8px;cursor:pointer">추가·수정 (행을 누르면 불러옵니다)</summary>
    <form class="grid2">${cols.map((c) => `<div><label>${esc(c.label)}</label>${c.type === "area" ? `<textarea name="${c.k}"></textarea>` :
      c.type === "check" ? `<input type="checkbox" name="${c.k}">` : `<input type="${c.type === "date" ? "date" : c.type === "num" ? "number" : "text"}" step="any" name="${c.k}">`}</div>`).join("")}</form>
    <div class="row"><button class="save">저장</button><button class="sub new">새로 입력</button><button class="danger del">삭제</button></div></details>`;
  const form = $("form", box); let cur = null;
  const fill = (r) => { cur = r?.id || null; cols.forEach((c) => { const el = form.elements[c.k]; if (c.type === "check") el.checked = !!r?.[c.k]; else el.value = r?.[c.k] ?? ""; }); };
  fill(opts.defaults || null);
  box.querySelectorAll("tr[data-id]").forEach((tr) => (tr.onclick = () => { fill(rows.find((r) => r.id === +tr.dataset.id)); $("details", box).open = true; opts.onSelect?.(rows.find((r) => r.id === +tr.dataset.id)); }));
  $(".new", box).onclick = () => fill(opts.defaults || null);
  $(".save", box).onclick = async () => {
    const d = { id: cur }; cols.forEach((c) => { const el = form.elements[c.k]; d[c.k] = c.type === "check" ? (el.checked ? 1 : 0) : c.type === "num" ? (el.value === "" ? null : +el.value) : el.value; });
    await post(`/api/t/${table}`, d); opts.onChange ? opts.onChange() : tableEditor(box, table, cols, opts);
  };
  $(".del", box).onclick = async () => { if (cur && confirm("삭제할까요?")) { await api(`/api/t/${table}/${cur}`, { method: "DELETE" }); opts.onChange ? opts.onChange() : tableEditor(box, table, cols, opts); } };
  return { form, rows, fill };
}

// ---------------------------------------------------------------- Sourcing: 딜 목록 + 지도
async function viewSourcing(box) {
  box.innerHTML = `<div class="card"><h3>딜 지도</h3><div id="map"></div><p class="muted">지도를 누르면 아래 입력칸의 위도·경도가 채워집니다. 마커를 누르면 구글지도로 열 수 있습니다.</p></div>
    <div class="card"><h3>딜 파이프라인 · 자문사 마케팅 현황</h3><div id="deals"></div></div>`;
  const cols = [{ k: "name", label: "딜(자산명)" }, { k: "asset_type", label: "유형" }, { k: "address", label: "주소" },
    { k: "lat", label: "위도", type: "num", show: false }, { k: "lng", label: "경도", type: "num", show: false },
    { k: "size", label: "규모" }, { k: "price", label: "가격" }, { k: "advisor", label: "매각자문사" },
    { k: "marketing", label: "마케팅 현황" }, { k: "stage", label: "진행단계" }, { k: "deadline", label: "마감일", type: "date" }, { k: "notes", label: "메모", type: "area", show: false }];
  if (typeof L === "undefined") {  // 지도 라이브러리(인터넷 필요)를 못 불러와도 딜 표는 쓸 수 있게
    $("#map").outerHTML = '<p class="warn">지도 라이브러리를 불러오지 못했습니다(인터넷 연결 필요). 딜 목록은 아래에서 그대로 쓸 수 있습니다.</p>';
    const draw = () => tableEditor($("#deals"), "deals", cols, { onChange: draw });
    return void (await draw());
  }
  const map = L.map("map").setView([37.45, 127.0], 10);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", { maxZoom: 19, attribution: "© OpenStreetMap" }).addTo(map);
  const layer = L.layerGroup().addTo(map);
  const draw = async () => {
    const ed = await tableEditor($("#deals"), "deals", cols, { onChange: draw });
    layer.clearLayers();
    const pts = [];
    ed.rows.filter((r) => r.lat && r.lng).forEach((r) => {
      pts.push([r.lat, r.lng]);
      L.marker([r.lat, r.lng]).addTo(layer).bindPopup(`<b>${esc(r.name)}</b><br>${esc(r.stage || "")} · ${esc(r.price || "")}<br>
        <a target="_blank" href="https://www.google.com/maps/search/?api=1&query=${r.lat},${r.lng}">구글지도에서 열기</a>`);
    });
    if (pts.length) map.fitBounds(pts, { maxZoom: 13, padding: [30, 30] });
    map.off("click").on("click", (e) => { ed.form.elements.lat.value = e.latlng.lat.toFixed(6); ed.form.elements.lng.value = e.latlng.lng.toFixed(6); $("#deals details").open = true; });
  };
  await draw();
  setTimeout(() => map.invalidateSize(), 200);
}

// ---------------------------------------------------------------- 일정관리: 캘린더 + 일정 표
async function viewSchedule(box) {
  box.innerHTML = `<div class="card"><h3>캘린더</h3><div class="row" style="margin-top:0"><button class="sub" id="prev">◀</button><b id="ym"></b><button class="sub" id="next">▶</button>
    <a class="btn sub" href="/api/events.ics" style="margin-left:auto">휴대폰·아웃룩 캘린더로 내보내기(.ics)</a></div><div class="cal" id="cal" style="margin-top:8px"></div></div>
    <div class="card"><h3>일정·마일스톤 목록</h3><div id="evs"></div></div>`;
  let cur = new Date(); cur.setDate(1);
  const cols = [{ k: "project", label: "프로젝트" }, { k: "date", label: "일자", type: "date" }, { k: "title", label: "내용" },
    { k: "kind", label: "구분(마케팅·마일스톤·회의·기타)" }, { k: "done", label: "완료", type: "check" }, { k: "notes", label: "메모", type: "area", show: false }];
  let events = [];
  const drawCal = () => {
    const y = cur.getFullYear(), m = cur.getMonth();
    $("#ym").textContent = `${y}년 ${m + 1}월`;
    const start = new Date(y, m, 1 - new Date(y, m, 1).getDay());
    const today = new Date().toISOString().slice(0, 10);
    let h = ["일", "월", "화", "수", "목", "금", "토"].map((d) => `<div class="h">${d}</div>`).join("");
    for (let i = 0; i < 42; i++) {
      const d = new Date(start); d.setDate(start.getDate() + i);
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
      const evs = events.filter((e) => e.date === key);
      h += `<div class="d ${d.getMonth() !== m ? "other" : ""} ${key === today ? "today" : ""}">${d.getDate()}${evs.map((e) => `<div class="ev ${e.done ? "done" : ""}" title="${esc(e.title)}">${esc(e.title)}</div>`).join("")}</div>`;
    }
    $("#cal").innerHTML = h;
  };
  const reload = async () => { const ed = await tableEditor($("#evs"), "events", cols, { onChange: reload, defaults: { kind: "마일스톤" } }); events = ed.rows; drawCal(); };
  $("#prev").onclick = () => { cur.setMonth(cur.getMonth() - 1); drawCal(); };
  $("#next").onclick = () => { cur.setMonth(cur.getMonth() + 1); drawCal(); };
  await reload();
}

// ---------------------------------------------------------------- 기관별 투자심의: Q&A
async function viewIC(box) {
  box.innerHTML = `<div class="card"><h3>기관별 Q&A 목록</h3><label>기관 필터</label><input type="text" id="instF" placeholder="예: 우리은행"><div id="qa" style="margin-top:8px"></div>
    <p class="muted">아래 '기관별 Q&A 대응 요약'을 생성하면 이 목록이 자동으로 함께 들어갑니다.</p></div>`;
  const cols = [{ k: "project", label: "프로젝트" }, { k: "institution", label: "기관" }, { k: "category", label: "구분" },
    { k: "question", label: "질문", type: "area" }, { k: "answer", label: "답변", type: "area" }, { k: "status", label: "상태(미답변·확인중·답변완료)" }, { k: "asked", label: "질문일", type: "date" }];
  const draw = () => tableEditor($("#qa"), "qa", cols, { onChange: draw, filter: (r) => !$("#instF").value || (r.institution || "").includes($("#instF").value) });
  $("#instF").oninput = draw;
  await draw();
}

// ---------------------------------------------------------------- 발주관리: RFP + 견적
async function viewProcurement(box) {
  box.innerHTML = `<div class="card"><h3>발주 목록(RFP)</h3><div id="rfps"></div></div>
    <div class="card"><h3>견적 <span id="qFor" class="muted"></span></h3><div id="quotes"></div>
    <p class="muted">견적을 입력한 뒤 아래 '견적 비교표·품의문'에서 RFP 번호를 넣고 생성하면 비교표와 품의문 초안이 만들어집니다.</p></div>`;
  let rfp = null;
  const qcols = [{ k: "rfp_id", label: "RFP #", type: "num" }, { k: "vendor", label: "업체" }, { k: "amount", label: "견적금액" },
    { k: "period", label: "수행기간" }, { k: "notes", label: "비고·범위 차이", type: "area" }, { k: "selected", label: "선정", type: "check" }];
  const drawQ = () => tableEditor($("#quotes"), "quotes", qcols, { onChange: drawQ, filter: (r) => !rfp || r.rfp_id === rfp, defaults: { rfp_id: rfp } });
  const rcols = [{ k: "project", label: "프로젝트" }, { k: "category", label: "구분(PM·FM·시장·재무·세무·회계·감사·보험)" },
    { k: "title", label: "용역명" }, { k: "scope", label: "과업 범위", type: "area", show: false }, { k: "due", label: "기한", type: "date" }, { k: "status", label: "상태(RFP발송·견적접수·품의·발주완료)" }];
  const drawR = () => tableEditor($("#rfps"), "rfps", rcols, { onChange: drawR, onSelect: (r) => { rfp = r.id; $("#qFor").textContent = `— RFP #${r.id} ${r.title}`; drawQ(); } });
  await drawR(); await drawQ();
}

// ---------------------------------------------------------------- 배포: 이메일 양식 채우기
async function viewDistribution(box) {
  box.innerHTML = `<div class="card"><h3>저장된 이메일 양식</h3><p class="muted">바뀌는 값은 {{변수명}}으로 적습니다. 예: {{대주기관명}}, {{입금금액}}</p><div id="tpls"></div></div>
    <div class="card"><h3>양식 채워서 일괄 작성</h3><label>양식 선택</label><select id="tplSel"></select>
      <div id="vars" class="muted" style="margin-top:6px"></div>
      <label>받는 사람 목록(CSV) — 첫 줄에 변수명, '이메일'·'참조' 열은 받는 사람·참조로 사용</label>
      <div class="row" style="margin-top:0"><input type="file" id="csv" accept=".csv"><button class="sub" id="csvSample">CSV 서식 내려받기</button></div>
      <label>또는 한 건만 직접 입력</label><div id="single" class="grid2"></div>
      <div class="row"><button id="fillBtn">채우기</button><button class="sub" id="emlBtn" disabled>.eml로 내려받기(아웃룩에서 열어 발송)</button></div>
      <div id="mails"></div></div>`;
  let tpls = [], mails = [];
  const cols = [{ k: "name", label: "양식 이름" }, { k: "category", label: "구분" }, { k: "subject", label: "제목", show: false }, { k: "body", label: "본문", type: "area", show: false }];
  const varsOf = (t) => [...new Set([...(t.subject + t.body).matchAll(/\{\{\s*([^}]+?)\s*\}\}/g)].map((m) => m[1]))];
  const refresh = async () => {
    const ed = await tableEditor($("#tpls"), "templates", cols, { onChange: refresh });
    tpls = ed.rows;
    $("#tplSel").innerHTML = tpls.map((t) => `<option value="${t.id}">${esc(t.name)}</option>`).join("");
    showVars();
  };
  const showVars = () => {
    const t = tpls.find((x) => x.id === +$("#tplSel").value); if (!t) return;
    const v = varsOf(t);
    $("#vars").textContent = "변수: " + (v.join(", ") || "없음");
    $("#single").innerHTML = ["이메일", ...v].map((k) => `<div><label>${esc(k)}</label><input type="text" data-v="${esc(k)}"></div>`).join("");
  };
  $("#tplSel").onchange = showVars;
  $("#csvSample").onclick = () => {
    const t = tpls.find((x) => x.id === +$("#tplSel").value); if (!t) return;
    const blob = new Blob(["﻿" + ["이메일", "참조", ...varsOf(t)].join(",") + "\n"], { type: "text/csv" });
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = `${t.name}_받는사람목록.csv`; a.click();
  };
  $("#fillBtn").onclick = async () => {
    const tid = +$("#tplSel").value;
    if ($("#csv").files[0]) { const fd = new FormData(); fd.append("template_id", tid); fd.append("file", $("#csv").files[0]); mails = await api("/api/fill_csv", { method: "POST", body: fd }); }
    else { const row = {}; $("#single").querySelectorAll("[data-v]").forEach((el) => (row[el.dataset.v] = el.value)); mails = await post("/api/fill", { template_id: tid, rows: [row] }); }
    $("#mails").innerHTML = `<p class="muted">${mails.length}건 작성됨</p>` + mails.map((m, i) => `<div class="mail"><b>받는 사람:</b> ${esc(m.to)} ${m.cc ? "· 참조: " + esc(m.cc) : ""}\n<b>제목:</b> ${esc(m.subject)}\n\n${esc(m.body)}
<div class="row"><a class="btn sub" href="mailto:${encodeURIComponent(m.to)}?subject=${encodeURIComponent(m.subject)}&body=${encodeURIComponent(m.body)}${m.cc ? "&cc=" + encodeURIComponent(m.cc) : ""}">메일 앱으로 열기</a><button class="sub" data-copy="${i}">본문 복사</button></div></div>`).join("");
    $("#mails").querySelectorAll("[data-copy]").forEach((b) => (b.onclick = () => navigator.clipboard.writeText(mails[+b.dataset.copy].body)));
    $("#emlBtn").disabled = !mails.length;
  };
  $("#emlBtn").onclick = async () => {
    const r = await fetch("/api/eml_zip", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ mails }) });
    const a = document.createElement("a"); a.href = URL.createObjectURL(await r.blob()); a.download = "mails.zip"; a.click();
  };
  await refresh();
}

// ---------------------------------------------------------------- 시작
(async () => {
  CFG = await api("/api/config");
  $("#modelTag").textContent = `Claude: ${CFG.model}`;
  $("#menuBtn").onclick = () => $("#nav").classList.toggle("open");
  window.onhashchange = route;
  route();
})();
