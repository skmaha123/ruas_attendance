/* RUAS Student Attendance & Learning Dashboard — runs fully in the browser */
"use strict";

const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const r1 = n => n == null || isNaN(n) ? null : Math.round(n * 10) / 10;
const mean = a => { const v = a.filter(x => x != null && !isNaN(x)); return v.length ? v.reduce((s, x) => s + x, 0) / v.length : null; };

/* ---------------- settings (defaults = example values; overridden by the workbook's Settings sheet) ---------------- */
const DEFAULTS = { cieMax: 50, seeMax: 100, quizMax: 10, wCie: 40, wSee: 50, wQuiz: 10, wTheory: 70,
  minAtt: 75, parentAtt: 65, fastScore: 75, fastAtt: 85, slowScore: 50, counselScore: 40 };
const SETTING_LABELS = {
  "cie maximum marks": "cieMax", "see maximum marks": "seeMax", "quiz maximum marks": "quizMax", "cie weight": "wCie", "see weight": "wSee",
  "quiz weight": "wQuiz", "theory attendance weight": "wTheory", "minimum attendance %": "minAtt", "parent notification attendance %": "parentAtt",
  "fast learner: minimum score %": "fastScore", "fast learner: minimum attendance %": "fastAtt", "slow learner: score % below": "slowScore", "counselling: score % below": "counselScore",
};
const RULE_FIELDS = [
  ["minAtt", "Minimum attendance %"], ["parentAtt", "Notify parents below attendance %"], ["counselScore", "Counselling below Score %"],
  ["slowScore", "Slow learner below Score %"], ["fastScore", "Fast learner from Score %"], ["fastAtt", "…and attendance from %"],
  ["cieMax", "CIE max marks"], ["seeMax", "SEE max marks"], ["quizMax", "Quiz max marks"],
  ["wCie", "CIE weight %"], ["wSee", "SEE weight %"], ["wQuiz", "Quiz weight %"], ["wTheory", "Theory share of attendance %"],
];
let S = { ...DEFAULTS };

const STATUS = {
  parent: { label: "Notify parents", color: "#d03b3b", icon: "⚑", order: 0 },
  counsel: { label: "Counselling", color: "#ec835a", icon: "◆", order: 1 },
  slow: { label: "Slow learner", color: "#fab219", icon: "▲", order: 2 },
  fast: { label: "Fast learner", color: "#0ca30c", icon: "★", order: 3 },
  avg: { label: "Average", color: "#9a96a6", icon: "●", order: 4 },
};
const ST_KEYS = ["parent", "counsel", "slow", "fast", "avg"];

const state = { raw: [], level: "student", filters: {}, items: [], sel: null, actionTab: "parent", facSort: "score" };

/* ================= XLSX / CSV reading ================= */
async function readWorkbook(buf) {
  const zip = await JSZip.loadAsync(buf);
  const txt = async p => (zip.file(p) ? zip.file(p).async("string") : null);
  const parse = s => new DOMParser().parseFromString(s, "application/xml");
  const wbx = parse(await txt("xl/workbook.xml"));
  const rels = parse(await txt("xl/_rels/workbook.xml.rels"));
  const relMap = {}; for (const r of rels.getElementsByTagName("Relationship")) relMap[r.getAttribute("Id")] = r.getAttribute("Target");
  const ss = []; const sst = await txt("xl/sharedStrings.xml");
  if (sst) for (const si of parse(sst).getElementsByTagName("si")) ss.push([...si.getElementsByTagName("t")].map(t => t.textContent).join(""));
  const sheets = {};
  for (const sh of wbx.getElementsByTagName("sheet")) {
    const rid = sh.getAttribute("r:id") || sh.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    let target = relMap[rid]; if (!target) continue;
    target = target.startsWith("/") ? target.slice(1) : "xl/" + target.replace(/^\.\//, "");
    const x = await txt(target); if (!x) continue;
    const rows = [];
    for (const row of parse(x).getElementsByTagName("row")) {
      const arr = [];
      for (const c of row.getElementsByTagName("c")) {
        const ref = c.getAttribute("r"); const col = ref.replace(/\d+/g, "").split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
        const t = c.getAttribute("t"); const v = c.getElementsByTagName("v")[0];
        let val = null;
        if (t === "s") val = v ? ss[+v.textContent] : null;
        else if (t === "inlineStr") val = [...c.getElementsByTagName("t")].map(n => n.textContent).join("");
        else if (t === "b") val = v ? v.textContent === "1" : null;
        else if (v) { const n = Number(v.textContent); val = isNaN(n) ? v.textContent : n; }
        arr[col] = val;
      }
      rows.push(arr);
    }
    sheets[sh.getAttribute("name")] = rows;
  }
  return sheets;
}
function parseCSV(text) {
  const rows = []; let row = [], cur = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"' && text[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; }
    else if (ch === '"') q = true; else if (ch === ",") { row.push(cur); cur = ""; }
    else if (ch === "\n" || ch === "\r") { if (ch === "\r" && text[i + 1] === "\n") i++; row.push(cur); rows.push(row); row = []; cur = ""; }
    else cur += ch;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows.map(r => r.map(v => (v !== "" && !isNaN(v) ? Number(v) : v)));
}

const norm = s => String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
const COLS = {
  ay: ["academicyear", "ay", "year"], faculty: ["faculty", "facultycode", "teacher", "facultyname"], dept: ["dept", "department"],
  program: ["program", "programme"], name: ["nameofstudent", "studentname", "name", "student"], reg: ["registration", "registrationno", "regno", "usn", "studentid", "id"],
  sem: ["sem", "semester"], course: ["coursetheory", "course", "theorycourse", "subject"], lab: ["lab", "labcourse"],
  tatt: ["theoryattendance", "theoryattendancepercent", "theoryattendance%", "attendance", "attendancepercent", "attendance%"],
  latt: ["labattendance", "labattendancepercent"], cie: ["ciemark", "cie", "ciemarks"], see: ["seemark", "see", "seemarks"], quiz: ["quizmark", "quiz", "quizmarks"],
};
function mapHeaders(hdr) {
  const n = hdr.map(norm), idx = {};
  for (const [k, al] of Object.entries(COLS)) { const i = n.findIndex(h => al.includes(h)); if (i >= 0) idx[k] = i; }
  if (idx.tatt == null) { const i = n.findIndex(h => h.includes("attendance") && !h.includes("lab") && !h.includes("overall")); if (i >= 0) idx.tatt = i; }
  return idx;
}
function loadRows(sheets) {
  let name = Object.keys(sheets).find(k => norm(k) === "attendancedata") || Object.keys(sheets).find(k => { const h = sheets[k][0] || []; const m = mapHeaders(h); return m.tatt != null && m.cie != null; });
  if (!name) throw new Error("No sheet with attendance and CIE columns was found. Use the template's column headers.");
  const rows = sheets[name]; const idx = mapHeaders(rows[0] || []);
  const missing = ["tatt", "cie", "course"].filter(k => idx[k] == null).concat(idx.reg == null && idx.name == null ? ["registration or name"] : []);
  if (missing.length) throw new Error("Missing columns: " + missing.join(", "));
  const num = v => (v === "" || v == null || isNaN(Number(v)) ? null : Number(v));
  const out = [];
  for (const r of rows.slice(1)) {
    if (!r || r.every(v => v == null || v === "")) continue;
    const g = k => (idx[k] != null ? r[idx[k]] : null);
    out.push({ ay: String(g("ay") ?? "—"), faculty: String(g("faculty") ?? "—"), dept: String(g("dept") ?? "—"), program: String(g("program") ?? "—"),
      name: String(g("name") ?? g("reg") ?? ""), reg: String(g("reg") ?? g("name") ?? ""), sem: String(g("sem") ?? "—"), course: String(g("course") ?? ""),
      lab: String(g("lab") ?? ""), tatt: num(g("tatt")), latt: num(g("latt")), cie: num(g("cie")), see: num(g("see")), quiz: num(g("quiz")) });
  }
  // Settings sheet → thresholds
  const st = sheets["Settings"];
  if (st) for (const row of st) {
    const key = SETTING_LABELS[String(row?.[0] ?? "").trim().toLowerCase()];
    if (key && typeof row[1] === "number") S[key] = row[1] <= 1 && key.startsWith("w") ? row[1] * 100 : row[1];
  }
  return out;
}

/* ================= computation ================= */
function rowMetrics(r) {
  const att = r.latt != null && r.tatt != null ? r.tatt * S.wTheory / 100 + r.latt * (1 - S.wTheory / 100) : r.tatt;
  const c = r.cie != null ? r.cie / S.cieMax : null, s = r.see != null ? r.see / S.seeMax : null, q = r.quiz != null ? r.quiz / S.quizMax : null;
  let score = null;
  if (c != null && s != null) score = (c * S.wCie + s * S.wSee + (q ?? 0) * S.wQuiz) / (S.wCie + S.wSee + (q != null ? S.wQuiz : 0)) * 100;
  else if (c != null) score = (c * S.wCie + (q ?? 0) * S.wQuiz) / (S.wCie + (q != null ? S.wQuiz : 0)) * 100;
  return { att: r1(att), score: r1(score) };
}
function decide(att, score, shortage = 0, lowCourses = 0, veryLow = 0) {
  const reasons = { parent: [], counsel: [], slow: [], fast: [] };
  if (att != null && att < S.parentAtt) reasons.parent.push(`attendance ${att}% is below ${S.parentAtt}%`);
  if (att != null && score != null && att < S.minAtt && score < S.counselScore) reasons.parent.push(`attendance shortage together with Score ${score}%`);
  if (veryLow >= 2) reasons.parent.push(`${veryLow} courses below ${S.parentAtt}% attendance`);
  if (att != null && att < S.minAtt) reasons.counsel.push(`attendance ${att}% is below ${S.minAtt}%`);
  if (score != null && score < S.counselScore) reasons.counsel.push(`Score ${score}% is below ${S.counselScore}%`);
  if (shortage >= 2) reasons.counsel.push(`attendance shortage in ${shortage} courses`);
  if (score != null && score < S.slowScore) reasons.slow.push(`Score ${score}% is below ${S.slowScore}%`);
  if (lowCourses >= 2) reasons.slow.push(`Score below ${S.counselScore}% in ${lowCourses} courses`);
  if (score != null && att != null && score >= S.fastScore && att >= S.fastAtt && shortage === 0) reasons.fast.push(`Score ${score}% and attendance ${att}%`);
  const tags = ["parent", "counsel", "slow", "fast"].filter(k => reasons[k].length);
  const primary = tags.find(k => k !== "fast" || tags.length === 1) || "avg";
  return { reasons, tags, primary: tags.length ? primary : "avg" };
}
function compute() {
  const f = state.filters;
  const rows = state.raw.filter(r => Object.entries(f).every(([k, v]) => !v || r[k] === v)).map(r => ({ ...r, ...rowMetrics(r) }));
  state.rows = rows;
  if (state.level === "course") {
    state.items = rows.map(r => ({ ...r, key: r.reg + "|" + r.course, label: `${r.name} — ${r.course}`, courses: [r], ...decide(r.att, r.score) }));
  } else {
    const g = new Map();
    for (const r of rows) { const k = r.reg + "|" + r.ay + "|" + r.sem; if (!g.has(k)) g.set(k, []); g.get(k).push(r); }
    state.items = [...g.values()].map(cs => {
      const a = r1(mean(cs.map(c => c.att))), s = r1(mean(cs.map(c => c.score)));
      const shortage = cs.filter(c => c.att != null && c.att < S.minAtt).length, low = cs.filter(c => c.score != null && c.score < S.counselScore).length;
      const veryLow = cs.filter(c => c.att != null && c.att < S.parentAtt).length;
      return { ...cs[0], key: cs[0].reg + "|" + cs[0].ay + "|" + cs[0].sem, label: cs[0].name, att: a, score: s, courses: cs, shortage, low, ...decide(a, s, shortage, low, veryLow) };
    });
  }
}

/* ================= charts (SVG) ================= */
const SVGNS = "http://www.w3.org/2000/svg";
function el(tag, attrs = {}, parent) { const e = document.createElementNS(SVGNS, tag); for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v); if (parent) parent.appendChild(e); return e; }
const tip = () => $("tip");
function showTip(e, html) { const t = tip(); t.innerHTML = html; t.hidden = false; const x = Math.min(e.clientX + 14, innerWidth - t.offsetWidth - 10); t.style.left = x + "px"; t.style.top = (e.clientY + 14) + "px"; }
function hideTip() { tip().hidden = true; }

function scatter() {
  const box = $("scatter"); box.innerHTML = "";
  const W = box.clientWidth || 760, H = 440, m = { l: 56, r: 20, t: 16, b: 50 };
  const pw = W - m.l - m.r, ph = H - m.t - m.b;
  const xmin = Math.min(40, ...state.items.map(i => i.att ?? 100)) - 2;
  const X = v => m.l + (v - xmin) / (100 - xmin) * pw, Y = v => m.t + (1 - v / 100) * ph;
  const svg = el("svg", { viewBox: `0 0 ${W} ${H}`, width: "100%", role: "img", "aria-label": "Scatter of attendance versus score" }, box);
  // quadrant shading (subtle)
  el("rect", { x: X(S.fastAtt), y: Y(100), width: X(100) - X(S.fastAtt), height: Y(S.fastScore) - Y(100), fill: "#0ca30c", opacity: .06 }, svg);
  el("rect", { x: X(xmin), y: Y(S.counselScore), width: X(S.minAtt) - X(xmin), height: Y(0) - Y(S.counselScore), fill: "#d03b3b", opacity: .06 }, svg);
  for (let v = 0; v <= 100; v += 20) {
    el("line", { x1: m.l, x2: m.l + pw, y1: Y(v), y2: Y(v), stroke: "#ece8f2" }, svg);
    el("text", { x: m.l - 8, y: Y(v) + 4, "text-anchor": "end", class: "axl" }, svg).textContent = v;
  }
  for (let v = Math.ceil(xmin / 10) * 10; v <= 100; v += 10) el("text", { x: X(v), y: m.t + ph + 18, "text-anchor": "middle", class: "axl" }, svg).textContent = v;
  el("text", { x: m.l + pw / 2, y: H - 8, "text-anchor": "middle", class: "axt" }, svg).textContent = "Attendance % →";
  el("text", { x: 14, y: m.t + ph / 2, "text-anchor": "middle", class: "axt", transform: `rotate(-90 14 ${m.t + ph / 2})` }, svg).textContent = "Score % (CIE + SEE + Quiz) →";
  const vline = (v, lab, col) => { el("line", { x1: X(v), x2: X(v), y1: m.t, y2: m.t + ph, stroke: col, "stroke-dasharray": "5 4", "stroke-width": 1.5 }, svg);
    el("text", { x: X(v) + 4, y: m.t + 12, class: "thl", fill: col }, svg).textContent = lab; };
  const hline = (v, lab, col) => { el("line", { x1: m.l, x2: m.l + pw, y1: Y(v), y2: Y(v), stroke: col, "stroke-dasharray": "5 4", "stroke-width": 1.5 }, svg);
    el("text", { x: m.l + pw - 4, y: Y(v) - 5, "text-anchor": "end", class: "thl", fill: col }, svg).textContent = lab; };
  vline(S.parentAtt, `${S.parentAtt}% notify`, "#b42318"); vline(S.minAtt, `${S.minAtt}% min`, "#8a5a00");
  hline(S.counselScore, `${S.counselScore}% counselling`, "#b42318"); hline(S.slowScore, `${S.slowScore}% slow`, "#8a5a00"); hline(S.fastScore, `${S.fastScore}% fast`, "#0a7a0a");
  el("text", { x: X(100) - 6, y: Y(100) + 30, "text-anchor": "end", class: "qlab", fill: "#0a7a0a" }, svg).textContent = "★ FAST LEARNERS";
  el("text", { x: X(xmin) + 6, y: Y(0) - 8, class: "qlab", fill: "#b42318" }, svg).textContent = "⚑ ACT NOW";
  el("text", { x: X(100) - 6, y: Y(0) - 8, "text-anchor": "end", class: "qlab", fill: "#8a5a00" }, svg).textContent = "▲ ATTENDS, BUT STRUGGLES";
  el("text", { x: X(xmin) + 6, y: Y(100) + 30, class: "qlab", fill: "#9a4b2a" }, svg).textContent = "◆ CAPABLE, BUT ABSENT";
  const pts = state.items.filter(i => i.att != null && i.score != null).sort((a, b) => STATUS[b.primary].order - STATUS[a.primary].order);
  const g = el("g", {}, svg);
  for (const p of pts) el("circle", { cx: X(p.att), cy: Y(p.score), r: 5, fill: STATUS[p.primary].color, stroke: "#fff", "stroke-width": 1.5, opacity: .9 }, g);
  // nearest-point hover layer
  const hit = el("rect", { x: m.l, y: m.t, width: pw, height: ph, fill: "transparent" }, svg);
  const ring = el("circle", { r: 8, fill: "none", stroke: "#331a6a", "stroke-width": 2, visibility: "hidden" }, svg);
  const nearest = e => {
    const pt = svg.createSVGPoint(); pt.x = e.clientX; pt.y = e.clientY; const loc = pt.matrixTransform(svg.getScreenCTM().inverse());
    let best = null, bd = 400; for (const p of pts) { const d = (X(p.att) - loc.x) ** 2 + (Y(p.score) - loc.y) ** 2; if (d < bd) { bd = d; best = p; } } return best;
  };
  hit.addEventListener("pointermove", e => {
    const p = nearest(e); if (!p) { hideTip(); ring.setAttribute("visibility", "hidden"); hit.style.cursor = "default"; return; }
    ring.setAttribute("cx", X(p.att)); ring.setAttribute("cy", Y(p.score)); ring.setAttribute("visibility", "visible"); hit.style.cursor = "pointer";
    showTip(e, `<b>${esc(p.label)}</b><br>${esc(p.reg)} · ${esc(p.program)} · Sem ${esc(p.sem)}<br><b>${p.att}%</b> attendance · <b>${p.score}%</b> score<br>${p.tags.map(t => `${STATUS[t].icon} ${STATUS[t].label}`).join(" · ") || "● Average"}<br><i>Click for details</i>`);
  });
  hit.addEventListener("pointerleave", () => { hideTip(); ring.setAttribute("visibility", "hidden"); });
  hit.addEventListener("click", e => { const p = nearest(e); if (p) openDetail(p); });
}

function hbars(boxId, data, opts = {}) {   // data: [{label, value, note, color}]
  const box = $(boxId); box.innerHTML = "";
  const max = opts.max || Math.max(1, ...data.map(d => d.value ?? 0));
  for (const d of data) {
    const row = document.createElement("div"); row.className = "hbar";
    row.innerHTML = `<span class="lbl"></span><div class="track"><div class="fill"></div></div><span class="val"></span>`;
    row.querySelector(".lbl").textContent = d.label; row.querySelector(".val").textContent = d.note ?? (d.value == null ? "—" : d.value);
    const fill = row.querySelector(".fill"); fill.style.width = (d.value ?? 0) / max * 100 + "%"; fill.style.background = d.color || "#331a6a";
    row.addEventListener("pointermove", e => showTip(e, `<b>${esc(d.note ?? d.value)}</b> ${esc(d.tip || d.label)}`)); row.addEventListener("pointerleave", hideTip);
    box.appendChild(row);
  }
}

function stacked(boxId, groups) {      // groups: [{label, counts:{parent,...}, total}]
  const box = $(boxId); box.innerHTML = "";
  for (const g of groups) {
    const row = document.createElement("div"); row.className = "srow";
    const lbl = document.createElement("span"); lbl.className = "lbl"; lbl.textContent = g.label; row.appendChild(lbl);
    const bar = document.createElement("div"); bar.className = "sbar";
    for (const k of ST_KEYS) {
      const n = g.counts[k] || 0; if (!n) continue;
      const seg = document.createElement("div"); seg.className = "sg"; seg.style.width = n / g.total * 100 + "%"; seg.style.background = STATUS[k].color;
      if (n / g.total > 0.07) seg.textContent = n;
      seg.addEventListener("pointermove", e => showTip(e, `<b>${n}</b> (${Math.round(n / g.total * 100)}%) ${STATUS[k].icon} ${STATUS[k].label}<br>${esc(g.label)}`));
      seg.addEventListener("pointerleave", hideTip); bar.appendChild(seg);
    }
    row.appendChild(bar);
    const tot = document.createElement("span"); tot.className = "val"; tot.textContent = g.total; row.appendChild(tot);
    box.appendChild(row);
  }
}

/* ================= render ================= */
function renderAll() {
  compute();
  const it = state.items, n = it.length;
  $("main").hidden = !state.raw.length; $("empty").hidden = !!state.raw.length;
  if (!state.raw.length) return;
  const cnt = k => it.filter(i => i.tags.includes(k)).length;
  const prim = k => it.filter(i => i.primary === k).length;
  const unit = state.level === "student" ? "students" : "course records";
  $("kpis").innerHTML = `
    <div class="kpi"><b>${n}</b><span>${unit}</span></div>
    <div class="kpi"><b>${r1(mean(it.map(i => i.att))) ?? "—"}%</b><span>average attendance</span></div>
    <div class="kpi"><b>${r1(mean(it.map(i => i.score))) ?? "—"}%</b><span>average score</span></div>
    ${["fast", "slow", "counsel", "parent"].map(k => `<div class="kpi st" style="--c:${STATUS[k].color}" data-act="${k}"><b>${STATUS[k].icon} ${cnt(k)}</b><span>${STATUS[k].label}</span></div>`).join("")}`;
  $("legend").innerHTML = ST_KEYS.map(k => `<span><i style="background:${STATUS[k].color}"></i>${STATUS[k].icon} ${STATUS[k].label} <b>${prim(k)}</b></span>`).join("");
  scatter();

  // decision mix by program
  const by = (key) => { const m = new Map(); for (const i of it) { const k = i[key]; if (!m.has(k)) m.set(k, { label: k, counts: {}, total: 0 }); const g = m.get(k); g.counts[i.primary] = (g.counts[i.primary] || 0) + 1; g.total++; } return [...m.values()]; };
  stacked("mix", by($("mixBy").value).sort((a, b) => ((b.counts.parent || 0) + (b.counts.counsel || 0)) / b.total - ((a.counts.parent || 0) + (a.counts.counsel || 0)) / a.total));

  // attendance band → score (course records, always)
  const bands = [["Below " + S.parentAtt + "%", 0, S.parentAtt], [`${S.parentAtt}–${S.minAtt - 1}%`, S.parentAtt, S.minAtt], [`${S.minAtt}–84%`, S.minAtt, 85], ["85–94%", 85, 95], ["95% and above", 95, 101]];
  const bd = bands.map(([l, a, b]) => { const rs = state.rows.filter(r => r.att != null && r.att >= a && r.att < b); const v = r1(mean(rs.map(r => r.score))); return { label: l, value: v, note: v == null ? "—" : v + "%", tip: `average score for ${rs.length} course records with attendance ${l}` }; });
  hbars("bands", bd, { max: 100 });
  const lo = bd[0].value ?? bd[1].value, hi = bd[4].value ?? bd[3].value;
  $("bandNote").textContent = lo != null && hi != null ? `Students attending 95%+ score ${r1(hi - lo)} points higher on average than those below ${S.parentAtt}%. Regular attendance matters.` : "";

  renderFaculty(); renderActions(); renderTable();
}

function renderFaculty() {
  const m = new Map();
  for (const r of state.rows) { if (!m.has(r.faculty)) m.set(r.faculty, []); m.get(r.faculty).push(r); }
  let fac = [...m.entries()].map(([f, rs]) => {
    const st = rs.map(r => decide(r.att, r.score));
    return { f, courses: [...new Set(rs.map(r => r.course))], n: rs.length, att: r1(mean(rs.map(r => r.att))), score: r1(mean(rs.map(r => r.score))),
      pass: Math.round(rs.filter(r => r.score != null && r.score >= S.counselScore).length / rs.length * 100),
      slow: st.filter(s => s.tags.includes("slow")).length, counsel: st.filter(s => s.tags.includes("counsel")).length, parent: st.filter(s => s.tags.includes("parent")).length };
  });
  const k = state.facSort; fac.sort((a, b) => k === "f" ? a.f.localeCompare(b.f) : (b[k] ?? 0) - (a[k] ?? 0));
  const ovAtt = r1(mean(state.rows.map(r => r.att))), ovScore = r1(mean(state.rows.map(r => r.score)));
  const cell = (v, avg, max = 100) => `<div class="mini"><div class="mtrack"><div class="mfill ${v != null && avg != null && v < avg ? "below" : ""}" style="width:${(v ?? 0) / max * 100}%"></div><span class="mavg" style="left:${avg / max * 100}%"></span></div><b>${v ?? "—"}</b></div>`;
  $("facTable").querySelector("tbody").innerHTML = fac.map(x => `<tr>
    <td><b>${esc(x.f)}</b></td><td class="small">${x.courses.map(esc).join("<br>")}</td><td class="c">${x.n}</td>
    <td>${cell(x.att, ovAtt)}</td><td>${cell(x.score, ovScore)}</td><td class="c">${x.pass}%</td>
    <td class="c">${x.slow ? `<span class="pill" style="--c:${STATUS.slow.color}">▲ ${x.slow}</span>` : "0"}</td>
    <td class="c">${x.counsel ? `<span class="pill" style="--c:${STATUS.counsel.color}">◆ ${x.counsel}</span>` : "0"}</td>
    <td class="c">${x.parent ? `<span class="pill" style="--c:${STATUS.parent.color}">⚑ ${x.parent}</span>` : "0"}</td></tr>`).join("");
  $("facNote").innerHTML = `Bars show each faculty member's class average; the thin line marks the overall average (attendance ${ovAtt}%, score ${ovScore}%). Bars in crimson are below it.`;
}

const ACTIONS = {
  parent: { title: "Notify parents", todo: "Send an attendance/performance letter; schedule a parent–mentor meeting." },
  counsel: { title: "Counselling session", todo: "Mentor or counsellor meeting within a week; agree an attendance and study plan." },
  slow: { title: "Slow learners — remedial support", todo: "Remedial classes, peer tutoring, extra practice and a re-test before SEE." },
  fast: { title: "Fast learners — enrichment", todo: "Advanced assignments, projects, competitions, research or MOOCs." },
};
function renderActions() {
  const k = state.actionTab;
  document.querySelectorAll("#actTabs button").forEach(b => b.classList.toggle("on", b.dataset.t === k));
  const list = state.items.filter(i => i.tags.includes(k)).sort((a, b) => k === "fast" ? b.score - a.score : a.att - b.att || a.score - b.score);
  $("actInfo").innerHTML = `<b>${list.length}</b> ${state.level === "student" ? "students" : "records"} · Suggested action: ${esc(ACTIONS[k].todo)}`;
  $("actTable").querySelector("tbody").innerHTML = list.map(i => `<tr data-k="${esc(i.key)}">
    <td><b>${esc(i.name)}</b><div class="small">${esc(i.reg)}</div></td><td class="small">${esc(i.program)} · Sem ${esc(i.sem)}${state.level === "course" ? "<br>" + esc(i.course) : ""}</td>
    <td class="c">${i.att ?? "—"}%</td><td class="c">${i.score ?? "—"}%</td><td class="small">${esc(i.reasons[k].join("; "))}</td></tr>`).join("") || `<tr><td colspan="5" class="muted">No one in this group for the current filters.</td></tr>`;
}

function renderTable() {
  const q = $("search").value.trim().toLowerCase();
  const list = state.items.filter(i => !q || (i.name + " " + i.reg + " " + i.course).toLowerCase().includes(q)).sort((a, b) => STATUS[a.primary].order - STATUS[b.primary].order || a.att - b.att).slice(0, 300);
  $("allTable").querySelector("tbody").innerHTML = list.map(i => `<tr data-k="${esc(i.key)}">
    <td><b>${esc(i.name)}</b><div class="small">${esc(i.reg)}</div></td><td class="small">${esc(i.dept)} · ${esc(i.program)}<br>Sem ${esc(i.sem)} · ${esc(i.ay)}</td>
    <td class="c">${i.att ?? "—"}%</td><td class="c">${i.score ?? "—"}%</td>
    <td>${(i.tags.length ? i.tags : ["avg"]).map(t => `<span class="pill" style="--c:${STATUS[t].color}">${STATUS[t].icon} ${STATUS[t].label}</span>`).join(" ")}</td></tr>`).join("");
  $("allCount").textContent = `${list.length}${list.length === 300 ? "+" : ""} shown`;
}

function openDetail(i) {
  state.sel = i;
  $("dName").textContent = i.name;
  $("dMeta").textContent = `${i.reg} · ${i.dept} · ${i.program} · Sem ${i.sem} · ${i.ay}`;
  $("dTags").innerHTML = (i.tags.length ? i.tags : ["avg"]).map(t => `<span class="pill big" style="--c:${STATUS[t].color}">${STATUS[t].icon} ${STATUS[t].label}</span>`).join(" ");
  $("dReasons").innerHTML = i.tags.map(t => `<li><b>${STATUS[t].label}:</b> ${esc(i.reasons[t].join("; "))}. <i>${esc(ACTIONS[t].todo)}</i></li>`).join("") || "<li>No action needed — keep monitoring.</li>";
  const box = $("dCourses"); box.innerHTML = "";
  for (const c of i.courses) {
    const row = document.createElement("div"); row.className = "crow";
    row.innerHTML = `<div class="cname"><b></b><small></small></div>
      <div class="pair"><span>Attendance</span><div class="track"><div class="fill att"></div><span class="tick" style="left:${S.minAtt}%"></span></div><b class="v1"></b></div>
      <div class="pair"><span>Score</span><div class="track"><div class="fill sc"></div><span class="tick" style="left:${S.slowScore}%"></span></div><b class="v2"></b></div>`;
    row.querySelector(".cname b").textContent = c.course; row.querySelector(".cname small").textContent = `${c.faculty}${c.lab ? " · Lab: " + c.lab : ""} · CIE ${c.cie ?? "—"}/${S.cieMax} · SEE ${c.see ?? "—"}/${S.seeMax} · Quiz ${c.quiz ?? "—"}/${S.quizMax}`;
    const fa = row.querySelector(".att"), fs = row.querySelector(".sc");
    fa.style.width = (c.att ?? 0) + "%"; fa.classList.toggle("bad", c.att < S.minAtt);
    fs.style.width = (c.score ?? 0) + "%"; fs.classList.toggle("bad", c.score < S.slowScore);
    row.querySelector(".v1").textContent = (c.att ?? "—") + "%"; row.querySelector(".v2").textContent = (c.score ?? "—") + "%";
    box.appendChild(row);
  }
  const low = i.courses.filter(c => c.att < S.minAtt).map(c => `${c.course} (${c.att}%)`);
  $("dLetter").value = i.tags.includes("parent") || i.tags.includes("counsel") ?
`Dear Parent/Guardian,

This is to inform you about the academic progress of your ward, ${i.name} (${i.reg}), ${i.program}, Semester ${i.sem}.

Overall attendance: ${i.att}% (required: ${S.minAtt}%).${low.length ? "\nCourses below the required attendance: " + low.join(", ") + "." : ""}
Average performance (CIE, SEE and quiz): ${i.score}%.

We request you to meet the mentor at the earliest to discuss a support plan.

Regards,
Mentor / Head of the Department` : "No notification needed for this student.";
  $("dlg").showModal();
}

/* ================= export ================= */
function colName(i) { let s = ""; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
const xe = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
async function exportXlsx() {
  const head = ["Name of Student", "Registration", "Dept", "Program", "Sem", "Academic Year", state.level === "course" ? "Course" : "Courses", "Attendance %", "Score %", "Decision(s)", "Reasons", "Suggested action"];
  const mk = list => [head, ...list.map(i => [i.name, i.reg, i.dept, i.program, i.sem, i.ay, state.level === "course" ? i.course : i.courses.length, i.att ?? "", i.score ?? "",
    (i.tags.length ? i.tags : ["avg"]).map(t => STATUS[t].label).join(", "), i.tags.map(t => i.reasons[t].join("; ")).join(" | "), i.tags.map(t => ACTIONS[t].todo).join(" ")])];
  const sheets = [["All", state.items], ["Notify parents", state.items.filter(i => i.tags.includes("parent"))], ["Counselling", state.items.filter(i => i.tags.includes("counsel"))],
    ["Slow learners", state.items.filter(i => i.tags.includes("slow"))], ["Fast learners", state.items.filter(i => i.tags.includes("fast"))]].map(([n, l]) => ({ name: n, rows: mk(l) }));
  sheets.push({ name: "Rules used", rows: [["Rule", "Value"], ...RULE_FIELDS.map(([k, l]) => [l, S[k]]), ["Generated", new Date().toLocaleString("en-IN")]] });
  const zip = new JSZip();
  const sx = rows => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${rows[0].map((_, i) => `<col min="${i + 1}" max="${i + 1}" width="${Math.min(60, Math.max(10, ...rows.map(r => String(r[i] ?? "").length + 2)))}" customWidth="1"/>`).join("")}</cols><sheetData>${rows.map((r, ri) => `<row r="${ri + 1}">${r.map((v, ci) => { const ref = colName(ci) + (ri + 1), st = ri === 0 ? ' s="1"' : ""; return typeof v === "number" ? `<c r="${ref}"${st}><v>${v}</v></c>` : `<c r="${ref}" t="inlineStr"${st}><is><t xml:space="preserve">${xe(v)}</t></is></c>`; }).join("")}</row>`).join("")}</sheetData></worksheet>`;
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${xe(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF331A6A"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"/></cellXfs></styleSheet>`);
  sheets.forEach((s, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, sx(s.rows)));
  const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = `RUAS_Attendance_Actions_${new Date().toISOString().slice(0, 10)}.xlsx`; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

/* ================= wiring ================= */
function buildFilters() {
  const defs = [["ay", "Academic Year"], ["dept", "Dept"], ["program", "Program"], ["sem", "Sem"], ["faculty", "Faculty"], ["course", "Course"]];
  $("filters").innerHTML = defs.map(([k, l]) => {
    const vals = [...new Set(state.raw.map(r => r[k]))].sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
    return `<label>${l}<select data-f="${k}"><option value="">All</option>${vals.map(v => `<option>${esc(v)}</option>`).join("")}</select></label>`;
  }).join("");
}
function buildRules() {
  $("rules").innerHTML = RULE_FIELDS.map(([k, l]) => `<label>${esc(l)}<input type="number" step="any" data-r="${k}" value="${S[k]}"/></label>`).join("");
}
async function loadBuffer(buf, name) {
  try {
    S = { ...DEFAULTS };
    const rows = name.toLowerCase().endsWith(".csv") ? loadRows({ Data: parseCSV(new TextDecoder().decode(buf)) }) : loadRows(await readWorkbook(buf));
    if (!rows.length) throw new Error("The sheet has no data rows.");
    state.raw = rows; state.filters = {};
    $("status").textContent = `Loaded ${rows.length} records from ${name}`; $("status").className = "status ok";
    buildFilters(); buildRules(); renderAll();
  } catch (e) { console.error(e); $("status").textContent = "Could not read the file: " + e.message; $("status").className = "status err"; }
}

function init() {
  $("file").addEventListener("change", async e => { const f = e.target.files[0]; if (f) await loadBuffer(await f.arrayBuffer(), f.name); e.target.value = ""; });
  const drop = $("drop");
  ["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", async e => { const f = e.dataTransfer.files[0]; if (f) await loadBuffer(await f.arrayBuffer(), f.name); });
  $("btnSample").addEventListener("click", async () => {
    try { const r = await fetch("RUAS_Attendance_Anonymised_Data.xlsx"); if (!r.ok) throw 0; await loadBuffer(await r.arrayBuffer(), "sample data"); }
    catch { $("status").textContent = "Sample data loads when the dashboard is opened from its website. Please upload the Excel file instead."; $("status").className = "status err"; }
  });
  $("filters").addEventListener("change", e => { const s = e.target.closest("[data-f]"); if (s) { state.filters[s.dataset.f] = s.value; renderAll(); } });
  $("rules").addEventListener("input", e => { const i = e.target.closest("[data-r]"); if (i && i.value !== "" && !isNaN(+i.value)) { S[i.dataset.r] = +i.value; renderAll(); } });
  $("btnReset").addEventListener("click", () => { state.filters = {}; document.querySelectorAll("[data-f]").forEach(s => (s.value = "")); renderAll(); });
  $("levelSeg").addEventListener("click", e => { const b = e.target.closest("button"); if (!b) return; state.level = b.dataset.l; document.querySelectorAll("#levelSeg button").forEach(x => x.classList.toggle("on", x === b)); renderAll(); });
  $("mixBy").addEventListener("change", renderAll);
  $("actTabs").addEventListener("click", e => { const b = e.target.closest("button"); if (b) { state.actionTab = b.dataset.t; renderActions(); } });
  $("kpis").addEventListener("click", e => { const k = e.target.closest("[data-act]"); if (k) { state.actionTab = k.dataset.act; renderActions(); $("actions").scrollIntoView({ behavior: "smooth" }); } });
  $("facTable").querySelector("thead").addEventListener("click", e => { const th = e.target.closest("[data-s]"); if (th) { state.facSort = th.dataset.s; renderFaculty(); } });
  $("search").addEventListener("input", renderTable);
  document.addEventListener("click", e => { const tr = e.target.closest("tr[data-k]"); if (tr) { const i = state.items.find(x => x.key === tr.dataset.k); if (i) openDetail(i); } });
  $("btnClose").addEventListener("click", () => $("dlg").close());
  $("btnCopy").addEventListener("click", () => { navigator.clipboard?.writeText($("dLetter").value); $("btnCopy").textContent = "Copied ✓"; setTimeout(() => ($("btnCopy").textContent = "Copy text"), 1500); });
  $("btnExport").addEventListener("click", exportXlsx);
  window.addEventListener("resize", () => state.raw.length && scatter());
}
document.addEventListener("DOMContentLoaded", init);
