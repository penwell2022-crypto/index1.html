/* =========================================================
   แผนภาพธาตุ 4 – อาการแยกตามธาตุ
   © 2026 [ชื่อผู้จัดทำ] สงวนลิขสิทธิ์ (All rights reserved)
   ห้ามคัดลอก ดัดแปลง หรือนำโค้ดไปใช้ต่อโดยไม่ได้รับอนุญาต
   ข้อมูล: B=รหัส C=ชื่อ D=คำอธิบาย E=ระดับ F=ภาวะ G..Z=อาการ 1-20 (เริ่มแถว 3)
   ========================================================= */

const DATA_URLS = ["data/data.xlsx", "data/aasyma.xlsx", "data.xlsx", "data/Data.xlsx"];
const TABS = ["cards", "ladder", "map", "explore", "quiz"];

// ธาตุหลัก (คำช่วยจำ = ตัวช่วยเรียนที่เพิ่มขึ้น ไม่ได้มาจากตำรา)
const ELEMENTS = [
  { key: "ไฟ",   code: "A01", title: "🔥 ไฟ (เตโช)",   prefix: "เตโชธาตุ", c: "var(--fire)",  l: "var(--fire-l)",  mn: "ร้อน • แสบ • ไข้ • ดี/กำเดา (ระบบเผาผลาญ)" },
  { key: "ดิน",  code: "A02", title: "⛰️ ดิน (ปถวี)",  prefix: "ปถวีธาตุ", c: "var(--earth)", l: "var(--earth-l)", mn: "ของแข็งในกาย: ผม ฟัน กระดูก อวัยวะ ท้อง-ไส้" },
  { key: "น้ำ",  code: "A03", title: "💧 น้ำ (อาโป)",  prefix: "อาโปธาตุ", c: "var(--water)", l: "var(--water-l)", mn: "ของเหลว: เสมหะ เลือด หนอง เหงื่อ ปัสสาวะ" },
  { key: "ลม",   code: "A04", title: "🌬️ ลม (วาโย)",  prefix: "วาโยธาตุ", c: "var(--wind)",  l: "var(--wind-l)",  mn: "การเคลื่อนไหว: หาวเรอ ผายลม เมื่อย วิงเวียน" },
  { key: "อากาศ", code: "A05", title: "🌌 อากาศ",       prefix: "อากาศธาตุ", c: "var(--space)", l: "var(--space-l)", mn: "ช่องว่างในกาย: หู ตา (ข้อมูลในตำราน้อย)" },
];
const STAGES = ["หย่อน", "กำเริบ", "พิการ", "ออกจากกาย", "แตก", "ขาด"];
const STAGE_COLORS = ["#FFF2CC", "#FFE699", "#F8CBAD", "#F4B183", "#E06666", "#990000"];

let ROWS = [];

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s).replace(/[&<>"]/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[m]));
const shuffle = (a) => { a = [...a]; for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
const realSymptoms = (r) => r.symptoms.filter((s) => s !== "/");
const elementOf = (desc) => ELEMENTS.filter((e) => desc.includes(e.key));
const asEl = (t) => (t && t.nodeType === 1 ? t : t && t.parentElement);

/* ---------- ลดการคัดลอก (ยับยั้งผู้ใช้ทั่วไปได้ระดับหนึ่ง) ---------- */
function protectPage() {
  const isField = (t) => { const el = asEl(t); return el && el.closest("input, select, textarea"); };
  document.addEventListener("contextmenu", (e) => e.preventDefault());
  ["copy", "cut"].forEach((ev) => document.addEventListener(ev, (e) => { if (!isField(e.target)) e.preventDefault(); }));
  document.addEventListener("selectstart", (e) => { if (!isField(e.target)) e.preventDefault(); });
  document.addEventListener("dragstart", (e) => { const el = asEl(e.target); if (!el || !el.closest(".dchip")) e.preventDefault(); });
  document.addEventListener("keydown", (e) => {
    const k = (e.key || "").toLowerCase(), ctrl = e.ctrlKey || e.metaKey;
    if (e.key === "F12" || (ctrl && ["u", "s"].includes(k)) || (ctrl && e.shiftKey && ["i", "j", "c"].includes(k))) e.preventDefault();
  });
}

/* ---------- คัดลอกลิงก์ ---------- */
async function copyLink() {
  const url = location.href;
  try { await navigator.clipboard.writeText(url); toast("🔗 คัดลอกลิงก์แล้ว ส่งให้เพื่อนเปิดอ่านทบทวนได้เลย"); }
  catch (e) { window.prompt("คัดลอกลิงก์นี้:", url); }
}
function toast(msg) {
  const t = $("toast"); t.textContent = msg; t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), 2200);
}

/* ---------- แท็บเมนู ---------- */
function showTab(id) {
  if (!TABS.includes(id)) id = "cards";
  document.querySelectorAll(".panel").forEach((p) => p.classList.toggle("active", p.id === id));
  document.querySelectorAll("nav a").forEach((a) => a.classList.toggle("active", a.dataset.tab === id));
  window.scrollTo({ top: 0, behavior: "smooth" });
}

/* ---------- โหลดข้อมูล ---------- */
function findDataSheet(wb) {
  const name = wb.SheetNames.find((n) => String((wb.Sheets[n]["B1"] || {}).v || "").trim().toUpperCase() === "ABCODE");
  return wb.Sheets[name || wb.SheetNames[0]];
}

function parseWorkbook(wb) {
  const grid = XLSX.utils.sheet_to_json(findDataSheet(wb), { header: 1, defval: "" });
  return grid.slice(2)
    .filter((r) => String(r[1]).trim() !== "")
    .map((r) => ({
      code: String(r[1]).trim(),
      name: String(r[2]).trim(),
      desc: String(r[3]).trim(),
      level: String(r[4]).trim(),
      condition: String(r[5]).trim(),
      symptoms: r.slice(6, 26).map((s) => String(s).trim()).filter((s) => s !== ""),
    }));
}

async function fetchFirst() {
  const tried = [];
  for (const url of DATA_URLS) {
    try {
      const res = await fetch(url, { cache: "no-cache" });
      if (res.ok) return await res.arrayBuffer();
      tried.push(url + " (" + res.status + ")");
    } catch (e) { tried.push(url + " (" + e.message + ")"); }
  }
  throw new Error("ไม่พบไฟล์ข้อมูล: " + tried.join(", "));
}

async function loadData() {
  try {
    if (typeof XLSX === "undefined") throw new Error("โหลดตัวอ่าน Excel (SheetJS) ไม่สำเร็จ");
    ROWS = parseWorkbook(XLSX.read(await fetchFirst(), { type: "array" }));
    if (!ROWS.length) throw new Error("อ่านไฟล์ได้ แต่ไม่พบแถวข้อมูล (ต้องมีหัวคอลัมน์ ABCODE ที่ B1 และข้อมูลเริ่มแถว 3)");
    $("loader").hidden = true;
    $("app").hidden = false;
    renderCards(); renderLadder(); renderMap(); setupExplore(); newQuiz(); newDrag();
    showTab(location.hash.slice(1));
  } catch (e) {
    console.warn(e);
    $("status").className = "status err";
    $("status").textContent = "⚠️ ยังโหลดข้อมูลไม่ได้ – " + e.message;
  }
}

/* ---------- ① การ์ดธาตุ ---------- */
function renderCards() {
  $("cardGrid").innerHTML = ELEMENTS.map((e) => {
    const rows = ROWS.filter((r) => r.code === e.code);
    const nSym = rows.reduce((s, r) => s + realSymptoms(r).length, 0);
    const first = rows[0];
    return `<div class="ecard" style="--c:${e.c};--l:${e.l}" data-code="${e.code}">
      <div class="head">${e.title}</div>
      <div class="body">
        <div class="code">${e.code}</div>
        <div class="big">${rows.length}</div><div class="lbl">ภาวะ/โรค</div>
        <div class="big">${nSym}</div><div class="lbl">ช่องอาการ*</div>
        <div class="mn">🔑 ${esc(e.mn)}</div>
        ${first ? `<div class="ex"><b>${esc(first.condition)}</b><br>${esc(realSymptoms(first).slice(0, 4).join(" • "))}</div>` : ""}
      </div></div>`;
  }).join("");
  document.querySelectorAll(".ecard").forEach((el) =>
    el.addEventListener("click", () => { $("groupSelect").value = el.dataset.code; $("searchBox").value = ""; renderExplore(); location.hash = "explore"; }));
}

/* ---------- ② บันไดความรุนแรง ---------- */
function renderLadder() {
  let h = `<tr><th>ธาตุ</th>${STAGES.map((s, i) => `<th style="background:${STAGE_COLORS[i]};color:${i >= 4 ? "#fff" : "#222"}">ขั้น ${i + 1}${i === 0 ? " (เบา)" : i === 5 ? " (หนักสุด) ☠" : ""} ➜<br><b>${s}</b></th>`).join("")}</tr>`;
  ELEMENTS.slice(0, 4).forEach((e) => {
    h += `<tr><td class="el" style="background:${e.l}">${e.title}</td>`;
    STAGES.forEach((s) => {
      const r = ROWS.find((x) => x.condition.startsWith(e.prefix + s));
      h += r ? `<td class="cell" data-cond="${esc(r.condition)}">${esc(realSymptoms(r).slice(0, 3).join(" • "))}</td>` : `<td>—</td>`;
    });
    h += "</tr>";
  });
  $("ladderTable").innerHTML = h;
  document.querySelectorAll(".ladder td.cell").forEach((td) =>
    td.addEventListener("click", () => showDetail(ROWS.find((r) => r.condition === td.dataset.cond))));
}

function showDetail(r) {
  if (!r) return;
  $("detailTitle").textContent = `${r.condition} (${r.code} · ${r.desc})`;
  $("detailBody").innerHTML = realSymptoms(r).map((s) => `<span class="chip">${esc(s)}</span>`).join("");
  $("detail").showModal();
}

/* ---------- ③ แผนที่กลุ่มธาตุ ---------- */
function groups() {
  const map = new Map();
  ROWS.forEach((r) => {
    if (!map.has(r.code)) map.set(r.code, { code: r.code, desc: r.desc, level: r.level, n: 0, sym: 0 });
    const g = map.get(r.code); g.n++; g.sym += realSymptoms(r).length;
  });
  return [...map.values()];
}

function renderMap() {
  const gs = groups();
  let h = `<tr><th>กลุ่มอาการ</th><th>รหัส</th><th>ระดับ</th><th>ภาวะ</th><th>ช่องอาการ</th>${ELEMENTS.map((e) => `<th style="background:${e.c}">${e.key}</th>`).join("")}</tr>`;
  gs.forEach((g) => {
    h += `<tr class="lv-${g.level}"><td>${esc(g.desc)}</td><td>${g.code}</td><td>${g.level}</td><td>${g.n}</td><td>${g.sym}</td>`;
    ELEMENTS.forEach((e) => {
      h += g.desc.includes(e.key) ? `<td style="background:${e.l};color:${e.c}" class="dot">●</td>` : "<td></td>";
    });
    h += "</tr>";
  });
  $("mapTable").innerHTML = h;

  const max = Math.max(...gs.map((g) => g.n));
  $("barChart").innerHTML = `<h4>จำนวนภาวะ/โรค ในแต่ละกลุ่มธาตุ</h4>` + gs.map((g) =>
    `<div class="bar-row"><span>${esc(g.desc)}</span><div class="bar" style="width:${(g.n / max) * 100}%"></div><b>${g.n}</b></div>`).join("");
}

/* ---------- ④ สำรวจเอง ---------- */
function setupExplore() {
  $("groupSelect").innerHTML = `<option value="">— ทุกกลุ่ม —</option>` +
    groups().map((g) => `<option value="${g.code}">${g.code} · ${esc(g.desc)}</option>`).join("");
  $("groupSelect").value = "A01";
  $("groupSelect").onchange = renderExplore;
  $("searchBox").oninput = renderExplore;
  renderExplore();
}

function renderExplore() {
  const code = $("groupSelect").value, q = $("searchBox").value.trim();
  const rows = ROWS.filter((r) => (!code || r.code === code) && (!q || r.condition.includes(q) || r.symptoms.some((s) => s.includes(q))));
  $("exploreCount").textContent = `${rows.length} ภาวะ`;
  $("exploreList").innerHTML = rows.map((r) => {
    const el = elementOf(r.desc)[0] || ELEMENTS[0];
    return `<div class="cond" style="--c:${el.c}"><div class="t">${esc(r.condition)} <small>(${r.code} · ${esc(r.desc)})</small></div>
      <div class="chips">${realSymptoms(r).map((s) => `<span class="chip ${q && s.includes(q) ? "hit" : ""}">${esc(s)}</span>`).join("")}</div></div>`;
  }).join("") || "<p>ไม่พบข้อมูล</p>";
}

/* ---------- ⑤ แบบทดสอบ ก) เลือกตอบ ---------- */
let quizState = { total: 0, right: 0 };

function newQuiz() {
  const pool = ROWS.filter((r) => ELEMENTS.some((e) => e.code === r.code) && realSymptoms(r).length >= 3);
  const qs = shuffle(pool).slice(0, 8);
  quizState = { total: qs.length, right: 0 };
  $("quizScore").textContent = `0 / ${qs.length}`;
  $("quizList").innerHTML = qs.map((r, i) => `
    <div class="q" data-i="${i}">
      <b>ข้อ ${i + 1}: อาการเหล่านี้เป็นของธาตุใด?</b>
      <div class="hints">${realSymptoms(r).slice(0, 4).map((s) => `<span class="hint">${esc(s)}</span>`).join("")}</div>
      <div class="choices">${ELEMENTS.map((e) => `<button class="choice" style="--c:${e.c};--l:${e.l}" data-k="${e.key}">${e.title}</button>`).join("")}</div>
      <div class="res"></div>
    </div>`).join("");
  document.querySelectorAll(".q").forEach((qEl) => {
    const r = qs[+qEl.dataset.i];
    qEl.querySelectorAll(".choice").forEach((btn) => btn.addEventListener("click", () => {
      const ok = r.desc.includes(btn.dataset.k);
      const res = qEl.querySelector(".res");
      if (ok) {
        if (!qEl.dataset.done && !qEl.dataset.tried) quizState.right++;
        qEl.dataset.done = "1";
        res.className = "res ok"; res.innerHTML = `✅ ถูก! เฉลย: ${esc(r.condition)}`;
        qEl.querySelectorAll(".choice").forEach((b) => (b.disabled = true));
      } else {
        qEl.dataset.tried = "1"; btn.disabled = true;
        res.className = "res bad"; res.textContent = "❌ ลองใหม่ (ตอบถูกในครั้งแรกเท่านั้นจึงได้คะแนน)";
      }
      $("quizScore").textContent = `${quizState.right} / ${quizState.total}`;
    }));
  });
}

/* ---------- ⑤ แบบทดสอบ ข) ลากวางจับคู่ ---------- */
let selectedChip = null;

function newDrag() {
  const main4 = ELEMENTS.slice(0, 4);
  const items = shuffle(main4.flatMap((e) => {
    const syms = ROWS.filter((r) => r.code === e.code).flatMap((r) => realSymptoms(r).filter((s) => s.length > 3));
    return shuffle([...new Set(syms)]).slice(0, 2).map((s) => ({ s, key: e.key }));
  }));
  selectedChip = null;
  $("dragScore").textContent = "–";
  $("dragPool").innerHTML = items.map((it, i) => `<span class="dchip" draggable="true" id="dc${i}" data-key="${it.key}">${esc(it.s)}</span>`).join("");
  $("buckets").innerHTML = main4.map((e) => `<div class="bucket" style="--c:${e.c};--l:${e.l}" data-key="${e.key}"><h4>${e.title}</h4></div>`).join("");

  document.querySelectorAll(".dchip").forEach((c) => {
    c.addEventListener("dragstart", (ev) => ev.dataTransfer.setData("text/plain", c.id));
    c.addEventListener("click", (ev) => {
      ev.stopPropagation();
      document.querySelectorAll(".dchip").forEach((x) => x.classList.remove("sel"));
      selectedChip = selectedChip === c ? null : c;
      if (selectedChip) c.classList.add("sel");
    });
  });
  [...document.querySelectorAll(".bucket"), $("dragPool")].forEach((b) => {
    b.addEventListener("dragover", (ev) => { ev.preventDefault(); b.classList.add("over"); });
    b.addEventListener("dragleave", () => b.classList.remove("over"));
    b.addEventListener("drop", (ev) => { ev.preventDefault(); b.classList.remove("over"); moveChip($(ev.dataTransfer.getData("text/plain")), b); });
    b.addEventListener("click", () => { if (selectedChip) moveChip(selectedChip, b); });
  });
}

function moveChip(chip, target) {
  if (!chip) return;
  chip.classList.remove("sel", "right", "wrong");
  target.appendChild(chip);
  selectedChip = null;
}

function checkDrag() {
  let right = 0, total = 0;
  document.querySelectorAll(".dchip").forEach((c) => {
    total++;
    const b = c.closest(".bucket");
    c.classList.remove("right", "wrong");
    if (!b) return;
    if (b.dataset.key === c.dataset.key) { c.classList.add("right"); right++; } else c.classList.add("wrong");
  });
  $("dragScore").textContent = `${right} / ${total}` + (right === total ? " 🎉 เยี่ยมมาก!" : "");
}

/* ---------- เริ่มต้น ---------- */
document.addEventListener("DOMContentLoaded", () => {
  protectPage();
  $("copyLink").onclick = copyLink;
  $("newQuiz").onclick = newQuiz;
  $("newDrag").onclick = newDrag;
  $("checkDrag").onclick = checkDrag;
  window.addEventListener("hashchange", () => showTab(location.hash.slice(1)));
  loadData();
});


