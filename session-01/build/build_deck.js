// Session 01 deck: Power BI Internals (PBIP, TMDL, PBIR).
// Rebuild:  NODE_PATH=<folder with pptxgenjs> node build_deck.js
// Each "Open interactive map" pill links to a view of the published map artifact (MAP + "#view").
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const MAP = "https://rishav-portfolio-i46g.onrender.com/training/power-bi-internals-map.html";
const OUT = process.env.OUT || path.join(__dirname, "Session01_Power_BI_Internals.pptx");
const ASSETS = path.join(__dirname, "assets");

const THEME = {
  name: "Power BI Internals",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "14171D", lt1: "FFFFFF", dk2: "1E2330", lt2: "F4F5F8",
    accent1: "5B3FD9", // report · PBIR
    accent2: "0B7C74", // model · TMDL
    accent3: "C27A00", // focus / live
    accent4: "2F8A3E", // ok
    accent5: "C23A2A", // stop
    accent6: "5A6170", // muted
    hlink: "F2B544", folHlink: "F2B544",
  },
};
const HEX = { line: "D9DDE4", lineDark: "3A4150" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Session 01 - Power BI Internals: PBIP, TMDL and PBIR";
const C = pres.SchemeColor;
const MONO = "Consolas";
const W = 13.333;

/* ---------------- layouts ---------------- */
const footer = (dark) => ({ text: { text: "Session 01 · Power BI Internals: PBIP, TMDL and PBIR", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 9, color: dark ? C.background2 : C.accent6, transparency: dark ? 45 : 0, margin: 0 } } });
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    footer(false),
    { placeholder: { options: { name: "eyebrow", type: "body", x: 0.6, y: 0.42, w: 9.6, h: 0.3, fontSize: 11, bold: true, color: C.accent6, charSpacing: 1.5, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.74, w: 9.8, h: 0.8, fontSize: 32, bold: true, color: C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
  ],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.3, fontSize: 9, color: "5A6170", align: "right" },
});
pres.defineSlideMaster({
  title: "LIVE",
  background: { color: C.text2 },
  objects: [
    footer(true),
    { placeholder: { options: { name: "eyebrow", type: "body", x: 0.6, y: 0.42, w: 9.6, h: 0.3, fontSize: 11, bold: true, color: C.accent3, charSpacing: 1.5, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.74, w: 12.1, h: 0.8, fontSize: 34, bold: true, color: C.background1, margin: 0, valign: "top", align: "left" }, text: "" } },
  ],
  slideNumber: { x: 12.23, y: 7.0, w: 0.5, h: 0.3, fontSize: 9, color: "9AA1AD", align: "right" },
});
pres.defineSlideMaster({
  title: "DARK",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "eyebrow", type: "body", x: 0.8, y: 1.55, w: 11.5, h: 0.35, fontSize: 13, bold: true, color: C.accent3, charSpacing: 2, margin: 0, valign: "top" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.95, w: 11.7, h: 1.25, fontSize: 52, bold: true, color: C.background1, margin: 0, valign: "top", align: "left" }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: 0.8, y: 3.25, w: 10.5, h: 0.9, fontSize: 22, color: C.background2, margin: 0, valign: "top" }, text: "" } },
  ],
});

/* ---------------- helpers ---------------- */
let sectionTitle = "Opening";
function slide(master, eyebrow, title, notes) {
  const s = pres.addSlide({ masterName: master, sectionTitle });
  if (eyebrow) s.addText(eyebrow, { placeholder: "eyebrow" });
  if (title) s.addText(title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}
function section(t) { sectionTitle = t; pres.addSection({ title: t }); }

function pill(s, view, dark = false) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.5, y: 0.5, w: 2.23, h: 0.44, rectRadius: 0.22, fill: { color: dark ? C.background1 : C.text1, transparency: dark ? 88 : 0 }, line: { color: dark ? C.accent3 : C.text1, width: 1 }, objectName: "Map link" });
  s.addText([{ text: "Open interactive map  ↗", options: { color: dark ? "F2B544" : C.background1, hyperlink: { url: `${MAP}#${view}`, tooltip: "Opens this map in the browser" } } }], { x: 10.5, y: 0.5, w: 2.23, h: 0.44, fontSize: 12, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Map link text" });
}
function card(s, x, y, w, h, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: o.fill || C.background2, transparency: o.t || 0 }, line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" }, objectName: o.name || "Card" });
}
function txt(s, text, x, y, w, h, o = {}) {
  s.addText(text, Object.assign({ x, y, w, h, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true }, o));
}
function num(s, x, y, n, color = C.accent3, d = 0.42) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { type: "none" } });
  s.addText(String(n), { x, y, w: d, h: d, fontSize: d > 0.45 ? 16 : 13, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
}
function arrow(s, x1, y1, x2, y2, color = C.accent6, wpt = 1.5, both = false) {
  const o = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001), line: { color, width: wpt, endArrowType: "triangle" } };
  if (both) o.line.beginArrowType = "triangle";
  if (x2 < x1) o.flipH = true;
  if (y2 < y1) o.flipV = true;
  s.addShape(pres.shapes.LINE, o);
}
// Syntax colouring for code blocks: TMDL keywords, property names, strings, comments, diff adds.
const KW = /^(\s*)(table|column|measure|partition|relationship|ref table|annotation)\b/;
function tokens(line, lang) {
  if (line.startsWith("+")) return [[line, "a"]];
  const segs = [];
  let rest = line;
  const cm = rest.indexOf("//") >= 0 && lang !== "json" ? rest.indexOf("//") : rest.indexOf("  ← ");
  let comment = "";
  if (cm > 0) { comment = rest.slice(cm); rest = rest.slice(0, cm); }
  if (lang === "tmdl") {
    const k = rest.match(KW);
    if (k) { segs.push([k[1], "x"], [k[2], "k"]); rest = rest.slice(k[0].length); }
    else { const p = rest.match(/^(\s*)(\w+)(:)/); if (p) { segs.push([p[1], "x"], [p[2], "p"], [":", "x"]); rest = rest.slice(p[0].length); } }
    rest.split(/("[^"]*")/).forEach(t => t && segs.push([t, t.startsWith('"') ? "s" : "x"]));
  } else if (lang === "json") {
    rest.split(/("[^"]*")/).forEach((t, i, arr) => { if (!t) return; if (t.startsWith('"')) segs.push([t, /^\s*:/.test(arr[i + 1] || "") ? "p" : "s"]); else segs.push([t, "x"]); });
  } else segs.push([rest, "x"]);
  if (comment) segs.push([comment, "c"]);
  return segs;
}
function code(s, x, y, w, h, lines, o = {}) {
  const dark = !!o.dark, size = o.size || 12, lang = o.lang || "tmdl";
  const col = { x: dark ? C.background1 : C.text1, k: dark ? "B6A6FF" : C.accent1, p: dark ? "5CD3C6" : C.accent2, s: dark ? "F2B544" : C.accent3, c: dark ? "9AA1AD" : C.accent6, a: dark ? "63C474" : C.accent4, h: C.accent3 };
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: dark ? C.text1 : C.background2 }, line: { color: dark ? HEX.lineDark : HEX.line, width: 0.75 }, objectName: "Code block" });
  const runs = [];
  lines.forEach((ln, i) => {
    // An empty run makes pptxgenjs write a second <a:pPr> into the paragraph, which PowerPoint refuses to open.
    const segs = (Array.isArray(ln) ? ln : tokens(ln, lang)).filter(([t]) => t !== "");
    if (!segs.length) segs.push([" ", "x"]);
    segs.forEach(([t, k], j) => runs.push({ text: t, options: { color: col[k] || col.x, bold: k === "h" || k === "a", breakLine: j === segs.length - 1 && i < lines.length - 1 } }));
  });
  s.addText(runs, { x: x + 0.2, y: y + 0.16, w: w - 0.4, h: h - 0.32, fontFace: MONO, fontSize: size, valign: "top", margin: 0, isTextBox: true, lineSpacingMultiple: 1.05, objectName: "Code" });
}
function bullets(items, o = {}) {
  return items.map((t, i) => ({ text: t, options: Object.assign({ bullet: o.num ? { type: "number" } : true, breakLine: i < items.length - 1, paraSpaceAfter: o.gap ?? 6 }, o.run || {}) }));
}
function liveSteps(s, steps, x = 0.6, y = 1.95, w = 7.0) {
  steps.forEach((st, i) => {
    const yy = y + i * 1.08;
    num(s, x, yy, i + 1, C.accent3, 0.5);
    txt(s, st[0], x + 0.75, yy - 0.02, w - 0.75, 0.4, { fontSize: 19, bold: true, color: C.background1 });
    if (st[1]) txt(s, st[1], x + 0.75, yy + 0.42, w - 0.75, 0.55, { fontSize: 14, color: C.background2, transparency: 25 });
  });
}
function askCard(s, x, y, w, h, qa) {
  card(s, x, y, w, h, { fill: C.background1, t: 92, line: HEX.lineDark, name: "Ask the room" });
  txt(s, "ASK THE ROOM", x + 0.3, y + 0.25, w - 0.6, 0.3, { fontSize: 11, bold: true, color: C.accent3, charSpacing: 1.5 });
  const runs = [];
  qa.forEach(([q, a], i) => {
    runs.push({ text: q, options: { bold: true, color: C.background1, breakLine: true } });
    runs.push({ text: a, options: { color: C.background2, transparency: 25, breakLine: i < qa.length - 1, paraSpaceAfter: 14 } });
  });
  txt(s, runs, x + 0.3, y + 0.65, w - 0.6, h - 0.85, { fontSize: 15 });
}

/* =========================================================
   OPENING
   ========================================================= */
section("Opening");
{
  const s = slide("DARK", "SESSION 01 · CLAUDE + POWER BI PROGRAMME", "Power BI Internals",
    "Open with the promise: by the end, everyone will know exactly what an AI is reading when it reads a Power BI report. Every example today comes from real projects, not invented syntax, and every status claim is checked against Microsoft Learn as of September 2026.\n\nThe arc has four steps: understand the project, understand its metadata, use AI to interpret it, then introduce MCP.");
  s.addText("PBIP, TMDL and PBIR: what an AI actually reads and changes", { placeholder: "sub" });
  const arc = ["Understand the project", "Understand its metadata", "Use AI to interpret it", "Introduce MCP"];
  arc.forEach((a, i) => {
    const x = 0.8 + i * 2.98;
    card(s, x, 4.75, 2.62, 0.9, { fill: C.background1, t: 91, line: HEX.lineDark, name: "Arc step" });
    txt(s, String(i + 1).padStart(2, "0"), x + 0.22, 4.88, 0.6, 0.3, { fontSize: 11, bold: true, color: C.accent3, fontFace: MONO });
    txt(s, a, x + 0.22, 5.16, 2.3, 0.4, { fontSize: 15, bold: true, color: C.background1 });
    if (i < arc.length - 1) arrow(s, x + 2.66, 5.2, x + 2.94, 5.2, C.accent3, 1.5);
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 6.15, w: 3.0, h: 0.5, rectRadius: 0.25, fill: { color: C.background1, transparency: 88 }, line: { color: C.accent3, width: 1 } });
  s.addText([{ text: "Open the interactive map  ↗", options: { color: "F2B544", hyperlink: { url: `${MAP}#overview`, tooltip: "Opens the interactive map" } } }], { x: 0.8, y: 6.15, w: 3.0, h: 0.5, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  txt(s, "Eight clickable maps built from the Meridian Medical Supply project", 4.0, 6.27, 7, 0.3, { fontSize: 13, color: C.background2, transparency: 35 });
}
{
  const s = slide("CONTENT", "THE PLAN · 120 MINUTES", "Today's two hours, module by module",
    "Two changes from the original draft: the AI exercises get 20 minutes instead of 15 (exercises 1 and 3 live, 2 and 4 as the lab), and Module 1 gives five minutes back because seniors already know PBIX.\n\nIf the programme stays at 60-minute sessions, split it: 01a = Modules 1 to 4 (the files), 01b = Modules 5 to 7 (reading them with AI). Recover the session by folding the old 'Model as code' session into 01a. Never compress 120 minutes into 60: Module 3 is what gets cut, and it is the most important.");
  const seg = [[10, "1", "Formats", C.accent6], [20, "2", "PBIP folder", C.text1], [25, "3", "TMDL", C.accent2], [10, "4", "Relationships", C.accent2], [15, "5", "PBIR", C.accent1], [20, "6", "AI without MCP", C.accent3], [15, "7", "MCP", C.accent3], [5, "Q", "Recap", C.accent6]];
  const x0 = 0.6, span = 12.13, per = span / 120; let x = x0, t = 0;
  seg.forEach(([m, n, label, color], i) => {
    const w = m * per;
    s.addShape(pres.shapes.RECTANGLE, { x: x + 0.02, y: 2.05, w: w - 0.04, h: 0.62, fill: { color }, line: { type: "none" } });
    txt(s, n, x + 0.02, 2.05, w - 0.04, 0.62, { fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
    const ly = 2.85 + (i % 2) * 0.62;
    s.addShape(pres.shapes.LINE, { x: x + 0.02, y: 2.67, w: 0, h: ly - 2.67 + 0.05, line: { color: HEX.line, width: 1 } });
    const last = i === seg.length - 1;
    txt(s, [{ text: `${String(t).padStart(2, "0")}–${String(t + m).padStart(2, "0")}  `, options: { color: C.accent6, fontFace: MONO } }, { text: label, options: { bold: true } }], last ? x0 + span - 2.6 : x + 0.1, ly, 2.6, 0.35, { fontSize: 13, align: last ? "right" : "left" });
    x += w; t += m;
  });
  card(s, 0.6, 4.35, 5.9, 1.85, { name: "Protect module 3" });
  num(s, 0.85, 4.6, 3, C.accent2);
  txt(s, "Protect Module 3", 1.42, 4.62, 4.8, 0.4, { fontSize: 18, bold: true });
  txt(s, "The semantic model is what every visual queries. If time runs short, cut from Module 1, never from TMDL.", 1.42, 5.08, 4.85, 1.3, { fontSize: 14, color: C.accent6 });
  card(s, 6.83, 4.35, 5.9, 1.85, { name: "60 minute split" });
  txt(s, "If sessions stay at 60 minutes", 7.1, 4.6, 5.4, 0.4, { fontSize: 18, bold: true });
  txt(s, [
    { text: "01a · The files  ", options: { bold: true } }, { text: "Modules 1–4, lab: three facts from the files alone", options: { color: C.accent6, breakLine: true, paraSpaceAfter: 8 } },
    { text: "01b · Reading them with AI  ", options: { bold: true } }, { text: "Modules 5–7, lab: exercises 2 and 4", options: { color: C.accent6 } },
  ], 7.1, 5.1, 5.4, 1.3, { fontSize: 14 });
}
{
  const s = slide("CONTENT", "WHY THIS COMES FIRST", "An AI can only change what it can read",
    "PBIX and PBIP hold the same report. PBIP spreads it across readable text files, and that is the only reason an AI can read and change it.\n\nSay out loud: Microsoft introduced PBIP for source control, readable diffs, merging, and CI/CD through Fabric Git integration. AI-assisted development is a consequence of that, not the original reason.");
  txt(s, "PBIP spreads a report across readable text files.", 0.6, 2.0, 5.4, 1.3, { fontSize: 28, bold: true, fontFace: "Arial" });
  txt(s, "That is the only reason an AI can read it and change it. Before we automate anything, we learn exactly what those files are.", 0.6, 3.45, 5.2, 1.6, { fontSize: 17, color: C.accent6 });
  const why = [["Source control", "Every change to the report or the model is a commit you can review."], ["Readable diffs", "A one-line change in Desktop is a one-line change in the file."], ["Merging", "Two developers can change different tables and Git merges them."], ["CI/CD", "Fabric Git integration deploys the same files across workspaces."]];
  why.forEach(([h, b], i) => {
    const x = 6.55 + (i % 2) * 3.14, y = 2.0 + Math.floor(i / 2) * 2.05;
    card(s, x, y, 2.98, 1.85);
    txt(s, h, x + 0.25, y + 0.25, 2.5, 0.4, { fontSize: 18, bold: true, color: i % 2 ? C.accent2 : C.accent1 });
    txt(s, b, x + 0.25, y + 0.72, 2.5, 1.0, { fontSize: 14, color: C.text1 });
  });
  txt(s, "AI-assisted development is a consequence of these, not the original reason.", 6.55, 6.2, 6.2, 0.4, { fontSize: 14, italic: true, color: C.accent6 });
}

{
  const s = slide("CONTENT", "THE DATASET FOR ALL EIGHT SESSIONS", "Meet Meridian Medical Supply",
    "One fictional company, used in every session. Meridian distributes medical consumables from 5 US distribution centres to 150 hospitals, clinics, pharmacies and care homes, Jan 2024 to Dec 2025. Fiscal year starts 1 July.\n\nThe data contains planted stories for participants to find. Keep them to yourself until they do: 2025 revenue +10.1%; flu-season peaks in PPE and Diagnostics; the Infusion Pump Tubing Set goes from +13% margin to -7% in 2025 as discounts rise; Pacific Rim Medical's on-time rate falls from about 90% to 30% in H2 2025 and stockouts follow; Swift Freight's on-time delivery drops to 73% in H2 2025.\n\nA separate data_quality_lab copy has 8 deliberate defects (orphans, duplicates, text in numbers, mixed dates) for the validation sessions, and sql/01_create_and_load.sql loads the same data into SQL Server for reconciliation.");
  const stats = [["18,470", "sales order lines"], ["2,073", "purchase order lines"], ["7,200", "month-end stock rows"], ["$27.0M", "revenue over two years"]];
  stats.forEach(([n, d], i) => {
    const y = 1.85 + i * 1.18;
    txt(s, n, 0.6, y, 3.6, 0.65, { fontSize: 36, bold: true, color: i === 3 ? C.accent3 : C.accent2, fontFace: "Arial" });
    txt(s, d, 0.6, y + 0.66, 3.6, 0.35, { fontSize: 14, color: C.accent6 });
  });
  const group = (x, y, w, h, title, col, items, note) => {
    card(s, x, y, w, h, { fill: col, t: 92, name: title });
    txt(s, title.toUpperCase(), x + 0.25, y + 0.18, w - 0.5, 0.3, { fontSize: 11, bold: true, color: col, charSpacing: 1.5 });
    items.forEach((it, i) => {
      const bx = x + 0.25 + i * ((w - 0.5 - (items.length - 1) * 0.2) / items.length + 0.2), bw = (w - 0.5 - (items.length - 1) * 0.2) / items.length;
      card(s, bx, y + 0.55, bw, 0.62, { fill: C.background1, line: col, name: it[0] });
      txt(s, it[0], bx, y + 0.55, bw, 0.36, { fontSize: 13, bold: true, fontFace: MONO, align: "center", valign: "bottom" });
      txt(s, it[1], bx, y + 0.9, bw, 0.25, { fontSize: 10.5, color: C.accent6, align: "center" });
    });
    if (note) txt(s, note, x + 0.25, y + h - 0.42, w - 0.5, 0.3, { fontSize: 12, color: C.accent6 });
  };
  group(4.6, 1.85, 8.13, 1.55, "Shared by every fact", C.accent2, [["Dim_Date", "731 days, fiscal July"], ["Dim_Product", "60 in 6 categories"], ["Dim_Warehouse", "5 regional DCs"]]);
  group(4.6, 3.55, 8.13, 1.55, "Three facts, three grains", C.text1, [["Fact_Sales", "order line"], ["Fact_Purchases", "PO line"], ["Fact_Inventory", "month-end snapshot"]]);
  group(4.6, 5.25, 8.13, 1.55, "Sales only, and a snowflake", C.accent1, [["Dim_Customer", "150 customers"], ["Dim_Carrier", "4 carriers"], ["Dim_Supplier", "12, via product"]]);
}
{
  const s = slide("CONTENT", "THE DATASET FOR ALL EIGHT SESSIONS", "The report we will take apart",
    "This is the real report, saved by Power BI Desktop as a PBIP project: three pages, 23 visuals, 30 measures, 15 relationships. Every file shown in this session comes from it.\n\nOpen pbip/MeridianSupplyChain.pbip. First time on a new machine: Transform data, Edit parameters, set DataFolder to the data folder, then Refresh.");
  pill(s, "report");
  const iw = 8.4, ih = iw * 1356 / 2395;
  card(s, 0.6, 1.8, iw, ih, { fill: C.background1, line: HEX.line, name: "Report frame" });
  s.addImage({ path: path.join(ASSETS, "report_executive.png"), x: 0.6, y: 1.8, w: iw, h: ih, altText: "Executive Overview page of the Meridian report in Power BI Desktop" });
  const tw = 3.38, th = tw * 1356 / 2395;
  [["report_supplier.png", "Supplier Performance"], ["report_inventory.png", "Inventory Health"]].forEach(([f, l], i) => {
    const y = 1.8 + i * (th + 0.6);
    card(s, 9.35, y, tw, th, { fill: C.background1, line: HEX.line, name: l });
    s.addImage({ path: path.join(ASSETS, f), x: 9.35, y, w: tw, h: th, altText: l + " page of the Meridian report" });
    txt(s, l, 9.35, y + th + 0.06, tw, 0.3, { fontSize: 12, bold: true });
  });
  txt(s, [{ text: "Executive Overview", options: { bold: true } }, { text: "   ·   MeridianSupplyChain.Report/definition/pages/ExecutiveOverview", options: { fontFace: MONO, color: C.accent6 } }], 0.6, 1.8 + ih + 0.08, iw, 0.3, { fontSize: 12 });
}

/* =========================================================
   PART 1 · THE FILES
   ========================================================= */
section("Part 1 · The files");
{
  const s = slide("DARK", "PART 1 · MODULES 1–5", "The files", "Part 1 is about what Power BI writes to disk. Open the real folder in VS Code for most of it; slides are only the headlines.");
  s.addText("What Power BI writes to disk, and which parts you can safely change", { placeholder: "sub" });
  ["1  File formats", "2  PBIP folder", "3  TMDL", "4  Relationships", "5  PBIR"].forEach((m, i) => txt(s, m, 0.8 + i * 2.4, 5.1, 2.3, 0.4, { fontSize: 16, bold: true, color: i === 2 || i === 3 ? "5CD3C6" : i === 4 ? "B6A6FF" : C.background1 }));
}
{
  const s = slide("CONTENT", "MODULE 1 · FILE FORMATS", "One report, two ways to save it",
    "Demo: open the same report as a PBIX and as a PBIP side by side. Rename the PBIX to .zip to show it is a binary container, then open the PBIP folder in VS Code.\n\nAsk: 'If two developers change different measures in the same PBIX, what happens when you merge?' You cannot merge; one overwrites the other. In PBIP, if the measures are in different tables, Git merges them cleanly.\n\nMisconception: 'PBIP is a different report.' It is the same report saved differently; Save as converts both ways.");
  pill(s, "formats");
  card(s, 0.6, 2.45, 3.0, 2.9, { line: HEX.line, name: "PBIX" });
  txt(s, "Sales.pbix", 0.9, 2.75, 2.5, 0.4, { fontSize: 18, bold: true, fontFace: MONO });
  txt(s, "model + data + report", 0.9, 3.25, 2.5, 0.35, { fontSize: 14, color: C.accent6 });
  txt(s, "One zipped binary file. Rename it to .zip and you still cannot read or diff it.", 0.9, 3.85, 2.45, 1.3, { fontSize: 14 });
  arrow(s, 3.75, 3.9, 5.25, 3.9, C.text1, 2);
  txt(s, "File → Save as\nPower BI project", 3.7, 3.15, 1.6, 0.65, { fontSize: 12, color: C.accent6, align: "center" });
  card(s, 5.4, 1.85, 7.33, 4.55, { fill: C.background1, line: C.text1, lw: 1.5, name: "PBIP folder" });
  txt(s, "Sales.pbip  +  two folders", 5.7, 2.05, 6, 0.4, { fontSize: 16, bold: true, fontFace: MONO });
  [[C.accent1, ".Report", "PBIR", "One JSON file per page and per visual", ["definition.pbir", "pages/…/page.json", "visuals/…/visual.json"]], [C.accent2, ".SemanticModel", "TMDL", "One text file per table", ["model.tmdl", "relationships.tmdl", "tables/*.tmdl"]]].forEach(([col, h, f, d, files], i) => {
    const x = 5.7 + i * 3.47;
    card(s, x, 2.6, 3.27, 3.55, { fill: col, t: 91, line: col, name: h });
    txt(s, h, x + 0.25, 2.8, 2.8, 0.4, { fontSize: 17, bold: true, color: col, fontFace: MONO });
    txt(s, f, x + 0.25, 3.25, 2.8, 0.4, { fontSize: 22, bold: true });
    txt(s, d, x + 0.25, 3.72, 2.8, 0.6, { fontSize: 14, color: C.accent6 });
    txt(s, files.join("\n"), x + 0.25, 4.45, 2.8, 1.4, { fontSize: 13, fontFace: MONO, paraSpaceAfter: 4 });
  });
  txt(s, "PBIP holds no data, apart from a local cache.abf that Git ignores.", 0.6, 6.55, 9, 0.35, { fontSize: 13, color: C.accent6 });
}
{
  const s = slide("CONTENT", "MODULE 1 · FILE FORMATS", "Four formats, two folders",
    "Misconception: 'TMDL and PBIR are the same thing.' TMDL describes the model; PBIR describes the report. Two formats, two folders, two toolchains.\n\nTMDL view in Desktop is GA; on the web it is still preview. PBIR is GA and the default report format as of September 2026.");
  pill(s, "formats");
  const H = (t, col) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: col } } });
  const R = (t, o = {}) => ({ text: t, options: o });
  const rows = [
    [H("", C.text1), H("PBIX", C.accent6), H("PBIP", C.text1), H("PBIR · report", C.accent1), H("TMDL · model", C.accent2)],
    [R("What it is", { bold: true }), R("One file: model, data and report"), R("A folder: .pbip file + .Report + .SemanticModel"), R("One JSON file per page and per visual"), R("One text file per table")],
    [R("Readable", { bold: true }), R("No"), R("Yes, folders of text files"), R("Yes, JSON"), R("Yes, indented text")],
    [R("Diff and merge in Git", { bold: true }), R("No"), R("Yes"), R("Yes, per visual"), R("Yes, per table")],
    [R("Edit outside Desktop", { bold: true }), R("No"), R("Yes, with care"), R("Yes, schema-validated"), R("Yes, VS Code or TMDL view")],
    [R("Status, Sept 2026", { bold: true }), R("Stable"), R("Save as → Power BI project"), R("GA and the default", { bold: true, color: C.accent1 }), R("TMDL view GA in Desktop", { bold: true, color: C.accent2 })],
    [R("Older alternative", { bold: true }), R("—"), R("—"), R("PBIR-Legacy: one report.json"), R("TMSL: one model.bim")],
  ];
  s.addTable(rows, { x: 0.6, y: 1.9, w: 12.13, colW: [2.25, 2.2, 2.78, 2.45, 2.45], fontSize: 13, color: "14171D", valign: "middle", border: { type: "solid", pt: 0.75, color: HEX.line }, rowH: [0.45, 0.72, 0.55, 0.55, 0.55, 0.62, 0.55], margin: [4, 8, 4, 8] });
}
{
  const s = slide("CONTENT", "MODULE 1 · FILE FORMATS", "Your reports are already changing",
    "The fact most participants will not know: since September 2026, saving a PBIR-Legacy report in Desktop silently converts it to PBIR, with no prompt. Desktop keeps a backup for 30 days under TempSaves\\Backups, and the Power BI service keeps a PBIR-Legacy backup for 28 days.\n\nAsk: 'Is PBIR something we should start using?' You already are: any report saved in a current Desktop build becomes PBIR.");
  const stats = [["Sept 2026", "PBIR became generally available and the default report format", C.accent1], ["30 days", "Desktop keeps the legacy report as a backup under TempSaves\\Backups", C.accent3], ["28 days", "The Power BI service keeps a PBIR-Legacy backup", C.accent2]];
  stats.forEach(([n, d, col], i) => {
    const x = 0.6 + i * 4.11;
    txt(s, n, x, 2.15, 3.8, 1.0, { fontSize: 54, bold: true, color: col, fontFace: "Arial" });
    txt(s, d, x, 3.25, 3.6, 1.0, { fontSize: 16, color: C.accent6 });
  });
  card(s, 0.6, 4.75, 12.13, 1.45, { fill: C.accent3, t: 88, name: "Callout" });
  txt(s, "Saving a PBIR-Legacy report in Desktop converts it to PBIR, with no prompt.", 0.95, 5.0, 11.4, 0.5, { fontSize: 22, bold: true });
  txt(s, "You are already using PBIR: any report saved in a current Desktop build becomes one.", 0.95, 5.55, 11.4, 0.45, { fontSize: 16, color: C.accent6 });
}
{
  const s = slide("CONTENT", "MODULE 2 · PBIP FOLDER ARCHITECTURE", "Two items joined by one relative path",
    "This is the actual MeridianSupplyChain structure as Power BI Desktop saved it. Page and visual folder names are readable here because the project was generated by script; Desktop itself names new pages and visuals with random IDs.\n\nMeridian has expressions.tmdl (the DataFolder parameter) and roles/ (RLS). Other projects also contain perspectives/, bookmarks/, diagramLayout.json, DAXQueries/ or TMDLScripts/. Absence is normal; do not tell participants every project looks identical.\n\nMisconception: 'The report contains the measures.' It holds only references to them by table and name.");
  pill(s, "folder");
  txt(s, [{ text: "MeridianSupplyChain.pbip", options: { fontFace: MONO, bold: true } }, { text: "  opens the project   ·   ", options: { color: C.accent6 } }, { text: ".gitignore", options: { fontFace: MONO, bold: true } }, { text: "  excludes cache.abf and localSettings.json", options: { color: C.accent6 } }], 0.6, 1.72, 12, 0.35, { fontSize: 13 });
  const rep = [
    [["MeridianSupplyChain.Report/", "k"]],
    [["├─ .pbi/localSettings.json", "x"], ["   never committed", "c"]],
    [["├─ .platform", "x"]],
    [["├─ definition.pbir", "h"], ["     → ../Meridian…SemanticModel", "s"]],
    [["├─ definition/", "x"]],
    [["│  ├─ report.json", "x"], ["           theme, settings", "c"]],
    [["│  ├─ version.json", "x"]],
    [["│  └─ pages/", "x"]],
    [["│     ├─ pages.json", "x"], ["        page order", "c"]],
    [["│     └─ ExecutiveOverview/page.json", "x"]],
    [["│        └─ visuals/marginProduct/", "x"]],
    [["│              visual.json", "x"]],
    [["└─ StaticResources/", "x"], ["       themes, images", "c"]],
  ];
  const mod = [
    [["MeridianSupplyChain.SemanticModel/", "p"]],
    [["├─ .pbi/cache.abf", "x"], ["        data, never committed", "c"]],
    [["├─ .platform  ·  definition.pbism", "x"]],
    [["└─ definition/", "x"]],
    [["   ├─ database.tmdl", "x"]],
    [["   ├─ model.tmdl", "x"], ["         ref table, ref role", "c"]],
    [["   ├─ expressions.tmdl", "h"], ["   DataFolder parameter", "c"]],
    [["   ├─ relationships.tmdl", "x"], [" 15 relationships", "c"]],
    [["   ├─ roles/West Region.tmdl", "x"], [" RLS", "c"]],
    [["   ├─ cultures/en-US.tmdl", "x"]],
    [["   └─ tables/", "x"]],
    [["      ├─ _Measures.tmdl  Dim_*.tmdl …", "x"]],
    [["      └─ Fact_Sales.tmdl", "h"], ["   one file per table", "c"]],
  ];
  code(s, 0.6, 2.2, 6.0, 4.6, rep, { size: 12.5 });
  code(s, 6.73, 2.2, 6.0, 4.6, mod, { size: 12.5 });
}
{
  const s = slide("CONTENT", "MODULE 2 · PBIP FOLDER ARCHITECTURE", "What to touch, and what to leave alone",
    "Name the four files Microsoft lists as not supported for external editing: the legacy single report.json, mobileState.json, semanticModelDiagramLayout.json and diagramLayout.json.\n\nTwo more rules: save every file as UTF-8 without BOM, and keep the full path under Windows' 260-character default limit.");
  pill(s, "folder");
  const cols = [
    [C.accent4, "Edit freely", "Where most real work happens", ["tables/*.tmdl", "relationships.tmdl", "StaticResources/ themes", ".gitignore"]],
    [C.accent3, "Edit with care", "Schema-validated or structural", ["definition.pbir", "report.json, page.json", "visual.json", "model.tmdl", ".platform (keep its ID)"]],
    [C.accent5, "Leave to Desktop", "Generated, local or versioned", [".pbi/cache.abf", ".pbi/localSettings.json", "database.tmdl, version.json", "definition.pbism", "cultures/*.tmdl"]],
  ];
  cols.forEach(([col, h, d, items], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.85, 3.93, 3.85, { name: h });
    s.addShape(pres.shapes.OVAL, { x: x + 0.28, y: 2.14, w: 0.24, h: 0.24, fill: { color: col }, line: { type: "none" } });
    txt(s, h, x + 0.66, 2.06, 3.1, 0.4, { fontSize: 19, bold: true });
    txt(s, d, x + 0.28, 2.55, 3.4, 0.35, { fontSize: 13, color: C.accent6 });
    txt(s, items.join("\n"), x + 0.28, 3.05, 3.5, 2.5, { fontSize: 14, fontFace: MONO, paraSpaceAfter: 6 });
  });
  card(s, 0.6, 5.88, 12.13, 0.95, { fill: C.accent5, t: 92, name: "Not supported" });
  txt(s, [{ text: "Not supported for external editing: ", options: { bold: true } }, { text: "legacy report.json, mobileState.json, semanticModelDiagramLayout.json, diagramLayout.json.  ", options: {} }, { text: "Save as UTF-8 without BOM.", options: { bold: true, color: C.accent5 } }], 0.9, 6.0, 11.6, 0.75, { fontSize: 14, valign: "middle" });
}
{
  const s = slide("LIVE", "LIVE DEMO · MODULE 2", "Walk the real folder in VS Code",
    "Do not teach the folder from a slide. Switch to VS Code with the Meridian pbip folder, committed to Git.\n\nShow definition.pbir: \"path\": \"../MeridianSupplyChain.SemanticModel\" is the entire connection between report and model.");
  liveSteps(s, [["Open the Meridian pbip folder in VS Code", "Already committed to Git"], ["Open definition.pbir", "\"path\": \"../MeridianSupplyChain.SemanticModel\" is the whole connection"], ["Open .gitignore", "localSettings.json and cache.abf are left out"], ["Open one visual.json", "Point out Entity and Property; there is no DAX"]]);
  askCard(s, 8.0, 1.95, 4.73, 3.5, [["I delete cache.abf and reopen. What happens?", "The project opens with the full model and no data until you refresh."], ["Why is localSettings.json excluded from Git?", "It holds one person's machine state; committing it conflicts on every save."]]);
}
{
  const s = slide("CONTENT", "MODULE 3 · THE SEMANTIC MODEL IN TMDL", "TMDL in nine rules",
    "The syntax rules come from Microsoft's TMDL reference, except rule 9: properties left at their default are not written. That one is observed in the real files rather than stated by Microsoft, and Module 4 relies on it.");
  pill(s, "tmdl");
  const rules = [
    ["Indentation is structure", "One tab per level; breaking it is a parse error", "table Fact_Sales\n    column Revenue"],
    ["\"=\" for the expression", "An object's default or expression property", "measure Revenue = SUM(…)"],
    ["\":\" for everything else", "Every other property", "formatString: #,##0"],
    ["A bare boolean is true", "Write just the property's name", "isHidden"],
    ["Quote awkward names", "Spaces, dots, = or : need single quotes", "measure 'Gross Margin %' = …"],
    ["/// is a description", "Directly above the object, no blank line", "/// Services Delivery Analytics"],
    ["ref points and orders", "Refers to another object and sets order", "ref table Fact_Sales"],
    ["Deeper means multi-line", "Expressions sit one level deeper", "source =\n        let …"],
    ["Defaults are not written", "Observed in the real files", "(no cardinality line)"],
  ];
  rules.forEach(([h, d, ex], i) => {
    const x = 0.6 + (i % 3) * 4.1, y = 1.82 + Math.floor(i / 3) * 1.68;
    card(s, x, y, 3.93, 1.52, { name: "Rule " + (i + 1) });
    txt(s, String(i + 1), x + 0.22, y + 0.17, 0.3, 0.3, { fontSize: 13, bold: true, color: C.accent3, fontFace: MONO });
    txt(s, h, x + 0.52, y + 0.15, 3.3, 0.35, { fontSize: 15, bold: true });
    txt(s, d, x + 0.52, y + 0.5, 3.3, 0.35, { fontSize: 12, color: C.accent6 });
    txt(s, ex, x + 0.52, y + 0.86, 3.3, 0.6, { fontSize: 12, fontFace: MONO, color: C.accent2 });
  });
}
{
  const s = slide("CONTENT", "MODULE 3 · THE SEMANTIC MODEL IN TMDL", "Anatomy of a table file",
    "Walk Fact_Sales.tmdl in VS Code. dataType values seen in these projects: string, int64, double, dateTime.\n\nThe partition reads DataFolder & \"Fact_Sales.csv\". DataFolder is a parameter in expressions.tmdl, so the path lives in one place. Ask: what happens when a colleague opens this on their machine? It fails to refresh until they change the parameter, which is one line, not nine.\n\nBlanksToNull turns empty CSV cells into nulls before the date columns are typed; without it, unshipped orders would error.");
  pill(s, "tmdl");
  code(s, 0.6, 1.85, 6.85, 4.95, [
    "table Fact_Sales",
    "    lineageTag: 749f08a3-b64f-5985-bb44-…",
    "",
    "    column OrderDate",
    "        dataType: dateTime",
    "        formatString: yyyy-mm-dd",
    "        summarizeBy: none",
    "        sourceColumn: OrderDate",
    "",
    "    column SalesAmount",
    "        dataType: double",
    "        isHidden",
    "        sourceColumn: SalesAmount",
    "",
    "    partition Fact_Sales = m",
    "        mode: import",
    [["        Csv.Document(File.Contents(DataFolder & ", "x"], ["\"Fact_Sales.csv\"", "s"], ["))", "x"]],
  ], { size: 12.5 });
  const notes = [
    ["sourceColumn:", "An imported column, loaded from the query. A calculated column uses = and DAX instead.", C.accent2],
    ["lineageTag:", "A stable identity that survives renames. Never copy one onto a new object.", C.accent2],
    ["summarizeBy: none", "Stops Power BI auto-summing the column in visuals.", C.accent2],
    ["partition … = m", "The Power Query that loads the data. It reads the DataFolder parameter, so moving the CSVs is a one-line change.", C.accent3],
  ];
  notes.forEach(([k, d, col], i) => {
    const y = 1.85 + i * 1.27;
    card(s, 7.7, y, 5.03, 1.12, { name: k });
    txt(s, k, 7.95, y + 0.15, 4.6, 0.32, { fontSize: 14, bold: true, fontFace: MONO, color: col });
    txt(s, d, 7.95, y + 0.5, 4.6, 0.6, { fontSize: 13 });
  });
}
{
  const s = slide("CONTENT", "MODULE 3 · THE SEMANTIC MODEL IN TMDL", "A measure, line by line",
    "Annotate on screen.\n\nMisconception: 'Renaming a measure in TMDL is safe because lineageTag tracks it.' Visuals reference measures by table and name, so a rename in the file breaks them. Rename inside Desktop, which updates the visuals.");
  pill(s, "tmdl");
  code(s, 0.6, 1.85, 12.13, 1.55, [
    "measure 'Revenue PY' = CALCULATE ( [Revenue], SAMEPERIODLASTYEAR ( Dim_Date[Date] ) )",
    "    formatString: \\$#,##0",
    "    displayFolder: 02 Time Intelligence",
  ], { size: 16 });
  const parts = [
    ["measure 'Revenue PY' =", "The name, in single quotes because it has a space. The DAX follows the =."],
    ["CALCULATE ( [Revenue], … )", "Evaluates the existing [Revenue] measure under a changed filter."],
    ["SAMEPERIODLASTYEAR ( Dim_Date[Date] )", "The same dates one year earlier. It needs a marked date table related to the facts."],
    ["formatString · displayFolder", "The backslash prints $ literally. The folder only organises the Fields pane."],
  ];
  parts.forEach(([k, d], i) => {
    const x = 0.6 + i * 3.08;
    card(s, x, 3.65, 2.93, 2.65, { name: "Part " + (i + 1) });
    num(s, x + 0.22, 3.86, i + 1, C.accent2, 0.4);
    txt(s, k, x + 0.22, 4.42, 2.5, 0.7, { fontSize: 13, bold: true, fontFace: MONO, color: C.accent2 });
    txt(s, d, x + 0.22, 5.12, 2.5, 1.1, { fontSize: 14 });
  });
  txt(s, "When an AI writes a measure for you, ask it to leave the lineageTag line out.", 0.6, 6.5, 12, 0.35, { fontSize: 13, italic: true, color: C.accent6 });
}
{
  const s = slide("CONTENT", "MODULE 3 · THE SEMANTIC MODEL IN TMDL", "Measures build on measures",
    "Meridian keeps all 30 measures in _Measures, in 5 display folders: a hidden table with one hidden placeholder column that exists only to hold measures. That is a convention, not a rule. A measure can belong to any table; its home only decides where it appears in the Fields pane.\n\nAsk: 'Does moving a measure to another table change its result?' No, the home table is organisational only.\n\nThe interactive map shows the sales measures, every dependency, and which visuals use each one.");
  pill(s, "tmdl");
  const bw = 2.55, bh = 0.62;
  const box = (x, y, label, kind) => {
    const col = kind === "col" ? C.accent2 : kind === "end" ? C.accent3 : C.text1;
    card(s, x, y, bw, bh, { fill: kind === "col" ? C.accent2 : C.background1, t: kind === "col" ? 90 : 0, line: col, lw: kind === "end" ? 2 : 1, name: label });
    txt(s, label, x + 0.15, y, bw - 0.3, bh, { fontSize: kind === "col" ? 12.5 : 14, bold: kind !== "col", fontFace: kind === "col" ? MONO : undefined, valign: "middle" });
  };
  const X = [0.6, 3.75, 6.9, 10.05];
  txt(s, "COLUMNS", X[0], 1.82, 2.5, 0.3, { fontSize: 11, bold: true, color: C.accent6, charSpacing: 1.5 });
  txt(s, "ON COLUMNS", X[1], 1.82, 2.5, 0.3, { fontSize: 11, bold: true, color: C.accent6, charSpacing: 1.5 });
  txt(s, "ON MEASURES", X[2], 1.82, 2.5, 0.3, { fontSize: 11, bold: true, color: C.accent6, charSpacing: 1.5 });
  txt(s, "WHAT THE VISUALS SHOW", X[3], 1.82, 2.7, 0.3, { fontSize: 11, bold: true, color: C.accent6, charSpacing: 1.5 });
  box(X[0], 2.25, "Fact_Sales[SalesAmount]", "col"); box(X[0], 3.3, "Fact_Sales[CostAmount]", "col"); box(X[0], 4.55, "Dim_Date[Date]", "col");
  box(X[1], 2.25, "Revenue", "m"); box(X[1], 3.3, "Cost", "m");
  box(X[2], 2.75, "Gross Margin", "m"); box(X[2], 4.55, "Revenue PY", "m");
  box(X[3], 2.75, "Gross Margin %", "end"); box(X[3], 4.55, "Revenue YoY %", "end");
  const mid = bh / 2;
  arrow(s, X[0] + bw, 2.25 + mid, X[1], 2.25 + mid); arrow(s, X[0] + bw, 3.3 + mid, X[1], 3.3 + mid);
  arrow(s, X[1] + bw, 2.25 + mid, X[2], 2.75 + 0.2); arrow(s, X[1] + bw, 3.3 + mid, X[2], 2.75 + 0.45);
  arrow(s, X[1] + bw, 2.25 + mid + 0.15, X[2], 4.55 + 0.18);
  arrow(s, X[0] + bw, 4.55 + mid, X[2], 4.55 + mid);
  arrow(s, X[2] + bw, 2.75 + mid, X[3], 2.75 + mid, C.accent3); arrow(s, X[2] + bw, 4.55 + mid, X[3], 4.55 + mid, C.accent3);
  txt(s, "DIVIDE ( [Gross Margin], [Revenue] )", X[3], 3.45, 2.7, 0.3, { fontSize: 11, fontFace: MONO, color: C.accent6 });
  txt(s, "DIVIDE ( [Revenue] - [Revenue PY], [Revenue PY] )", X[3], 5.25, 2.8, 0.5, { fontSize: 11, fontFace: MONO, color: C.accent6 });
  card(s, 0.6, 5.95, 12.13, 0.85, { fill: C.accent2, t: 92, name: "Convention note" });
  txt(s, [{ text: "All 30 Meridian measures live in a hidden _Measures table. ", options: { bold: true } }, { text: "That is a convention, not a rule: a measure can belong to any table. Simplified here; the map shows every dependency.", options: { color: C.accent6 } }], 0.9, 6.03, 11.6, 0.7, { fontSize: 14, valign: "middle" });
}
{
  const s = slide("LIVE", "LIVE DEMO · MODULE 3", "The file is the model",
    "The fastest way to prove the file is the model.\n\nMicrosoft's own pages disagree here: the PBIP pages describe an 'Apply external changes' banner that reloads edits without reopening; the TMDL view page says to restart Desktop. Test on the training build a week before. Show the banner if it appears; if not, close and reopen. Say plainly that the docs differ.");
  liveSteps(s, [["Change a measure's formatString in VS Code", "In _Measures.tmdl, for example Revenue"], ["Save as UTF-8 without BOM", "Tabs stay tabs"], ["Apply external changes in Desktop", "Or close and reopen if the banner does not appear"], ["Watch the card update", "Nothing was touched in Desktop itself"]]);
  askCard(s, 8.0, 1.95, 4.73, 3.5, [["Which columns here are calculated?", "Only those with = followed by DAX. sourceColumn: means imported."], ["Does moving a measure to another table change its result?", "No. The home table is organisational only."]]);
}
{
  const s = slide("CONTENT", "MODULE 4 · RELATIONSHIPS", "A relationship is three lines",
    "relationships.tmdl is the whole model diagram as text. Most relationships are only three lines, because a standard many-to-one, single-direction, active relationship writes nothing beyond its two columns.\n\nMeridian's readable names (Sales_OrderDate) were written by the build script; Desktop names relationships it creates with a GUID. Meridian relates Fact_Sales to Dim_Date twice and Fact_Purchases to Dim_Date three times. Only one can be active per pair of tables, so the others carry isActive: false and a measure switches them on with USERELATIONSHIP, as Revenue by Ship Date does.\n\nDim_Product to Dim_Supplier is a snowflake. Fact_Purchases.SupplierKey deliberately has no relationship: a second path to Dim_Supplier would be ambiguous.");
  pill(s, "relationships");
  code(s, 0.6, 1.85, 6.1, 3.05, [
    "relationship Sales_OrderDate",
    "    fromColumn: Fact_Sales.OrderDate",
    "    toColumn: Dim_Date.Date",
    "",
    "relationship Sales_ShipDate",
    "    isActive: false",
    "    fromColumn: Fact_Sales.ShipDate",
    "    toColumn: Dim_Date.Date",
    "",
    "relationship Sales_Product",
    "    fromColumn: Fact_Sales.ProductKey",
    "    toColumn: Dim_Product.ProductKey",
  ], { size: 12 });
  txt(s, "Used by the inactive relationship", 0.6, 5.1, 6.1, 0.3, { fontSize: 12, bold: true, color: C.accent6 });
  code(s, 0.6, 5.42, 6.1, 1.38, ["measure 'Revenue by Ship Date' =", "    CALCULATE ( [Revenue],", "        USERELATIONSHIP ( Fact_Sales[ShipDate],", "                          Dim_Date[Date] ) )"], { size: 12 });
  const tbl = (x, y, name, w = 2.1) => { card(s, x, y, w, 0.7, { fill: C.accent2, t: 88, line: C.accent2, name }); txt(s, name, x, y, w, 0.7, { fontSize: 14, bold: true, fontFace: MONO, align: "center", valign: "middle" }); };
  tbl(8.75, 3.55, "Fact_Sales", 2.3);
  tbl(7.05, 1.95, "Dim_Date"); tbl(10.6, 1.95, "Dim_Product"); tbl(8.85, 5.45, "Dim_Warehouse");
  arrow(s, 8.1, 2.65, 9.25, 3.55, C.accent2, 1.75);
  s.addShape(pres.shapes.LINE, { x: 7.55, y: 2.65, w: 1.2, h: 1.1, line: { color: C.accent6, width: 1.5, dashType: "dash", endArrowType: "triangle" } });
  txt(s, "ShipDate, inactive", 6.95, 3.62, 1.7, 0.3, { fontSize: 11, italic: true, color: C.accent6, align: "right" }); arrow(s, 11.65, 2.65, 10.55, 3.55, C.accent2, 1.75); arrow(s, 9.9, 5.45, 9.9, 4.25, C.accent2, 1.75);
  txt(s, "1", 7.85, 2.68, 0.3, 0.3, { fontSize: 14, bold: true, color: C.accent3 }); txt(s, "*", 9.05, 3.25, 0.3, 0.3, { fontSize: 16, bold: true, color: C.accent3 });
  txt(s, "1", 11.75, 2.68, 0.3, 0.3, { fontSize: 14, bold: true, color: C.accent3 }); txt(s, "*", 10.62, 3.25, 0.3, 0.3, { fontSize: 16, bold: true, color: C.accent3 });
  txt(s, "1", 10.02, 5.1, 0.3, 0.3, { fontSize: 14, bold: true, color: C.accent3 }); txt(s, "*", 10.02, 4.3, 0.3, 0.3, { fontSize: 16, bold: true, color: C.accent3 });
  txt(s, [{ text: "from", options: { bold: true, fontFace: MONO, color: C.accent2 } }, { text: " = many side (the fact)    ", options: {} }, { text: "to", options: { bold: true, fontFace: MONO, color: C.accent2 } }, { text: " = one side (the dimension)", options: {} }], 7.05, 6.4, 5.7, 0.35, { fontSize: 13 });
}
{
  const s = slide("CONTENT", "MODULE 4 · RELATIONSHIPS", "Defaults are not written",
    "Apart from isActive, none of the TMDL files in the sample projects uses these properties, so this table is illustrative. The property names follow the Tabular Object Model, which TMDL writes in camelCase. Prove each one live with the git diff demo on the next slide rather than trusting the table.\n\nMisconception: 'One-to-many and many-to-one are different relationship types.' They are the same relationship described from opposite ends.");
  pill(s, "relationships");
  const H = t => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } });
  const M = t => ({ text: t, options: { fontFace: MONO, color: C.accent2 } });
  const rows = [[H("In Desktop"), H("What appears in TMDL"), H("Default when absent")],
    [{ text: "Many-to-one / one-to-many" }, { text: "nothing: one relationship seen from either end" }, M("from = many, to = one")],
    [{ text: "One-to-one" }, M("fromCardinality: one"), { text: "many" }],
    [{ text: "Many-to-many" }, M("toCardinality: many"), { text: "one" }],
    [{ text: "Cross-filter direction: Both" }, M("crossFilteringBehavior: bothDirections"), M("oneDirection")],
    [{ text: "Security filter in both directions" }, M("securityFilteringBehavior: bothDirections"), M("oneDirection")],
    [{ text: "Inactive" }, M("isActive: false"), { text: "active" }],
    [{ text: "Assume referential integrity (DirectQuery)" }, M("relyOnReferentialIntegrity"), { text: "off" }]];
  s.addTable(rows, { x: 0.6, y: 1.85, w: 12.13, colW: [4.0, 5.0, 3.13], fontSize: 14, color: "14171D", valign: "middle", border: { type: "solid", pt: 0.75, color: HEX.line }, rowH: 0.52, margin: [4, 10, 4, 10] });
  txt(s, "Illustrative: no sample file changes these defaults. Prove each one with git diff in class.", 0.6, 6.25, 12, 0.35, { fontSize: 13, italic: true, color: C.accent6 });
}
{
  const s = slide("LIVE", "LIVE DEMO · MODULE 4", "Desktop just writes files",
    "This teaches more than the table does: participants see that Desktop is just a UI that writes these files, and that a one-line change in the UI is a one-line change in the file.\n\nThe position of the added line in the diff may differ from what the slide shows; trust the live diff.");
  liveSteps(s, [["Commit the project", "A clean starting point"], ["Set one relationship to filter Both", "Manage relationships, then save"], ["Run git diff relationships.tmdl", "One new line appears under that relationship"], ["Set it back and save", "The diff is empty again"]], 0.6, 1.95, 6.6);
  code(s, 7.45, 1.95, 5.28, 1.85, ["relationship Sales_Warehouse", "+   crossFilteringBehavior: bothDirections", "    fromColumn: Fact_Sales.WarehouseKey", "    toColumn: Dim_Warehouse.WarehouseKey"], { dark: true, size: 12.5 });
  askCard(s, 7.45, 4.0, 5.28, 2.5, [["Find every inactive relationship in a 40-table model without Desktop?", "Search relationships.tmdl for isActive: false."]]);
}
{
  const s = slide("CONTENT", "MODULE 5 · THE REPORT FOLDER AND PBIR", "The report holds no calculations",
    "Reading it in plain English: a clustered bar chart in the bottom-left of the Executive Overview, one bar per product, bar length showing gross margin percentage, sorted lowest first. With 2025 selected, the Infusion Pump Tubing Set's bar goes below zero.\n\nEntity + Property (table name + field name) is the only link back to the semantic model. visualContainerObjects holds the frame (title, background, border); filterConfig holds the visual's filters.\n\nMisconception: 'A visual's filters are always in its file.' Automatic filters are written to visual.json only after the filter pane has been expanded once.\n\nMicrosoft's limits: 1,000 pages per report, 1,000 visuals per page, 300 MB for all report files.");
  pill(s, "report");
  code(s, 0.6, 1.85, 6.3, 4.95, [
    [["\"visualType\"", "p"], [": ", "x"], ["\"clusteredBarChart\"", "s"]],
    [["\"queryState\"", "p"], [": {", "x"]],
    [["  \"Category\"", "p"], [": { \"Column\": {", "x"]],
    [["      \"Entity\"", "p"], [": ", "x"], ["\"Dim_Product\"", "h"]],
    [["      \"Property\"", "p"], [": ", "x"], ["\"ProductName\"", "h"], [" } },", "x"]],
    [["  \"Y\"", "p"], [": { \"Measure\": {", "x"]],
    [["      \"Entity\"", "p"], [": ", "x"], ["\"_Measures\"", "h"]],
    [["      \"Property\"", "p"], [": ", "x"], ["\"Gross Margin %\"", "h"], [" } }", "x"]],
    [["},", "x"]],
    [["\"sortDefinition\"", "p"], [": … ", "x"], ["\"Ascending\"", "s"]],
    [["\"title\"", "p"], [": … ", "x"], ["\"Gross margin % by product…\"", "s"]],
    [["", "x"]],
    [["// no DAX anywhere in this file", "c"]],
  ], { size: 14, lang: "json" });
  // Wireframe of the real page, visuals at their stored positions (1280 x 720).
  const px = 7.15, py = 1.85, k = 5.58 / 1280;
  card(s, px, py, 1280 * k, 720 * k, { fill: C.background2, line: HEX.line, name: "Page wireframe" });
  const vis = [[24, 12, 640, 64, "Text box"], [704, 8, 184, 80, "Slicer"], [888, 8, 184, 80, "Slicer"], [1072, 8, 184, 80, "Slicer"], [24, 96, 1232, 112, "Card: 4 KPIs"], [24, 224, 728, 232, "Line chart"], [768, 224, 488, 232, "Bar chart"], [768, 472, 488, 232, "Bar chart"], [24, 472, 728, 232, "marginProduct · Gross margin % by product"]];
  vis.forEach(([x, y, w, h, l], i) => {
    const hl = i === vis.length - 1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px + x * k, y: py + y * k, w: w * k, h: h * k, rectRadius: 0.03, fill: { color: hl ? C.accent3 : C.background1, transparency: hl ? 82 : 0 }, line: { color: hl ? C.accent3 : C.accent1, width: hl ? 1.75 : 0.75 } });
    if (l) txt(s, l, px + x * k + 0.06, py + y * k + 0.04, w * k - 0.1, 0.25, { fontSize: 10, bold: hl, color: hl ? C.text1 : C.accent6 });
  });
  txt(s, "Executive Overview · 1280 × 720 · 9 visuals, drawn at their stored positions", px, py + 720 * k + 0.1, 5.6, 0.3, { fontSize: 11, color: C.accent6 });
  card(s, px, 5.55, 5.58, 1.25, { fill: C.accent1, t: 91, name: "Reading" });
  txt(s, [{ text: "Entity + Property = table + field. ", options: { bold: true } }, { text: "The only link back to the model. The DAX for Gross Margin % lives in _Measures.tmdl.", options: {} }], px + 0.25, 5.65, 5.1, 1.05, { fontSize: 14, valign: "middle" });
}
{
  const s = slide("CONTENT", "MODULE 5 · THE REPORT FOLDER AND PBIR", "How a number reaches the page",
    "Walk the chain once, slowly. Then open the interactive map and press 'Trace a number': it animates the same nine steps on the real files.\n\nAsk: 'Where is the DAX for this chart's values?' Not in the report at all; visual.json only names the measure. The DAX is in _Measures.tmdl.\n\nAsk: 'How would you list every visual that uses Gross Margin % across 20 pages?' Search the pages folder for \"Property\": \"Gross Margin %\".");
  pill(s, "overview");
  const steps = [
    ["MeridianSupplyChain.pbip", "opens the report folder", C.accent6],
    ["definition.pbir", "byPath → ../MeridianSupplyChain.SemanticModel", C.accent1],
    ["page.json", "Executive Overview, 1280 × 720", C.accent1],
    ["visual.json · marginProduct", "asks for _Measures / Gross Margin %", C.accent1],
    ["_Measures.tmdl", "DIVIDE ( [Gross Margin], [Revenue] )", C.accent2],
    ["Fact_Sales.tmdl → CSV", "SUM of SalesAmount and CostAmount, via DataFolder", C.accent2],
  ];
  steps.forEach(([h, d, col], i) => {
    const y = 1.85 + i * 0.83;
    num(s, 0.6, y + 0.12, i + 1, col, 0.46);
    txt(s, h, 1.25, y + 0.04, 4.9, 0.35, { fontSize: 16, bold: true, fontFace: MONO });
    txt(s, d, 1.25, y + 0.4, 4.9, 0.35, { fontSize: 13, color: C.accent6 });
    if (i < steps.length - 1) s.addShape(pres.shapes.LINE, { x: 0.83, y: y + 0.6, w: 0, h: 0.33, line: { color: HEX.line, width: 1.5 } });
  });
  const iw = 6.4, ih = iw * 1102 / 1508;
  card(s, 6.33, 1.75, iw + 0.0, ih + 0.0, { fill: C.background1, line: HEX.line, name: "Map preview frame" });
  s.addImage({ path: path.join(ASSETS, "map_overview.png"), x: 6.33, y: 1.75, w: iw, h: ih, altText: "Interactive map: the report and semantic model folders, with the bar chart's visual.json highlighted" });
  s.addText([{ text: "Interactive version: press Trace a number  ↗", options: { color: C.accent1, hyperlink: { url: `${MAP}#overview`, tooltip: "Opens the map" } } }], { x: 6.33, y: 1.75 + ih + 0.08, w: iw, h: 0.3, fontSize: 12, bold: true, margin: 0, isTextBox: true, align: "right" });
}
{
  const s = slide("LIVE", "LIVE DEMO · MODULE 5", "Rename a measure, break a visual",
    "That single demonstration explains why AI edits to the model must be checked against the report, and why renames belong in Desktop.\n\nThe interactive map's Report view has the same rename as a button: the two visuals that use Gross Margin % turn red.");
  liveSteps(s, [["Commit to Git", "Always the first step"], ["Rename 'Gross Margin %' to 'GM %'", "In _Measures.tmdl, then save"], ["Apply the changes in Desktop", "The bar chart breaks"], ["git checkout the file", "Everything is restored"]]);
  card(s, 8.0, 1.95, 4.73, 4.55, { fill: C.background1, t: 92, line: HEX.lineDark, name: "Why it breaks" });
  txt(s, "WHY IT BREAKS", 8.3, 2.2, 4.2, 0.3, { fontSize: 11, bold: true, color: C.accent3, charSpacing: 1.5 });
  txt(s, "visual.json still asks for", 8.3, 2.6, 4.2, 0.35, { fontSize: 15, color: C.background2, transparency: 25 });
  txt(s, "_Measures / Gross Margin %", 8.3, 2.98, 4.2, 0.4, { fontSize: 16, bold: true, fontFace: MONO, color: "F2B544" });
  txt(s, "which no longer exists. Visuals find measures by table and name, not by lineageTag.", 8.3, 3.45, 4.2, 1.1, { fontSize: 15, color: C.background2, transparency: 25 });
  txt(s, "Renames belong in Desktop, which updates every visual that uses the measure.", 8.3, 4.75, 4.2, 1.2, { fontSize: 16, bold: true, color: C.background1 });
}

/* =========================================================
   PART 2 · READING IT WITH AI
   ========================================================= */
section("Part 2 · Reading it with AI");
{
  const s = slide("DARK", "PART 2 · MODULES 6–7", "Reading it with AI", "Once you can find the right file, any AI assistant can read it. Copy-paste works today with no setup, and it shows exactly what is tedious about it, which is what MCP fixes.");
  s.addText("Copy-paste first, so everyone sees what MCP is for", { placeholder: "sub" });
  ["6  AI without MCP", "7  Why MCP"].forEach((m, i) => txt(s, m, 0.8 + i * 3.0, 5.1, 2.9, 0.4, { fontSize: 16, bold: true, color: "F2B544" }));
}
{
  const s = slide("CONTENT", "MODULE 6 · AI WITHOUT MCP", "Four rules before anyone pastes",
    "Rule 1 is a governance rule. Model files carry more than measures: expressions.tmdl holds server and database names, partitions hold file paths, RLS roles hold email logic. The Medical-Legal sample's expressions.tmdl names a SQL Server instance, a database and a stored procedure. Never paste a client's model into a public AI tool, and check the company's AI policy first.\n\nRule 4: the 'Store 360 â€”' text in the sample is what a wrong encoding does. Show it.");
  const rules = [["Training model only", "Model files carry server names, file paths and RLS logic. Never paste a client's model into a public AI tool.", C.accent5], ["Commit before any edit", "Every change can then be reviewed with git diff and undone in one command.", C.accent2], ["AI output is a draft", "Every claim it makes gets checked against Desktop before anyone relies on it.", C.accent1], ["Paste back cleanly", "Tab indentation, straight quotes, and UTF-8 without BOM.", C.accent3]];
  rules.forEach(([h, d, col], i) => {
    const x = 0.6 + i * 3.08;
    card(s, x, 1.95, 2.93, 3.5, { name: h });
    txt(s, String(i + 1), x + 0.28, 2.15, 1, 0.8, { fontSize: 44, bold: true, color: col, fontFace: "Arial" });
    txt(s, h, x + 0.28, 3.05, 2.45, 0.7, { fontSize: 18, bold: true });
    txt(s, d, x + 0.28, 3.8, 2.45, 1.5, { fontSize: 14, color: C.accent6 });
  });
  card(s, 0.6, 5.7, 12.13, 1.1, { fill: C.accent5, t: 92, name: "Governance example" });
  txt(s, [{ text: "Why rule 1 matters: ", options: { bold: true } }, { text: "one sample project's ", options: {} }, { text: "expressions.tmdl", options: { fontFace: MONO, bold: true } }, { text: " names a SQL Server instance, a database and a stored procedure. That is what a paste would share.", options: {} }], 0.9, 5.8, 11.6, 0.9, { fontSize: 15, valign: "middle" });
}
{
  const s = slide("CONTENT", "MODULE 6 · AI WITHOUT MCP", "Four exercises, each with a check",
    "Run exercises 1 and 3 live; set 2 and 4 as the lab. Every prompt ends with a constraint ('do not assume any table that is not shown', 'only report what is in the JSON') because the most common AI error is inventing detail.\n\nThe prompts are in the trainer's guide and on the map's AI page, each with a Copy button.");
  pill(s, "ailoop");
  const ex = [
    ["Exercise 1 · live", "Explain a measure", "Turn Revenue YoY % into business language and list where it returns blank.", "Year = 2024 shows blank: there is no 2023 data."],
    ["Exercise 2 · lab", "Explain the model", "From relationships.tmdl alone: three facts, role-playing dates, the snowflake, anything odd.", "Model view, including the dotted inactive lines."],
    ["Exercise 3 · live", "Read a visual", "From marginProduct's visual.json: type, fields, wells, sort, title.", "Compare with Desktop; watch for invented titles or colours."],
    ["Exercise 4 · lab", "Generate and apply", "Revenue QTD as TMDL, pasted into _Measures.tmdl and applied.", "March equals the Q1 total; git diff shows one measure."],
  ];
  ex.forEach(([e, h, d, v], i) => {
    const x = 0.6 + (i % 2) * 6.14, y = 1.85 + Math.floor(i / 2) * 2.5;
    card(s, x, y, 5.99, 2.32, { name: h });
    txt(s, e.toUpperCase(), x + 0.3, y + 0.22, 5, 0.3, { fontSize: 11, bold: true, color: e.includes("live") ? C.accent3 : C.accent6, charSpacing: 1.5 });
    txt(s, h, x + 0.3, y + 0.52, 5.4, 0.42, { fontSize: 19, bold: true });
    txt(s, d, x + 0.3, y + 0.98, 5.4, 0.62, { fontSize: 14 });
    txt(s, [{ text: "Check  ", options: { bold: true, color: C.accent4 } }, { text: v, options: { color: C.accent6 } }], x + 0.3, y + 1.68, 5.4, 0.5, { fontSize: 13 });
  });
}
{
  const s = slide("LIVE", "LIVE · MODULE 6", "Exercises 1 and 3, on screen",
    "Exercise 1 expected output: year-on-year revenue growth; blank unless exactly one year is selected (HASONEVALUE), because across all years last year is only part of the data; blank for 2024 because there is no 2023; DIVIDE avoids a divide-by-zero; the format shows one decimal place as a percentage. For 2025 the answer is +10.1%.\n\nExercise 3 expected output: a clustered bar chart with Dim_Product.ProductName on the category axis and the Gross Margin % measure as values, sorted ascending, titled 'Gross margin % by product, lowest first'. Nothing about colours or labels: the JSON does not set them.");
  liveSteps(s, [["Copy the three measures", "Revenue, Revenue PY and Revenue YoY % from _Measures.tmdl"], ["Ask for meaning, steps and edge cases", "'Do not assume any table or column that is not shown'"], ["Check: set Year to 2024", "Revenue YoY % should be blank; 2025 shows +10.1%"], ["Now paste marginProduct's visual.json", "Ask for type, fields, wells, sort and title"]]);
  card(s, 8.0, 1.95, 4.73, 3.6, { fill: C.background1, t: 92, line: HEX.lineDark, name: "Watch for" });
  txt(s, "WATCH FOR", 8.3, 2.2, 4.2, 0.3, { fontSize: 11, bold: true, color: C.accent3, charSpacing: 1.5 });
  txt(s, "The most common AI error", 8.3, 2.6, 4.2, 0.4, { fontSize: 19, bold: true, color: C.background1 });
  txt(s, "describing a title, colour or axis that the JSON never mentions.", 8.3, 3.05, 4.2, 1.0, { fontSize: 16, color: C.background2, transparency: 20 });
  txt(s, "The fix is in the prompt: only report what is in the file; say \"not specified\" for anything else.", 8.3, 4.35, 4.2, 1.4, { fontSize: 15, color: "F2B544" });
}
{
  const s = slide("CONTENT", "MODULE 7 · WHY MCP", "In Module 6, you were the integration",
    "Ask: 'In Exercise 4, which steps were you doing that a tool could do?' Reading the model, checking names exist, writing the measure, testing the value.\n\nThe same task both ways: copy-paste took seven steps; an AI client with an MCP server takes one request plus approvals.");
  pill(s, "ailoop");
  const steps = ["Commit", "Copy the right file", "Prompt the AI", "Check every name exists", "Paste with tabs, UTF-8 no BOM", "Apply in Desktop", "Validate March = Q1, git diff"];
  const auto = [false, true, false, true, true, true, true];
  steps.forEach((t, i) => {
    const y = 1.85 + i * 0.68;
    card(s, 0.6, y, 6.0, 0.56, { fill: auto[i] ? C.accent3 : C.background2, t: auto[i] ? 86 : 0, name: "Step " + (i + 1) });
    txt(s, String(i + 1), 0.8, y, 0.4, 0.56, { fontSize: 14, bold: true, fontFace: MONO, color: C.accent6, valign: "middle" });
    txt(s, t, 1.25, y, 4.0, 0.56, { fontSize: 15, bold: true, valign: "middle" });
    if (auto[i]) txt(s, "a tool can do this", 4.6, y, 1.85, 0.56, { fontSize: 12, color: C.accent3, bold: true, valign: "middle", align: "right" });
  });
  card(s, 7.0, 1.85, 5.73, 4.64, { fill: C.text1, name: "MCP version" });
  txt(s, "WITH AN AI CLIENT + MCP SERVER", 7.35, 2.15, 5.1, 0.3, { fontSize: 11, bold: true, color: "F2B544", charSpacing: 1.5 });
  txt(s, "One request,\nplus your approvals", 7.35, 2.55, 5.1, 1.2, { fontSize: 30, bold: true, color: C.background1 });
  txt(s, bullets(["Reads the model as it is now, not a stale paste", "Checks names against the real model", "Writes only through a tool call you approve", "Runs a DAX query and returns the actual number"], { gap: 8 }), 7.35, 3.95, 5.1, 2.3, { fontSize: 15, color: C.background2 });
}
{
  const s = slide("CONTENT", "MODULE 7 · WHY MCP", "Host, client, server, tools",
    "MCP, the Model Context Protocol, is an open standard for connecting AI applications to external tools and data, published by Anthropic in November 2024.\n\nHow a call happens: the host starts the server (a local process over standard input and output) or connects over HTTP. They exchange JSON-RPC messages and the server lists its tools. The model decides a tool would help and requests it, the host asks the person to approve, the server runs it, and the result goes back into the conversation.\n\nWhy it exists: without a standard, every AI application needs a custom integration for every tool. One Power BI server works with any client that speaks the protocol.\n\nAsk: 'Why start with a read-only connection?' Write access becomes a decision, not a default.");
  pill(s, "mcp");
  const lanes = [["You", "the developer", C.accent6], ["Host", "Claude Code in VS Code", C.text1], ["MCP client", "one per server", C.text1], ["MCP server", "Power BI modeling", C.accent2], ["Semantic model", "Desktop or Fabric", C.accent2]];
  lanes.forEach(([h, d, col], i) => {
    const x = 0.6 + i * 2.5;
    card(s, x, 1.9, 2.13, 1.05, { fill: C.background1, line: col, lw: 1.25, name: h });
    txt(s, h, x, 2.02, 2.13, 0.4, { fontSize: 16, bold: true, align: "center" });
    txt(s, d, x, 2.45, 2.13, 0.35, { fontSize: 12, color: C.accent6, align: "center" });
    if (i < lanes.length - 1) arrow(s, x + 2.15, 2.42, x + 2.48, 2.42, C.accent6, 1.5, i === 2);
  });
  txt(s, "JSON-RPC over stdio or HTTP", 5.3, 3.02, 3.2, 0.3, { fontSize: 11, color: C.accent6, align: "center" });
  const rows = [
    [{ text: "Term", options: { bold: true, color: C.background1, fill: { color: C.text1 } } }, { text: "Meaning", options: { bold: true, color: C.background1, fill: { color: C.text1 } } }, { text: "Power BI example", options: { bold: true, color: C.background1, fill: { color: C.text1 } } }],
    [{ text: "Host", options: { bold: true } }, "The application the person uses", "Claude Code in VS Code, Claude Desktop"],
    [{ text: "Client", options: { bold: true } }, "The connection inside the host, one per server", "The link from Claude Code to the Power BI server"],
    [{ text: "Server", options: { bold: true } }, "A program that exposes capabilities", "Microsoft's Power BI Authoring MCP server"],
    [{ text: "Tools", options: { bold: true } }, "Actions the model can ask to run", "List tables, run a DAX query, create a measure"],
    [{ text: "Resources", options: { bold: true } }, "Data the application can read", "Model metadata, a file"],
    [{ text: "Prompts", options: { bold: true } }, "Reusable instructions the server offers", "\"Document this model\""],
  ];
  s.addTable(rows, { x: 0.6, y: 3.5, w: 12.13, colW: [1.9, 4.9, 5.33], fontSize: 13, color: "14171D", valign: "middle", border: { type: "solid", pt: 0.75, color: HEX.line }, rowH: 0.45, margin: [3, 10, 3, 10] });
}
{
  const s = slide("CONTENT", "MODULE 7 · WHY MCP", "Three kinds of automation",
    "Do not let participants assume every MCP server can do everything. These are separate capabilities.\n\nThe live-model server is Microsoft's Power BI Authoring MCP server (package @microsoft/powerbi-modeling-mcp; Windows only for a model open in Desktop). Register it read-only for a first session.\n\nClaude and Codex are separate clients: each has its own MCP configuration, and each server's capabilities differ. Check the client and the server independently.\n\nMisconception: 'MCP means the AI controls Power BI.' The server exposes specific tools, the host asks for approval, and read-only mode removes writes completely.");
  pill(s, "mcp");
  const kinds = [["Files", "Editing TMDL and PBIR on disk", "An AI client with file access needs no MCP server at all. Module 6, automated, and it stays under Git.", C.accent1], ["Live model", "The model open in Desktop or Fabric", "The Power BI Authoring MCP server, through the Analysis Services engine. Start it read-only.", C.accent2], ["Report", "Pages and visuals", "Separate report tooling and the PBIR files. The modeling server does not change visuals.", C.accent6]];
  kinds.forEach(([h, sub, d, col], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.85, 3.93, 3.3, { name: h });
    num(s, x + 0.28, 2.1, i + 1, col, 0.5);
    txt(s, h, x + 0.95, 2.13, 2.8, 0.45, { fontSize: 22, bold: true });
    txt(s, sub, x + 0.28, 2.85, 3.4, 0.4, { fontSize: 15, bold: true, color: col });
    txt(s, d, x + 0.28, 3.35, 3.4, 1.6, { fontSize: 14, color: C.accent6 });
  });
  card(s, 0.6, 5.4, 12.13, 1.4, { fill: C.accent3, t: 86, name: "Session 2 warning" });
  txt(s, "CARRY THIS INTO SESSION 2", 0.95, 5.55, 6, 0.3, { fontSize: 11, bold: true, color: C.accent3, charSpacing: 1.5 });
  txt(s, [{ text: "A measure written to the live model lands outside Desktop's project state ", options: { bold: true } }, { text: "and can surface as \"Something's wrong with one or more fields\". Changes belong in the files, under Git. That is why this session came first.", options: {} }], 0.95, 5.88, 11.5, 0.85, { fontSize: 15 });
}

/* =========================================================
   CLOSE
   ========================================================= */
section("Close");
{
  const s = slide("CONTENT", "ASSIGNMENT · BEFORE SESSION 02", "Read the project, then change it safely",
    "About 60 minutes, on the Meridian PBIP (or any PBIP of their own that holds no client data).\n\nMarking: Part A answers traced to a file and line; names checked before applying; value checked against an independent figure; the measure opens in Desktop with no errors, tabs, no BOM; a clean diff showing only the intended change; an honest note on what the AI got wrong.");
  const cols = [
    ["Part A", "Read it from the files", "VS Code only, no Desktop", ["Where does definition.pbir point?", "Which tables are facts?", "Which column is the many side?", "Which measure uses another?", "Which measure does a visual show?", "Would the data path work for a colleague?", "Which files do not belong in Git?"], C.accent1],
    ["Part B", "One AI measure, validated", "Commit first", ["Paste only TMDL, never data", "Check every name the AI uses", "Add it with tabs, UTF-8 no BOM", "Open in Desktop, no errors", "Check the value another way", "Commit, then git diff HEAD~1"], C.accent2],
    ["Submit", "What to hand in", "Before Session 02", ["Part A answers, each with file and line", "The prompt and the AI's answer", "The git diff output", "One sentence: right first time, or what you changed?"], C.accent3],
  ];
  cols.forEach(([e, h, sub, items, col], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.85, 3.93, 4.95, { name: e });
    txt(s, e.toUpperCase(), x + 0.28, 2.07, 3.4, 0.3, { fontSize: 11, bold: true, color: col, charSpacing: 1.5 });
    txt(s, h, x + 0.28, 2.38, 3.4, 0.42, { fontSize: 19, bold: true });
    txt(s, sub, x + 0.28, 2.82, 3.4, 0.3, { fontSize: 12, color: C.accent6 });
    txt(s, bullets(items, { num: i < 2, gap: 9 }), x + 0.28, 3.3, 3.45, 3.4, { fontSize: 15 });
  });
}
{
  const s = slide("DARK", "NEXT · SESSION 02", "Connecting Claude, safely",
    "Recap in three lines, then hand out the assignment. Share the interactive map link (it is private until you share it from its Share menu) and point out the offline copy in the session folder.");
  s.addText("Read-only first, files under Git, every change validated", { placeholder: "sub" });
  const take = [["A PBIP is text", "A report and a model, joined by one relative path"], ["Visuals hold names", "The DAX lives in TMDL, found by table and name"], ["Commit before AI edits", "Change, apply, validate, diff, then commit"]];
  take.forEach(([h, d], i) => {
    const x = 0.8 + i * 3.98;
    card(s, x, 4.45, 3.7, 1.35, { fill: C.background1, t: 91, line: HEX.lineDark, name: h });
    txt(s, h, x + 0.25, 4.62, 3.3, 0.4, { fontSize: 18, bold: true, color: C.background1 });
    txt(s, d, x + 0.25, 5.05, 3.3, 0.65, { fontSize: 14, color: C.background2, transparency: 25 });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 6.2, w: 3.0, h: 0.5, rectRadius: 0.25, fill: { color: C.background1, transparency: 88 }, line: { color: C.accent3, width: 1 } });
  s.addText([{ text: "Open the interactive map  ↗", options: { color: "F2B544", hyperlink: { url: `${MAP}#overview`, tooltip: "Opens the interactive map" } } }], { x: 0.8, y: 6.2, w: 3.0, h: 0.5, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  txt(s, "Offline copy: Power_BI_Internals_Map.html, in the same folder as this deck", 4.0, 6.32, 8, 0.3, { fontSize: 13, color: C.background2, transparency: 35 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  if (!process.env.SKIP_THEME) await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
