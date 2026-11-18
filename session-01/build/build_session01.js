// Session 01 deck: Connecting Claude to Power BI (60 minutes). Reuses the theme and helpers of the deep-dive deck.
// Rebuild:  NODE_PATH=<folder with pptxgenjs> node build_deck.js
// Each "Open interactive map" pill links to a view of the published map artifact (MAP + "#view").
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const MAP = "https://rishav-portfolio-i46g.onrender.com/training/power-bi-internals-map.html";
const OUT = process.env.OUT || path.join(__dirname, "Session01_Connecting_Claude_to_Power_BI.pptx");
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
pres.title = "Session 01 - Connecting Claude to Power BI";
const C = pres.SchemeColor;
const MONO = "Consolas";
const W = 13.333;

/* ---------------- layouts ---------------- */
const footer = (dark) => ({ text: { text: "Session 01 · Connecting Claude to Power BI", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 9, color: dark ? C.background2 : C.accent6, transparency: dark ? 45 : 0, margin: 0 } } });
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
   SESSION 01 · CONNECTING CLAUDE TO POWER BI  (60 minutes)
   ========================================================= */
section("Opening");
{
  const s = slide("DARK", "SESSION 01 · WEEK 1 · CLAUDE + POWER BI PROGRAMME", "Connecting Claude to Power BI",
    "OPENING (about 2 minutes)\n\nSay: Good morning, everyone, and welcome to session one of eight. By the end of this programme, Claude will help you understand models, write and check DAX, document reports and validate numbers. Today we do the first and most important step: we connect Claude to a Power BI model, safely, and ask it questions.\n\nSay: Three things happen in the next hour. We install one piece: the Power BI MCP server. We connect Claude to a report open in Power BI Desktop, in read-only mode, so it cannot change anything. And each of you asks the model three questions, then checks every answer against Desktop. That last part matters most: we never trust an answer we have not checked.");
  s.addText("Ask your model questions in plain English, read-only, and check every answer", { placeholder: "sub" });
  const arc = ["Understand MCP", "Install and register", "Connect read-only", "Ask and verify"];
  arc.forEach((a, i) => {
    const x = 0.8 + i * 2.98;
    card(s, x, 4.75, 2.62, 0.9, { fill: C.background1, t: 91, line: HEX.lineDark, name: "Step" });
    txt(s, String(i + 1).padStart(2, "0"), x + 0.22, 4.88, 0.6, 0.3, { fontSize: 11, bold: true, color: C.accent3, fontFace: MONO });
    txt(s, a, x + 0.22, 5.16, 2.3, 0.4, { fontSize: 15, bold: true, color: C.background1 });
    if (i < arc.length - 1) arrow(s, x + 2.66, 5.2, x + 2.94, 5.2, C.accent3, 1.5);
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 6.15, w: 3.0, h: 0.5, rectRadius: 0.25, fill: { color: C.background1, transparency: 88 }, line: { color: C.accent3, width: 1 } });
  s.addText([{ text: "Open the interactive map  ↗", options: { color: "F2B544", hyperlink: { url: `${MAP}#mcp`, tooltip: "Opens the interactive map" } } }], { x: 0.8, y: 6.15, w: 3.0, h: 0.5, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  txt(s, "60 minutes · hands-on · Meridian Medical Supply sample", 4.0, 6.27, 8, 0.3, { fontSize: 13, color: C.background2, transparency: 35 });
}
{
  const s = slide("CONTENT", "PROGRAMME OBJECTIVES · THE NEXT 4–5 SESSIONS", "Where these sessions take us",
    "PROGRAMME OBJECTIVES (about 2 minutes)\n\nSay: Before we start, here's where the next four to five sessions are heading. Four objectives.\n\nSay: One: automate Power BI reporting. The routine work, answering questions about a model, checking it, making repeat changes, handed to Claude, with you approving and checking every step.\n\nSay: Two: understand Power BI metadata. How a report and its model are actually stored, as PBIP, TMDL and PBIR, and how that structure lets us rebuild and reproduce reporting instead of redoing it by hand.\n\nSay: Three: build Skills with Claude on TMDL. Once we can read TMDL, we write Skills that work on it, so a procedure runs the same way every time, for anyone on the team.\n\nSay: Four: troubleshoot and understand any report. Skills that explain a report, find what's broken and show why, so repetitive checks and fixes take minutes, not hours.\n\nSay: Today is step one of that path: we connect Claude to a model, read-only, and we build our first Skill before the hour ends.");
  const objs = [
    ["Automate Power BI reporting", "Hand routine report work to Claude: questions, checks and repeat changes, with you approving every step.", C.accent3],
    ["Understand Power BI metadata", "How a report and its model are stored (PBIP, TMDL, PBIR), and how that structure lets us rebuild and reproduce reporting.", C.accent2],
    ["Build Skills with Claude on TMDL", "Read the TMDL structure of a report, then write Skills that work on it, so a procedure runs the same way every time.", C.accent1],
    ["Troubleshoot and understand any report", "Skills that explain a report, find what's broken and show why, so repetitive tasks take minutes, not hours.", C.text1],
  ];
  objs.forEach(([h, d, col], i) => {
    const x = 0.6 + (i % 2) * 6.14, y = 1.85 + Math.floor(i / 2) * 2.12;
    card(s, x, y, 5.99, 1.95, { name: h });
    num(s, x + 0.28, y + 0.28, i + 1, col, 0.52);
    txt(s, h, x + 1.0, y + 0.26, 4.8, 0.55, { fontSize: 19, bold: true, valign: "middle" });
    txt(s, d, x + 1.0, y + 0.88, 4.75, 0.95, { fontSize: 14, color: C.accent6 });
  });
  card(s, 0.6, 6.15, 12.13, 0.68, { fill: C.accent3, t: 88, name: "Today" });
  txt(s, [{ text: "Today, step one:  ", options: { bold: true } }, { text: "connect Claude to a model read-only, check its answers, and build our first Skill.", options: {} }], 0.9, 6.15, 11.6, 0.68, { fontSize: 15, valign: "middle" });
}
{
  const s = slide("CONTENT", "OUR SAMPLE FOR ALL EIGHT SESSIONS", "Meet Meridian Medical Supply",
    "THE SAMPLE (about 2 minutes)\n\nSay: We need a real model to talk to. Meridian Medical Supply is a fictional distributor of medical supplies: gloves, masks, test kits, IV sets. Five warehouses, a hundred and fifty healthcare customers, two years of data, January 2024 to December 2025.\n\nSay: You each have the same file: MeridianSupplyChain.pbix. Open it in Power BI Desktop now, if you haven't already, and leave it open. Claude will connect to exactly what you see.\n\nSay: There are business stories hidden in this data. Problems a real supply-chain manager would want to find. We'll find some of them today with Claude, and more in later sessions.\n\nDo not reveal the stories (trainer only): 2025 revenue +10.1%; flu-season peaks; the Infusion Pump Tubing Set loses money in 2025; Pacific Rim Medical's on-time rate collapses in H2 2025; Swift Freight's deliveries slip in H2 2025.");
  pill(s, "report");
  const stats = [["3", "fact tables: sales, purchases, stock"], ["30", "measures, already written"], ["15", "relationships, 3 of them inactive"], ["$27.0M", "revenue, Jan 2024 – Dec 2025"]];
  stats.forEach(([n, d], i) => {
    const y = 1.85 + i * 1.18;
    txt(s, n, 0.6, y, 3.6, 0.65, { fontSize: 36, bold: true, color: i === 3 ? C.accent3 : C.accent2, fontFace: "Arial" });
    txt(s, d, 0.6, y + 0.66, 3.6, 0.35, { fontSize: 14, color: C.accent6 });
  });
  const iw = 8.0, ih = iw * 1356 / 2395;
  card(s, 4.73, 1.85, iw, ih, { fill: C.background1, line: HEX.line, name: "Report frame" });
  s.addImage({ path: path.join(ASSETS, "report_executive.png"), x: 4.73, y: 1.85, w: iw, h: ih, altText: "Executive Overview page of the Meridian report" });
  txt(s, [{ text: "MeridianSupplyChain.pbix", options: { bold: true, fontFace: MONO } }, { text: "   ·   open it in Power BI Desktop and leave it open", options: { color: C.accent6 } }], 4.73, 1.85 + ih + 0.1, iw, 0.3, { fontSize: 12 });
}

section("What MCP is");
{
  const s = slide("CONTENT", "WHAT MCP IS", "How Claude reaches a Power BI model",
    "THE BIG PICTURE (about 4 minutes)\n\nSay: Read this left to right. You type a question in Claude Code, inside VS Code. That's the host. Inside the host is an MCP client, a connection that knows how to talk to one server.\n\nSay: The server is a small program from Microsoft called the Power BI modeling MCP server. It runs on your laptop. It knows how to find the Power BI Desktop you have open, and how to talk to the model engine inside it: the same Analysis Services engine that DAX Studio and Tabular Editor connect to.\n\nSay: So when you ask 'which measures use Revenue?', Claude doesn't guess. It asks the server, the server asks the model, and the real answer comes back.\n\nSay: Two things to notice. One: Claude never opens your PBIX file directly; it talks to the running model through the server. Two: every step that touches the model is a tool call, and you see it and approve it.\n\nAsk: 'What tool would you have used before to look inside a running model?' Answer: DAX Studio or Tabular Editor. MCP gives Claude the same kind of access.");
  pill(s, "mcp");
  const lanes = [["You", "ask in plain English", C.accent6], ["Claude Code", "the host, in VS Code", C.text1], ["MCP client", "one per server", C.text1], ["Power BI MCP server", "runs on your laptop", C.accent2], ["Power BI Desktop", "the model you have open", C.accent2]];
  lanes.forEach(([h, d, col], i) => {
    const x = 0.6 + i * 2.5;
    card(s, x, 2.0, 2.13, 1.25, { fill: C.background1, line: col, lw: 1.5, name: h });
    txt(s, h, x + 0.08, 2.14, 1.97, 0.55, { fontSize: 15, bold: true, align: "center", valign: "middle" });
    txt(s, d, x + 0.08, 2.72, 1.97, 0.4, { fontSize: 12, color: C.accent6, align: "center" });
    if (i < lanes.length - 1) arrow(s, x + 2.15, 2.62, x + 2.48, 2.62, C.accent6, 1.5, true);
  });
  txt(s, "JSON-RPC over standard input/output", 5.3, 3.35, 3.2, 0.3, { fontSize: 11, color: C.accent6, align: "center" });
  txt(s, "local connection to the model engine", 9.85, 3.35, 2.9, 0.3, { fontSize: 11, color: C.accent6, align: "center" });
  card(s, 0.6, 4.05, 12.13, 2.65, { name: "Example" });
  txt(s, "WHAT HAPPENS WHEN YOU ASK A QUESTION", 0.9, 4.25, 8, 0.3, { fontSize: 11, bold: true, color: C.accent6, charSpacing: 1.5 });
  const steps = [["You ask", "\"Which supplier was late most often in 2025?\""], ["Claude picks a tool", "run a DAX query on the model"], ["You approve", "you see the tool and its inputs first"], ["The model answers", "Pacific Rim Medical, 61.2% on time"]];
  steps.forEach(([h, d], i) => {
    const x = 0.9 + i * 2.95;
    num(s, x, 4.7, i + 1, i === 2 ? C.accent3 : C.accent2, 0.44);
    txt(s, h, x + 0.58, 4.7, 2.2, 0.4, { fontSize: 15, bold: true });
    txt(s, d, x + 0.58, 5.12, 2.25, 1.2, { fontSize: 13, color: C.accent6 });
  });
}
{
  const s = slide("CONTENT", "WHAT MCP IS", "One standard plug between AI and tools",
    "MCP IN ONE SLIDE (about 3 minutes)\n\nSay: MCP stands for Model Context Protocol. It's an open standard for connecting AI applications to tools and data, published by Anthropic in November 2024 and now used by many AI apps.\n\nSay: Before MCP, every AI app needed custom code for every tool. With MCP, a tool is built once as a server, and any app that speaks MCP can use it. That's why Microsoft could build one Power BI server that works with Claude Code, Claude Desktop, VS Code Copilot and others.\n\nSay: A server offers three kinds of things. Tools: actions Claude can ask to run, like list tables or run a DAX query. Resources: data Claude can read. Prompts: ready-made instructions. Today we only need tools.\n\nSay: And the safety rule is built into the conversation: Claude asks, you approve, then it runs. Plus, the Power BI server has a read-only mode, which we'll switch on.");
  pill(s, "mcp");
  txt(s, "Model Context Protocol", 0.6, 1.95, 6, 0.6, { fontSize: 26, bold: true, fontFace: "Arial" });
  txt(s, "An open standard, published by Anthropic in November 2024, for connecting AI applications to tools and data. Build a tool once as a server; any MCP-capable app can use it.", 0.6, 2.6, 5.6, 1.5, { fontSize: 16, color: C.accent6 });
  card(s, 0.6, 4.25, 5.6, 2.45, { fill: C.accent3, t: 88, name: "Before and after" });
  txt(s, "WITHOUT A STANDARD", 0.9, 4.45, 5, 0.3, { fontSize: 11, bold: true, color: C.accent3, charSpacing: 1.5 });
  txt(s, "Every AI app × every tool = custom code.", 0.9, 4.78, 5.1, 0.4, { fontSize: 15, bold: true });
  txt(s, "WITH MCP", 0.9, 5.4, 5, 0.3, { fontSize: 11, bold: true, color: C.accent3, charSpacing: 1.5 });
  txt(s, "One Power BI server works with Claude Code, Claude Desktop and other MCP clients.", 0.9, 5.73, 5.1, 0.8, { fontSize: 15, bold: true });
  const kinds = [["Tools", "Actions Claude can ask to run", "list tables · get a measure · run a DAX query", C.accent2], ["Resources", "Data Claude can read", "files, query results", C.accent1], ["Prompts", "Reusable instructions the server offers", "\"document this model\"", C.accent6]];
  kinds.forEach(([h, d, e, col], i) => {
    const y = 1.95 + i * 1.6;
    card(s, 6.6, y, 6.13, 1.42, { name: h });
    txt(s, h, 6.9, y + 0.18, 3, 0.45, { fontSize: 20, bold: true, color: col });
    txt(s, d, 6.9, y + 0.62, 5.6, 0.35, { fontSize: 14 });
    txt(s, e, 6.9, y + 0.97, 5.6, 0.35, { fontSize: 12, fontFace: MONO, color: C.accent6 });
  });
}
{
  const s = slide("CONTENT", "WHAT MCP IS · UNDER THE HOOD", "Inside the server: TOM and ADOMD.NET",
    "TOM AND ADOMD.NET (about 2 minutes)\n\nSay: So what is the Power BI server actually doing? It's a .NET program built on two Microsoft libraries. Both ship inside powerbi-modeling-mcp.exe; I checked the package.\n\nSay: The first is TOM, the Tabular Object Model. It reads and changes the model's structure: tables, columns, measures with their DAX, relationships. If you've used Tabular Editor, you've used TOM. When Claude lists tables or shows a measure's DAX, that's TOM reading. When Claude creates a measure, that's TOM writing, and that's exactly what --readonly blocks.\n\nSay: The second is ADOMD.NET. It runs DAX queries and brings the rows back. If you've used DAX Studio, you've used ADOMD.NET. Every number Claude reports, like 2025 revenue or a supplier's on-time % by quarter, comes back through ADOMD.NET.\n\nSay: Both connect to the same place: the Analysis Services engine running inside Power BI Desktop, on a local port. So MCP doesn't invent a new way into your model. It gives Claude the same access Tabular Editor and DAX Studio already have, and asks you before every step.\n\nIf asked how we know: the DLL names Microsoft.AnalysisServices.Tabular and Microsoft.AnalysisServices.AdomdClient are inside the server's exe. Which tool calls which library is inferred from what each library is designed for.");
  pill(s, "mcp");
  card(s, 0.6, 2.05, 3.0, 2.6, { fill: C.background1, line: C.accent2, lw: 1.5, name: "MCP server" });
  txt(s, "Power BI MCP server", 0.8, 2.3, 2.6, 0.45, { fontSize: 17, bold: true });
  txt(s, "powerbi-modeling-mcp.exe", 0.8, 2.8, 2.6, 0.35, { fontSize: 11.5, fontFace: MONO, color: C.accent6 });
  txt(s, "A .NET program. Both libraries are bundled inside it.", 0.8, 3.3, 2.6, 1.1, { fontSize: 13 });
  const libs = [
    ["TOM", "Microsoft.AnalysisServices.Tabular", "Reads and changes the model's structure", "same library as Tabular Editor", C.accent1, 1.9],
    ["ADOMD.NET", "Microsoft.AnalysisServices.AdomdClient", "Runs DAX queries and returns rows", "same library as DAX Studio", C.accent3, 3.45],
  ];
  libs.forEach(([h, dll, d, same, col, y]) => {
    card(s, 4.25, y, 4.55, 1.38, { fill: col, t: 90, line: col, name: h });
    txt(s, h, 4.45, y + 0.12, 4.1, 0.42, { fontSize: 19, bold: true, color: col });
    txt(s, dll, 4.45, y + 0.52, 4.2, 0.3, { fontSize: 10.5, fontFace: MONO, color: C.accent6 });
    txt(s, [{ text: d, options: { bold: true, breakLine: true } }, { text: same, options: { color: C.accent6 } }], 4.45, y + 0.8, 4.2, 0.55, { fontSize: 12.5 });
    arrow(s, 3.62, 3.35, 4.23, y + 0.69, C.accent6, 1.5);
    arrow(s, 8.82, y + 0.69, 9.43, 3.35, C.accent6, 1.5);
  });
  card(s, 9.45, 2.05, 3.28, 2.6, { fill: C.background1, line: C.accent2, lw: 1.5, name: "Engine" });
  txt(s, "Analysis Services engine", 9.65, 2.3, 2.9, 0.75, { fontSize: 17, bold: true });
  txt(s, "inside Power BI Desktop", 9.65, 3.05, 2.9, 0.35, { fontSize: 13, color: C.accent6 });
  txt(s, "localhost:<port>", 9.65, 3.55, 2.9, 0.35, { fontSize: 13, fontFace: MONO, color: C.accent2 });
  const H = t => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } });
  const rows = [[H("You ask Claude"), H("MCP tool"), H("Library")],
    [{ text: "\"List the tables\" · \"Show me the DAX for Gross Margin %\"" }, { text: "table_ / measure_operations", options: { fontFace: MONO } }, { text: "TOM · read", options: { bold: true, color: C.accent1 } }],
    [{ text: "\"Create Revenue QTD\"  (blocked when read-only)" }, { text: "measure_operations · Create", options: { fontFace: MONO } }, { text: "TOM · write", options: { bold: true, color: C.accent1 } }],
    [{ text: "\"What was revenue in 2025?\" · \"By quarter?\"" }, { text: "dax_query_operations", options: { fontFace: MONO } }, { text: "ADOMD.NET · DAX", options: { bold: true, color: C.accent3 } }],
  ];
  s.addTable(rows, { x: 0.6, y: 5.0, w: 12.13, colW: [5.6, 3.9, 2.63], fontSize: 13, color: "14171D", valign: "middle", border: { type: "solid", pt: 0.75, color: HEX.line }, rowH: 0.42, margin: [3, 10, 3, 10] });
}

section("Setup");
{
  const s = slide("CONTENT", "SETUP · STEP 1 OF 4", "Open the project, then check what's installed",
    "SETUP 1: OPEN AND CHECK (about 3 minutes, everyone does it with you)\n\nSay: Before we install anything, let's open the right folder and see what you already have.\n\nDo: In VS Code: File > Open Folder > choose Meridian_SupplyChain. Then open a terminal with Ctrl + backtick. Make sure the terminal says PowerShell (top right of the terminal panel).\n\nDo: In Power BI Desktop, open MeridianSupplyChain.pbix and leave it open.\n\nSay: Now type these three commands with me, one at a time.\n\nnode --version  -> v22 or higher is fine.\nclaude --version  -> something like 2.1.x (Claude Code) means Claude is installed.\nclaude doctor  -> runs a health check; no errors means Claude is healthy.\n\nSay: If any of them says 'is not recognized', that piece isn't installed yet. Hands up, and we fix it on the next slide.\n\nWhy this folder matters: Claude loads project Skills from the folder you start it in. We'll use that in the Skills demo.\n\nNote: the Power BI MCP server connects to Power BI Desktop on Windows. Mac users pair with a Windows user.");
  txt(s, "OPEN FIRST", 0.6, 1.85, 4, 0.3, { fontSize: 11, bold: true, color: C.accent6, charSpacing: 1.5 });
  const open = [["VS Code", "File → Open Folder → Meridian_SupplyChain"], ["Terminal", "Ctrl + `  ·  check it says PowerShell"], ["Power BI Desktop", "open MeridianSupplyChain.pbix, leave it open"]];
  open.forEach(([h, d], i) => {
    const y = 2.2 + i * 1.12;
    card(s, 0.6, y, 4.9, 0.98, { name: h });
    num(s, 0.82, y + 0.26, i + 1, C.accent2, 0.44);
    txt(s, h, 1.42, y + 0.12, 3.9, 0.4, { fontSize: 16, bold: true });
    txt(s, d, 1.42, y + 0.52, 3.95, 0.38, { fontSize: 12.5, color: C.accent6 });
  });
  txt(s, "THEN CHECK · TYPE EACH LINE", 5.85, 1.85, 6, 0.3, { fontSize: 11, bold: true, color: C.accent6, charSpacing: 1.5 });
  const rows = [
    [{ text: "Type this", options: { bold: true, color: C.background1, fill: { color: C.text1 } } }, { text: "Healthy result", options: { bold: true, color: C.background1, fill: { color: C.text1 } } }],
    [{ text: "node --version", options: { fontFace: MONO, bold: true } }, { text: "v22 or higher" }],
    [{ text: "claude --version", options: { fontFace: MONO, bold: true } }, { text: "2.x.x (Claude Code)" }],
    [{ text: "claude doctor", options: { fontFace: MONO, bold: true } }, { text: "no errors" }],
  ];
  s.addTable(rows, { x: 5.85, y: 2.2, w: 6.88, colW: [3.3, 3.58], fontSize: 15, color: "14171D", valign: "middle", border: { type: "solid", pt: 0.75, color: HEX.line }, rowH: [0.5, 0.72, 0.72, 0.72], margin: [4, 12, 4, 12] });
  card(s, 5.85, 5.1, 6.88, 1.2, { fill: C.accent3, t: 88, name: "Not recognized" });
  txt(s, [{ text: "See  ", options: {} }, { text: "'… is not recognized'", options: { bold: true, fontFace: MONO } }, { text: " ?\nThat piece isn't installed yet. Next slide.", options: {} }], 6.1, 5.1, 6.4, 1.2, { fontSize: 15, valign: "middle" });
}
{
  const s = slide("CONTENT", "SETUP · STEP 2 OF 4", "Install Claude Code and sign in",
    "SETUP 2: INSTALL AND SIGN IN (about 4 minutes; only for people whose check failed)\n\nSay: If claude wasn't recognised, three steps.\n\nOne: in the PowerShell terminal, paste the install line. It's Anthropic's official installer and needs no administrator rights.\n\nTwo: close the terminal (the bin icon) and open a new one with Ctrl + backtick. The new terminal is what finds the claude command. Then check: claude --version.\n\nThree: type claude and press Enter. A browser window opens; sign in with your Claude account, then come back to the terminal. Type /exit to leave Claude for now.\n\nSay: Claude Code needs a Pro, Max, Team or Enterprise account. The free plan doesn't include it.\n\nSay: If you're in CMD rather than PowerShell, you'll see 'irm is not recognized': use the CMD line at the bottom instead. winget and npm also work; npm needs Node 22 or later.\n\nSay: Some of you have the Claude extension in VS Code. Keep it, but the extension doesn't add the claude command to your terminal, and we need that command next.");
  const steps = [
    ["Install", "PowerShell · no admin rights needed", "irm https://claude.ai/install.ps1 | iex"],
    ["Open a NEW terminal", "close it (bin icon), then Ctrl + `", "claude --version"],
    ["Sign in", "a browser opens · sign in · come back", "claude"],
  ];
  steps.forEach(([h, d, cmd], i) => {
    const y = 1.85 + i * 1.22;
    card(s, 0.6, y, 8.0, 1.08, { name: h });
    num(s, 0.82, y + 0.3, i + 1, C.accent3, 0.46);
    txt(s, h, 1.48, y + 0.12, 2.65, 0.42, { fontSize: 17, bold: true });
    txt(s, d, 1.48, y + 0.56, 2.65, 0.42, { fontSize: 11.5, color: C.accent6 });
    card(s, 4.2, y + 0.22, 4.25, 0.64, { fill: C.text1, name: "Command" });
    txt(s, cmd, 4.33, y + 0.22, 4.05, 0.64, { fontSize: 12, fontFace: MONO, color: "F2B544", valign: "middle", bold: true });
  });
  card(s, 8.85, 1.85, 3.88, 3.52, { fill: C.accent5, t: 92, name: "Account" });
  txt(s, "ACCOUNT", 9.1, 2.05, 3.4, 0.3, { fontSize: 11, bold: true, color: C.accent5, charSpacing: 1.5 });
  txt(s, "Pro, Max, Team or Enterprise", 9.1, 2.4, 3.45, 0.75, { fontSize: 18, bold: true });
  txt(s, "The free plan doesn't include Claude Code.", 9.1, 3.2, 3.45, 0.7, { fontSize: 14 });
  txt(s, "Have the VS Code extension? Keep it, but install this too: the extension doesn't add the claude command to the terminal.", 9.1, 3.95, 3.45, 1.3, { fontSize: 12.5, color: C.accent6 });
  txt(s, "OTHER WAYS TO INSTALL", 0.6, 5.62, 6, 0.3, { fontSize: 11, bold: true, color: C.accent6, charSpacing: 1.5 });
  const alts = [["In CMD", "curl -fsSL https://claude.ai/install.cmd -o install.cmd && install.cmd && del install.cmd"], ["winget", "winget install Anthropic.ClaudeCode"], ["npm · Node 22+", "npm install -g @anthropic-ai/claude-code"]];
  txt(s, alts.map(([k, v], i) => [{ text: k + "   ", options: { bold: true } }, { text: v, options: { fontFace: MONO, color: C.accent6, breakLine: i < alts.length - 1 } }]).flat(), 0.6, 5.95, 12.13, 0.9, { fontSize: 11.5, paraSpaceAfter: 2 });
}
{
  const s = slide("CONTENT", "SETUP · STEP 3 OF 4", "Register the Power BI server, read-only",
    "SETUP 3: ONE COMMAND (about 4 minutes)\n\nSay: This one command tells Claude Code about the Power BI server. Let's read it before we run it, because reading before running is a habit we want.\n\nSay: claude mcp add: register a server. --scope user: register it once for your user account, so it works in every folder. powerbi-modeling is just the name we give it. The two dashes are in quotes on purpose: in PowerShell a bare -- can be swallowed before Claude sees it, and you'd get 'unknown option -y'. Everything after the dashes is how to start it: cmd /c npx downloads and runs Microsoft's package; on Windows, npx needs cmd /c in front.\n\nSay: Two flags at the end. --start starts the server. --readonly is the important one: Microsoft calls it safe mode. It prevents any write operation to your model. Today, Claude can look but not touch.\n\nDo: Everyone runs it, then claude mcp list. First run takes 20-30 seconds while it downloads. Wait for powerbi-modeling to show as connected.\n\nIf the server asks you to accept its licence terms on first use: read it and accept. It blocks other tool calls until you do.\n\nIf it says the name already exists: claude mcp remove powerbi-modeling, then run the add command again.");
  code(s, 0.6, 1.9, 12.13, 1.05, [
    [["claude mcp add --scope user powerbi-modeling '--' cmd /c npx -y @microsoft/powerbi-modeling-mcp@latest ", "x"], ["--start --readonly", "h"]],
    [["claude mcp list", "x"]],
  ], { size: 12, lang: "text" });
  const parts = [
    ["claude mcp add", "Register a server with Claude Code.", C.text1],
    ["--scope user", "Once for your account: works in every folder.", C.text1],
    ["'--'", "In quotes, so PowerShell passes it through.", C.accent2],
    ["cmd /c npx …", "Download and run Microsoft's server package.", C.accent2],
    ["--readonly", "Safe mode: no changes to your model.", C.accent3],
  ];
  const cw = (12.13 - 4 * 0.15) / 5;
  parts.forEach(([k, d, col], i) => {
    const x = 0.6 + i * (cw + 0.15);
    card(s, x, 3.25, cw, 2.2, { fill: i === 4 ? C.accent3 : C.background2, t: i === 4 ? 88 : 0, name: k });
    txt(s, k, x + 0.2, 3.45, cw - 0.4, 0.45, { fontSize: 14, bold: true, fontFace: MONO, color: col });
    txt(s, d, x + 0.2, 4.0, cw - 0.4, 1.3, { fontSize: 14 });
  });
  card(s, 0.6, 5.7, 12.13, 0.95, { fill: C.accent4, t: 90, name: "Check" });
  txt(s, [{ text: "Check:  ", options: { bold: true } }, { text: "claude mcp list", options: { fontFace: MONO, bold: true } }, { text: "  shows  ", options: {} }, { text: "powerbi-modeling", options: { fontFace: MONO, bold: true } }, { text: "  as connected. The first run takes 20–30 seconds.", options: {} }], 0.9, 5.7, 11.6, 0.95, { fontSize: 15, valign: "middle" });
}
{
  const s = slide("CONTENT", "SETUP · IF SOMETHING FAILS", "Quick fixes for the setup",
    "QUICK FIXES (keep this slide handy; show it when someone is stuck)\n\nSay: If anything went wrong, find your error message in the left column. The fix is next to it.\n\nThe two you'll see most: 'claude is not recognized' (they didn't open a new terminal after installing) and 'unknown option -y' (the dashes weren't in quotes).\n\nDon't let one person hold up the room: pair them with a neighbour and keep going. They can catch up during the hands-on.");
  const H = t => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } });
  const E = t => ({ text: t, options: { fontFace: MONO, bold: true, color: C.accent5 } });
  const rows = [[H("You see"), H("Fix")],
    [E("'claude' is not recognized"), { text: "Install Claude Code (step 2), then open a NEW terminal" }],
    [E("'irm' is not recognized"), { text: "You're in CMD: use the CMD install line, or switch the terminal to PowerShell" }],
    [E("error: unknown option '-y'"), { text: "Put the two dashes in quotes:  '--'" }],
    [E("'node' / 'npx' is not recognized"), { text: "Install Node.js 22 or later from nodejs.org, then restart VS Code" }],
    [E("powerbi-modeling not connected"), { text: "claude mcp remove powerbi-modeling, then run the add command again" }],
    [E("sign-in: no access to Claude Code"), { text: "Free plan. Claude Code needs Pro, Max, Team or Enterprise" }],
    [E("server blocks every tool call"), { text: "Accept the licence terms it shows the first time" }],
    [E("no Power BI Desktop instance found"), { text: "Open the .pbix and wait until the visuals have loaded" }],
  ];
  s.addTable(rows, { x: 0.6, y: 1.85, w: 12.13, colW: [4.7, 7.43], fontSize: 13.5, color: "14171D", valign: "middle", border: { type: "solid", pt: 0.75, color: HEX.line }, rowH: 0.5, margin: [3, 12, 3, 12] });
}
{
  const s = slide("LIVE", "SETUP · STEP 4 OF 4 · EVERYONE DOES THIS", "Connect Claude to the open report",
    "SETUP 4: CONNECT (about 5 minutes)\n\nDo: With MeridianSupplyChain.pbix open in Desktop, type claude in the terminal (started in the Meridian_SupplyChain folder), then type exactly: Connect to 'MeridianSupplyChain' in Power BI Desktop. This wording comes from Microsoft's own examples.\n\nSay: Watch what Claude does. First it asks the server to list Power BI Desktop instances running on this machine. Then it connects to the one called MeridianSupplyChain. Each of those is a tool call you approve.\n\nSay: Read the approval box before you click. It shows the tool name and its inputs. Choose Yes, not 'don't ask again', so you see every request today. This habit, reading what you approve, is the most important skill in this programme.\n\nDo (30 seconds): type /mcp and open powerbi-modeling to show its tool list. Say: Remember on slide 5 I said the server sends Claude a list of tools? Here it is. Claude can only use what's on this list.\n\nReal output from this file: 'Found 1 local Power BI Desktop instance', then 'Connection established'. Your port number will be different; that's normal.\n\nIf no instance is found: the report must be open in Desktop and finished loading, not just Desktop's start page.");
  liveSteps(s, [["Desktop: MeridianSupplyChain.pbix open", "Wait until the visuals have loaded"], ["Terminal: type  claude", "Started in the Meridian_SupplyChain folder"], ["Ask it to connect", "Connect to 'MeridianSupplyChain' in Power BI Desktop"], ["Read, then approve, each tool call", "Choose Yes, not 'don't ask again'"]], 0.6, 1.95, 6.6);
  code(s, 7.45, 1.95, 5.28, 3.1, [
    [["> ", "c"], ["Connect to 'MeridianSupplyChain'", "x"]],
    [["  in Power BI Desktop", "x"]],
    [["", "x"]],
    [["● ", "s"], ["connection_operations · ListLocalInstances", "p"]],
    [["  Found 1 local Power BI Desktop instance", "c"]],
    [["  MeridianSupplyChain · localhost:55295", "c"]],
    [["", "x"]],
    [["● ", "s"], ["connection_operations · Connect", "p"]],
    [["  Connection established", "a"]],
  ], { dark: true, size: 12.5 });
  txt(s, "Real output from this file. Your port number will differ. Then try  /mcp  to see the server's tool list.", 7.45, 5.15, 5.28, 0.6, { fontSize: 12, italic: true, color: C.background2, transparency: 35 });
}

section("Ask and verify");
{
  const s = slide("CONTENT", "READING THE MODEL", "Four things Claude can read",
    "WHAT CLAUDE CAN READ (about 4 minutes)\n\nSay: Once connected, Claude can read the four building blocks of every semantic model. Tables. Columns, with their data types. Measures, including the full DAX. And relationships: which tables connect, which side is many, and which ones are inactive.\n\nSay: Under each card is a question you can ask in plain English, and the tool Claude will use to answer it. You don't need to know the tool names; Claude picks them. But you will see them in the approval box, so it helps to recognise them.\n\nDo: Live, ask the first one: 'List the tables in this model, and say which ones are facts and which are dimensions.' Read the answer out. Then ask: 'Show me the DAX for Gross Margin % and the table it lives in.'\n\nReal answer: Gross Margin % lives in _Measures, folder 01 Sales: DIVIDE ( [Gross Margin], [Revenue] ), format 0.0%.");
  pill(s, "tmdl");
  const cards = [
    ["Tables", "10", "\"List the tables, and say which are facts and which are dimensions.\"", "table_operations · List", C.accent2],
    ["Columns", "84", "\"What columns does Fact_Sales have, and which are hidden?\"", "column_operations · List", C.accent2],
    ["Measures", "30", "\"Show me the DAX for Gross Margin % and where it lives.\"", "measure_operations · Get", C.accent3],
    ["Relationships", "15", "\"Which relationships are inactive, and why would a model need them?\"", "relationship_operations · List", C.accent1],
  ];
  cards.forEach(([h, n, q, tool, col], i) => {
    const x = 0.6 + (i % 2) * 6.14, y = 1.85 + Math.floor(i / 2) * 2.5;
    card(s, x, y, 5.99, 2.32, { name: h });
    txt(s, n, x + 0.3, y + 0.2, 1.6, 0.7, { fontSize: 34, bold: true, color: col, fontFace: "Arial" });
    txt(s, h, x + 1.85, y + 0.32, 3.8, 0.5, { fontSize: 20, bold: true });
    txt(s, q, x + 0.3, y + 1.0, 5.4, 0.7, { fontSize: 14, italic: true });
    txt(s, tool, x + 0.3, y + 1.78, 5.4, 0.3, { fontSize: 12, fontFace: MONO, color: C.accent6 });
  });
}
{
  const s = slide("LIVE", "HANDS-ON · 12 MINUTES · EVERYONE", "Ask three questions, check three answers",
    "HANDS-ON (about 12 minutes)\n\nSay: Your turn. Ask Claude these three questions, one at a time. Before you check anything, write down Claude's answer. Then find the same fact in Power BI Desktop. Only then compare.\n\nSay: Why write it down first? Because once you've seen the right answer, it's very easy to believe Claude got it right too.\n\nExpected answers (from the real model):\n1. 10 tables, 30 measures, 15 relationships. Check: the Data pane, and Modeling > Manage relationships.\n2. Three inactive relationships: Sales_ShipDate, Purchases_PromisedDate, Purchases_ReceivedDate. They exist because Fact_Sales and Fact_Purchases have more than one date column pointing at the same date table; only one can be active, and measures switch the others on with USERELATIONSHIP. Check: Model view, dotted lines.\n3. 2025 revenue: $14.13M (14,134,114.90). Worst supplier on-time % in 2025: Pacific Rim Medical, 61.2%. Check: Executive Overview card with Year 2025; Supplier Performance page with Year 2025.\n\nAsk at the end: 'Did anyone get a different answer from Claude? What did it do differently?' Discuss. This is the moment to say: Claude is usually right on metadata, and a DAX query result is exact, but it can misread a question. That is why we check.\n\nStory hook for question 3: 'Pacific Rim is on 61% for the year. Was it always that bad? Ask Claude to break it down by quarter.' (It drops sharply in H2 2025. Don't tell them; let them find it.)");
  const qs = [
    ["How many tables, measures and relationships does this model have?", "Data pane · Manage relationships"],
    ["Which relationships are inactive, and why might a model need them?", "Model view · dotted lines"],
    ["What was revenue in 2025, and which supplier had the worst on-time % that year?", "Year = 2025 · Executive Overview and Supplier Performance pages"],
  ];
  qs.forEach(([q, chk], i) => {
    const y = 1.95 + i * 1.45;
    num(s, 0.6, y + 0.08, i + 1, C.accent3, 0.5);
    txt(s, q, 1.3, y, 7.0, 0.75, { fontSize: 18, bold: true, color: C.background1 });
    txt(s, [{ text: "Check in Desktop: ", options: { bold: true, color: "F2B544" } }, { text: chk, options: {} }], 1.3, y + 0.78, 7.0, 0.4, { fontSize: 13, color: C.background2, transparency: 20 });
  });
  card(s, 8.85, 1.95, 3.88, 4.4, { fill: C.background1, t: 92, line: HEX.lineDark, name: "The rule" });
  txt(s, "THE RULE", 9.15, 2.2, 3.4, 0.3, { fontSize: 11, bold: true, color: C.accent3, charSpacing: 1.5 });
  txt(s, "Write down Claude's answer first.", 9.15, 2.6, 3.4, 0.8, { fontSize: 19, bold: true, color: C.background1 });
  txt(s, "Then find the same fact in Desktop. Then compare.", 9.15, 3.45, 3.4, 0.9, { fontSize: 15, color: C.background2, transparency: 20 });
  txt(s, "If they differ, ask Claude how it got its answer.", 9.15, 4.5, 3.4, 0.9, { fontSize: 15, color: "F2B544" });
}

section("Skills");
{
  const s = slide("CONTENT", "SKILLS · THE IDEA", "What a Skill is",
    "WHAT A SKILL IS (about 2 minutes)\n\nSay: You've just typed three questions and checked three answers. Now imagine doing that same check on every report, every morning, or before every release. You'd write it down as a checklist, a standard procedure, so anyone could follow it. A Skill is exactly that, written for Claude.\n\nSay: A Skill is a saved procedure. Claude follows it the same way every time, for anyone who runs it.\n\nSay (point at the code): It's a folder with one file in it, called SKILL.md. Three parts: a name, a description that tells Claude when to use it, and the steps in plain English, the way you'd explain them to a new colleague.\n\nSay (point at the right): Where you save it decides who gets it. Inside the project, it travels with the project through Git, so the whole team gets the same Skill. In your home folder, it's yours in every project.\n\nSay: And you call it back by name with a slash, or just by asking in your own words.");
  card(s, 0.6, 1.85, 12.13, 0.95, { fill: C.accent2, t: 88, name: "Definition" });
  txt(s, [{ text: "A Skill is a saved procedure.  ", options: { bold: true, color: C.accent2 } }, { text: "You write the steps once, in plain English. Claude follows them the same way every time, for anyone on the team.", options: {} }], 0.9, 1.85, 11.6, 0.95, { fontSize: 17, valign: "middle" });
  txt(s, "A folder with one file: SKILL.md", 0.6, 3.05, 5.6, 0.45, { fontSize: 19, bold: true, fontFace: "Arial" });
  code(s, 0.6, 3.55, 5.6, 2.1, [
    [["---", "c"]],
    [["name", "p"], [": model-health-check", "x"]],
    [["description", "p"], [": Read-only health check", "x"]],
    [["  of the connected model. Use when…", "x"]],
    [["---", "c"]],
    [["# Steps in plain English", "k"]],
    [["1. Count tables, measures, relationships…", "x"]],
  ], { size: 13 });
  const parts = [["name", "what you call it"], ["description", "tells Claude when to use it"], ["steps", "the procedure, in plain English"]];
  parts.forEach(([k, d], i) => txt(s, [{ text: k + "  ", options: { bold: true, fontFace: MONO, color: C.accent2 } }, { text: d, options: { color: C.accent6 } }], 0.6, 5.78 + i * 0.32, 5.6, 0.3, { fontSize: 13 }));
  const where = [["Project Skill", ".claude/skills/<name>/SKILL.md", "Inside the project. Shared with the team through Git.", C.accent2], ["Personal Skill", "~/.claude/skills/<name>/SKILL.md", "In your home folder. Yours, in every project.", C.accent1]];
  where.forEach(([h, p, d, col], i) => {
    const y = 3.05 + i * 1.3;
    card(s, 6.55, y, 6.18, 1.17, { name: h });
    txt(s, h, 6.8, y + 0.12, 2.6, 0.4, { fontSize: 17, bold: true, color: col });
    txt(s, p, 9.15, y + 0.15, 3.5, 0.35, { fontSize: 11.5, fontFace: MONO, color: C.accent6 });
    txt(s, d, 6.8, y + 0.56, 5.7, 0.5, { fontSize: 14 });
  });
  card(s, 6.55, 5.7, 6.18, 1.05, { fill: C.accent3, t: 88, name: "Recall" });
  txt(s, [{ text: "Call it back:  ", options: { bold: true } }, { text: "/model-health-check", options: { bold: true, fontFace: MONO, color: C.accent3 } }, { text: "   or just ask: \"Give this model a health check.\"", options: {} }], 6.8, 5.7, 5.8, 1.05, { fontSize: 14, valign: "middle" });
}
{
  const s = slide("CONTENT", "SKILLS · WHAT THEY DO AT WORK", "What Skills can do for a Power BI team",
    "WHAT SKILLS CAN DO (about 3 minutes; keep it brief, each one gets its own session later)\n\nSay: Here's why Skills matter for anyone who builds or checks Power BI reports every day. Six jobs, all things you do by hand today. Every one of them works with what we've set up, and we'll build each one properly in a later session.\n\n1. Say: Daily health check. Every morning: do the key measures return values, without errors? Anything missing a description or a format? That's the one I'll show you in two minutes.\n\n2. Say: Source against report. Soon Meridian's data moves from Excel to SQL Server, and we'll connect a SQL Server MCP server next to the Power BI one. Then a Skill can ask the same question twice, revenue by month from SQL and from the model, and flag any month that doesn't match.\n\n3. Say: Threshold alerts. You write the thresholds into the Skill: revenue above a target, a supplier's on-time rate below 80 percent. Open the report, run the Skill, and it lists what crossed a line. Be clear: it runs when you run it. It's not a push notification on its own.\n\n4. Say: What changed. Once the report is saved as PBIP, which is next session, every change is a file change. A Skill can read the Git diff and summarise it: which measures, columns, relationships or visuals changed since the last version.\n\n5. Say: Visual and formatting check. The report pages are files too. A Skill can find visuals that point at a field that no longer exists, missing titles, and colours or fonts that don't match the theme.\n\n6. Say: Performance. A Skill can list the biggest, highest-cardinality columns and time the slowest measures. And once we're on SQL Server, it can read the Power Query steps and flag the ones likely to break query folding, like custom columns built with functions SQL doesn't have. You then confirm in Power Query with View Native Query. Claude points you to the step; you check it.\n\nSay: Today, just the idea. We'll build these one by one.\n\nIf asked: the SQL Server MCP and PBIP are not part of today's setup. Items 1, 3 and 6 (model side) work with today's read-only connection.");
  const apps = [
    ["Daily health check", "Measures return values, with no errors. Missing descriptions and formats flagged.", "Power BI MCP · works today", C.accent3],
    ["Source vs report", "Ask SQL Server and the model the same question, such as revenue by month, and flag any gap.", "SQL Server MCP + Power BI MCP", C.accent2],
    ["Threshold alerts", "Thresholds live in the Skill: revenue above target, on-time below 80%. Run it, see what crossed.", "DAX queries · runs when you run it", C.accent3],
    ["What changed", "Summarise the measures, columns, relationships and visuals that changed since the last version.", "PBIP files + Git diff · Session 2", C.accent1],
    ["Visual and format check", "Visuals pointing at missing fields, missing titles, colours or fonts off the theme.", "PBIR report files", C.accent1],
    ["Performance", "Biggest columns, slowest measures, and Power Query steps likely to break query folding.", "Model stats + Power Query steps", C.accent2],
  ];
  apps.forEach(([h, d, use, col], i) => {
    const x = 0.6 + (i % 3) * 4.09, y = 1.85 + Math.floor(i / 3) * 2.05;
    card(s, x, y, 3.95, 1.9, { name: h });
    num(s, x + 0.22, y + 0.22, i + 1, col, 0.44);
    txt(s, h, x + 0.82, y + 0.2, 3.0, 0.48, { fontSize: 16, bold: true, valign: "middle" });
    txt(s, d, x + 0.22, y + 0.78, 3.55, 0.72, { fontSize: 12.5 });
    txt(s, use, x + 0.22, y + 1.5, 3.55, 0.3, { fontSize: 10.5, bold: true, color: col });
  });
  card(s, 0.6, 6.0, 12.13, 0.8, { fill: C.accent3, t: 88, name: "Path" });
  txt(s, [{ text: "Where this goes:  ", options: { bold: true } }, { text: "Meridian moves from Excel to SQL Server, we connect a SQL Server MCP next to Power BI's, and we build these Skills one session at a time.", options: {} }], 0.9, 6.0, 11.6, 0.8, { fontSize: 14, valign: "middle" });
}
{
  const s = slide("CONTENT", "SKILLS · ONE YOU CAN USE TOMORROW", "A Skill for every morning: report-snapshot",
    "REPORT-SNAPSHOT (about 2 minutes)\n\nSay: Here's a real Skill, built for this report, that I'll run in a minute. It's called report-snapshot, and it answers five questions you'd otherwise check by hand every morning.\n\nSay (point at each row): One: when was the data last refreshed? Two: what are the key KPIs this year against last year, and has any crossed a threshold we set? Three: how many relationships are there, what types, and which are inactive? Four: how long does each page take to render, and which visual is the slowest? Five: what did Power Query do to each table?\n\nSay (point at the grey tags): Each answer comes from a different place. One to three are DAX queries through the Power BI server: that's ADOMD.NET from slide 6. Four comes from Performance Analyzer: you record it once in Desktop and export it, because Claude can't click through Desktop for you. Five comes from the model's TMDL files.\n\nSay (point at the dark panel): And this is what it found on Meridian, for real. Two flags: on-time delivery is 89.8%, under our 90% line, and 45 stockouts, over our 40 line.\n\nIf asked: the Skill is a SKILL.md plus two small Python scripts that read the Performance Analyzer file and the Power Query steps, so those two parts give the same output every run. Thresholds are a table in SKILL.md; edit it for your team.\n\nChecked against the live model: last refresh 9 Nov 2026 12:25 UTC (17:55 IST). 2025 vs 2024: Revenue 14,134,115 vs 12,840,263 (+10.1%); Gross Margin 37.1% vs 38.4%; On-Time Delivery 89.8% vs 94.2%; Supplier On-Time 91.0% vs 93.4%; OTIF 87.6% vs 91.0%; Stockouts 45 vs 43. Relationships 15: 12 active, 3 inactive, all many-to-one, single direction. Page render: Executive Overview 1,060 ms, Supplier Performance 1,012 ms, Inventory Health 792 ms; slowest visual the KPI card on Executive Overview, 954 ms. Power Query: 9 tables read CSV files, promote headers, change types; Fact_Sales and Fact_Purchases also replace blanks with null.");
  const checks = [
    ["Last refreshed", "When each table last loaded", "DAX · INFO.PARTITIONS"],
    ["Key KPI values", "This year vs last, flagged against thresholds you set", "DAX query · ADOMD.NET"],
    ["Relationships", "Total, types, active and inactive", "DAX · INFO.VIEW.RELATIONSHIPS"],
    ["Render time per page and visual", "Each page, then its visuals slowest first", "Performance Analyzer export + PBIR files"],
    ["Power Query steps", "Each table's source and applied steps, in plain English", "TMDL files"],
  ];
  checks.forEach(([h, d, src], i) => {
    const y = 1.85 + i * 0.82;
    card(s, 0.6, y, 6.6, 0.72, { name: h });
    num(s, 0.75, y + 0.15, i + 1, i < 3 ? C.accent3 : C.accent1, 0.42);
    txt(s, h, 1.32, y + 0.04, 3.4, 0.36, { fontSize: 14.5, bold: true });
    txt(s, d, 1.32, y + 0.37, 3.5, 0.32, { fontSize: 11.5, color: C.accent6 });
    txt(s, src, 4.75, y, 2.35, 0.72, { fontSize: 10.5, bold: true, color: i < 3 ? C.accent3 : C.accent1, valign: "middle", align: "right" });
  });
  code(s, 7.45, 1.85, 5.28, 4.0, [
    [["report-snapshot · Meridian", "k"]],
    [["1 Last refreshed  ", "p"], ["9 Nov 2026, 17:55 IST", "x"]],
    [["2 Revenue 2025    ", "p"], ["14.13M  (+10.1%)", "x"]],
    [["  On-time deliv.  ", "p"], ["89.8%  ", "x"], ["FLAG < 90%", "c"]],
    [["  Stockouts       ", "p"], ["45     ", "x"], ["FLAG > 40", "c"]],
    [["3 Relationships   ", "p"], ["15: 12 active, 3 inactive", "x"]],
    [["                  ", "p"], ["all many-to-one, single", "x"]],
    [["4 Slowest visual  ", "p"], ["KPI card, 954 ms", "x"]],
    [["  Pages (ms)      ", "p"], ["1,060 · 1,012 · 792", "x"]],
    [["5 Power Query     ", "p"], ["9 tables: CSV, headers,", "x"]],
    [["                  ", "p"], ["types; 2 facts blank→null", "x"]],
  ], { dark: true, size: 13.5 });
  card(s, 0.6, 6.05, 12.13, 0.75, { fill: C.accent3, t: 88, name: "Where" });
  txt(s, [{ text: "In the project:  ", options: { bold: true } }, { text: ".claude/skills/report-snapshot/", options: { fontFace: MONO, color: C.accent3, bold: true } }, { text: "  holds SKILL.md and two small scripts. Read-only: it never changes the model.", options: {} }], 0.9, 6.05, 11.6, 0.75, { fontSize: 13.5, valign: "middle" });
}
{
  const s = slide("LIVE", "LIVE DEMO · 4 MINUTES", "Save, recall and run a Skill",
    "LIVE DEMO: REPORT-SNAPSHOT (about 4 minutes)\n\nBefore the session: open the PBIP project folder (1st Session) in VS Code, with MeridianSupplyChain.pbip open in Power BI Desktop. Record Performance Analyzer once (Optimize > Performance analyzer > Start recording > Refresh visuals > click Supplier Performance and Inventory Health > Stop > Export) and save PowerBIPerformanceData.json next to the .pbip. It's already there from rehearsal; re-record it if you change the report.\n\nDo: In VS Code, open .claude/skills/report-snapshot/SKILL.md. Scroll through it: the description, the read-only rule, the five steps, the threshold table, the output format.\n\nSay: This is all a Skill is. Saved here, in the project, it travels with the project through Git. To make your own, ask Claude Code to draft one, read it, then save it.\n\nDo: Start a new Claude Code session (/exit, then claude) so it picks up the Skill. Type /report-snapshot. If your version doesn't list it, ask: 'Give me today's report snapshot.' Approve each read-only tool call, reading the DAX in each approval out loud.\n\nExpected result (checked against the live model): last refresh 9 Nov 2026 12:25 UTC (17:55 IST); 2025 Revenue 14,134,115, +10.1% on 2024; flags on On-Time Delivery 89.8% (below 90%) and Stockouts 45 (above 40); 15 relationships, 12 active, 3 inactive (Sales_ShipDate, Purchases_PromisedDate, Purchases_ReceivedDate), used by Revenue by Ship Date and Spend by Received Date; pages 1,060 / 1,012 / 792 ms, slowest the KPI card on Executive Overview at 954 ms; 9 tables from CSV, 2 with blank-to-null steps. Render times change a little each recording.\n\nSay: Same five checks, same format, every morning, for anyone who opens this project.");
  liveSteps(s, [["Save it in the project", ".claude/skills/report-snapshot/: SKILL.md + 2 scripts"], ["Record performance once", "Optimize › Performance analyzer › Record › visit each page › Export"], ["Restart Claude Code", "/exit, then claude, so it loads the Skill"], ["Recall it", "/report-snapshot, or \"Give me today's report snapshot\""]], 0.6, 1.95, 6.6);
  code(s, 7.45, 1.95, 5.28, 3.55, [
    [["---", "c"]],
    [["name", "p"], [": report-snapshot", "x"]],
    [["description", "p"], [": Read-only daily snapshot", "x"]],
    [["  of the open Power BI report…", "x"]],
    [["---", "c"]],
    [["## Rules", "k"]],
    [["- Read only. Never change anything.", "x"]],
    [["## Steps", "k"]],
    [["1. Last refreshed   2. Key KPIs", "x"]],
    [["3. Relationships    4. Render times", "x"]],
    [["5. Power Query steps", "x"]],
    [["## Output: 5 headings + Flags", "k"]],
  ], { dark: true, size: 12 });
  txt(s, [{ text: "On Meridian it flags: ", options: { bold: true, color: "F2B544" } }, { text: "On-Time Delivery 89.8% (below 90%) and 45 stockouts (above 40).", options: {} }], 7.45, 5.62, 5.28, 0.6, { fontSize: 13, color: C.background2 });
}

section("Under the hood");
{
  const s = slide("CONTENT", "WHAT CLAUDE IS ACTUALLY READING", "Underneath: the model as files",
    "UNDER THE HOOD, OVERVIEW ONLY (about 3 minutes)\n\nSay: Today Claude read the model live, from the Desktop you have open. Next session we go one level deeper.\n\nSay: The PBIX you opened today is one sealed file: model, data and report together. Power BI can also save the same report as a project, a PBIP. Then it becomes a folder of plain text files: one folder for the report, stored in a format called PBIR, and one for the semantic model, stored in a format called TMDL, with one text file per table.\n\nSay: Why should you care? Because text files can be read by Claude, compared in Git, and undone with one command. In session two, we save Meridian as a PBIP, open it in VS Code, have Claude make a change across the model, read the change as a Git diff, and undo it.\n\nSay: If you want to go deeper before then, the interactive map has the whole structure: every file, what it holds, and what's safe to touch. The links are on the right.\n\nDon't go into TMDL syntax today; that's session two.");
  card(s, 0.6, 2.0, 3.1, 3.2, { line: HEX.line, name: "PBIX" });
  txt(s, "MeridianSupplyChain.pbix", 0.85, 2.25, 2.7, 0.4, { fontSize: 13, bold: true, fontFace: MONO });
  txt(s, "Today", 0.85, 2.7, 2.6, 0.35, { fontSize: 13, bold: true, color: C.accent3 });
  txt(s, "One sealed file: model, data and report together. Claude reads the live model through the MCP server.", 0.85, 3.1, 2.65, 1.9, { fontSize: 14 });
  arrow(s, 3.8, 3.6, 4.55, 3.6, C.text1, 2);
  txt(s, "Save as\nPBIP", 3.75, 2.95, 0.85, 0.6, { fontSize: 11, color: C.accent6, align: "center" });
  card(s, 4.65, 1.85, 4.55, 3.55, { fill: C.background1, line: C.text1, lw: 1.5, name: "PBIP" });
  txt(s, "Next session: a folder of text files", 4.9, 2.0, 4.1, 0.35, { fontSize: 13, bold: true, color: C.accent3 });
  [[C.accent1, ".Report", "PBIR · a JSON file per page and per visual"], [C.accent2, ".SemanticModel", "TMDL · a text file per table"]].forEach(([col, h, d], i) => {
    const y = 2.5 + i * 1.4;
    card(s, 4.9, y, 4.05, 1.2, { fill: col, t: 90, line: col, name: h });
    txt(s, h, 5.1, y + 0.12, 3.7, 0.4, { fontSize: 16, bold: true, fontFace: MONO, color: col });
    txt(s, d, 5.1, y + 0.55, 3.7, 0.5, { fontSize: 13 });
  });
  card(s, 9.55, 1.85, 3.18, 3.55, { fill: C.accent3, t: 90, name: "Go deeper" });
  txt(s, "GO DEEPER BEFORE SESSION 2", 9.8, 2.05, 2.8, 0.5, { fontSize: 11, bold: true, color: C.accent3, charSpacing: 1.2 });
  const links = [["The project folder", "folder"], ["The model in TMDL", "tmdl"], ["Relationships", "relationships"], ["The report in PBIR", "report"], ["How a number reaches the page", "overview"]];
  links.forEach(([l, v], i) => s.addText([{ text: l + "  ↗", options: { color: C.accent1, hyperlink: { url: `${MAP}#${v}`, tooltip: "Opens the interactive map" } } }], { x: 9.8, y: 2.6 + i * 0.52, w: 2.85, h: 0.4, fontSize: 13, bold: true, margin: 0, isTextBox: true }));
  card(s, 0.6, 5.7, 12.13, 1.0, { fill: C.background2, name: "Why" });
  txt(s, [{ text: "Why it matters: ", options: { bold: true } }, { text: "text files can be read by Claude, compared in Git and undone with one command. That is Session 2.", options: { color: C.accent6 } }], 0.9, 5.7, 11.6, 1.0, { fontSize: 16, valign: "middle" });
}

section("Close");
{
  const s = slide("CONTENT", "BEFORE YOU USE THIS AT WORK", "Four rules for Claude on a Power BI model",
    "THE RULES (about 2 minutes)\n\nSay: Four rules, and they apply from today.\n\nOne: read-only first. On any new model, start with --readonly. Write access is a decision you make later, on purpose.\n\nTwo: read every approval. The tool name and its inputs tell you exactly what's about to happen.\n\nThree: check every answer. Metadata answers are usually right; a DAX query result is exact; but Claude can still misread your question. Check against Desktop, like you did today.\n\nFour: data leaves your machine. Microsoft's own documentation says it plainly: AI assistance doesn't expand your data access, but it may send the data it reads to the AI provider. So: training data or approved models only, never a client's model without permission, and follow our AI policy.\n\nAsk: 'Which rule would have saved someone in your team trouble already?' Take one or two answers.");
  const rules = [["Read-only first", "Start every new model in read-only mode. Write access is a later, deliberate choice.", C.accent3], ["Read every approval", "The tool name and its inputs say exactly what will happen.", C.accent2], ["Check every answer", "Compare with Desktop. Claude can misread a question.", C.accent1], ["Data leaves your machine", "What Claude reads may go to the AI provider. Approved models only.", C.accent5]];
  rules.forEach(([h, d, col], i) => {
    const x = 0.6 + i * 3.08;
    card(s, x, 1.95, 2.93, 3.6, { name: h });
    txt(s, String(i + 1), x + 0.28, 2.15, 1, 0.8, { fontSize: 44, bold: true, color: col, fontFace: "Arial" });
    txt(s, h, x + 0.28, 3.05, 2.45, 0.7, { fontSize: 18, bold: true });
    txt(s, d, x + 0.28, 3.8, 2.45, 1.6, { fontSize: 14, color: C.accent6 });
  });
  card(s, 0.6, 5.8, 12.13, 0.95, { fill: C.accent5, t: 92, name: "Quote" });
  txt(s, [{ text: "Microsoft's own warning: ", options: { bold: true } }, { text: "\"AI assistance does not expand data access, but it may transmit accessed data to the LLM provider.\"", options: { italic: true } }], 0.9, 5.8, 11.6, 0.95, { fontSize: 15, valign: "middle" });
}
{
  const s = slide("DARK", "NEXT · SESSION 02 · THE MODEL AS FILES", "Next: the model as files",
    "CLOSE (about 2 minutes)\n\nSay: Three things to remember from today. One: MCP is the standard plug that lets Claude reach a Power BI model, through a server running on your laptop. Two: we start read-only, and we read every approval. Three: we check every answer against Desktop.\n\nSay: Next session, we save Meridian as a PBIP project, open the folder in VS Code, and ask Claude to rename a column across every measure that uses it. We'll read the change as a Git diff, then undo it with one command.\n\nSay: Before then, two optional things. Ask Claude three more questions about Meridian and check them. And if you're curious, explore the interactive map: it shows exactly what's inside a PBIP project.\n\nDo: Share the map link in the chat. Leave this slide up for questions.");
  s.addText("Save as PBIP, open it in VS Code, let Claude change it, read the Git diff, undo it", { placeholder: "sub" });
  const take = [["MCP is the plug", "Claude reaches the model through a server on your laptop"], ["Read-only first", "and read every approval before you click"], ["Check every answer", "against Desktop, before you believe it"]];
  take.forEach(([h, d], i) => {
    const x = 0.8 + i * 3.98;
    card(s, x, 4.45, 3.7, 1.35, { fill: C.background1, t: 91, line: HEX.lineDark, name: h });
    txt(s, h, x + 0.25, 4.62, 3.3, 0.4, { fontSize: 18, bold: true, color: C.background1 });
    txt(s, d, x + 0.25, 5.05, 3.3, 0.65, { fontSize: 14, color: C.background2, transparency: 25 });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 6.2, w: 3.0, h: 0.5, rectRadius: 0.25, fill: { color: C.background1, transparency: 88 }, line: { color: C.accent3, width: 1 } });
  s.addText([{ text: "Explore the interactive map  ↗", options: { color: "F2B544", hyperlink: { url: `${MAP}#overview`, tooltip: "Opens the interactive map" } } }], { x: 0.8, y: 6.2, w: 3.0, h: 0.5, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  txt(s, "Optional before Session 2: ask Claude three more questions about Meridian, and check them", 4.0, 6.32, 8.7, 0.3, { fontSize: 13, color: C.background2, transparency: 35 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  if (!process.env.SKIP_THEME) await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
