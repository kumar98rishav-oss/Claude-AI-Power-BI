// Session 02 deck: Power BI service, paginated reports and Skills (90 minutes). Same theme and helpers as Session 01.
// Rebuild:  NODE_PATH=<folder with pptxgenjs> node build_deck.js
// Each "Open interactive map" pill links to a view of the published map artifact (MAP + "#view").
const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const MAP = "https://rishav-portfolio-i46g.onrender.com/training/power-bi-internals-map.html";
const GUIDE = "https://rishav-portfolio-i46g.onrender.com/training/session-02-guide.html";
const OUT = process.env.OUT || path.join(__dirname, "Session02_PowerBI_Service_Paginated_Skills.pptx");
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
pres.title = "Session 02 - Claude Skills for the Power BI Service";
const C = pres.SchemeColor;
const MONO = "Consolas";
const W = 13.333;

/* ---------------- layouts ---------------- */
const footer = (dark) => ({ text: { text: "Session 02 · Claude Skills for the Power BI Service", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 9, color: dark ? C.background2 : C.accent6, transparency: dark ? 45 : 0, margin: 0 } } });
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
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.5, y: 0.5, w: 2.23, h: 0.44, rectRadius: 0.22, fill: { color: dark ? C.background1 : C.text1, transparency: dark ? 88 : 0 }, line: { color: dark ? C.accent3 : C.text1, width: 1 }, objectName: "Guide link" });
  s.addText([{ text: "Open guide  ↗", options: { color: dark ? "F2B544" : C.background1, hyperlink: { url: `${GUIDE}#${view}`, tooltip: "Opens this step in the Session 02 guide" } } }], { x: 10.5, y: 0.5, w: 2.23, h: 0.44, fontSize: 12, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "Guide link text" });
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
   SESSION 02 · POWER BI SERVICE, PAGINATED REPORTS, SKILLS (90 min)
   ========================================================= */
const H = t => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } });
const TBL = { fontSize: 13, color: "14171D", valign: "middle", border: { type: "solid", pt: 0.75, color: HEX.line }, margin: [4, 10, 4, 10] };
function stepsLight(s, steps, x, y, w, gap = 0.98) {
  steps.forEach((st, i) => {
    const yy = y + i * gap;
    num(s, x, yy, i + 1, C.accent2, 0.46);
    txt(s, st[0], x + 0.68, yy - 0.02, w - 0.68, 0.4, { fontSize: 17, bold: true });
    if (st[1]) txt(s, st[1], x + 0.68, yy + 0.38, w - 0.68, 0.55, { fontSize: 13.5, color: C.accent6 });
  });
}
function box(s, x, y, w, h, title, sub, col, o = {}) {
  card(s, x, y, w, h, { fill: o.fill || C.background1, line: col, lw: 1.5, name: title });
  txt(s, title, x + 0.15, y + 0.12, w - 0.3, 0.4, { fontSize: o.ts || 16, bold: true, color: col, align: "center" });
  if (sub) txt(s, sub, x + 0.15, y + 0.55, w - 0.3, h - 0.65, { fontSize: o.ss || 12.5, color: C.text1, align: "center" });
}

section("Opening");
{
  const s = slide("DARK", "SESSION 02 · CLAUDE + POWER BI PROGRAMME", "Claude Skills for Power BI",
    "OPENING (about 2 minutes)\n\nScreen: this slide. Do: wait for people to join.\n\nSay: Hi everyone, welcome back! Last time Claude read the Meridian report on our own laptops. Today the report goes online, into the Power BI service, and the star of the session is Skills: saved procedures that check, print and act for us, and that can run on a schedule while we do something else.\n\nSay: Today is very hands-on. I'll share one of your screens at a time and guide you, and everyone else follows along on their own machine.\n\nDo: open the Session 02 guide link (bottom of this slide) in a browser tab next to you. Every command in today's deck is in that guide with a copy button.");
  s.addText("Connect Claude to the published model, then turn every routine check into a Skill that runs itself", { placeholder: "sub" });
  const parts = [["Connect", "Claude on the published model"], ["Check", "the service-check Skill"], ["Print", "the paginated report Skill"], ["Automate", "Skills on a schedule"]];
  parts.forEach(([h, d], i) => {
    const x = 0.8 + i * 2.98;
    card(s, x, 4.5, 2.75, 1.25, { fill: C.background1, t: 91, line: HEX.lineDark, name: h });
    num(s, x + 0.22, 4.7, i + 1, C.accent3, 0.42);
    txt(s, h, x + 0.8, 4.7, 1.9, 0.42, { fontSize: 18, bold: true, color: C.background1, valign: "middle" });
    txt(s, d, x + 0.22, 5.22, 2.4, 0.45, { fontSize: 13, color: C.background2, transparency: 25 });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 6.2, w: 3.2, h: 0.5, rectRadius: 0.25, fill: { color: C.background1, transparency: 88 }, line: { color: C.accent3, width: 1 } });
  s.addText([{ text: "Open the Session 02 guide  ↗", options: { color: "F2B544", hyperlink: { url: `${GUIDE}#before`, tooltip: "Opens the step-by-step guide" } } }], { x: 0.8, y: 6.2, w: 3.2, h: 0.5, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  txt(s, "90 minutes · read-only first · sample data only", 4.2, 6.32, 8, 0.3, { fontSize: 13, color: C.background2, transparency: 35 });
}
{
  const s = slide("CONTENT", "RECAP · WHERE WE ARE", "From your laptop to the Service",
    "RECAP (about 3 minutes)\n\nSay (point left): Last session, Claude reached the model on your laptop through the Power BI MCP server. We kept it read-only, read every approval, and checked every answer against Desktop. And we built our first Skill, report-snapshot.\n\nSay (point middle): Today the same model goes to the Power BI service, and we spend most of the session on Skills: what they can do with the Service, how to build one, and how to make them run by themselves.\n\nSay (point right): And by the end, our checks become Skills that can run without us at the keyboard.\n\nAsk: 'Did anyone try report-snapshot or ask Claude more questions since last time? What happened?' Take one answer.");
  pill(s, "p0");
  const cols = [
    ["SESSION 01 · DONE", C.accent6, ["Claude Code + MCP server installed", "Connected to Desktop, read-only", "Checked answers against Desktop", "First Skill: report-snapshot"]],
    ["SESSION 02 · TODAY", C.accent3, ["Publish Meridian to a workspace", "Claude reads the published model", "What Skills can do for Power BI", "Three Skills: check, print, act", "Skills that run on a schedule"]],
    ["NEXT", C.accent2, ["The model as files: PBIP + TMDL", "Claude edits files, Git shows the diff", "Later: SQL Server + SQL MCP"]],
  ];
  cols.forEach(([h, col, items], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.9, 3.9, 4.6, { fill: i === 1 ? C.accent3 : C.background2, t: i === 1 ? 90 : 0, line: i === 1 ? C.accent3 : undefined, name: h });
    txt(s, h, x + 0.3, 2.1, 3.4, 0.35, { fontSize: 12, bold: true, color: col, charSpacing: 1.5 });
    txt(s, bullets(items, { gap: 10 }), x + 0.3, 2.6, 3.4, 3.7, { fontSize: 15.5 });
  });
  arrow(s, 4.52, 4.2, 4.68, 4.2, C.accent6, 2);
  arrow(s, 8.62, 4.2, 8.78, 4.2, C.accent6, 2);
}
{
  const s = slide("CONTENT", "WHY SKILLS", "Stop repeating yourself: save it as a Skill",
    "WHY SKILLS (about 3 minutes)\n\nSay: Think about your week. How many times do you open a report to check it refreshed, look at the same KPIs, compare a number with last week, or tidy up a report before it goes out? Same steps, every time, by hand.\n\nSay (point left): Without a Skill, you retype the same prompts, everyone checks a bit differently, and steps get forgotten. And it only happens when you remember.\n\nSay (point right): With a Skill, it's one command. Same steps and the same output every time. It's shared with the team through Git. The safety rules are written into it, like read-only. And because it's written down, it can run on a schedule without you.\n\nSay (point at the bottom): A Skill is just a folder with one file, SKILL.md: when to use it, the rules, the steps and what the output looks like. Today we use three of them and build one ourselves.\n\nAsk: 'What's one check you do by hand every week on a Power BI report?' Take two answers; they're your homework Skills.");
  pill(s, "p5");
  const cols = [["WITHOUT A SKILL", C.accent5, ["Retype the same prompts every time", "Everyone checks a bit differently", "Steps get forgotten", "Only happens when you remember"]],
    ["WITH A SKILL", C.accent4, ["One command: /service-check", "Same steps, same output, every time", "Shared with the team through Git", "Rules built in, like read-only", "Can run on a schedule, without you"]]];
  cols.forEach(([h, col, items], i) => {
    const x = 0.6 + i * 6.14;
    card(s, x, 1.9, 5.99, 3.6, { line: col, lw: 1.5, fill: C.background1, name: h });
    txt(s, h, x + 0.3, 2.08, 5.4, 0.3, { fontSize: 13, bold: true, color: col, charSpacing: 1.5 });
    txt(s, bullets(items, { gap: 10 }), x + 0.3, 2.55, 5.4, 2.8, { fontSize: 17 });
  });
  arrow(s, 6.62, 3.7, 6.72, 3.7, C.accent6, 2);
  card(s, 0.6, 5.8, 12.13, 0.85, { fill: C.accent2, t: 90, name: "Definition" });
  txt(s, [{ text: "A Skill = a folder with SKILL.md: ", options: { bold: true } }, { text: "when to use it, the rules, the steps and the output. Claude follows it the same way every time.", options: {} }], 0.9, 5.8, 11.6, 0.85, { fontSize: 16, valign: "middle" });
}
{
  const s = slide("CONTENT", "WHAT SKILLS CAN DO WITH THE POWER BI SERVICE", "Eight jobs a Skill can take off your plate",
    "WHAT SKILLS CAN DO (about 4 minutes)\n\nSay: Here's what's realistic today with Claude, the Power BI MCP server and a little PowerShell. Every one of these works on the published model in the Service. Go through them quickly: one sentence each.\n\nSay: One, a morning health check: last refresh, failures, key KPIs. Two, a refresh watcher: the refresh history, how long each took, and the error text when one fails. Three, KPI alerts: thresholds written into the Skill, so revenue or on-time delivery crossing a line gets flagged. Four, drift: compare the published model with your Desktop copy and list what's different.\n\nSay: Five, an RLS test: run the KPIs as a security role, like our West Region role, to see what those users would see. Six, documentation: a measure dictionary with DAX, descriptions and folders, written to a Markdown file. Seven, the paginated report editor we'll use today. Eight, refresh on demand: start a refresh, wait, and report back.\n\nSay (point at the tags): The colour tells you the kind. Green only reads, so it's safe to schedule. Amber edits a file, after a backup. Red acts on the Service, so it stays manual.\n\nIf asked about RLS: the DAX query tool can run a query as a role (impersonation). You need enough permission on the model for that.");
  pill(s, "p5");
  const jobs = [["Morning health check", "Last refresh, failed refreshes, key KPIs, one Flags list", "R", "MCP + REST API"],
    ["Refresh watcher", "Refresh history, duration, error text of failures", "R", "REST API"],
    ["KPI alerts", "Revenue, on-time %, stockouts against your thresholds", "R", "DAX queries"],
    ["Desktop vs published drift", "Measures, relationships and DAX that differ", "R", "MCP"],
    ["RLS test", "Run the KPIs as a role, e.g. West Region", "R", "DAX as a role"],
    ["Model documentation", "Measure dictionary with DAX and folders, to Markdown", "E", "MCP + a file"],
    ["Paginated report editor", "Format and fix .rdl files with schema rules", "E", ".rdl file"],
    ["Refresh on demand", "Start a refresh, wait, report the result", "A", "REST API · manual"]];
  const tag = { R: ["READS", C.accent4], E: ["EDITS A FILE", C.accent3], A: ["ACTS · MANUAL", C.accent5] };
  jobs.forEach(([h, d, k, uses], i) => {
    const x = 0.6 + (i % 4) * 3.07, y = 1.9 + Math.floor(i / 4) * 2.2;
    card(s, x, y, 2.9, 2.05, { name: h });
    txt(s, tag[k][0], x + 0.22, y + 0.17, 2.5, 0.25, { fontSize: 10.5, bold: true, color: tag[k][1], charSpacing: 1.2 });
    txt(s, h, x + 0.22, y + 0.45, 2.5, 0.5, { fontSize: 15.5, bold: true });
    txt(s, d, x + 0.22, y + 0.95, 2.5, 0.7, { fontSize: 12.5, color: C.accent6 });
    txt(s, uses, x + 0.22, y + 1.67, 2.5, 0.28, { fontSize: 11, bold: true, color: C.accent2 });
  });
  txt(s, [{ text: "Reads ", options: { bold: true, color: C.accent4 } }, { text: "can run on a schedule · ", options: {} }, { text: "Edits ", options: { bold: true, color: C.accent3 } }, { text: "make a backup first · ", options: {} }, { text: "Acts ", options: { bold: true, color: C.accent5 } }, { text: "stay manual", options: {} }], 0.6, 6.35, 12.13, 0.35, { fontSize: 14 });
}
{
  const s = slide("CONTENT", "TODAY'S JOURNEY", "From a report to Skills that run themselves",
    "THE JOURNEY (about 2 minutes)\n\nSay: Here's today on one line. Left to right.\n\nSay: We publish from Desktop to the Service. Claude connects to the published model through the MCP server. Then we wrap our checks into Skills, and the Skills run on a schedule. At the end of the line is you: you only look at the flags and decide what to do.\n\nSay (point at the bottom): That's the big idea of today. Power BI does the refreshing. Skills do the checking. You do the deciding.");
  pill(s, "p0");
  const nodes = [["Power BI Desktop", "your laptop", C.accent6], ["Power BI service", "the published model", C.accent1], ["Claude (MCP)", "reads it, read-only", C.accent2], ["Skills", "check · print · act", C.accent4], ["Schedule", "runs without you", C.accent3], ["You", "read the flags, decide", C.text1]];
  nodes.forEach(([h, d, col], i) => {
    const x = 0.6 + i * 2.1;
    box(s, x, 2.3, 1.8, 1.45, h, d, col, { ts: 14.5, ss: 12 });
    if (i < nodes.length - 1) arrow(s, x + 1.82, 3.02, x + 2.08, 3.02, C.accent6, 1.75);
  });
  const parts = [["Part 1", "Connect", 2.7], ["Parts 2–4", "Skills", 6.9], ["Part 5", "Automate", 9.0]];
  parts.forEach(([p, d, x]) => txt(s, [{ text: p + "  ", options: { bold: true, color: C.accent3 } }, { text: d, options: { color: C.accent6 } }], x, 3.95, 2.0, 0.4, { fontSize: 12.5, align: "center" }));
  const roles = [["Power BI", "refreshes", C.accent1], ["Skills", "check", C.accent4], ["You", "decide", C.text1]];
  roles.forEach(([h, d, col], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 4.75, 3.9, 1.75, { name: h });
    txt(s, h, x + 0.3, 4.95, 3.3, 0.5, { fontSize: 24, bold: true, color: col });
    txt(s, d, x + 0.3, 5.55, 3.3, 0.6, { fontSize: 20 });
  });
}
{
  const s = slide("CONTENT", "BEFORE WE START", "Two tracks, depending on your licence",
    "TRACKS (about 3 minutes)\n\nAsk: 'Who has a workspace in the Power BI service that isn't My workspace? Who started the Fabric trial?' Note the names.\n\nSay: Claude reaches a published model through the XMLA endpoint. Microsoft supports that only when the workspace is on Premium, Premium Per User or Fabric capacity. The free 60-day Fabric trial gives PPU-level rights, so it counts.\n\nSay: If you're on a free account with only My workspace, you're Track B. You'll still publish, build the paginated report and create Skills. For Claude reading the published model, you watch my screen.\n\nSay: Either way, check you have Report Builder installed and the Power BI PowerShell module. The install lines are at the bottom of the slide.");
  pill(s, "before");
  const tr = [["TRACK A", "Capacity workspace", C.accent1, ["Premium, PPU or Fabric capacity (Fabric trial counts)", "Does every part hands-on", "Needs Write or Build permission on the model"]],
    ["TRACK B", "Free account", C.accent6, ["Only My workspace", "Publishes, builds the paginated report, creates Skills", "Watches the Service connection on the trainer's screen"]]];
  tr.forEach(([e, h, col, items], i) => {
    const x = 0.6 + i * 6.14;
    card(s, x, 1.9, 5.99, 3.05, { line: col, lw: 1.5, name: h, fill: C.background1 });
    txt(s, e, x + 0.3, 2.08, 5.4, 0.3, { fontSize: 12, bold: true, color: col, charSpacing: 1.5 });
    txt(s, h, x + 0.3, 2.4, 5.4, 0.5, { fontSize: 22, bold: true });
    txt(s, bullets(items, { gap: 8 }), x + 0.3, 3.0, 5.4, 1.8, { fontSize: 15 });
  });
  txt(s, "EVERYONE INSTALLS (ONE TIME)", 0.6, 5.18, 6, 0.3, { fontSize: 11, bold: true, color: C.accent6, charSpacing: 1.5 });
  code(s, 0.6, 5.5, 12.13, 1.2, [
    [["winget install --id Microsoft.PowerBIReportBuilder -e", "x"], ["      # Power BI Report Builder (free)", "c"]],
    [["Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser", "x"], ["      # Power BI PowerShell module", "c"]],
  ], { size: 14, lang: "text" });
}

section("Part 1 · Connect");
{
  const s = slide("CONTENT", "PART 1 · DESKTOP VS SERVICE", "What changes when the model goes online",
    "DESKTOP VS SERVICE (about 3 minutes)\n\nSay: Before we click Publish, let's be clear what changes. Read the table row by row.\n\nSay: Where it lives: on your laptop it runs inside Power BI Desktop; in the Service it lives in a workspace that other people can open.\n\nSay: Who refreshes: on your laptop, you click Refresh. In the Service, a scheduled refresh does it.\n\nSay: How Claude connects: last week, through a local port on your laptop. Today, through the XMLA endpoint, signed in with your Microsoft account.\n\nSay: And the safety rule changes too. In Desktop, a mistake is easy to undo. In the Service, it's live for everyone, and a write over XMLA stops you downloading the .pbix again. So today: read only.");
  pill(s, "p1");
  const rows = [[H(""), H("Power BI Desktop"), H("Power BI service")],
    [{ text: "Where the model lives", options: { bold: true } }, "Inside Desktop, on your laptop", "In a workspace, shared with others"],
    [{ text: "Who refreshes it", options: { bold: true } }, "You click Refresh", "A scheduled refresh in the Service"],
    [{ text: "How Claude connects", options: { bold: true } }, "MCP server → local port (localhost)", "MCP server → XMLA endpoint, Microsoft sign-in"],
    [{ text: "What it needs", options: { bold: true } }, "Nothing extra", "Premium, PPU or Fabric capacity workspace"],
    [{ text: "If something goes wrong", options: { bold: true } }, "Only you see it; easy to undo", { text: "Live for everyone; a write blocks .pbix download", options: { color: C.accent5, bold: true } }],
  ];
  s.addTable(rows, Object.assign({ x: 0.6, y: 1.9, w: 12.13, colW: [3.0, 4.4, 4.73], rowH: 0.62, fontSize: 14.5 }, TBL, { fontSize: 14.5 }));
}
{
  const s = slide("LIVE", "PART 1 · EVERYONE DOES THIS", "Publish Meridian and find its IDs",
    "PUBLISH (about 7 minutes)\n\nDo: share one participant's screen.\n\nCoach: 'Click Publish on the Home ribbon. Save first if it asks. Pick your workspace: Track A your capacity workspace, Track B My workspace.'\n\nCoach: 'When it says Success, click Open in Power BI. Check the numbers match Desktop: 2025 revenue 14.1 million.'\n\nCoach: 'Now go to the workspace list and open the item with the semantic model icon, not the report icon. Look at the address bar.'\n\nSay (point at the code box): Three IDs live in these addresses. The code after groups is the workspace. After reports is the report. And the semantic model shows after datasets, or after modeling if you open the web modelling view. Our Skills need the workspace ID and the semantic model ID, not the report ID. That's a common mix-up.\n\nCheck: everyone has the workspace ID and semantic model ID in Notepad.\n\nIf it fails: 'You don't have permission to publish' means they need Contributor in that workspace, or use My workspace. Publish greyed out: sign in to Desktop first.");
  pill(s, "p1", true);
  liveSteps(s, [["Home › Publish", "Save first if asked, then pick the workspace"], ["Open it in the Service", "Numbers must match Desktop: 2025 revenue 14.1M"], ["Open the semantic model item", "The model icon, not the report icon"], ["Copy two IDs into Notepad", "Workspace ID and semantic model ID"]], 0.6, 1.95, 5.6);
  code(s, 6.55, 1.95, 6.18, 3.0, [
    [["app.powerbi.com/groups/", "x"], ["<workspace-id>", "h"]],
    [["  /reports/", "x"], ["<report-id>", "s"], ["/ExecutiveOverview", "x"]],
    [[" ", "x"]],
    [["app.powerbi.com/groups/", "x"], ["<workspace-id>", "h"]],
    [["  /datasets/", "x"], ["<semantic-model-id>", "a"], ["/details", "x"]],
    [["  or /modeling/", "x"], ["<semantic-model-id>", "a"], ["/modelView", "x"]],
  ], { dark: true, size: 13.5, lang: "text" });
  txt(s, [{ text: "Skills need ", options: {} }, { text: "workspace ID", options: { bold: true, color: "F2B544" } }, { text: " + ", options: {} }, { text: "semantic model ID", options: { bold: true, color: "63C474" } }, { text: ". The report ID is not the model ID.", options: {} }], 6.55, 5.1, 6.18, 0.7, { fontSize: 14, color: C.background2 });
}

{
  const s = slide("CONTENT", "PART 1 · HOW IT CONNECTS", "How Claude reaches a published model",
    "HOW IT CONNECTS (about 3 minutes)\n\nSay: Same MCP server as last week, same read-only flag. What changes is the address. Instead of a local port on your laptop, the server talks to the Service over the internet, through the XMLA endpoint, and you sign in with your Microsoft account.\n\nSay (point at the requirement cards): Three conditions. The workspace must be on Premium, PPU or Fabric capacity: that's the Track A question. You need Write or Build permission on the model. And you sign in with the same account that owns the workspace.\n\nSay: Underneath, it's the same two libraries from Session 1: TOM reads the structure, ADOMD.NET runs the DAX queries. They just connect to the Service instead of Desktop.\n\nIf asked: the MCP server's README says it 'connects to semantic models through the XMLA endpoint'. Microsoft's docs list Premium, PPU and Embedded workspaces as the ones with an XMLA endpoint.");
  pill(s, "p2");
  const n = [["You", "ask in plain English", C.accent6], ["Claude Code", "in VS Code", C.accent3], ["MCP server", "powerbi-modeling, read-only", C.accent2], ["XMLA endpoint", "Microsoft sign-in", C.accent1], ["Semantic model", "in your workspace", C.text1]];
  n.forEach(([h, d, col], i) => {
    const x = 0.6 + i * 2.52;
    box(s, x, 1.95, 2.2, 1.35, h, d, col, { ts: 15, ss: 12 });
    if (i < n.length - 1) arrow(s, x + 2.22, 2.62, x + 2.5, 2.62, C.accent6, 1.75);
  });
  const req = [["Capacity workspace", "Premium, PPU or Fabric (trial counts). Not a Free My workspace."], ["Permission", "Write or Build on the semantic model"], ["Sign-in", "Browser pop-up, same account that owns the workspace"]];
  req.forEach(([h, d], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 3.75, 3.9, 1.5, { name: h });
    num(s, x + 0.25, 3.95, i + 1, C.accent1, 0.42);
    txt(s, h, x + 0.85, 3.95, 2.9, 0.42, { fontSize: 16, bold: true, valign: "middle" });
    txt(s, d, x + 0.25, 4.5, 3.45, 0.7, { fontSize: 13.5, color: C.accent6 });
  });
  card(s, 0.6, 5.55, 12.13, 0.95, { fill: C.accent2, t: 90, name: "Same libraries" });
  txt(s, [{ text: "Same engine as Session 1: ", options: { bold: true } }, { text: "TOM reads the structure, ADOMD.NET runs the DAX. Only the address changes, from localhost to the Service.", options: {} }], 0.9, 5.55, 11.6, 0.95, { fontSize: 15, valign: "middle" });
}
{
  const s = slide("LIVE", "PART 1 · TRACK A HANDS-ON · TRACK B WATCHES", "Connect, then compare with Desktop",
    "CONNECT AND COMPARE (about 7 minutes)\n\nDo: Track A participant shares their screen. Track B watch yours or theirs.\n\nCoach: 'Start a new Claude session in your project folder: type claude. Then type /mcp. You should see powerbi-modeling and nothing else that says Power BI.' If they see a plugin called powerbi-authoring, stop: that one can write. Disable it with: claude plugin disable powerbi-authoring@fabric-collection, then start a new session.\n\nCoach: 'Now type the connect sentence. A browser asks you to sign in: use the account that owns the workspace. Read each approval box out loud before Yes.'\n\nCoach: 'Paste the compare prompt.' Expected: 10 tables, 30 measures, 15 relationships on both sides, and the two refresh times.\n\nSay: This is the first real use of the Service connection: catching drift between what's on your laptop and what users actually see.\n\nIf it fails: an XMLA or 'not supported' error means the workspace isn't on capacity, so switch them to Track B. A sign-in loop: wrong account.");
  pill(s, "p2", true);
  liveSteps(s, [["New session, check /mcp", "Only powerbi-modeling (read-only) should appear"], ["Connect to the published model", "Sign in when the browser opens; read every approval"], ["Compare with Desktop", "Tables, measures, relationships, DAX, refresh time"]], 0.6, 1.95, 5.4);
  code(s, 6.3, 1.95, 6.43, 1.05, [[["Connect to semantic model 'MeridianSupplyChain'", "x"]], [["in Fabric Workspace 'Data Analysis_dev'", "x"]]], { dark: true, size: 13.5, lang: "text" });
  code(s, 6.3, 3.2, 6.43, 2.2, [
    [["Compare the published MeridianSupplyChain model", "x"]],
    [["with the copy open in Power BI Desktop:", "x"]],
    [["number of tables, measures and relationships,", "x"]],
    [["any measure whose DAX differs, and the last", "x"]],
    [["refresh time of each. Read only, change nothing.", "s"]],
  ], { dark: true, size: 13.5, lang: "text" });
  txt(s, [{ text: "Expect: ", options: { bold: true, color: "F2B544" } }, { text: "10 tables · 30 measures · 15 relationships on both sides", options: {} }], 6.3, 5.6, 6.43, 0.5, { fontSize: 14, color: C.background2 });
}
{
  const s = slide("CONTENT", "PART 1 · THE ONE RULE FOR THE SERVICE", "Read the published model. Never write to it.",
    "THE RULE (about 2 minutes)\n\nSay: When Claude connected to my workspace, it listed everything it could do: read, query, and also create measures, change relationships, edit roles. Don't let that tempt you.\n\nSay (point at the quote): This is Microsoft's own wording. If you change a Desktop-authored model through the XMLA endpoint, you can never download it back as a .pbix. Your editable copy is gone from the Service.\n\nSay (point at the right): So the workflow is: change the Desktop or PBIP copy, then republish. The Service gets read-only questions: compare, query, check refreshes.\n\nSay: And Claude saying it 'can' change things is a hint that you're on a read-write server. Check /mcp.");
  pill(s, "p2");
  card(s, 0.6, 1.95, 12.13, 1.25, { fill: C.accent5, t: 90, line: C.accent5, name: "Quote" });
  txt(s, [{ text: "Microsoft: ", options: { bold: true, color: C.accent5 } }, { text: "\"a write operation on a semantic model authored in Power BI Desktop prevents it from being downloaded back as a PBIX file.\"", options: { italic: true } }], 0.9, 1.95, 11.6, 1.25, { fontSize: 17, valign: "middle" });
  const cols = [["DO IN THE SERVICE", C.accent4, ["Compare published vs Desktop", "Run DAX queries: KPIs, top/bottom lists", "Read refresh times and history", "Read measures, relationships, roles"]],
    ["DON'T DO IN THE SERVICE", C.accent5, ["Create, edit or delete measures", "Change relationships or roles", "Let Claude 'sync' or overwrite anything", "Test a fix on the live model"]],
    ["INSTEAD", C.accent2, ["Change Desktop or the PBIP files", "Check it in Desktop", "Republish", "Then compare again"]]];
  cols.forEach(([h, col, items], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 3.5, 3.9, 3.0, { name: h });
    txt(s, h, x + 0.3, 3.7, 3.4, 0.3, { fontSize: 12, bold: true, color: col, charSpacing: 1.2 });
    txt(s, bullets(items, { gap: 8, num: i === 2 }), x + 0.3, 4.15, 3.4, 2.2, { fontSize: 15 });
  });
}

section("Part 2 · Your first Skill");
{
  const s = slide("CONTENT", "PART 2 · ANATOMY OF A SKILL", "What's inside a Skill",
    "ANATOMY (about 3 minutes)\n\nSay: Let's open one up. This is service-check, the Skill we'll run in a minute. It's a folder with SKILL.md and a scripts folder.\n\nSay (point at each number): One, the name and description. The description is how Claude knows when to use it: if you ask 'is the published report OK?', Claude matches that to this description. Two, settings: your workspace and semantic model IDs, from Part 1. Three, rules: read only, no row-level data. That's our safety built into the Skill. Four, the steps: connect, a DAX query for the last refresh, a PowerShell script for the refresh history, the KPI query. Five, the output: the same headings every time, ending with a Flags list.\n\nSay (point at the folder): Saved in .claude/skills inside the project, it travels with the project through Git, so your whole team gets it.\n\nSay: Nothing here is code you need to write by hand. Next part, skill-creator writes this for us.");
  pill(s, "files");
  code(s, 0.6, 1.9, 7.3, 4.75, [
    [["---", "c"]],
    [["name", "p"], [": service-check", "x"]],
    [["description", "p"], [": Read-only health check of the", "x"]],
    [["  published Meridian model. Use when asked", "x"]],
    [["  \"is the published report OK\" ...", "x"]],
    [["---", "c"]],
    [["## Settings", "k"]],
    [["- Workspace ID / Semantic model ID", "x"]],
    [["## Rules", "k"]],
    [["- Read only. Never change anything.", "h"]],
    [["- Totals and lists only, never row-level data.", "x"]],
    [["## Steps", "k"]],
    [["1. Connect  2. Last refresh (DAX)", "x"]],
    [["3. Refresh history (scripts/service_check.ps1)", "x"]],
    [["4. KPIs with thresholds", "x"]],
    [["## Output", "k"]],
    [["Connection · Refresh · KPIs · Flags", "x"]],
  ], { size: 13 });
  const call = [["Name + description", "When Claude should use it"], ["Settings", "Your workspace + model IDs"], ["Rules", "Read-only, no row data: safety built in"], ["Steps", "MCP, DAX and a small script"], ["Output", "Same headings every time + Flags"]];
  call.forEach(([h, d], i) => {
    const y = 1.95 + i * 0.78;
    num(s, 8.25, y, i + 1, C.accent2, 0.42);
    txt(s, h, 8.8, y - 0.03, 3.9, 0.35, { fontSize: 15.5, bold: true });
    txt(s, d, 8.8, y + 0.3, 3.9, 0.35, { fontSize: 13, color: C.accent6 });
  });
  code(s, 8.25, 5.9, 4.48, 0.75, [[[".claude/skills/service-check/", "x"]], [["  SKILL.md  ·  scripts/service_check.ps1", "c"]]], { size: 11.5, lang: "text" });
}
{
  const s = slide("CONTENT", "PART 2 · THREE KINDS OF SKILL", "Read, edit, act: three kinds of Skill",
    "THREE SKILLS (about 3 minutes)\n\nSay: Last session we had one Skill. Today three, and each shows a different kind.\n\nSay (point at each row): service-check only reads: last refresh, refresh history, KPIs. Safe to run any time, even on a schedule.\n\nSay: paginated-report edits a file, but only the .rdl, and it makes a backup first. It carries the rules we learned, like the Border element.\n\nSay: refresh-now changes the Service: it starts a refresh. So it has one extra line in its frontmatter: disable-model-invocation true. That means Claude never runs it on its own. You must type /refresh-now, and even then it asks for a yes first.\n\nSay: The rule of thumb: the more a Skill can change, the more control you keep.");
  pill(s, "p5");
  const rows = [[H("Skill"), H("What it does"), H("Changes anything?"), H("Claude may start it itself?")],
    [{ text: "service-check", options: { fontFace: MONO, bold: true } }, "Last refresh, refresh history, KPIs, Flags", { text: "No", options: { color: C.accent4, bold: true } }, "Yes"],
    [{ text: "paginated-report", options: { fontFace: MONO, bold: true } }, "Edits .rdl layout with RDL rules", "Only the .rdl, after a backup", "Yes"],
    [{ text: "refresh-now", options: { fontFace: MONO, bold: true } }, "Starts a refresh, reports the result", { text: "Yes: the Service", options: { color: C.accent5, bold: true } }, { text: "No: only /refresh-now", options: { bold: true } }]];
  s.addTable(rows, Object.assign({}, TBL, { x: 0.6, y: 1.9, w: 12.13, colW: [2.6, 4.2, 2.6, 2.73], rowH: 0.62, fontSize: 14 }));
  txt(s, "THE LINE THAT KEEPS refresh-now MANUAL", 0.6, 4.6, 6, 0.3, { fontSize: 12, bold: true, color: C.accent6, charSpacing: 1.5 });
  code(s, 0.6, 4.95, 6.2, 1.55, [[["---", "c"]], [["name", "p"], [": refresh-now", "x"]], [["description", "p"], [": Start an on-demand refresh ...", "x"]], [["disable-model-invocation", "p"], [": true", "h"]], [["---", "c"]]], { size: 13.5 });
  card(s, 7.1, 4.95, 5.63, 1.55, { fill: C.accent3, t: 88, name: "Rule" });
  txt(s, [{ text: "Rule of thumb: ", options: { bold: true } }, { text: "the more a Skill can change, the more control you keep. Read-only Skills can run on a schedule; anything that acts stays manual.", options: {} }], 7.35, 4.95, 5.2, 1.55, { fontSize: 14.5, valign: "middle" });
}
{
  const s = slide("LIVE", "PART 2 · TRACK A HANDS-ON · TRACK B WATCHES", "Skill 1: run service-check",
    "RUN SERVICE-CHECK (about 7 minutes)\n\nDo: Track A participant shares their screen. Make sure the service-check folder is in their project's .claude/skills, with their workspace ID and semantic model ID in the Settings section.\n\nCoach: 'Start a new Claude session in the project folder and type /service-check. Approve each step after reading it: the connection, the DAX query, and the PowerShell script. The first time, a browser asks you to sign in to Power BI.'\n\nSay (point at the output): Four headings, always in the same order. Connection: which model. Refresh: the last five refreshes, with how long they took and the error text if one failed. KPIs: this year against last year, with our thresholds. And Flags: the only part you really need to read every morning.\n\nSay: When I first ran this, the refresh history was empty. Publishing doesn't count as a refresh. Click Refresh now on the semantic model in the Service, then run it again.\n\nExpected KPIs on Meridian: On-Time Delivery 89.8% (below our 90% line) and Stockouts 45 (above 40) are flagged.\n\nIf it fails: 'could not read refresh history' means they need Write permission on the model. A sign-in loop: wrong account. 'Running scripts is disabled': Set-ExecutionPolicy -Scope CurrentUser RemoteSigned.");
  pill(s, "p5", true);
  liveSteps(s, [["Add your IDs to Settings", "Workspace ID + semantic model ID from Part 1"], ["New session, type /service-check", "Or just ask: \"Is the published report OK?\""], ["Approve each read", "Connection · DAX query · PowerShell script"], ["Read the Flags", "The one part to read every morning"]], 0.6, 1.95, 5.4);
  code(s, 6.3, 1.95, 6.43, 0.6, [[["/service-check", "k"]]], { dark: true, size: 14, lang: "text" });
  code(s, 6.3, 2.75, 6.43, 3.45, [
    [["## Connection", "k"]],
    [["MeridianSupplyChain · Data Analysis_dev", "x"]],
    [["## Refresh", "k"]],
    [["OnDemand  Completed  1.2 min", "a"]],
    [["## KPIs (2025 vs 2024)", "k"]],
    [["Revenue 14.13M (+10.1%)", "x"]],
    [["On-Time Delivery 89.8%   FLAG < 90%", "s"]],
    [["Stockouts 45             FLAG > 40", "s"]],
    [["## Flags", "k"]],
    [["2 KPI thresholds broken", "s"]],
  ], { dark: true, size: 12.5, lang: "text" });
  txt(s, "Example output · your numbers and times will differ", 6.3, 6.3, 6.43, 0.3, { fontSize: 12, color: C.background2, transparency: 35 });
}

section("Part 3 · Paginated report Skill");
{
  const s = slide("CONTENT", "PART 3 · PAGINATED REPORTS", "When you need it on paper",
    "PAGINATED VS POWER BI (about 3 minutes)\n\nSay: A Power BI report is for exploring: click, filter, drill. A paginated report is for printing and exporting: invoices, statements, a supplier scorecard that goes out as a PDF every month. Every row, laid out exactly, page after page.\n\nSay (point at the table): It's built in a different tool, Power BI Report Builder, which is free. And the file is different: an .rdl file, which is XML text. That's the important part for us: because it's text, Claude can read and edit it, the same way it edits TMDL.\n\nSay: Licences: you can publish a paginated report to My workspace without Pro. Other workspaces need Pro or PPU. Viewers need Pro or PPU unless the workspace is on F64 or P1 or bigger.");
  pill(s, "p4");
  const rows = [[H(""), H("Power BI report"), H("Paginated report")],
    [{ text: "Made for", options: { bold: true } }, "Exploring: click, filter, drill", "Printing and exporting: every row, exact layout"],
    [{ text: "Built in", options: { bold: true } }, "Power BI Desktop", "Power BI Report Builder (free)"],
    [{ text: "File", options: { bold: true } }, ".pbix / PBIR folder", { text: ".rdl: XML text Claude can edit", options: { bold: true, color: C.accent2 } }],
    [{ text: "Data", options: { bold: true } }, "The semantic model", "The same semantic model, with a DAX query"],
    [{ text: "Typical use", options: { bold: true } }, "Dashboards, analysis", "Scorecards, invoices, monthly PDFs, Excel exports"],
    [{ text: "Publish", options: { bold: true } }, "Any workspace you can edit", "My workspace: free · others: Pro or PPU"]];
  s.addTable(rows, Object.assign({}, TBL, { x: 0.6, y: 1.9, w: 12.13, colW: [2.4, 4.6, 5.13], rowH: 0.6, fontSize: 14.5 }));
}
{
  const s = slide("CONTENT", "PART 3 · WHO DOES WHAT", "Who does what in a paginated report",
    "WHO DOES WHAT (about 2 minutes)\n\nSay: This split is the most important idea in this part. Four steps, four owners.\n\nSay: One, you make the connection in Report Builder: Power BI Semantic Model Connection, then a dataset with the DAX query copied from Performance Analyzer. Claude doesn't touch connections or credentials.\n\nSay: Two, Claude edits the layout in the .rdl file: titles, borders, formats, page size, parameters.\n\nSay: Three, Report Builder is the judge. Claude can't see the report; it only edits text. So we always reopen the file and click Run.\n\nSay: Four, you publish and export from Report Builder or the Service.\n\nAsk: 'Why don't we let Claude write the connection too?' Answer: it would have to guess connection strings and IDs; Report Builder does it correctly in two clicks, and credentials never end up in a text file.");
  pill(s, "p4");
  const o = [["You", "Connect", "Report Builder: Power BI Semantic Model Connection + DAX dataset", C.accent6],
    ["Claude", "Edit layout", "Titles, borders, formats, page size, parameters in the .rdl", C.accent2],
    ["Report Builder", "Check", "Reopen and Run. It rejects anything invalid, with the line number", C.accent3],
    ["You", "Publish", "File › Publish, then export to PDF or Excel in the Service", C.accent1]];
  o.forEach(([who, h, d, col], i) => {
    const x = 0.6 + i * 3.1;
    card(s, x, 2.0, 2.85, 3.3, { line: col, lw: 1.5, fill: C.background1, name: h });
    num(s, x + 0.25, 2.2, i + 1, col, 0.46);
    txt(s, who.toUpperCase(), x + 0.85, 2.27, 1.9, 0.3, { fontSize: 11.5, bold: true, color: col, charSpacing: 1.2 });
    txt(s, h, x + 0.25, 2.85, 2.4, 0.5, { fontSize: 21, bold: true });
    txt(s, d, x + 0.25, 3.45, 2.4, 1.7, { fontSize: 14, color: C.accent6 });
    if (i < 3) arrow(s, x + 2.87, 3.65, x + 3.08, 3.65, C.accent6, 1.75);
  });
  card(s, 0.6, 5.65, 12.13, 0.85, { fill: C.accent3, t: 88, name: "Rule" });
  txt(s, [{ text: "Rule: ", options: { bold: true } }, { text: "Claude edits the file but can't see the report. Report Builder is always the last check.", options: {} }], 0.9, 5.65, 11.6, 0.85, { fontSize: 16, valign: "middle" });
}
{
  const s = slide("LIVE", "PART 3 · EVERYONE DOES THIS", "Build the Supplier Scorecard",
    "BUILD THE SCORECARD (about 12 minutes)\n\nCoach step 1: 'In Desktop: Optimize › Performance analyzer › Start recording › Refresh visuals. Find Supplier scorecard, expand it, click Copy query.'\n\nCoach step 2: 'Open Report Builder, sign in top right. Report Data › New › Power BI Semantic Model Connection › MeridianSupplyChain. Right-click it › Add Dataset › Query Designer: DAX selected, Design Mode off. Paste, delete any TOPN, Execute Query, OK. Insert › Table › Table Wizard, drag all fields in. Save as reports\\Supplier_Scorecard.rdl in the project folder.'\n\nSay: Don't style it. That's Claude's job.\n\nCoach step 3: 'In Claude, paste the paginated-report prompt.' The Skill makes a backup first.\n\nCoach step 4: 'Back in Report Builder, open the file again and click Run.' If it shows an error, copy the whole message into Claude.\n\nCoach step 5: 'File › Publish › your workspace.'\n\nCheck: bold dark header, borders, 0.0% on the on-time column, A4 landscape, opens in the Service.\n\nIf it fails: no Semantic Model Connection option means not signed in to Report Builder. The DAX errors: remove TOPN and unused __DS0 variables, or ask Claude to simplify the query.");
  pill(s, "p4", true);
  liveSteps(s, [["Copy the DAX", "Desktop › Performance analyzer › Supplier scorecard › Copy query"], ["Plain table in Report Builder", "Semantic Model Connection › DAX dataset › Table Wizard › save .rdl"], ["Claude styles it", "/paginated-report prompt (right) · backup first"], ["Reopen and Run", "Report Builder is the judge"], ["Publish", "File › Publish › your workspace"]], 0.6, 1.9, 5.6);
  code(s, 6.45, 1.95, 6.28, 3.35, [
    [["/paginated-report", "k"], [" In reports/Supplier_", "x"]],
    [["Scorecard.rdl: add a title \"Meridian", "x"]],
    [["Supplier Scorecard\" and a page footer", "x"]],
    [["with the page number, make the header row", "x"]],
    [["bold with a dark background and white", "x"]],
    [["text, a thin grey border on every cell,", "x"]],
    [["format on-time % as 0.0%, and set the", "x"]],
    [["page to A4 landscape.", "x"]],
  ], { dark: true, size: 13, lang: "text" });
}
{
  const s = slide("CONTENT", "PART 3 · A REAL ERROR FROM OUR TEST RUN", "Turn a mistake into a rule",
    "THE BORDERSTYLE ERROR (about 3 minutes)\n\nSay: When we tested this, Report Builder refused the file Claude wrote. This is the exact message. Don't panic at a wall of red text; read it in three pieces.\n\nSay (point at the highlights): What's wrong: 'invalid child element BorderStyle'. What's allowed: the list after 'possible elements expected', and Border is in it. Where: line 150, position 370.\n\nSay (point at the code): Borders in RDL are a Border element with Color, Style and Width inside. BorderStyle doesn't exist. Claude invented a name that sounds right.\n\nSay: Two lessons. One: paste the whole error into Claude and it fixes it in seconds. Two: we wrote the rule into the paginated-report Skill, so it won't make that mistake again. That's what Skills are for: turning a lesson into a rule.");
  pill(s, "p4");
  code(s, 0.6, 1.9, 12.13, 1.6, [
    [["Deserialization failed: The element 'Style' ... has ", "x"], ["invalid child element 'BorderStyle'", "h"], [".", "x"]],
    [["List of possible elements expected: '", "x"], ["Border", "a"], [", TopBorder, BottomBorder, LeftBorder, RightBorder,", "x"]],
    [["BackgroundColor, ... FontWeight, Format, TextAlign, ... PaddingBottom ...'. ", "x"], ["Line 150, position 370.", "h"]],
  ], { size: 13, lang: "text" });
  card(s, 0.6, 3.75, 5.9, 2.75, { line: C.accent5, lw: 1.5, fill: C.background1, name: "Wrong" });
  txt(s, "WRONG · REPORT BUILDER REJECTS IT", 0.85, 3.9, 5.4, 0.3, { fontSize: 11.5, bold: true, color: C.accent5, charSpacing: 1.2 });
  code(s, 0.85, 4.35, 5.4, 0.9, [[["<Style>", "k"]], [["  <BorderStyle>Solid</BorderStyle>", "h"]], [["</Style>", "k"]]], { size: 13, lang: "text" });
  txt(s, "Sounds right, but the RDL schema has no BorderStyle element.", 0.85, 5.45, 5.4, 0.8, { fontSize: 13.5, color: C.accent6 });
  card(s, 6.83, 3.75, 5.9, 2.75, { line: C.accent4, lw: 1.5, fill: C.background1, name: "Right" });
  txt(s, "RIGHT · A BORDER ELEMENT INSIDE STYLE", 7.08, 3.9, 5.4, 0.3, { fontSize: 11.5, bold: true, color: C.accent4, charSpacing: 1.2 });
  code(s, 7.08, 4.35, 5.4, 2.0, [[["<Style>", "k"]], [["  <Border>", "a"]], [["    <Color>#cfd4da</Color>", "x"]], [["    <Style>Solid</Style>", "x"]], [["    <Width>1pt</Width>", "x"]], [["  </Border>", "a"]], [["</Style>", "k"]]], { size: 12.5, lang: "text" });
}

section("Part 4 · Build your own Skill");
{
  const s = slide("LIVE", "PART 4 · EVERYONE DOES THIS", "Create a Skill with skill-creator",
    "SKILL-CREATOR (about 10 minutes)\n\nSay: skill-creator is a Skill that writes Skills. It isn't built into Claude Code; it comes from Anthropic's skills marketplace, so we install it once.\n\nCoach step 1: 'In a terminal Claude session, type the first plugin line, wait, then the second. If it asks where, choose user.' If /plugin isn't recognised (VS Code chat panel), use the claude plugin lines in a normal terminal instead.\n\nCoach step 2: 'Type /skill and look at the suggestions.' Plugin Skills are named plugin:skill, so expect /example-skills:skill-creator. On my machine it's /anthropic-skills:skill-creator from a different plugin. Use whatever the list shows.\n\nCoach step 3: 'Paste the prompt from the guide (Part 5, step 3). Answer its questions, then read the SKILL.md it writes before you accept it.'\n\nCoach step 4: 'Type /service-check.' Expected: four headings and a Flags list. Then commit it to Git.\n\nSay: New Skills load without a restart. If one doesn't show up, run /reload-skills.\n\nShort on time: copy the ready-made service-check folder from the guide's Skill files section instead.");
  pill(s, "p5", true);
  liveSteps(s, [["Install skill-creator (once)", "From Anthropic's skills marketplace"], ["Find its name", "Type /skill · expect /example-skills:skill-creator"], ["Describe the Skill you want", "Prompt in the guide · read what it writes"], ["Run it, then commit", "/service-check · git add + commit"]], 0.6, 1.95, 5.6);
  code(s, 6.45, 1.95, 6.28, 1.2, [[["/plugin marketplace add anthropics/skills", "x"]], [["/plugin install example-skills@anthropic-agent-skills", "x"]]], { dark: true, size: 12.5, lang: "text" });
  code(s, 6.45, 3.3, 6.28, 1.35, [[["/example-skills:skill-creator", "k"], [" Create a project", "x"]], [["skill called service-check ... read-only ...", "x"]], [["ends with a Flags list.", "x"]]], { dark: true, size: 12.5, lang: "text" });
  code(s, 6.45, 4.8, 6.28, 1.3, [[["/service-check", "k"]], [["git add .claude/skills/service-check", "x"]], [["git commit -m \"Add service-check skill\"", "x"]]], { dark: true, size: 12.5, lang: "text" });
  txt(s, "Skills reload automatically · if not: /reload-skills", 6.45, 6.25, 6.28, 0.35, { fontSize: 13, color: C.background2, transparency: 25 });
}

section("Part 5 · Automate");
{
  const s = slide("CONTENT", "PART 5 · RUN SKILLS AUTOMATICALLY", "Four ways to run a Skill on a timer",
    "FOUR OPTIONS (about 3 minutes)\n\nSay: Four ways to run a Skill without typing it. What matters is where each one runs, because our checks need Power BI Desktop or a Power BI sign-in on this laptop.\n\nSay (point at each card): /loop runs inside your open Claude session: great for a demo, but it stops when you close the session, and recurring loops expire after seven days.\n\nSay: Claude Desktop scheduled tasks run on your laptop from the Claude Desktop app, while the app is open and the laptop is awake.\n\nSay: Windows Task Scheduler runs claude -p, a one-shot headless run, at a set time and writes a log. That's a standard Windows feature, not an Anthropic recipe, so test it by hand first.\n\nSay: Cloud /schedule runs in Anthropic's cloud on a fresh copy of your Git repo. It can't see your laptop, Desktop or your files, so it's not for this.\n\nSay: And remember: the daily refresh itself is Power BI's job. We schedule the checking.");
  pill(s, "p6");
  const o = [["/loop", "Your open session", "Demos, watching for an hour", "Stops with the session; expires after 7 days", C.accent3],
    ["Desktop scheduled task", "Your laptop, Claude Desktop app", "A daily check, no scripts", "App open, laptop awake", C.accent2],
    ["Windows Task Scheduler", "Your laptop: Windows runs claude -p", "A dependable morning run with a log", "Not an Anthropic recipe: test by hand", C.accent1],
    ["Cloud /schedule", "Anthropic's cloud, repo copy", "Work that only needs the repo", "Can't see your laptop: not for this", C.accent5]];
  o.forEach(([h, where, good, catchT, col], i) => {
    const x = 0.6 + i * 3.1;
    card(s, x, 1.95, 2.85, 4.55, { line: col, lw: 1.5, fill: C.background1, name: h });
    txt(s, h, x + 0.25, 2.12, 2.4, 0.75, { fontSize: 17, bold: true, color: col });
    txt(s, "RUNS", x + 0.25, 2.95, 2.4, 0.25, { fontSize: 10.5, bold: true, color: C.accent6, charSpacing: 1.2 });
    txt(s, where, x + 0.25, 3.2, 2.4, 0.6, { fontSize: 13.5 });
    txt(s, "GOOD FOR", x + 0.25, 3.9, 2.4, 0.25, { fontSize: 10.5, bold: true, color: C.accent6, charSpacing: 1.2 });
    txt(s, good, x + 0.25, 4.15, 2.4, 0.7, { fontSize: 13.5 });
    txt(s, "CATCH", x + 0.25, 4.95, 2.4, 0.25, { fontSize: 10.5, bold: true, color: C.accent6, charSpacing: 1.2 });
    txt(s, catchT, x + 0.25, 5.2, 2.4, 1.1, { fontSize: 13.5, color: i === 3 ? C.accent5 : C.text1, bold: i === 3 });
  });
}
{
  const s = slide("CONTENT", "PART 5 · AUTOMATION IN PRACTICE", "An automated morning, end to end",
    "AN AUTOMATED MORNING (about 3 minutes)\n\nSay: Let's put it all together. This is what a normal morning can look like.\n\nSay (follow the timeline): At 7, Power BI's own scheduled refresh runs. That's Power BI's job, not Claude's. At 9, Windows Task Scheduler starts Claude in headless mode with one command: run /service-check. A couple of minutes later, the result is in a log file, ending with the Flags list. At quarter past, you open the log. No flags? You're done in ten seconds. A flag? You decide: fix it in Desktop and republish, or type /refresh-now yourself.\n\nSay (point at the bottom): Notice who does what. Power BI refreshes, the Skill checks, you decide. Nothing that changes the Service runs unattended.\n\nSay: One practical catch: an unattended run against the Service needs a sign-in that's still valid. If it asks you to sign in every morning, the long-term fix is a service principal from IT. Against Power BI Desktop, it needs nothing extra, which is why today's homework schedules report-snapshot.");
  pill(s, "p6");
  const t = [["07:00", "Scheduled refresh", "Power BI", C.accent1], ["09:00", "Task Scheduler runs claude -p \"/service-check\"", "Windows + Claude", C.accent3], ["09:02", "Log written, ending with Flags", "The Skill", C.accent4], ["09:15", "You read the flags", "You", C.text1], ["If flagged", "Fix in Desktop + republish, or /refresh-now", "You, by hand", C.accent5]];
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 2.55, w: 11.5, h: 0, line: { color: HEX.line, width: 2 } });
  t.forEach(([time, what, who, col], i) => {
    const x = 0.6 + i * 2.48;
    s.addShape(pres.shapes.OVAL, { x: x + 0.95, y: 2.4, w: 0.3, h: 0.3, fill: { color: col }, line: { type: "none" } });
    txt(s, time, x, 1.85, 2.2, 0.4, { fontSize: 17, bold: true, color: col, align: "center" });
    card(s, x, 2.95, 2.2, 1.85, { name: time });
    txt(s, what, x + 0.15, 3.1, 1.9, 1.15, { fontSize: 13.5, bold: true, align: "center" });
    txt(s, who, x + 0.15, 4.3, 1.9, 0.35, { fontSize: 11.5, color: col, bold: true, align: "center" });
  });
  const roles = [["Power BI refreshes", C.accent1], ["The Skill checks", C.accent4], ["You decide", C.text1]];
  roles.forEach(([h, col], i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 5.1, 3.9, 0.75, { fill: C.background1, line: col, lw: 1.5, name: h });
    txt(s, h, x, 5.1, 3.9, 0.75, { fontSize: 18, bold: true, color: col, align: "center", valign: "middle" });
  });
  txt(s, "Unattended runs against the Service need a sign-in that's still valid (a service principal from IT, long term). Against Desktop, nothing extra.", 0.6, 6.05, 12.13, 0.55, { fontSize: 13, color: C.accent6 });
}
{
  const s = slide("LIVE", "PART 5 · LIVE + HOMEWORK", "Schedule a read-only Skill",
    "SCHEDULE IT (about 7 minutes)\n\nDo step 1 live: type the /loop line. Say: Every 30 minutes, while this session is open, Claude runs report-snapshot. Stop it by saying 'cancel the report-snapshot loop'.\n\nSay step 2: claude -p runs one prompt and exits. A /skill-name in the prompt runs that Skill. --allowedTools lists what it may use without asking, because nobody is there to click Yes. Try it by hand first.\n\nSay step 3: save run-snapshot.cmd in the project folder (it's in the guide), then one schtasks line schedules it for 9 am on weekdays. schtasks /Run tests it now; type logs\\snapshot.log shows the result.\n\nSay (point at the warning): Only schedule read-only Skills. A scheduled run approves its own tool calls, so never schedule refresh-now. And for the published model, unattended runs need a sign-in that's still valid; if it asks you to sign in every morning, that needs a service principal from IT. So today we schedule report-snapshot against Desktop, which needs no sign-in.\n\nHomework: schedule it for tomorrow morning and send me the log.");
  pill(s, "p6", true);
  liveSteps(s, [["Live: /loop", "Runs every 30 min while the session is open"], ["Try headless by hand", "claude -p runs one Skill and exits"], ["Schedule it with Windows", "run-snapshot.cmd + schtasks · read logs\\snapshot.log"]], 0.6, 1.95, 5.2);
  code(s, 6.1, 1.95, 6.63, 0.6, [[["/loop 30m /report-snapshot", "k"]]], { dark: true, size: 13, lang: "text" });
  code(s, 6.1, 2.7, 6.63, 1.0, [[["claude -p \"/report-snapshot\" --allowedTools", "x"]], [["  \"mcp__powerbi-modeling__*\" \"Bash(python *)\" --max-turns 30", "x"]]], { dark: true, size: 12, lang: "text" });
  code(s, 6.1, 3.85, 6.63, 1.3, [[["schtasks /Create /TN \"Meridian report snapshot\"", "x"]], [["  /SC WEEKLY /D MON,TUE,WED,THU,FRI /ST 09:00", "x"]], [["  /TR \"C:\\Training\\pbip-project\\run-snapshot.cmd\"", "x"]]], { dark: true, size: 12, lang: "text" });
  card(s, 6.1, 5.35, 6.63, 1.2, { fill: C.accent5, t: 85, line: C.accent5, name: "Warning" });
  txt(s, [{ text: "Only schedule read-only Skills. ", options: { bold: true, color: "FF9A8F" } }, { text: "A scheduled run approves its own tool calls. Never schedule refresh-now.", options: {} }], 6.35, 5.35, 6.2, 1.2, { fontSize: 14, color: C.background1, valign: "middle" });
}

section("Close");
{
  const s = slide("CONTENT", "BEFORE YOU USE THIS AT WORK", "Five rules for Claude in the Service",
    "RULES (about 2 minutes)\n\nSay: Five rules, building on last week's four.\n\nSay: One, read-only first: only the --readonly server in these sessions; check /mcp. Two, never write to a published model: change Desktop or PBIP, then republish. Three, the refresh belongs to Power BI; Claude checks it. Four, for paginated reports, Claude edits and Report Builder judges. Five, schedule only read-only Skills; anything that acts stays manual with disable-model-invocation.\n\nSay (point at the quote): And the data rule still applies: query results go to the AI provider. Sample data only, never a client's model without permission, always our AI policy.");
  pill(s, "licences");
  const r = [["Read-only first", "Only the --readonly server. Check /mcp in every new session."], ["Never write to the Service", "Change Desktop or PBIP, then republish. A write blocks the .pbix download."], ["Power BI refreshes, Claude checks", "Scheduled refresh in the Service; service-check confirms it."], ["Claude edits, Report Builder judges", "Reopen every .rdl Claude touched and click Run."], ["Schedule only read-only Skills", "Anything that acts stays manual: disable-model-invocation: true."]];
  r.forEach(([h, d], i) => {
    const y = 1.9 + i * 0.78;
    num(s, 0.6, y, i + 1, C.accent2, 0.46);
    txt(s, h, 1.3, y - 0.02, 4.2, 0.45, { fontSize: 17, bold: true, valign: "middle" });
    txt(s, d, 5.6, y - 0.02, 7.1, 0.5, { fontSize: 14.5, color: C.accent6, valign: "middle" });
  });
  card(s, 0.6, 5.9, 12.13, 0.75, { fill: C.accent5, t: 92, name: "Quote" });
  txt(s, [{ text: "Still true: ", options: { bold: true } }, { text: "\"AI assistance does not expand data access, but it may transmit accessed data to the LLM provider.\" Sample data only.", options: { italic: true } }], 0.9, 5.9, 11.6, 0.75, { fontSize: 14, valign: "middle" });
}
{
  const s = slide("DARK", "NEXT · SESSION 03 · THE MODEL AS FILES", "Next: the model as files",
    "CLOSE (about 3 minutes)\n\nSay: Three takeaways. One: Claude can read the published model, but only on Premium, PPU or Fabric capacity, and we never write to it. Two: a Skill is a saved procedure: same steps, same output, shared through Git, and it can check, print or act. Three: read-only Skills can run on a schedule, and anything that acts stays manual.\n\nSay: Homework: add a Year parameter to the Supplier Scorecard with Claude and reopen it in Report Builder; schedule run-snapshot.cmd for tomorrow and send me the log; Track A, run /service-check after the next scheduled refresh.\n\nSay: Next session: we save Meridian as a PBIP project, open the folder in VS Code, and let Claude change the model as files, then read the Git diff and undo it.\n\nDo: paste the guide link in the chat. Questions?");
  s.addText("Save as PBIP, let Claude edit the files, read the Git diff and undo it", { placeholder: "sub" });
  const take = [["Read, never write", "Claude on the published model; changes go through Desktop"], ["A Skill is a saved procedure", "Same steps, same output, shared through Git"], ["Read-only runs itself", "Schedule checks; anything that acts stays manual"]];
  take.forEach(([h, d], i) => {
    const x = 0.8 + i * 3.98;
    card(s, x, 4.45, 3.7, 1.35, { fill: C.background1, t: 91, line: HEX.lineDark, name: h });
    txt(s, h, x + 0.25, 4.62, 3.3, 0.4, { fontSize: 18, bold: true, color: C.background1 });
    txt(s, d, x + 0.25, 5.05, 3.3, 0.65, { fontSize: 13.5, color: C.background2, transparency: 25 });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 6.2, w: 3.2, h: 0.5, rectRadius: 0.25, fill: { color: C.background1, transparency: 88 }, line: { color: C.accent3, width: 1 } });
  s.addText([{ text: "Open the Session 02 guide  ↗", options: { color: "F2B544", hyperlink: { url: `${GUIDE}#p7`, tooltip: "Opens the guide" } } }], { x: 0.8, y: 6.2, w: 3.2, h: 0.5, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
  txt(s, "Homework: Year parameter · schedule run-snapshot.cmd · Track A: /service-check after the next refresh", 4.2, 6.32, 8.5, 0.3, { fontSize: 13, color: C.background2, transparency: 35 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  if (!process.env.SKIP_THEME) await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
