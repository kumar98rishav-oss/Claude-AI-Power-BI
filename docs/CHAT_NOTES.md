# Chat notes: building Session 01

The important decisions, questions and answers from the Claude sessions that built Session 01 (November 2026), in the order they happened. Raw transcripts are not stored here; this is the part worth keeping.

## 1. The programme plan (approved)

- Sessions are **1 hour** each. Session 1 gives the basic understanding; depth comes later.
- **Session 1, "Connecting Claude to Power BI":** what MCP is; Claude Code in VS Code (vs Claude Desktop); connect read-only to an open .pbix; read tables, columns, measures and relationships. Hands-on: install the Power BI MCP server, register it with Claude Code, connect, ask three questions, check the answers against Desktop.
- **Session 2, "The model as files: PBIP and TMDL":** save as PBIP and open in VS Code; Claude renames a column across measures; read the change as a Git diff, then undo it.
- Programme objectives (slide 2): automate Power BI reporting; understand Power BI metadata (PBIP, TMDL, PBIR); build Skills with Claude on TMDL; troubleshoot and understand any report.

## 2. Deck evolution (Session 01)

| Change | Why |
| --- | --- |
| Rebuilt as a short, basic deck (the 32-slide deep dive moved to `session-01/for-session-02/`) | Session 1 is overview only |
| Added Skills: create, save, recall, applications | User asked for Skills in Session 1 |
| Expanded setup into check → install + sign in → register → quick fixes → connect | Make install/login easy for participants |
| Slide 2 changed to programme objectives | User's rough brief, structured |
| Added slide "Inside the server: TOM and ADOMD.NET" | User asked how the MCP server works |
| Skills split into "What a Skill is" + "What Skills can do for a Power BI team" (6 applications) | Explain the idea before the demo |
| Removed "Claude Code vs Claude Desktop" slide | User request |
| Added slide 16 "report-snapshot" and made the live demo run it | User wanted one real Skill with 5 checks |
| Transcript rewritten in a casual, friendly tone, with the user's own intro | User request |

Final deck: **20 slides**, `session-01/Session01_Connecting_Claude_to_Power_BI.pptx`.

## 3. Key technical answers

### Installing and registering

- Claude Code install (official, no admin): `irm https://claude.ai/install.ps1 | iex` in PowerShell. CMD alternative: `curl -fsSL https://claude.ai/install.cmd -o install.cmd && install.cmd && del install.cmd`. Needs a **Pro, Max, Team or Enterprise** plan (not Free). Check with `claude --version` and `claude doctor`.
- Register the server **once for every folder**:

  ```
  claude mcp add --scope user powerbi-modeling '--' cmd /c npx -y @microsoft/powerbi-modeling-mcp@latest --start --readonly
  ```

- **Error `unknown option '-y'`:** PowerShell swallows a bare `--` when `claude` is the npm `claude.ps1` shim. Fix: quote it as `'--'` (or use CMD / `claude.cmd`).
- Without `--scope user`, `claude mcp add` registers for the **current folder only** (local scope).
- `--readonly` = Microsoft's "safe mode": no writes to the model.
- Connect prompt (from Microsoft's docs): `Connect to 'MeridianSupplyChain' in Power BI Desktop`.

### Uninstalling for a live demo

- Remove: `claude mcp remove powerbi-modeling --scope user`. If it still shows up, there's also a **folder-only** copy: run `claude mcp remove powerbi-modeling --scope local` in that folder. `claude mcp remove` without `--scope` removes from whichever scope it finds.
- The `powerbi-authoring` plugin brings its own (read-write) Power BI server. Turn it off for the demo: `claude plugin disable powerbi-authoring@fabric-collection`; back on: `claude plugin enable powerbi-authoring@fabric-collection`.
- Start a **new** Claude session after any change; open sessions keep their old servers.
- `npx` caches the package, so a re-install on the trainer's laptop is faster than participants' first install.

### Signing in

- The VS Code extension and the `claude` terminal command both sign in through the browser (claude.ai: email link/code, Google, or SSO → **Authorize**). **No Claude Desktop app needed.**
- The extension does **not** add `claude` to the terminal, so participants still run the installer. Servers registered in the terminal show up in the extension too (same settings).

### TOM and ADOMD.NET

- `powerbi-modeling-mcp.exe` bundles **Microsoft.AnalysisServices.Tabular (TOM)** and **Microsoft.AnalysisServices.AdomdClient (ADOMD.NET)** (DLL names verified inside the exe).
- **TOM** reads and changes model structure: tables, columns, measures + DAX, relationships. Same library as Tabular Editor. `--readonly` blocks its writes.
- **ADOMD.NET** runs DAX queries and returns rows. Same library as DAX Studio.
- Both connect to the Analysis Services engine inside Power BI Desktop on a local port.
- Which MCP tool uses which library is **inferred** from what each library does; Microsoft doesn't document it.

### What Claude sees (data safety)

- Connecting does **not** upload the model. Claude only sees what each tool call returns.
- Metadata calls (TOM) return names, types, DAX, relationships. No rows. Row counts need a DAX query.
- DAX queries (ADOMD.NET) return rows, and **Claude reads every row a query returns**. "Distinct customer names" means Claude sees all the names.
- **Read-only protects the model, not the data.** Query results go into the conversation, which the AI provider processes.
- Row-level security doesn't filter you as the model's author in Desktop.
- Safety habits: check company policy, practise on sample data, start read-only, read every approval (it shows the DAX), ask for aggregates/top-N, keep personal data out of prompts.

### Skills

- A Skill = a folder with `SKILL.md` (frontmatter `name`, `description`, then plain-English steps). Project Skill: `.claude/skills/<name>/SKILL.md` (shared via Git). Personal: `~/.claude/skills/<name>/SKILL.md`.
- Recall with `/<name>` or plain words. Restart Claude after adding a Skill.
- Six realistic Power BI applications (slide 15): daily health check; source vs report (SQL Server MCP + Power BI MCP); threshold alerts (run on demand, not push); what changed (PBIP + Git diff); visual/format check (PBIR files); performance (biggest columns, slowest measures, Power Query steps likely to break folding, confirmed with View Native Query).

## 4. The report-snapshot Skill

In `pbip-project/.claude/skills/report-snapshot/`. Read-only. Five checks:

| Check | How |
| --- | --- |
| Last refreshed | DAX `INFO.PARTITIONS()` → `RefreshedTime` (UTC) |
| Key KPIs | One DAX query, latest year vs previous, thresholds in SKILL.md |
| Relationships | DAX `INFO.VIEW.RELATIONSHIPS()` + measures using `USERELATIONSHIP` |
| Page/visual render time | `scripts/perf_summary.py` reads a Performance Analyzer export; `visualId` = PBIR visual folder name; pages split by "User Action" events |
| Power Query steps | `scripts/pq_steps.py` reads M code from the TMDL files |

Verified results on Meridian: last refresh 9 Nov 2026 12:25 UTC; Revenue 2025 14,134,115 (+10.1%); flags On-Time Delivery 89.8% (< 90%) and Stockouts 45 (> 40); 15 relationships (12 active, 3 inactive, all many-to-one, single direction); pages 1,060 / 1,012 / 792 ms, slowest visual KPI card 954 ms; 9 tables from CSV with headers + types, 2 facts also blank → null.

**Performance Analyzer can't be driven by Claude.** Someone records it once: Optimize › Performance analyzer › Start recording › Refresh visuals › visit every page › Stop › Export, saved next to the `.pbip`.

## 5. Meridian facts to know

- 10 tables, 84 columns, 30 measures, 15 relationships, 3 inactive (Sales_ShipDate, Purchases_PromisedDate, Purchases_ReceivedDate); 2 measures use USERELATIONSHIP (Revenue by Ship Date, Spend by Received Date).
- 2025 revenue 14,134,114.90. Pacific Rim Medical 61.2% on time in 2025; by quarter 2024: 100 / 87.0 / 96.0 / 96.2%, 2025: 88.9 / 88.9 / **28.0 / 33.3%**.
- Gross Margin % = `DIVIDE ( [Gross Margin], [Revenue] )`, `_Measures`, folder 01 Sales, format `0.0%`.
- 10 of 30 measures have no description. All have format strings.
- The data is **CSV files** today (not Excel). Plan: move to SQL Server and add a SQL Server MCP.

## 6. Gotchas hit while building

- Building the deck while it's open in PowerPoint fails with `EBUSY`. Close it, or build to another file with the `OUT` env variable.
- Opening Power BI Desktop through automation can start a second, blank Desktop. Keep only one open so Claude doesn't have to choose.
- The live map is served from the portfolio site; it deploys on push to that repo.

## 7. Power BI service and paginated reports (tested after Session 01)

- Claude connected to the published model in workspace `Data Analysis_dev` with `Connect to semantic model 'MeridianSupplyChain' in Fabric Workspace '<workspace>'` (XMLA endpoint).
- XMLA needs a Premium, PPU or Fabric capacity workspace (the Fabric trial gives PPU-level rights). A Free account on My workspace can't connect.
- Microsoft: "a write operation on a semantic model authored in Power BI Desktop prevents it from being downloaded back as a PBIX file." Never write to the published model in training.
- Gateway status and refresh history aren't MCP tools; they come from the Power BI REST API (`MicrosoftPowerBIMgmt`, `Invoke-PowerBIRestMethod`). Refresh history needs Write permission. Meridian's CSV source needs a gateway (personal mode is fine for training).
- Paginated reports: Report Builder (free, winget `Microsoft.PowerBIReportBuilder`) makes the connection; Claude edits the .rdl XML; Report Builder validates. Real error hit: `BorderStyle` is not a valid child of `Style`; borders use a `<Border>` element.
- Skills: skill-creator is not built in; install with `/plugin marketplace add anthropics/skills` + `/plugin install example-skills@anthropic-agent-skills`. Skills reload live (`/reload-skills` if needed). `claude -p "/skill"` runs a Skill headless; `--allowedTools "mcp__powerbi-modeling__*"`. Scheduling: `/loop` (session), Desktop scheduled tasks (app open), Windows Task Scheduler + `claude -p` (not an Anthropic recipe), cloud `/schedule` can't reach the laptop.
