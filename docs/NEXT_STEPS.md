# Next steps

## Before Session 01 (Tuesday)

- [ ] Re-record Performance Analyzer on `pbip-project/MeridianSupplyChain.pbip` and save `PowerBIPerformanceData.json` next to it.
- [ ] Remove the MCP server and the `powerbi-authoring` plugin before class so the install demo starts clean (commands in [CHAT_NOTES.md](CHAT_NOTES.md#uninstalling-for-a-live-demo)).
- [ ] One full run-through with a stopwatch, including `/report-snapshot` in a fresh Claude session.
- [ ] Send the participant prep message (in the transcript, "Message to send everyone the day before").

## Session 02: The model as files (PBIP and TMDL)

Approved plan:

1. Save Meridian as **PBIP** and open the folder in VS Code.
2. Tour the folder: `*.Report` (PBIR: `definition/pages/<page>/visuals/<visual>/visual.json`) and `*.SemanticModel` (TMDL: one file per table, `relationships.tmdl`, `model.tmdl`, `expressions.tmdl`). `.pbi/cache.abf` and `localSettings.json` are local-only (gitignored).
3. `git init` + first commit.
4. Claude **renames a column across every measure** that uses it, editing the TMDL files.
5. Read the change as a **Git diff**.
6. **Undo** it with one command (`git restore .` or `git checkout -- .`).

Material to reuse:

- `session-01/for-session-02/Session01_Power_BI_Internals.pptx`: 32-slide deep dive on PBIP, TMDL and PBIR.
- `session-01/for-session-02/Session 01 Speaker Script.pdf`: its speaker script.
- The interactive map (tabs 00–05 cover formats, the project folder, TMDL, relationships and PBIR).
- `pbip-project/` is already a working PBIP of Meridian.

To do:

- [ ] Pick the column to rename (it needs to be used in several measures; check `_Measures.tmdl`).
- [ ] Build the Session 02 deck (short, like Session 01) and transcript in the same casual tone.
- [ ] Decide whether Session 2 uses read-write MCP or only file edits (safer: file edits + Git).

## Later sessions (ideas already agreed)

- Move Meridian's data from CSV to **SQL Server** (`meridian-dataset/sql/` has the load scripts) and add a **SQL Server MCP** next to the Power BI one.
- Skills from slide 15: source vs report reconciliation, threshold alerts, what changed (Git diff), visual/format check (PBIR), performance + query folding.
- `meridian-dataset/data_quality_lab/` is ready for a data-quality session.
