---
name: report-snapshot
description: Read-only daily snapshot of the open Power BI report - last refresh, key KPI values with threshold flags, relationships (total, types, active and inactive), page and visual render times slowest first, and the Power Query steps behind every table. Use when asked for a report snapshot, a daily check, or "how is the report doing today".
---

# Report snapshot

A read-only check of the Power BI report in this project, run the same way every time.

## Rules
- Read only. Never create, change, refresh or delete anything in the model or the files.
- Use the Power BI modeling MCP server for items 1-3, and the two scripts in `scripts/` for items 4-5.
- Return totals and lists, never row-level data (no customer names, no transaction rows).
- If a step can't run, say why in one line and carry on with the next step.

## Setup
1. Connect to the open model: `connection_operations` ListLocalInstances, then Connect to the instance whose window title is this report.
2. Find the project files: the `*.SemanticModel/definition` and `*.Report/definition` folders next to the `.pbip` file.

## Steps

### 1. Last refreshed
Run with `dax_query_operations` Execute:
```dax
EVALUATE SELECTCOLUMNS ( INFO.PARTITIONS (), "Table", [Name], "RefreshedTime", [RefreshedTime] )
```
Report the latest `RefreshedTime` as the last refresh. It is UTC: show it in UTC and in this computer's local time. Ignore the `_Measures` table. If any table's refresh time is more than 1 hour older than the newest one, list it.

### 2. Key KPI values
Run:
```dax
DEFINE VAR LatestYear = YEAR ( MAX ( Fact_Sales[OrderDate] ) )
EVALUATE
SUMMARIZECOLUMNS (
    Dim_Date[Year],
    TREATAS ( { LatestYear - 1, LatestYear }, Dim_Date[Year] ),
    "Revenue", [Revenue],
    "Revenue YoY %", [Revenue YoY %],
    "Gross Margin %", [Gross Margin %],
    "On-Time Delivery %", [On-Time Delivery %],
    "Supplier On-Time %", [Supplier On-Time %],
    "OTIF %", [OTIF %],
    "Inventory Value", [Inventory Value],
    "Stockouts", [Stockouts]
)
ORDER BY Dim_Date[Year]
```
Show one row per KPI: latest year, previous year, change. Flag any KPI that breaks its threshold:

| KPI | Flag when |
|---|---|
| Revenue YoY % | below 0% |
| Gross Margin % | below 35% |
| On-Time Delivery % | below 90% |
| Supplier On-Time % | below 90% |
| OTIF % | below 85% |
| Stockouts | above 40 |

Edit this table to change the KPIs or thresholds for your team.

### 3. Relationships
Run:
```dax
EVALUATE SELECTCOLUMNS ( INFO.VIEW.RELATIONSHIPS (),
    "Name", [Name], "From", [FromTable] & "[" & [FromColumn] & "]", "To", [ToTable] & "[" & [ToColumn] & "]",
    "FromCard", [FromCardinality], "ToCard", [ToCardinality], "CrossFilter", [CrossFilteringBehavior], "Active", [IsActive] )
```
Report: the total; how many active and inactive; counts by type (many-to-one, one-to-one, many-to-many) and by filter direction (single, both). List every inactive relationship, and any many-to-many or both-direction one. Then run:
```dax
EVALUATE SELECTCOLUMNS ( FILTER ( INFO.VIEW.MEASURES (), CONTAINSSTRING ( [Expression], "USERELATIONSHIP" ) ), "Measure", [Name] )
```
and name the measures that switch an inactive relationship on.

### 4. Page and visual render times
This needs a Performance Analyzer recording, which only a person can make in Power BI Desktop:
Optimize > Performance analyzer > Start recording > Refresh visuals > click through every page > Stop > Export, saved next to the `.pbip` file.

Use the newest `PowerBIPerformanceData*.json` next to the `.pbip` file and run:
```
python .claude/skills/report-snapshot/scripts/perf_summary.py <the json file> <the *.Report/definition folder>
```
Show each page's render time, then its visuals slowest first with DAX query, visual display and other time. Other is mostly waiting for other visuals on the page to load. Name the three slowest visuals. Say how old the recording is. If there is no recording, give the steps above and skip to step 5.

### 5. Power Query transformations
Run:
```
python .claude/skills/report-snapshot/scripts/pq_steps.py <the *.SemanticModel/definition folder>
```
For each table, give its source and its steps in plain English. Point out anything unusual: a step that only one table has, a merge, a custom column, or a buffer.

## Output format
Use these headings, in this order, with short tables:

1. **Last refreshed**: one line, then any stale tables
2. **Key KPIs**: KPI | previous year | latest year | change | flag
3. **Relationships**: totals line, then the inactive and unusual ones
4. **Render times**: one table per page, slowest visual first
5. **Power Query**: one line per table: source, then steps

End with **Flags**: every threshold break, stale table, unusual relationship and slow visual (over 1,000 ms), or "No flags today."
