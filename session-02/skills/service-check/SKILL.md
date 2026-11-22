---
name: service-check
description: Read-only health check of the Meridian semantic model published in the Power BI service - last refresh, refresh history and failures, data sources and gateway, key KPIs with thresholds. Use when asked for a service check, "is the published report OK", or "did last night's refresh work".
---

# Service check (read-only)

Checks the **published** model, the same way every time. Never changes anything.

## Settings (edit for your workspace)

- Workspace name: `Data Analysis_dev`
- Workspace ID: `<paste from the URL: app.powerbi.com/groups/<workspace-id>/...>`
- Semantic model name: `MeridianSupplyChain`
- Semantic model ID: `<paste from the URL: .../datasets/<model-id>/details>`

## Rules

- Read only. Never create, change, refresh, publish or delete anything.
- Totals and lists only, never row-level data.
- If a step fails, say why in one line and carry on.

## Steps

1. **Connect** with the Power BI modeling MCP server:
   `Connect to semantic model 'MeridianSupplyChain' in Fabric Workspace 'Data Analysis_dev'`
2. **Last refresh (model side):** run
   `EVALUATE SELECTCOLUMNS ( INFO.PARTITIONS (), "Table", [Name], "RefreshedTime", [RefreshedTime] )`
   Report the newest time (UTC and local). Ignore `_Measures`.
3. **Refresh history, data sources and gateway (service side):** run
   `powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/service-check/scripts/service_check.ps1 -WorkspaceId <id> -DatasetId <id>`
   Show the last 5 refreshes (type, status, start, duration) and the error text of any failure. List each data source and whether it goes through a gateway.
4. **Key KPIs:** run the KPI query from the report-snapshot Skill (latest year vs previous) and apply its threshold table.

## Output

Four headings in this order: **Connection**, **Refresh**, **Data sources and gateway**, **KPIs**. End with **Flags**: any failed refresh, a refresh older than 24 hours, a data source without a gateway, any KPI threshold break. Or "No flags today."
