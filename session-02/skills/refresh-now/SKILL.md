---
name: refresh-now
description: Start an on-demand refresh of the published Meridian semantic model and report whether it succeeded. Only when the user explicitly asks to refresh now.
disable-model-invocation: true
---

# Refresh now (changes the service: run only on request)

`disable-model-invocation: true` means Claude never runs this on its own. You must type `/refresh-now`.

## Settings

- Workspace ID: `<workspace-id>`
- Semantic model ID: `<model-id>`

## Steps

1. Ask the user to confirm: "Start a refresh of MeridianSupplyChain in the service now?" Stop unless they say yes.
2. Start the refresh:
   `powershell -NoProfile -ExecutionPolicy Bypass -File .claude/skills/refresh-now/scripts/refresh_now.ps1 -WorkspaceId <workspace-id> -DatasetId <model-id>`
3. Wait about a minute, then run `.claude/skills/service-check/scripts/service_check.ps1` (same IDs) to read the latest refresh status. Repeat up to 5 times until the status is Completed or Failed.
4. Report: status, how long it took, and the error text if it failed. If the error mentions a gateway or credentials, say: "check the gateway is running and the data source credentials in the semantic model's settings".

## Notes

- On shared (non-Premium) capacity, the API allows about 8 refresh requests per day.
- Power BI's own **scheduled refresh** is the right tool for daily refreshes. Use this Skill only for "refresh it now".
