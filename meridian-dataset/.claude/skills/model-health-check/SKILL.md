---
name: model-health-check
description: Read-only health check of the Power BI semantic model Claude is connected to through the Power BI MCP server. Use when asked to check, review, audit or health-check a model.
---

# Model health check (read-only)

Run this against the model you are already connected to. If you are not connected, ask which open Power BI Desktop file to connect to first.

## Rules

- Read only. Never create, change, rename or delete anything, even if a fix looks obvious.
- Use the MCP tools to read the model. Do not guess from file names or earlier answers.

## Steps

1. Count the tables, columns, measures and relationships.
2. List every measure that has no description.
3. List every measure that has no format string.
4. List the inactive relationships, and the measures that switch them on with USERELATIONSHIP.

## Report

Reply with one table:

| Check | Result | Items |
| --- | --- | --- |

Then three lines: what is healthy, what needs attention first, and one suggestion. Suggest fixes in words only; do not apply them.
