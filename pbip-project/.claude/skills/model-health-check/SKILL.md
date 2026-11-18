---
name: model-health-check
description: Run the model health check on the open MeridianSupplyChain Power BI model. Use when the user says "check model_health", "model health check" or similar. Answers - what is the revenue in 2025, and which supplier had the worst on-time % that year.
---

# Model Health Check

Always run the checks live against the open model. Never answer from earlier results.

## Steps
1. Connect to the open Power BI Desktop model: `connection_operations` -> `ListLocalInstances`, then `Connect` (skip if already connected).
2. Answer these two questions with DAX queries (`dax_query_operations`, `Execute`, `resultMode: Inline`):
   1. **What is the revenue in 2025?**
   2. **Which supplier had the worst on-time % in 2025?** Ignore suppliers with a blank on-time value, and sort explicitly with `ORDER BY` (TOPN alone does not return sorted rows).
3. Reply with the revenue and the worst supplier with its on-time %, noting any supplier excluded for having no value.
4. Check when was the report last refreshed by querying the `ReportRefresh` table. If the last refresh was more than 7 days ago, include a warning in your response.
5. Any viusal failures or errors in the model should be reported as well. If there are none, state that the model is healthy.