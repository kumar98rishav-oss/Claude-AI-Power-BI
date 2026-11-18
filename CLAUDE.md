# Context for Claude

This repo holds a training programme: **Claude + Power BI**, for Programmers.io. The trainer is a Power BI developer / data engineer. Sessions are **1 hour**, practical, and use one fictional dataset: **Meridian Medical Supply** (`meridian-dataset/`, PBIP copy in `pbip-project/`).

## Status

- Session 01 "Connecting Claude to Power BI" is finished: `session-01/` (20-slide deck, transcript, map).
- Session 02 "The model as files: PBIP and TMDL" is next. Plan and to-do list: `docs/NEXT_STEPS.md`.
- Decisions, verified commands and facts from building Session 01: `docs/CHAT_NOTES.md`. Read it before answering questions about MCP setup, TOM/ADOMD.NET, data safety or the report-snapshot Skill.

## How to write for this programme

- Transcripts: casual, friendly, spoken English. Quote blocks = what the trainer says; **Do** lines = clicks; **Ask/Answer**, **Expected** with real numbers.
- Decks: short and basic. Depth goes into later sessions.
- Only claim what is verified. When something is inferred (e.g. which MCP tool uses TOM vs ADOMD.NET), say so.
- Facts about Meridian live in `docs/CHAT_NOTES.md` section 5. Don't invent numbers.

## Rules

- Never commit client data, API keys or personal details. This repo is public.
- Read-only first when connecting Claude to any model.
- PPTX files can't be rebuilt from a phone; draft outlines and transcripts as Markdown in `docs/` or `session-NN/`, and rebuild decks on the laptop with `session-01/build/`.
