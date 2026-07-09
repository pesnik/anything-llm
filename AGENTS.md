# Agents Guide

## Banglalink PPTX Skill

### Source of Truth
**Source repo:** `~/Pesnik/skills/anythingllm-banglalink-skills/`

This is where all code changes must be made. The source structure:
- `handler.js` — main orchestrator (Flow 1-4, QA loop)
- `pptx/section-agent.js` — per-section subagent prompt
- `pptx/lib/qa.js` — structural + visual QA
- `pptx/lib/qa-reviewer.js` — QA fix recommendation
- `pptx/lib/slides.js` — per-slide-type rendering
- `pptx/lib/ghost-logo.js` — post-processing XML injection
- `shared/brand/index.js` — brand tokens
- `plugin.json` — skill config

### Build & Deploy
After editing source files:
```bash
cd ~/Pesnik/skills/anythingllm-banglalink-skills
node scripts/build.js
```

The build script (`scripts/build.js`):
1. Copies root files (handler.js, plugin.json) to `dist/`
2. Copies `pptx/` contents (lib/, scripts/, section-agent.js) to `dist/`
3. Copies `shared/` to `dist/shared/`
4. Rewrites require paths (`./pptx/lib/` → `./lib/`)

### Deploy to AnythingLLM
```bash
rm -rf ~/Tools/anything-llm/server/storage/plugins/agent-skills/create-banglalink-pptx-presentation/*
cp -r ~/Pesnik/skills/anythingllm-banglalink-skills/dist/create-banglalink-pptx-presentation/* \
  ~/Tools/anything-llm/server/storage/plugins/agent-skills/create-banglalink-pptx-presentation/
```

Then restart with `yarn dev`.

### Key Architecture Notes
- Chat attachments arrive as base64, not file paths — filesystem agent can't access them
- Main agent makes TWO calls: first empty `{}` (triggers research), then with data (hasData bypass)
- Section-break slides have orange backgrounds — QA must skip them to avoid false positives
- KPI cards are 2.8"x2.5" — values ≤12 chars, labels ≤20 chars
- `originalSections` deep-copy prevents QA reviewer from fabricating data

## Teradata Skills (Viewpoint + EXPLAIN & Stats)

### Source of Truth
**Source repo:** `~/Pesnik/skills/anythingllm-teradata-skills/`

Monorepo containing both Teradata skills. Each skill has its own subdirectory with its own `plugin.json`, `handler.js`, and scripts.

```
anythingllm-teradata-skills/
├── scripts/build.js          ← builds both skills
├── teradata-viewpoint/        ← skill #1
│   ├── handler.js             Node wrapper → fetch_sessions.py + analyze_sessions.py
│   ├── plugin.json            accepts VIEWPOINT_USER, VIEWPOINT_PASS, VIEWPOINT_URL
│   ├── scripts/
│   │   ├── fetch_sessions.py  fetch active sessions from Viewpoint REST API
│   │   ├── analyze_sessions.py structured analysis (summary, top-cpu, blocked, skew, etc.)
│   │   └── jq_queries.sh      ad-hoc jq one-liners for quick drilldowns
│   └── references/
│       └── session_fields.md  reference of all 43 session fields
└── teradata-explain-stats/    ← skill #2
    ├── handler.js             Node wrapper → explain_sql.sh
    ├── plugin.json            accepts TD_HOST, TD_USER, TD_PASS
    └── scripts/
        ├── explain_sql.sh     run EXPLAIN via BTEQ (Docker or local)
        └── extract_and_explain.py  orchestrate EXPLAIN for sessions from JSON
```

### Build & Deploy
```bash
# Build both skills
cd ~/Pesnik/skills/anythingllm-teradata-skills
node scripts/build.js

# Deploy both to AnythingLLM
rm -rf ~/Tools/anything-llm/server/storage/plugins/agent-skills/teradata-viewpoint
rm -rf ~/Tools/anything-llm/server/storage/plugins/agent-skills/teradata-explain-stats
cp -r ~/Pesnik/skills/anythingllm-teradata-skills/dist/teradata-viewpoint \
  ~/Tools/anything-llm/server/storage/plugins/agent-skills/teradata-viewpoint
cp -r ~/Pesnik/skills/anythingllm-teradata-skills/dist/teradata-explain-stats \
  ~/Tools/anything-llm/server/storage/plugins/agent-skills/teradata-explain-stats
```

Then reload the AnythingLLM UI. Configure credentials under `@agent` → Agent Skills for each skill.

### Chained Workflow
1. Viewpoint fetches sessions → `/tmp/td_sessions.json`
2. User passes `session_id` from viewpoint output to `teradata-explain-stats` skill
3. EXPLAIN stats skill reads `/tmp/td_sessions.json` to get `sql_text` for that session

### EXPLAIN Execution Priority
The `explain_sql.sh` script auto-selects:
1. **Docker (teradata/tpt)** — preferred if Docker is running
2. **Local BTEQ** — fallback if Docker is unavailable

### Safety
- EXPLAIN is read-only by design
- Script blocks INSERT/UPDATE/DELETE/MERGE/UPSERT
- Only SELECT/WITH (CTE) queries are permitted
