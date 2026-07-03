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
