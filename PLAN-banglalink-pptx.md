# Banglalink PPTX Generator — AnythingLLM Custom Skill

## Overview
Enterprise-grade Banglalink-branded PowerPoint generator as an AnythingLLM custom skill.
Builds presentations with section-level research via sub-agents, strict brand enforcement,
and Playwright-powered HTML rendering. Published on the Community Hub.

## Skill Metadata
- **hubId**: `banglalink-pptx`
- **Display name**: Banglalink PPTX Generator
- **Description**: Generate Banglalink-branded PowerPoint presentations with per-section AI research, Playwright HTML rendering, and strict brand compliance.
- **Trigger keywords**: `presentation`, `ppt`, `pptx`, `slides`, `powerpoint`, `banglalink`, `deck`
- **Author**: Banglalink / AnythingLLM Community Hub

## File Structure
```
storage/plugins/agent-skills/banglalink-pptx/
├── plugin.json              # Skill metadata, trigger keywords, description
├── package.json             # Dependencies: playwright, puppeteer-core
├── handler.js               # Entry point — orchestrates the full workflow
├── brand.js                 # Brand tokens, base64-encoded PNG logos
├── html/
│   └── renderer.js          # Playwright launch → HTML→DOM-dimension extraction
├── pptx/
│   ├── builder.js           # pptxgenjs instance creation, slide orchestration
│   ├── slides.js            # Per-slide-type layout functions
│   ├── shapes.js            # Reusable shape helpers (branded tables, bullet lists, headers)
│   ├── ghost-logo.js        # jszip post-processing: unzip → inject XML → rezip
│   └── themes.js            # Color substitutions, accent mapping, style rules
├── section-agent.js         # Child AIbitat spawner (port of built-in pattern)
├── utils.js                 # Helpers: logo loading, Playwright detection, error handling
└── assets/                  # Included for reference (logos embedded as base64 in brand.js)
```

## Dependencies
| Package | Source | Size | Purpose |
|---------|--------|------|---------|
| `playwright` | npm install | ~300MB (with Chromium postinstall) | Primary HTML rendering engine |
| `puppeteer-core` | npm install | ~5MB | System Chrome fallback detection |
| `pptxgenjs` | Server node_modules (v4.0.1) | 0 (pre-installed) | PPTX generation |
| `jszip` | Server node_modules | 0 (pre-installed) | Ghost logo XML injection |

## Workflow (handler.js)

```
User: "Create a presentation about 5G expansion in Bangladesh"
  │
  ├─ 1. Parse request → { topic, structure, audience, outputFileName }
  ├─ 2. Instantiate pptxgenjs (13.333×7.5 in, Banglalink theme)
  ├─ 3. Build title slide (orange bg, white logo, centered title)
  ├─ 4. FOR each section:
  │      ├─ Spawn child AIbitat sub-agent (web search + scraping tools)
  │      │   └─ path: require("../../../../../server/utils/agents/aibitat/index.js")
  │      ├─ Sub-agent researches section topic
  │      ├─ Sub-agent calls submit-section-slides({ type, content, layout })
  │      └─ Orchestrator renders slides for that section
  │           └─ IF HTML rendering needed → Playwright extracts DOM positions
  ├─ 5. Build closing slide
  ├─ 6. jszip post-process: inject ghost logo XML into closing slide
  ├─ 7. Save to storage/generated-files/{outputFileName}.pptx
  └─ 8. Return file URL to user
```

## Slide Layouts

| Slide Type | Layout | Background | Logo |
|------------|--------|------------|------|
| **Title** | Full orange panel, centered title/subtitle | #F56F1F | `logo-full-white.png` (top-right) |
| **Section Divider** | Orange left panel (1/3) + white right panel (2/3) | Split #F56F1F / #FFFFFF | Ghost swoosh (left) |
| **Bullet Content** | Title bar + bullet list body | #FFFFFF | `logo-icon.png` (top-left) |
| **Two-Column** | Left: text bullets, Right: image/data callout | #FFFFFF | `logo-icon.png` |
| **Table** | Brand-styled table, orange header row | Alternating gray/white | `logo-icon.png` |
| **Quote** | Large text with orange left accent border | #FFFFFF | None |
| **Closing** | Ghost swoosh (jszip injected) + CTA + contact | #F56F1F gradient | Ghost logo |

## Brand Enforcement (brand.js)

```js
module.exports = {
  colors: {
    orange: 'F56F1F',
    dark: '1A1A1A',
    white: 'FFFFFF',
    gray: 'F5F5F5',
    darkGray: '333333',
    mediumGray: '666666',
    tableHeader: 'F56F1F',
    tableRowAlt: 'FFF3EB',
  },
  fonts: {
    heading: 'Arial',
    body: 'Arial',
    mono: 'Courier New',
  },
  slideSize: { width: 13.333, height: 7.5 },  // inches
  logos: {
    title: 'data:image/png;base64,...',   // logo-full-white.png
    content: 'data:image/png;base64,...',  // logo-icon.png
  },
  layout: {
    marginX: 0.5,
    marginY: 0.4,
    logoWidth: 1.8,   // title slide
    logoIconWidth: 0.5,  // content slides
    headerHeight: 0.8,
    footerHeight: 0.3,
  }
};
```

## Playwright/Chromium Strategy (html/renderer.js)

```
detectBrowser():
  1. Try `playwright.chromium.launch()` → success → return browser
  2. On failure: try `puppeteer-core.launch({ executablePath })` with platform paths:
     - macOS: /Applications/Google Chrome.app/.../Google Chrome
     - Linux: /usr/bin/chromium-browser, /usr/bin/google-chrome
     - Windows: C:\Program Files\Google\Chrome\Application\chrome.exe
  3. On failure: return null (graceful degradation)

renderHTML(html, slideWidth, slideHeight):
  1. Launch browser via detectBrowser()
  2. If no browser: return null (skip HTML, use basic shapes)
  3. Create page, set content to html
  4. Extract element positions via page.evaluate()
  5. Return { elements: [{ type, x, y, w, h, text, style }] }
  6. Close browser
```

## Ghost Logo Injection (pptx/ghost-logo.js)

After pptxgenjs generates the PPTX buffer:

```
async function injectGhostLogo(pptxBuffer) {
  1. const zip = await JSZip.loadAsync(pptxBuffer)
  2. Find last slide: zip.file(/ppt\/slides\/slide\d+\.xml/).sort().last()
  3. Read existing slide XML
  4. Inject <p:sp> element from ghost-logo-left.xml into the slide's <p:spTree>
  5. Zip.generateAsync({ type: 'nodebuffer' })
  6. Return new buffer
}
```

The ghost-logo-left.xml (1097 lines) contains a complete `<p:sp>` with custom geometry path
data — the Banglalink swoosh shape. This is injected directly as OOXML into the closing slide.

## Sub-Agent Spawning (section-agent.js)

Port of built-in `section-agent.js` pattern:

```js
const AIbitat = require("../../../../../server/utils/agents/aibitat/index.js");
const webBrowsing = require("../../../../../server/utils/agents/aibitat/plugins/web-browsing");
const webScraping = require("../../../../../server/utils/agents/aibitat/plugins/web-scraping");

async function spawnSectionAgent(section, parentHandler) {
  const child = new AIbitat({
    provider: parentHandler.super.config?.provider,
    model: parentHandler.super.config?.model,
    handlerProps: { ...parentHandler.handlerProps },
    maxToolCalls: 5,
  });

  child.use(webBrowsing.plugin());
  child.use(webScraping.plugin());

  child.function({
    super: child,
    name: "submit-section-slides",
    description: "Submit rendered slide data for this section",
    parameters: {
      type: "object",
      properties: {
        slides: { type: "array", items: { type: "object" } }
      }
    },
    handler: (params) => {
      // Store slide data for orchestrator
      return JSON.stringify({ accepted: true, slideCount: params.slides.length });
    }
  });

  const messages = [
    { role: "system", content: buildSectionPrompt(section) },
    { role: "user", content: `Research and build slides for: ${section.title}` }
  ];

  return child.handleAsyncExecution(messages, child.functions, section.title);
}
```

## Community Hub Publishing (plugin.json)

```json
{
  "hubId": "banglalink-pptx",
  "displayName": "Banglalink PPTX Generator",
  "description": "Generate Banglalink-branded PowerPoint presentations...",
  "author": "AnythingLLM Community",
  "version": "1.0.0",
  "trigger": {
    "keywords": ["presentation", "ppt", "pptx", "slides", "powerpoint", "banglalink", "deck"],
    "description": "Generate a Banglalink-branded presentation"
  },
  "params": {
    "topic": { "type": "string", "required": true },
    "sections": { "type": "array", "required": false },
    "audience": { "type": "string", "required": false },
    "outputFileName": { "type": "string", "required": false }
  },
  "dependencies": {
    "playwright": "^1.61.0",
    "puppeteer-core": "^24.0.0"
  }
}
```

## Build Order (to implement)

1. **brand.js** — Brand tokens, base64 logo data, theme constants
2. **pptx/themes.js** — Color substitution maps, style generators
3. **pptx/shapes.js** — Reusable shapes: branded tables, bullet lists, headers, accent bars
4. **pptx/slides.js** — Per-type slide layout builders (title, content, divider, table, quote, closing)
5. **pptx/builder.js** — Orchestrator: create pptxgenjs instance, compose slides from data
6. **pptx/ghost-logo.js** — jszip post-processing for ghost logo XML injection
7. **html/renderer.js** — Playwright detection + system Chrome fallback + DOM extraction
8. **section-agent.js** — Child AIbitat spawner with web search/scraping tools
9. **utils.js** — Shared helpers, path constants, error handling
10. **handler.js** — Main entry point tying everything together
11. **plugin.json + package.json** — Skill metadata and dependencies
12. **Integration test** — Verify end-to-end with a real AnythingLLM instance

---

## Appendix: Presenton (github.com/presenton/presenton) Insights

Studied 2026-06-30 — 8.6k stars, Apache 2.0, TypeScript+Python AI presentation generator.

### How Presenton Generates PPTX (Different Approach)

Presenton does NOT use pptxgenjs. Their pipeline:

```
User prompt → FastAPI AI generation → Next.js renders slides as React TSX templates
→ Puppeteer screenshots each rendered slide → Python binary (PyInstaller) assembles
PPTX with python-pptx using screenshots as slide background images
```

Their PPTX pipeline has three tiers:
1. **Next.js** (TypeScript/React 19): TSX component templates per slide type, Zod schemas define structured data fields with validation, constraints, and AI-facing descriptions
2. **Puppeteer/Playwright** (Node.js): Renders the React template as HTML+CSS, takes full-size screenshots
3. **PyInstaller binary** (separate private repo): Uses `python-pptx` to place screenshot images as slide backgrounds, handles font embedding

### Key Insights for Our Skill

| Insight | What Presenton Does | What We Should Do |
|---------|--------------------|-------------------|
| **Template Schema Pattern** | Each slide type = Zod schema defining fields + constraints + defaults + AI descriptions (`.meta({ layoutId, displayName, description })`) | Adopt simplified schema-per-slide pattern in `slides.js`. Each layout exports field definitions the sub-agent uses to generate structured content |
| **Image-Based PPTX** | All slides = full-page background images (screenshots of React renders). Pros: pixel-perfect. Cons: no text selection/editing in PPTX | **We stay with native pptxgenjs shapes** — text remains editable. Only use Playwright for specific HTML-rich elements (complex tables, charts) |
| **Oklch Color Generation** | Dynamically generates color palettes using Oklch color space | **Not needed** — our brand colors are fixed. Presenton's approach validates that theme-driven palette generation is a solved problem |
| **Zod + AI Content Generation** | AI generates content that matches the Zod schema structure, which is then rendered by the React template component | We can adapt: define a JSON schema per slide type → sub-agent generates content matching schema → pptxgenjs renders it natively |
| **Playwright Confirmation** | Uses Puppeteer/Playwright in the export pipeline to render HTML → screenshot | Validates our Playwright approach is industry-standard for AI PPTX generation |
| **python-pptx Choice** | Chose python-pptx over pptxgenjs for the final assembly (compiled as PyInstaller binary) | Reinforces that pptxgenjs (our choice) is fine for native-shape approach, since Presenton uses python-pptx mainly because their whole pipeline is Python-based |

### Template Schema Pattern (Adaptable)

Presenton's Zod-based per-slide-type schema pattern, simplified for our use:

```js
// Our equivalent of Presenton's Zod schema — plain JSON
const titleSlideSchema = {
  type: 'title',
  fields: {
    title:     { type: 'string', min: 3, max: 100, default: '', description: 'Main presentation title' },
    subtitle:  { type: 'string', min: 0, max: 200, default: '', description: 'Optional subtitle or tagline' },
    presenter: { type: 'string', min: 0, max: 100, default: '', description: 'Presenter name' },
    date:      { type: 'date', default: '', description: 'Presentation date' },
  }
};
```

The sub-agent receives the schema with field descriptions and generates structured content. The orchestrator then passes the content to the appropriate `slides.js` layout function.

### What NOT to Copy
- PyInstaller compilation (we're pure JS)
- React TSX template rendering (we use pptxgenjs native shapes)
- Oklch dynamic color generation (brand colors are fixed)
- `python-pptx` (we use pptxgenjs)
- Separate binary for export (our skill runs in-process)
