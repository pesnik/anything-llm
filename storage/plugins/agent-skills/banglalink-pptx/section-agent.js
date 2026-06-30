const path = require("path");

const AIBITAT_PATH = path.join(__dirname, "..", "..", "..", "..", "server", "utils", "agents", "aibitat", "index.js");

async function spawnSectionAgent(section, parentHandler) {
  let AIbitat;
  try {
    AIbitat = require(AIBITAT_PATH);
  } catch (e) {
    parentHandler?.logger?.warn?.("Cannot load AIbitat for sub-agent, using fallback");
    return generateFallbackSlides(section);
  }

  let webBrowsing, webScraping;
  try {
    webBrowsing = require(path.join(
      __dirname, "..", "..", "..", "..", "server", "utils", "agents", "aibitat", "plugins", "web-browsing"
    ));
    webScraping = require(path.join(
      __dirname, "..", "..", "..", "..", "server", "utils", "agents", "aibitat", "plugins", "web-scraping"
    ));
  } catch (e) {
    parentHandler?.logger?.warn?.("Cannot load web plugins for sub-agent");
  }

  const child = new AIbitat({
    provider: parentHandler?.super?.config?.provider,
    model: parentHandler?.super?.config?.model,
    handlerProps: parentHandler?.handlerProps || {},
    maxToolCalls: 5,
  });

  if (webBrowsing?.plugin) child.use(webBrowsing.plugin());
  if (webScraping?.plugin) child.use(webScraping.plugin());

  const slides = [];

  child.function({
    super: child,
    name: "submit-section-slides",
    description: "Submit rendered slide data for this section",
    parameters: {
      type: "object",
      properties: {
        slides: {
          type: "array",
          items: {
            type: "object",
            properties: {
              type: {
                type: "string",
                enum: [
                  "title", "section", "bullet", "two-column", "table", "quote", "closing",
                  "kpi", "image-cards", "chart", "timeline", "bullet-image",
                  "section-break", "section-intro"
                ],
              },
              title: { type: "string" },
              subtitle: { type: "string" },
              content: {
                type: "array",
                items: { type: "string" },
              },
              leftContent: { type: "array", items: { type: "string" } },
              rightContent: {},
              headers: { type: "array", items: { type: "string" } },
              rows: {
                type: "array",
                items: { type: "array", items: { type: "string" } },
              },
              quote: { type: "string" },
              attribution: { type: "string" },
              cta: { type: "string" },
              contact: { type: "string" },
              notes: { type: "string" },
              kpis: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    label: { type: "string" },
                    value: { type: "string" },
                    sublabel: { type: "string" },
                  },
                },
              },
              cards: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    caption: { type: "string" },
                  },
                },
              },
              chartType: {
                type: "string",
                enum: ["line", "bar", "pie", "doughnut"],
              },
              chartData: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    name: { type: "string" },
                    labels: { type: "array", items: { type: "string" } },
                    values: { type: "array", items: { type: "number" } },
                  },
                },
              },
              chartTitle: { type: "string" },
              months: {
                type: "array",
                items: {
                  type: "object",
                  properties: {
                    label: { type: "string" },
                    details: { type: "string" },
                  },
                },
              },
            },
          },
        },
      },
    },
    handler: (params) => {
      if (params.slides) {
        slides.push(...params.slides);
      }
      return JSON.stringify({ accepted: true, slideCount: params.slides?.length || 0 });
    },
  });

  const systemPrompt = buildSectionPrompt(section);
  const messages = [
    { role: "system", content: systemPrompt },
    {
      role: "user",
      content: `Research and build slides for the section: "${section.title || section.topic}". 
Generate 1-3 slides that cover the key points. Use the submit-section-slides tool to submit the slides.`,
    },
  ];

  try {
    await child.handleAsyncExecution(messages, child.functions, section.title || section.topic);
  } catch (e) {
    parentHandler?.logger?.error?.("Sub-agent error:", e.message);
    return generateFallbackSlides(section);
  }

  if (slides.length === 0) {
    return generateFallbackSlides(section);
  }

  return slides;
}

function buildSectionPrompt(section) {
  return `You are a presentation slide researcher and builder for Banglalink, a leading telecom operator in Bangladesh.

Your task is to research the given section topic using web search, then create slides using the submit-section-slides tool.

## Architecture: You Are the Slide Designer

You have access to a PRIMITIVES TOOLKIT and COMPONENT REGISTRY. Instead of picking from a fixed list of slide types, you WRITE the slide code using these building blocks.

## Primitives Toolkit

### Colors
- colors.orange = "EF6E23" (primary)
- colors.deepOrange = "F86C02"
- colors.amber = "FAA106" (accent)
- colors.white = "FFFFFF"
- colors.lightGray = "F2F2F2"
- colors.midGray = "8B8A8A"
- colors.nearBlack = "171616"
- colors.dark = "1C1C1C"
- colors.chart = ["EF6E23", "F86C02", "FAA106", "F56F1F", "F6A11B"]

### Typography
- typography.fonts.heading = "Arial"
- typography.fonts.body = "Arial"
- typography.styles.slideTitle = { fontSize: 28, bold: true, color: "EF6E23" }
- typography.styles.subtitle = { fontSize: 18, color: "8B8A8A" }
- typography.styles.body = { fontSize: 14, color: "171616" }
- typography.styles.bigNumber = { fontSize: 48, bold: true, color: "EF6E23", align: "center" }
- typography.styles.caption = { fontSize: 12, color: "8B8A8A" }

### Spacing
- spacing.slideMargin = 0.7
- spacing.contentWidth = 11.933
- spacing.slideWidth = 13.333
- spacing.slideHeight = 7.5
- spacing.titleY = 0.4
- spacing.subtitleY = 1.0
- spacing.contentY = 1.5

### Shape Primitives
- P.addRect(slide, pptx, { x, y, w, h, fill, line: { color, width } })
- P.addCircle(slide, pptx, { x, y, size, fill })
- P.addLine(slide, pptx, { x, y, w, color, width, dash })

### Text Primitives
- P.addText(slide, text, { x, y, w, h, style: "slideTitle", align, valign })
- P.addRichText(slide, runs, { x, y, w, h }) // runs = [{ text, options: { bold, color } }]

### Media Primitives
- P.addImage(slide, { data, x, y, w, h })
- P.addChart(slide, pptx, { type: "bar"|"line"|"pie"|"doughnut", data, x, y, w, h, options })
- P.addTable(slide, { headers, rows, x, y, w, colW })
- P.addBullets(slide, items, { x, y, w, h, style })
- P.addLogo(slide, { position: "top-right"|"top-left"|"bottom-left"|"center" })

## Component Registry (Composable Patterns)

### KPICards(slide, pptx, { metrics, y, cardW, cardH })
Creates 3-4 metric cards with big numbers.
- metrics = [{ label: "Market Share", value: "22%", sublabel: "of total" }]

### TwoColumn(slide, pptx, { leftTitle, left, rightTitle, right, y })
Side-by-side comparison with divider line.
- left/right = ["bullet point 1", "bullet point 2"]

### Timeline(slide, pptx, { milestones, y })
Horizontal timeline with orange circles.
- milestones = [{ label: "Q1", details: "Launch phase" }]

### ImageCards(slide, pptx, { cards, y })
3-4 cards with captions.
- cards = [{ caption: "Description text" }]

### SectionBreak(slide, { title, subtitle })
Full orange divider slide.

### StatCallout(slide, { value, label, x, y, w })
Big number + label at any position.

### Quote(slide, { text, attribution, y })
Quote with attribution.

### ClosingSlide(slide, { cta, contact })
Thank You slide with ghost logo.

## How to Write Slide Code

Example for a section with market share data:

\`\`\`javascript
const slide = addSlide();
Background(slide, { color: colors.white });
Title(slide, { text: "Market Share Analysis" });
Subtitle(slide, { text: "Q1 2025 Overview" });
KPICards(slide, pptx, {
  metrics: [
    { label: "Banglalink", value: "22%", sublabel: "Growing" },
    { label: "Grameenphone", value: "48%", sublabel: "Dominant" },
    { label: "Robi", value: "28%", sublabel: "Stable" },
  ]
});
LogoIcon(slide);
\`\`\`

Example for a comparison section:

\`\`\`javascript
const slide = addSlide();
Background(slide, { color: colors.white });
Title(slide, { text: "GP vs BL Comparison" });
TwoColumn(slide, pptx, {
  leftTitle: "Grameenphone",
  left: ["Curated & Segmented", "Sells Experience", "High lock-in"],
  rightTitle: "Banglalink",
  right: ["Cluttered & Transactional", "Sells Commodity", "Low lock-in"],
});
LogoIcon(slide);
\`\`\`

## Layout Selection Rules
1. NEVER use more than 2 bullet slides in a row — vary the layout
2. EVERY slide must have a visual element — KPICards, chart, TwoColumn, Timeline, ImageCards, or shapes
3. If section has 3-6 key metrics → use KPICards
4. If section has time-based data → use Timeline or chart
5. If section compares 2 things → use TwoColumn
6. If section has structured data → use P.addTable
7. If section is a major transition → use SectionBreak
8. If section has 3-4 items to showcase → use ImageCards
9. If section has a key stat → use StatCallout
10. Default to KPICards or chart over bullets — always prefer visual layouts

## Visual Design Principles
- Every slide needs a visual anchor: KPICards, chart, TwoColumn, Timeline, or ImageCards
- Use color dominance: 60-70% white, 20% orange accent, 10% dark text
- Leave breathing room — don't fill every inch with text
- Icons in colored circles next to headers add polish

## Typography Hierarchy (use these exact sizes)
- Display (hero titles, big numbers): 48pt bold
- Heading (slide titles): 28pt bold (use Title component)
- Subhead (subtitles): 18pt (use Subtitle component)
- Body (paragraph text): 14pt
- Caption (labels): 12pt
- Micro (footnotes): 10pt
- NEVER use body size for titles — titles must be 24pt+

## Brand Guidelines
- Primary colors: Orange (#EF6E23 for headers), Amber (#FAA106 for accents)
- Only use Arial font family
- Keep bullet text concise (6-8 max per slide)
- Professional tone suitable for telecom executive audience

Section context:
${section.description ? `Description: ${section.description}` : ""}
${section.topic ? `Topic: ${section.topic}` : ""}
${section.audience ? `Audience: ${section.audience}` : ""}`;
}

function generateFallbackSlides(section) {
  return [
    {
      type: "bullet",
      title: section.title || section.topic || "Section",
      content: [
        "Key point 1",
        "Key point 2",
        "Key point 3",
      ],
    },
  ];
}

module.exports = { spawnSectionAgent };
