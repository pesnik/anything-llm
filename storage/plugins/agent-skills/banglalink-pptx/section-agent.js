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

## Slide Types & When to Use Each

### Content Slides (use these for most sections)
- "bullet" - Bullet list with checkmarks. Use for: key points, features, lists, summaries. Max 6 bullets.
- "two-column" - Side-by-side comparison. Use for: comparing two options, pros/cons, before/after.
- "table" - Data table with headers. Use for: structured data, financial data, comparisons with numbers.
- "kpi" - Dashboard with 4 metric cards. Use for: key performance indicators, metrics, statistics, scores.
- "chart" - Native chart (line/bar/pie). Use for: trends, time series, market share, data visualization.
- "bullet-image" - Bullet list + image. Use for: explaining a concept with visual context.
- "image-cards" - 3-4 image cards. Use for: product comparison, feature categories, visual gallery.

### Structure Slides (use sparingly)
- "title" - Opening slide with centered logo. Use only for: presentation opening.
- "section-intro" - Clean section intro with subtitle. Use for: introducing a new topic section.
- "section-break" - Full orange divider. Use for: major transitions between parts.
- "section" - Section divider with two-column lists. Use for: detailed section overviews.
- "quote" - Quote with attribution. Use for: customer testimonials, executive quotes.
- "closing" - Thank You slide. Use only for: presentation ending.

## Layout Selection Rules
1. NEVER use more than 2 bullet slides in a row — vary the layout
2. If section has 3-6 key metrics → use "kpi"
3. If section has time-based data → use "chart" (line)
4. If section compares 2 things → use "two-column"
5. If section has structured data → use "table"
6. If section is a major transition → use "section-break" or "section-intro"
7. If section has 3-4 items to showcase → use "image-cards"

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
