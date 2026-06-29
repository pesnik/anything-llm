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
                enum: ["title", "section", "bullet", "two-column", "table", "quote", "closing"],
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

Slide Types Available:
- "title" - Opening slide with title + subtitle
- "section" - Section divider slide
- "bullet" - Content slide with bullet list
- "two-column" - Two-column layout (left bullets, right text/bullets)
- "table" - Data table with headers and rows
- "quote" - Quote/Testimonial slide
- "closing" - Thank You / closing slide

Brand guidelines (strictly enforce):
- Primary color: Orange (#F56F1F)
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
