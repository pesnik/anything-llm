const path = require("path");
const fs = require("fs");

const brand = require("./brand");
const builder = require("./pptx/builder");
const sectionAgent = require("./section-agent");
const utils = require("./utils");

const CREATE_FILES_LIB_PATH = path.join(
  __dirname, "..", "..", "..", "..", "server", "utils", "agents", "aibitat",
  "plugins", "create-files", "lib.js"
);

function getSaveFunction() {
  try {
    const lib = require(CREATE_FILES_LIB_PATH);
    if (typeof lib.saveGeneratedFile === "function") return lib.saveGeneratedFile;
    if (typeof lib.saveFile === "function") return lib.saveFile;
  } catch (e) {}

  return async function saveGeneratedFileFallback(fileName, buffer) {
    const dir = utils.ensureGeneratedDir();
    const filePath = path.join(dir, fileName);
    fs.writeFileSync(filePath, buffer);
    return `/api/files/generated/${fileName}`;
  };
}

function inferSections(topic, audience) {
  return [
    {
      title: topic,
      slides: [
        {
          type: "title",
          title: topic,
          subtitle: audience ? `For ${audience}` : "",
        },
      ],
    },
    {
      title: "Overview",
      slides: [
        {
          type: "section",
          title: "Overview",
          subtitle: "Key highlights and context",
        },
        {
          type: "bullet",
          title: "Introduction",
          content: [
            `Overview of ${topic}`,
            "Key objectives and goals",
            "Strategic importance",
          ],
        },
      ],
    },
    {
      title: "Details",
      slides: [
        {
          type: "bullet",
          title: "Key Points",
          content: [
            "Detail point one",
            "Detail point two",
            "Detail point three",
          ],
        },
      ],
    },
    {
      title: "Summary",
      slides: [
        {
          type: "closing",
          cta: "Thank You",
          contact: audience || "",
        },
      ],
    },
  ];
}

async function handler({ handlerProps, logger, super: superAgent }) {
  const context = { handlerProps, logger, super: superAgent };

  const params = handlerProps?.params || {};
  const topic = params.topic || "Presentation";
  const audience = params.audience || "";
  const outputFileName = params.outputFileName || utils.slugify(topic);
  const userSections = params.sections || [];
  const theme = params.theme || "dark";
  const enableResearch = params.enableResearch !== false;

  logger?.info?.(`Banglalink PPTX Generator: generating "${topic}"`);

  const sections = [];

  if (userSections.length > 0) {
    for (const section of userSections) {
      let slides;

      if (enableResearch && section.topic) {
        logger?.info?.(`Researching section: ${section.title || section.topic}`);
        try {
          slides = await sectionAgent.spawnSectionAgent(section, context);
        } catch (e) {
          logger?.warn?.(`Research failed for "${section.title}", using defaults`);
          slides = null;
        }
      }

      if (!slides) {
        slides = [
          {
            type: section.isTitle ? "title" : section.isClosing ? "closing" : "bullet",
            title: section.title || section.topic,
            subtitle: section.subtitle,
            content: section.bulletPoints || ["Content placeholder"],
            cta: section.isClosing ? "Thank You" : undefined,
          },
        ];
      }

      sections.push({
        title: section.title || section.topic,
        slides,
      });
    }
  } else {
    logger?.info?.(`No sections provided — auto-generating structure for "${topic}"`);

    if (enableResearch) {
      const overviewSection = {
        title: "Overview",
        topic,
        description: `Overview of ${topic}`,
        audience,
      };
      try {
        const overviewSlides = await sectionAgent.spawnSectionAgent(overviewSection, context);
        sections.push({ title: "Overview", slides: overviewSlides });
      } catch (e) {
        sections.push({
          title: "Overview",
          slides: [{ type: "bullet", title: `About ${topic}`, content: ["Key information"] }],
        });
      }
    } else {
      const inferredSections = inferSections(topic, audience);
      for (const s of inferredSections) {
        sections.push(s);
      }
    }
  }

  logger?.info?.(`Building PPTX with ${sections.length} sections...`);

  let buffer;
  try {
    buffer = await builder.buildPresentation(sections, { theme });
  } catch (e) {
    logger?.error?.(`PPTX build failed: ${e.message}`);
    return `Error: PPTX generation failed — ${e.message}`;
  }

  const fileName = `${outputFileName}.pptx`;

  try {
    const saveFile = getSaveFunction();
    const fileUrl = await saveFile(fileName, Buffer.from(buffer));
    logger?.info?.(`PPTX saved: ${fileUrl}`);

    const slideCount = sections.reduce((count, s) => count + (s.slides?.length || 0), 0);

    return JSON.stringify({
      success: true,
      message: `Presentation "${topic}" generated successfully`,
      file: fileUrl,
      fileName,
      slideCount,
      sections: sections.length,
    });
  } catch (e) {
    logger?.error?.(`Failed to save PPTX: ${e.message}`);
    return `Error: Failed to save presentation — ${e.message}`;
  }
}

module.exports = { handler, runtime: { handler } };
