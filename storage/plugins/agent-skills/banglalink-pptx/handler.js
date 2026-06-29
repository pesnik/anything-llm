const path = require("path");
const fs = require("fs");
const builder = require("./pptx/builder");
const sectionAgent = require("./section-agent");
const utils = require("./utils");

const createFilesLib = require(path.join(
  __dirname, "..", "..", "..", "..",
  "server", "utils", "agents", "aibitat",
  "plugins", "create-files", "lib.js"
));

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

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

async function handler(args) {
  console.log(`[banglalink-pptx-handler] STARTED at ${Date.now()} pid=${process.pid}`);
  const superAgent = this.super;
  const logger = this.logger;
  if (typeof logger === "function") logger("[banglalink-pptx-handler] Handler invoked");

  const params = args || {};
  const topic = params.topic || "Presentation";
  const audience = params.audience || "";
  const outputFileName = params.outputFileName || utils.slugify(topic);
  const userSections = params.sections || [];
  const theme = params.theme || "dark";
  const enableResearch = params.enableResearch !== false;

  if (typeof logger === "function")
    logger(`[banglalink-pptx] Generating "${topic}"`);

  const sections = [];

  if (userSections.length > 0) {
    for (const section of userSections) {
      let slides;
      if (enableResearch && section.topic) {
        try {
          slides = await sectionAgent.spawnSectionAgent(section, {
            handlerProps: superAgent.handlerProps,
            logger,
            super: superAgent,
          });
        } catch (e) {
          slides = null;
        }
      }
      if (!slides) {
        slides = [{
          type: section.isTitle ? "title" : section.isClosing ? "closing" : "bullet",
          title: section.title || section.topic,
          subtitle: section.subtitle,
          content: section.bulletPoints || ["Content placeholder"],
          cta: section.isClosing ? "Thank You" : undefined,
        }];
      }
      sections.push({ title: section.title || section.topic, slides });
    }
  } else {
    if (enableResearch) {
      const overviewSection = { title: "Overview", topic, description: `Overview of ${topic}`, audience };
      try {
        const overviewSlides = await sectionAgent.spawnSectionAgent(overviewSection, {
          handlerProps: superAgent.handlerProps, logger, super: superAgent,
        });
        sections.push({ title: "Overview", slides: overviewSlides });
      } catch (e) {
        sections.push({ title: "Overview", slides: [{ type: "bullet", title: `About ${topic}`, content: ["Key information"] }] });
      }
    } else {
      for (const s of inferSections(topic, audience)) sections.push(s);
    }
  }

  console.log(`[banglalink-pptx-handler] Building with ${sections.length} sections`);
  let buffer;
  try {
    buffer = await builder.buildPresentation(sections, { theme });
  } catch (e) {
    console.log(`[banglalink-pptx-handler] Build ERROR: ${e.message}`);
    return `Error: PPTX generation failed — ${e.message}`;
  }

  const slideCount = sections.reduce((count, s) => count + (s.slides?.length || 0), 0);
  const displayFilename = `${outputFileName}.pptx`;

  try {
    const savedFile = await createFilesLib.saveGeneratedFile({
      fileType: "pptx", extension: "pptx", buffer: Buffer.from(buffer), displayFilename,
    });

    console.log(`[banglalink-pptx-handler] saveGeneratedFile returned: storagePath=${savedFile.storagePath}`);
    const fileExists = fs.existsSync(savedFile.storagePath);
    console.log(`[banglalink-pptx-handler] File exists on disk: ${fileExists}, size: ${fileExists ? fs.statSync(savedFile.storagePath).size : 0}`);

    createFilesLib.registerOutput(superAgent, "PptxFileDownload", {
      filename: savedFile.displayFilename, storageFilename: savedFile.filename, fileSize: savedFile.fileSize,
    });

    const downloadPath = `/agent-skills/generated-files/${savedFile.filename}`;

    superAgent?.socket?.send?.("fileDownloadCard", {
      filename: savedFile.displayFilename,
      storageFilename: savedFile.filename,
      fileSize: savedFile.fileSize,
    });

    superAgent?.introspect?.(
      `"${topic}" ready — download at ${downloadPath}`
    );

    return `Successfully created Banglalink-branded presentation "${topic}" with ${slideCount} slides using the ${theme} theme.`;
  } catch (e) {
    console.log(`[banglalink-pptx-handler] Save ERROR: ${e.message}`);
    return `Error: Failed to save presentation — ${e.message}`;
  }
}

module.exports = { handler, runtime: { handler } };
