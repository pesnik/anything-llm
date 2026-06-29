const PptxGenJS = require("pptxgenjs");
const brand = require("../brand");
const slides = require("./slides");
const ghostLogo = require("./ghost-logo");

async function buildPresentation(sections, options) {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "BL_WIDESCREEN", width: brand.slideSize.width, height: brand.slideSize.height });
  pptx.layout = "BL_WIDESCREEN";

  const totalSlides = sections.reduce((sum, s) => sum + (s.slides ? s.slides.length : 1), 0);
  let slideNumber = 0;

  for (let sectionIdx = 0; sectionIdx < sections.length; sectionIdx++) {
    const section = sections[sectionIdx];
    const slideDataList = section.slides || [section];

    for (let slideIdx = 0; slideIdx < slideDataList.length; slideIdx++) {
      const data = slideDataList[slideIdx];
      slideNumber++;

      const slide = pptx.addSlide();

      switch (data.type) {
        case "title":
          slides.renderTitleSlide(slide, pptx, data);
          break;
        case "section":
          slides.renderSectionSlide(slide, pptx, data);
          break;
        case "bullet":
          slides.renderBulletSlide(slide, pptx, data);
          break;
        case "two-column":
        case "twocolumn":
          slides.renderTwoColumnSlide(slide, pptx, data);
          break;
        case "table":
          slides.renderTableSlide(slide, pptx, data);
          break;
        case "quote":
          slides.renderQuoteSlide(slide, pptx, data);
          break;
        case "closing":
          slides.renderClosingSlide(slide, pptx, data);
          break;
        case "kpi":
          slides.renderKpiSlide(slide, pptx, data);
          break;
        case "image-cards":
          slides.renderImageCardsSlide(slide, pptx, data);
          break;
        case "chart":
          slides.renderChartSlide(slide, pptx, data);
          break;
        case "timeline":
          slides.renderTimelineSlide(slide, pptx, data);
          break;
        case "bullet-image":
          slides.renderBulletImageSlide(slide, pptx, data);
          break;
        case "section-break":
          slides.renderSectionBreakSlide(slide, pptx, data);
          break;
        case "section-intro":
          slides.renderSectionIntroSlide(slide, pptx, data);
          break;
        default:
          slides.renderBulletSlide(slide, pptx, data);
          break;
      }

      if (data.notes) {
        slide.addNotes(data.notes);
      }
    }
  }

  let buffer = await pptx.write({ outputType: "nodebuffer" });

  buffer = await ghostLogo.injectGhostLogo(buffer, sections);

  return buffer;
}

module.exports = { buildPresentation };
