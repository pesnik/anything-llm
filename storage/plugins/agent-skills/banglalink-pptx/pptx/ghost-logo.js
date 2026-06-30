const JSZip = require("jszip");
const brand = require("../brand");

async function injectGhostLogo(buffer, sections) {
  if (!brand.ghostLogoXML) {
    return buffer;
  }

  const zip = await JSZip.loadAsync(buffer);

  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/)[1], 10);
      const nb = parseInt(b.match(/slide(\d+)/)[1], 10);
      return na - nb;
    });

  if (slideFiles.length === 0) return buffer;

  // Find section-break slides and closing slide
  const sectionBreakSlides = [];
  let slideIdx = 0;
  for (const section of sections) {
    for (const slideData of (section.slides || [])) {
      if (slideData.type === "section-break") {
        sectionBreakSlides.push(slideIdx);
      }
      slideIdx++;
    }
  }

  // Inject ghost logo into section-break slides and closing slide
  const slidesToInject = [...sectionBreakSlides, slideFiles.length - 1];
  const uniqueSlides = [...new Set(slidesToInject)];

  for (const slideNum of uniqueSlides) {
    if (slideNum >= slideFiles.length) continue;
    const slideFile = slideFiles[slideNum];
    const slideXml = await zip.file(slideFile).async("string");

    const spTreeMatch = slideXml.match(/<p:spTree[\s\S]*?<\/p:spTree>/);
    if (!spTreeMatch) continue;

    const spTree = spTreeMatch[0];
    const closeTag = "</p:spTree>";
    const insertPos = spTree.lastIndexOf(closeTag);

    if (insertPos === -1) continue;

    const modifiedSpTree =
      spTree.slice(0, insertPos) + "\n" + brand.ghostLogoXML + "\n" + spTree.slice(insertPos);

    const modifiedSlideXml = slideXml.replace(spTree, modifiedSpTree);
    zip.file(slideFile, modifiedSlideXml);
  }

  const finalBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  return finalBuffer;
}

module.exports = { injectGhostLogo };
