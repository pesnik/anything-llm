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

  const closingSlideFile = slideFiles[slideFiles.length - 1];
  const slideXml = await zip.file(closingSlideFile).async("string");

  const spTreeMatch = slideXml.match(/<p:spTree[\s\S]*?<\/p:spTree>/);
  if (!spTreeMatch) return buffer;

  const spTree = spTreeMatch[0];
  const closeTag = "</p:spTree>";
  const insertPos = spTree.lastIndexOf(closeTag);

  if (insertPos === -1) return buffer;

  const modifiedSpTree =
    spTree.slice(0, insertPos) + "\n" + brand.ghostLogoXML + "\n" + spTree.slice(insertPos);

  const modifiedSlideXml = slideXml.replace(spTree, modifiedSpTree);

  zip.file(closingSlideFile, modifiedSlideXml);

  const finalBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  return finalBuffer;
}

module.exports = { injectGhostLogo };
