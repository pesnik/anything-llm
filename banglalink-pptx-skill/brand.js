const path = require("path");
const fs = require("fs");

const ASSETS_DIR = path.join(__dirname, "assets");

function loadLogoBase64(filename) {
  const filePath = path.join(ASSETS_DIR, filename);
  if (!fs.existsSync(filePath)) return null;
  const ext = path.extname(filename).slice(1);
  const mime = ext === "svg" ? "image/svg+xml" : `image/${ext}`;
  const data = fs.readFileSync(filePath);
  return `data:${mime};base64,${data.toString("base64")}`;
}

function loadFile(filename) {
  const filePath = path.join(ASSETS_DIR, filename);
  if (!fs.existsSync(filePath)) return null;
  return fs.readFileSync(filePath, "utf-8");
}

module.exports = {
  colors: {
    primaryOrange: "F56F1F",
    deepOrange: "F86C02",
    amber: "FAA106",
    checkmarkAmber: "F6A11B",
    logoOrange: "F26E21",
    logoAmber: "F9A11B",
    nearBlack: "171616",
    dark: "1C1C1C",
    midGray: "8B8A8A",
    dividerGray: "D9D9D9",
    rowAlt: "BFBFBF",
    lightGray: "F2F2F2",
    yellowHighlight: "FFF2CC",
    white: "FFFFFF",
    chartBlue: "4472C4",
  },

  fonts: {
    body: "Arial",
    heading: "Arial",
    mono: "Courier New",
  },

  fontSize: {
    heroTitle: 44,
    sectionTitle: 36,
    slideTitle: 28,
    subtitle: 24,
    bodyLarge: 20,
    body: 18,
    bodySmall: 16,
    caption: 14,
    footer: 12,
    footnote: 11,
  },

  slideSize: {
    width: 13.333,
    height: 7.5,
  },

  table: {
    headerFill: "FAA106",
    headerText: "FFFFFF",
    headerFontBold: true,
    rowFill: "FFFFFF",
    rowAltFill: "F2F2F2",
    noBorders: true,
  },

  logo: {
    logoFullWhite: loadLogoBase64("logo-full-white.png"),
    logoIcon: loadLogoBase64("logo-icon.png"),
  },

  ghostLogoXML: loadFile("ghost-logo-left.xml"),

  titleSlide: {
    titleX: 1.0,
    titleY: 0.8,
    titleW: 11.333,
    titleH: 2.0,
    subtitleX: 1.0,
    subtitleY: 2.8,
    subtitleW: 11.333,
    subtitleH: 1.0,
    logoFullLeft: 5.107,
    logoFullTop: 1.9,
    logoFullWidth: 3.12,
    logoFullHeight: 2.37,
  },

  contentSlide: {
    marginX: 0.7,
    marginY: 0.4,
    contentW: 11.933,
    logoIconLeft: 12.417,
    logoIconTop: 0.06,
    logoIconWidth: 0.669,
    logoIconHeight: 0.742,
    headerHeight: 0.8,
    footerHeight: 0.3,
  },

  sectionDivider: {
    orangePanelW: 4.444,
    titleX: 5.333,
    titleY: 2.5,
    titleW: 7.0,
    titleH: 2.0,
  },

  closing: {
    ghostLeft: 0,
    ghostTop: 0.28,
    ghostWidth: 3.81,
    ghostHeight: 7.06,
  },

  isDarkColor(hexColor) {
    const hex = (hexColor || "FFFFFF").replace("#", "");
    const r = parseInt(hex.substr(0, 2), 16);
    const g = parseInt(hex.substr(2, 2), 16);
    const b = parseInt(hex.substr(4, 2), 16);
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.5;
  },

  hexToRGB(hex) {
    const h = hex.replace("#", "");
    return {
      r: parseInt(h.substr(0, 2), 16),
      g: parseInt(h.substr(2, 2), 16),
      b: parseInt(h.substr(4, 2), 16),
    };
  },
};
