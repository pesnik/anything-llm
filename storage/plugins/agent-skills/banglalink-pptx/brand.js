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

function loadJSON(filename) {
  const filePath = path.join(ASSETS_DIR, filename);
  if (!fs.existsSync(filePath)) return null;
  try {
    return JSON.parse(fs.readFileSync(filePath, "utf-8"));
  } catch {
    return null;
  }
}

const tokens = loadJSON("brand-tokens.json");

module.exports = {
  colors: {
    primaryOrange: (tokens?.colors?.primary_orange || "#F56F1F").replace("#", ""),
    swooshOrange: (tokens?.colors?.swoosh_orange || "#EF6E23").replace("#", ""),
    deepOrange: (tokens?.colors?.deep_orange || "#F86C02").replace("#", ""),
    amber: (tokens?.colors?.amber || "#FAA106").replace("#", ""),
    checkmarkAmber: (tokens?.colors?.checkmark_amber || "#F6A11B").replace("#", ""),
    logoOrange: (tokens?.colors?.logo_orange || "#F26E21").replace("#", ""),
    logoAmber: (tokens?.colors?.logo_amber || "#F9A11B").replace("#", ""),
    nearBlack: (tokens?.colors?.near_black || "#171616").replace("#", ""),
    dark: (tokens?.colors?.dark || "#1C1C1C").replace("#", ""),
    midGray: (tokens?.colors?.mid_gray || "#8B8A8A").replace("#", ""),
    dividerGray: (tokens?.colors?.divider_gray || "#D9D9D9").replace("#", ""),
    rowAlt: (tokens?.colors?.row_alt || "#BFBFBF").replace("#", ""),
    lightGray: (tokens?.colors?.light_gray || "#F2F2F2").replace("#", ""),
    yellowHighlight: (tokens?.colors?.yellow_highlight || "#FFF2CC").replace("#", ""),
    white: (tokens?.colors?.white || "#FFFFFF").replace("#", ""),
    chartBlue: (tokens?.colors?.chart_blue || "#4472C4").replace("#", ""),
  },

  fonts: {
    body: tokens?.typography?.primary_font || "Arial",
    heading: tokens?.typography?.primary_font || "Arial",
    mono: "Courier New",
  },

  fontSize: {
    heroTitle: tokens?.typography?.roles?.hero_title || 44,
    sectionTitle: tokens?.typography?.roles?.section_title || 36,
    slideTitle: tokens?.typography?.roles?.slide_title || 28,
    subtitle: tokens?.typography?.roles?.subtitle || 24,
    bodyLarge: tokens?.typography?.roles?.body_large || 20,
    body: tokens?.typography?.roles?.body || 18,
    bodySmall: tokens?.typography?.roles?.body_small || 16,
    caption: tokens?.typography?.roles?.caption || 14,
    footer: tokens?.typography?.roles?.footer || 12,
    footnote: tokens?.typography?.roles?.footnote || 11,
  },

  slideSize: {
    width: tokens?.slide_size?.width_in || 13.333,
    height: tokens?.slide_size?.height_in || 7.5,
  },

  table: {
    headerFill: (tokens?.table?.header_fill || "#FAA106").replace("#", ""),
    headerText: (tokens?.table?.header_text || "#FFFFFF").replace("#", ""),
    headerFontBold: tokens?.table?.header_font_bold ?? true,
    rowFill: (tokens?.table?.row_fill || "#FFFFFF").replace("#", ""),
    rowAltFill: (tokens?.table?.row_alt_fill || "#F2F2F2").replace("#", ""),
    noBorders: tokens?.table?.no_borders ?? true,
  },

  logo: {
    logoFullWhite: loadLogoBase64("logo-full-white.png"),
    logoFull: loadLogoBase64("logo-full.png"),
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

  kpiSlide: {
    cardW: 2.8,
    cardH: 2.5,
    cardGap: 0.3,
    cardY: 2.0,
    numberFontSize: 36,
  },

  imageCards: {
    cardW: 3.5,
    cardH: 3.0,
    cardGap: 0.3,
    cardY: 1.8,
    captionH: 0.6,
  },

  timeline: {
    colW: 1.8,
    colGap: 0.2,
    colY: 1.8,
    iconSize: 0.5,
    colH: 3.5,
  },

  // ── Color palette weight hierarchy (60/30/10 rule) ──
  palette: {
    // 60% dominant: backgrounds, large surfaces
    dominant: {
      light: (tokens?.colors?.white || "#FFFFFF").replace("#", ""),
      dark: (tokens?.colors?.dark || "#1C1C1C").replace("#", ""),
    },
    // 30% secondary: headers, cards, secondary elements
    secondary: {
      primary: (tokens?.colors?.swoosh_orange || "#EF6E23").replace("#", ""),
      amber: (tokens?.colors?.amber || "#FAA106").replace("#", ""),
    },
    // 10% accent: highlights, CTAs, small emphasis
    accent: {
      deep: (tokens?.colors?.deep_orange || "#F86C02").replace("#", ""),
      highlight: (tokens?.colors?.yellow_highlight || "#FFF2CC").replace("#", ""),
    },
    // Text hierarchy
    text: {
      primary: (tokens?.colors?.near_black || "#171616").replace("#", ""),
      secondary: (tokens?.colors?.mid_gray || "#8B8A8A").replace("#", ""),
      muted: (tokens?.colors?.divider_gray || "#D9D9D9").replace("#", ""),
      inverse: (tokens?.colors?.white || "#FFFFFF").replace("#", ""),
    },
  },

  // ── Typography hierarchy (Anthropic-inspired) ──
  typography: {
    // Display: hero titles, big numbers
    display: { size: 48, weight: "bold", tracking: -0.5 },
    // Heading: slide titles, section headers
    heading: { size: 28, weight: "bold", tracking: 0 },
    // Subhead: subtitles, card headers
    subhead: { size: 18, weight: "bold", tracking: 0 },
    // Body: paragraph text
    body: { size: 14, weight: "normal", tracking: 0 },
    // Caption: labels, metadata
    caption: { size: 12, weight: "normal", tracking: 0 },
    // Micro: footnotes, fine print
    micro: { size: 10, weight: "normal", tracking: 0 },
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
