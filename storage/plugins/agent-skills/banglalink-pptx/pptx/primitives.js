/**
 * Primitives Toolkit — Raw building blocks for slide generation.
 * 
 * The agent uses these primitives to write custom slide code.
 * Each primitive wraps pptxgenjs with brand constraints.
 * 
 * Pattern from: artifact-kit/pptxgenjs-jsx (component model)
 */

const brand = require("../brand");

// ── Color primitives ──
const colors = {
  // Primary palette
  orange: brand.colors.swooshOrange,
  deepOrange: brand.colors.deepOrange,
  amber: brand.colors.amber,
  primaryOrange: brand.colors.primaryOrange,
  
  // Neutrals
  white: brand.colors.white,
  lightGray: brand.colors.lightGray,
  midGray: brand.colors.midGray,
  darkGray: "4A4A4A",
  nearBlack: brand.colors.nearBlack,
  dark: brand.colors.dark,
  
  // Functional
  divider: brand.colors.dividerGray,
  rowAlt: brand.colors.rowAlt,
  highlight: brand.colors.yellowHighlight,
  
  // Chart palette
  chart: [brand.colors.swooshOrange, brand.colors.deepOrange, brand.colors.amber, brand.colors.primaryOrange, brand.colors.checkmarkAmber],
};

// ── Typography primitives ──
const typography = {
  fonts: {
    heading: brand.fonts.heading,
    body: brand.fonts.body,
    mono: brand.fonts.mono,
  },
  
  // Size hierarchy (Anthropic pattern)
  sizes: {
    display: 48,    // Hero titles, big numbers
    title: 28,      // Slide titles
    subtitle: 18,   // Subtitles
    body: 14,       // Paragraph text
    caption: 12,    // Labels
    micro: 10,      // Footnotes
  },
  
  // Pre-defined text styles
  styles: {
    heroTitle: {
      fontSize: 48,
      bold: true,
      color: colors.orange,
      fontFace: brand.fonts.heading,
    },
    slideTitle: {
      fontSize: 28,
      bold: true,
      color: colors.orange,
      fontFace: brand.fonts.heading,
    },
    subtitle: {
      fontSize: 18,
      color: colors.midGray,
      fontFace: brand.fonts.body,
    },
    body: {
      fontSize: 14,
      color: colors.nearBlack,
      fontFace: brand.fonts.body,
    },
    caption: {
      fontSize: 12,
      color: colors.midGray,
      fontFace: brand.fonts.body,
    },
    bigNumber: {
      fontSize: 48,
      bold: true,
      color: colors.orange,
      fontFace: brand.fonts.heading,
      align: "center",
    },
    label: {
      fontSize: 12,
      color: colors.midGray,
      fontFace: brand.fonts.body,
      align: "center",
    },
  },
};

// ── Spacing primitives ──
const spacing = {
  slideMargin: 0.7,
  contentWidth: 11.933,
  slideWidth: brand.slideSize.width,
  slideHeight: brand.slideSize.height,
  
  // Vertical positions
  titleY: 0.4,
  titleHeight: 0.65,
  subtitleY: 1.0,
  contentY: 1.5,
  
  // Card spacing
  cardGap: 0.3,
  cardPadding: 0.2,
};

// ── Shape primitives ──
function addRect(slide, pptx, { x, y, w, h, fill, line, shadow, radius }) {
  const opts = { x, y, w, h };
  if (fill) opts.fill = { color: fill };
  if (line) opts.line = { color: line.color || fill, width: line.width || 1 };
  if (shadow) opts.shadow = shadow;
  if (radius) opts.rectRadius = radius;
  slide.addShape(pptx.shapes.RECTANGLE, opts);
}

function addCircle(slide, pptx, { x, y, size, fill, line }) {
  const opts = { x, y, w: size, h: size };
  if (fill) opts.fill = { color: fill };
  if (line) opts.line = { color: line.color || fill, width: line.width || 1 };
  slide.addShape(pptx.shapes.OVAL, opts);
}

function addLine(slide, pptx, { x, y, w, color, width: lineWidth, dash }) {
  slide.addShape(pptx.shapes.LINE, {
    x, y, w, h: 0,
    line: { color: color || colors.divider, width: lineWidth || 1, dashType: dash || "solid" },
  });
}

// ── Text primitives ──
function addText(slide, text, { x, y, w, h, style, align, valign, wrap, margin }) {
  const opts = { x, y, w, h, ...(typography.styles[style] || {}), text };
  if (align) opts.align = align;
  if (valign) opts.valign = valign;
  if (wrap !== undefined) opts.wrap = wrap;
  if (margin !== undefined) opts.margin = margin;
  slide.addText(text, opts);
}

function addRichText(slide, runs, { x, y, w, h, align, valign }) {
  const opts = { x, y, h, align: align || "left", valign: valign || "top" };
  // runs is array of { text, options: { bold, color, fontSize, ... } }
  slide.addText(runs, opts);
}

// ── Image primitives ──
function addImage(slide, { data, path, x, y, w, h, sizing }) {
  const opts = { x, y, w, h };
  if (data) opts.data = data;
  if (path) opts.path = path;
  if (sizing) opts.sizing = sizing;
  slide.addImage(opts);
}

// ── Chart primitives ──
function addChart(slide, pptx, { type, data, x, y, w, h, options }) {
  const chartType = pptx.charts[type.toUpperCase()] || pptx.charts.BAR;
  const opts = { x, y, w, h, chartColors: colors.chart, ...options };
  slide.addChart(chartType, data, opts);
}

// ── Table primitives ──
function addTable(slide, { headers, rows, x, y, w, colW, headerStyle, rowStyle }) {
  const tableData = [];
  
  if (headers) {
    tableData.push(headers.map(h => ({
      text: h,
      options: {
        bold: true,
        fontSize: 14,
        fontFace: brand.fonts.body,
        color: colors.white,
        fill: { color: colors.amber },
        align: "left",
        valign: "middle",
        margin: [4, 8, 4, 8],
        ...headerStyle,
      },
    })));
  }
  
  if (rows) {
    rows.forEach((row, idx) => {
      tableData.push(row.map(cell => ({
        text: String(cell),
        options: {
          fontSize: 13,
          fontFace: brand.fonts.body,
          color: colors.nearBlack,
          fill: { color: idx % 2 === 1 ? colors.lightGray : colors.white },
          align: "left",
          valign: "middle",
          margin: [4, 8, 4, 8],
          ...rowStyle,
        },
      })));
    });
  }
  
  const colCount = tableData[0]?.length || 1;
  slide.addTable(tableData, {
    x,
    y,
    w,
    colW: colW || w / colCount,
    rowH: 0.4,
    border: { type: "solid", pt: 0.5, color: colors.divider },
  });
}

// ── Bullet list primitive ──
function addBullets(slide, items, { x, y, w, h, style }) {
  const runs = items.map(text => ({
    text,
    options: {
      bullet: { characterCode: "2713" },
      breakLine: true,
      ...(typography.styles[style] || typography.styles.body),
    },
  }));
  
  slide.addText(runs, { x, y, w, h, valign: "top" });
}

// ── Logo primitive ──
function addLogo(slide, { position = "top-right" }) {
  const logo = brand.logo.logoIcon;
  if (!logo) return;
  
  const positions = {
    "top-right": { x: 12.417, y: 0.06, w: 0.669, h: 0.742 },
    "top-left": { x: 0.7, y: 0.06, w: 0.669, h: 0.742 },
    "bottom-left": { x: 0.7, y: 6.7, w: 0.669, h: 0.742 },
    "center": { x: 5.83, y: 3.38, w: 1.67, h: 1.86 },
  };
  
  const pos = positions[position] || positions["top-right"];
  slide.addImage({ data: logo, ...pos });
}

// ── Export all primitives ──
module.exports = {
  colors,
  typography,
  spacing,
  
  // Shape primitives
  addRect,
  addCircle,
  addLine,
  
  // Text primitives
  addText,
  addRichText,
  
  // Media primitives
  addImage,
  addChart,
  addTable,
  addBullets,
  
  // Brand primitives
  addLogo,
  
  // Raw access to pptxgenjs
  raw: {
    addShape: (slide, pptx, type, opts) => slide.addShape(type, opts),
    addText: (slide, text, opts) => slide.addText(text, opts),
    addImage: (slide, opts) => slide.addImage(opts),
    addChart: (slide, type, data, opts) => slide.addChart(type, data, opts),
    addTable: (slide, rows, opts) => slide.addTable(rows, opts),
  },
};
