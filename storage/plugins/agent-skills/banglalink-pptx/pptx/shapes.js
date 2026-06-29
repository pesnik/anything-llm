const brand = require("../brand");
const themes = require("./themes");

const ALLOWED_COLORS = new Set(Object.values(brand.colors));

// ── Factory functions (prevents PptxGenJS object mutation corruption) ──
function makeShadow(opts = {}) {
  return {
    type: opts.type || "outer",
    color: opts.color || "000000",
    blur: opts.blur ?? 6,
    offset: opts.offset ?? 2,
    angle: opts.angle ?? 135,
    opacity: opts.opacity ?? 0.15,
  };
}

function makeCardShadow() {
  return makeShadow({ blur: 8, offset: 3, opacity: 0.12 });
}

function makeSubtleShadow() {
  return makeShadow({ blur: 4, offset: 1, opacity: 0.08 });
}

function makeChartOpts(overrides = {}) {
  const c = brand.colors;
  return {
    x: brand.contentSlide.marginX,
    y: 1.8,
    w: 9.5,
    h: 5.0,
    showLegend: true,
    legendPos: "r",
    legendFontSize: brand.fontSize.footnote,
    showTitle: false,
    chartColors: [c.swooshOrange, c.deepOrange, c.amber, c.primaryOrange, c.checkmarkAmber],
    catAxisLabelFontSize: brand.fontSize.footnote,
    valAxisLabelFontSize: brand.fontSize.footnote,
    dataLabelFontSize: brand.fontSize.footnote,
    dataLabelColor: c.nearBlack,
    catAxisOrientation: "minMax",
    valGridLine: { style: "dash", color: c.dividerGray, size: 0.5 },
    ...overrides,
  };
}

function makeTableOpts() {
  return {
    margin: [4, 8, 4, 8],
  };
}

function assertBrandColor(hex) {
  const clean = hex.replace("#", "").toUpperCase();
  if (!ALLOWED_COLORS.has(clean)) {
    const mapped = mapToBrand(clean);
    return mapped;
  }
  return clean;
}

function mapToBrand(hex) {
  const c = hex.toUpperCase();
  const brandColors = [
    { from: /^3B82F6$|^2563EB$/, to: brand.colors.primaryOrange },
    { from: /^1D4ED8$/, to: "EA580C" },
    { from: /^93C5FD$/, to: "FDBA74" },
    { from: /^06B6D4$/, to: brand.colors.amber },
    { from: /^0891B2$/, to: "D97706" },
    { from: /^67E8F9$/, to: "FCD34D" },
    { from: /^14B8A6$/, to: brand.colors.deepOrange },
    { from: /^0D9488$/, to: "EA580C" },
    { from: /^0F172A$/, to: brand.colors.nearBlack },
    { from: /^1E293B$/, to: brand.colors.dark },
    { from: /^334155$/, to: "2D2D2D" },
  ];
  for (const mapping of brandColors) {
    if (mapping.from.test(c)) return mapping.to;
  }
  return c;
}

function addBrandedRect(slide, pptx, x, y, w, h, colorHex, opts) {
  opts = opts || {};
  const fill = assertBrandColor(colorHex);
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w,
    h,
    fill: { color: fill, transparency: opts.transparency || 0 },
    line: opts.line
      ? { color: assertBrandColor(opts.line.color || fill), width: opts.line.width || 0 }
      : { color: fill, width: 0 },
    rectRadius: opts.radius || 0,
  });
}

function addBrandedText(slide, text, x, y, w, h, opts) {
  opts = opts || {};
  const color = opts.color ? assertBrandColor(opts.color) : brand.colors.nearBlack;
  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontSize: opts.fontSize || brand.fontSize.body,
    fontFace: opts.fontFace || brand.fonts.body,
    color,
    bold: opts.bold || false,
    italic: opts.italic || false,
    align: opts.align || "left",
    valign: opts.valign || "top",
    wrap: opts.wrap !== false,
    paraSpaceAfter: opts.paraSpaceAfter || 0,
    paraSpaceBefore: opts.paraSpaceBefore || 0,
    ...(opts.shadingColor ? { fill: { color: assertBrandColor(opts.shadingColor) } } : {}),
  });
}

function addBrandedTable(slide, pptx, headers, rows, x, y, w, colW) {
  const tableTheme = themes.buildTableTheme();

  const tableRows = [];

  if (headers && headers.length > 0) {
    tableRows.push(
      headers.map((h) => ({
        text: h,
        options: {
          bold: tableTheme.headerBold,
          fontSize: tableTheme.headerFontSize,
          fontFace: brand.fonts.body,
          color: tableTheme.headerColor,
          fill: { color: tableTheme.headerFill },
          align: "left",
          valign: "middle",
          margin: [4, 8, 4, 8],
        },
      }))
    );
  }

  if (rows && rows.length > 0) {
    rows.forEach((row, idx) => {
      tableRows.push(
        row.map((cell) => ({
          text: String(cell),
          options: {
            fontSize: tableTheme.bodyFontSize,
            fontFace: brand.fonts.body,
            color: tableTheme.bodyColor,
            fill: { color: idx % 2 === 1 ? tableTheme.bodyAltFill : tableTheme.bodyFill },
            align: "left",
            valign: "middle",
            margin: [4, 8, 4, 8],
          },
        }))
      );
    });
  }

  if (tableRows.length === 0) return;

  const colCount = tableRows[0].length;
  slide.addTable(tableRows, {
    x,
    y,
    w,
    colW: colW || w / colCount,
    rowH: 0.4,
    border: { type: "solid", pt: 0.5, color: tableTheme.borderColor },
  });
}

function addBulletList(slide, items, x, y, w, h) {
  if (!items || items.length === 0) return;

  const bulletPoints = items.map((text) => ({
    text,
    options: {
      fontSize: brand.fontSize.body,
      color: brand.colors.nearBlack,
      fontFace: brand.fonts.body,
      bullet: { characterCode: "2713" },
      paraSpaceAfter: 6,
    },
  }));

  slide.addText(bulletPoints, {
    x,
    y,
    w,
    h,
    valign: "top",
    lineSpacingMultiple: 1.2,
  });
}

function addLogoIcon(slide, pptx) {
  const logo = brand.logo.logoIcon;
  if (!logo) return;
  const l = brand.contentSlide;
  slide.addImage({
    data: logo,
    x: l.logoIconLeft,
    y: l.logoIconTop,
    w: l.logoIconWidth,
    h: l.logoIconHeight,
  });
}

function addLogoFull(slide) {
  const logo = brand.logo.logoFullWhite;
  if (!logo) return;
  const t = brand.titleSlide;
  slide.addImage({
    data: logo,
    x: t.logoFullLeft,
    y: t.logoFullTop,
    w: t.logoFullWidth,
    h: t.logoFullHeight,
  });
}

function addAccentBar(slide, pptx, x, y, w, colorHex) {
  addBrandedRect(slide, pptx, x, y, w, 0.04, colorHex || brand.colors.primaryOrange);
}

// ── Icon generation (react-icons → PNG → base64) ──
let _react, _ReactDOMServer, _sharp;
async function _loadIconDeps() {
  if (!_react) {
    try {
      _react = require("react");
      _ReactDOMServer = require("react-dom/server");
      _sharp = require("sharp");
    } catch (e) {
      return false;
    }
  }
  return true;
}

async function generateIcon(IconComponent, color, size = 256) {
  const loaded = await _loadIconDeps();
  if (!loaded || !_react || !_ReactDOMServer || !_sharp) return null;

  const svg = _ReactDOMServer.renderToStaticMarkup(
    _react.createElement(IconComponent, { color: color || "#000000", size: String(size) })
  );
  const pngBuffer = await _sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

function addIconSlide(slide, pptx, iconData, x, y, w, h) {
  if (!iconData) return;
  slide.addImage({ data: iconData, x, y, w, h });
}

// ── Visual motif: colored circle with icon ──
function addIconCircle(slide, pptx, iconData, x, y, size, bgColor) {
  const c = brand.colors;
  slide.addShape(pptx.ShapeType.ellipse, {
    x,
    y,
    w: size,
    h: size,
    fill: { color: bgColor || c.swooshOrange },
  });
  if (iconData) {
    const padding = size * 0.2;
    slide.addImage({ data: iconData, x: x + padding, y: y + padding, w: size - padding * 2, h: size - padding * 2 });
  }
}

// ── Stat callout (big number + label) ──
function addStatCallout(slide, value, label, x, y, w, opts = {}) {
  const c = brand.colors;
  slide.addText(value, {
    x,
    y,
    w,
    h: opts.valueHeight || 0.8,
    fontSize: opts.valueFontSize || 48,
    bold: true,
    color: opts.valueColor || c.swooshOrange,
    fontFace: brand.fonts.heading,
    align: "center",
    valign: "bottom",
  });
  slide.addText(label, {
    x,
    y: y + (opts.valueHeight || 0.8),
    w,
    h: opts.labelHeight || 0.4,
    fontSize: opts.labelFontSize || brand.fontSize.caption,
    color: opts.labelColor || c.midGray,
    fontFace: brand.fonts.body,
    align: "center",
    valign: "top",
  });
}

module.exports = {
  assertBrandColor,
  mapToBrand,
  addBrandedRect,
  addBrandedText,
  addBrandedTable,
  addBulletList,
  addLogoIcon,
  addLogoFull,
  addAccentBar,
  makeShadow,
  makeCardShadow,
  makeSubtleShadow,
  makeChartOpts,
  makeTableOpts,
  generateIcon,
  addIconSlide,
  addIconCircle,
  addStatCallout,
};
