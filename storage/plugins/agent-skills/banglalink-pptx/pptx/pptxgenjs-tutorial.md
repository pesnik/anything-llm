# PptxGenJS Tutorial for Slide Generation

## Setup

```javascript
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3" × 7.5"
pres.author = "Your Name";
pres.title = "Presentation Title";
```

## Layout Dimensions

| Layout | Width | Height |
|--------|-------|--------|
| `LAYOUT_16x9` | 10" | 5.625" |
| `LAYOUT_16x10` | 10" | 6.25" |
| `LAYOUT_4x3` | 10" | 7.5" |
| `LAYOUT_WIDE` | 13.3" | 7.5" |

All coordinates (x, y, w, h) are in inches.

## Text

```javascript
// Basic text
slide.addText("Title", {
  x: 0.5, y: 0.5, w: 8, h: 1,
  fontSize: 28, fontFace: "Arial", bold: true,
  color: "EF6E23", align: "center", valign: "middle"
});

// CharSpacing for headings (not letterSpacing — silently ignored)
slide.addText("HEADING", { charSpacing: 6 });

// Margin: 0 for precise alignment with shapes
slide.addText("Title", { x: 1, y: 1, w: 5, h: 1, margin: 0 });

// Multi-line text (breakLine: REQUIRED)
slide.addText([
  { text: "Line 1", options: { breakLine: true } },
  { text: "Line 2", options: { breakLine: true } },
  { text: "Line 3" }
], { x: 0.5, y: 0.5, w: 8, h: 2 });

// Rich text
slide.addText([
  { text: "Bold ", options: { bold: true, breakLine: true } },
  { text: "Normal" }
], { x: 0.5, y: 0.5, w: 8, h: 1 });
```

## Bullets

```javascript
// ✅ CORRECT: Use bullet option
slide.addText([
  { text: "First", options: { bullet: true, breakLine: true } },
  { text: "Second", options: { bullet: true, breakLine: true } },
  { text: "Third", options: { bullet: true } }
], { x: 0.5, y: 0.5, w: 8, h: 3 });

// ❌ WRONG: Unicode bullets create double bullets
slide.addText("• First item", { ... });

// Numbered lists
{ text: "First", options: { bullet: { type: "number" }, breakLine: true } }

// Sub-items
{ text: "Sub-item", options: { bullet: true, indentLevel: 1 } }
```

## Shapes

```javascript
// Rectangle
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 0.8, w: 1.5, h: 3.0,
  fill: { color: "FF0000" },
  line: { color: "000000", width: 2 }
});

// With transparency
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "0088CC", transparency: 50 }
});

// Rounded rectangle (don't use with accent bars)
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" }, rectRadius: 0.1
});

// Shadow (never use negative offset)
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" },
  shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 }
});

// Upward shadow (use angle: 270)
const upShadow = { type: "outer", blur: 4, offset: 2, angle: 270, color: "000000", opacity: 0.1 };
```

## Charts

```javascript
// Bar chart (vertical columns — use barDir: "col")
slide.addChart(pres.charts.BAR, [
  { name: "Sales", labels: ["Q1","Q2","Q3","Q4"], values: [45,55,62,71] }
], {
  x: 0.5, y: 1, w: 9, h: 4,
  barDir: "col",
  chartColors: ["EF6E23", "F86C02"],
  catGridLine: { style: "none" },
  valGridLine: { color: "D9D9D9", size: 0.5 },
  showValue: true,
  dataLabelPosition: "outEnd",
  dataLabelColor: "1E293B",
  chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },
  showLegend: true,
  legendPos: "r"
});

// Line chart (smooth curves)
slide.addChart(pres.charts.LINE, chartData, {
  lineSize: 3,
  lineSmooth: true,
});

// Pie chart (show percentages)
slide.addChart(pres.charts.PIE, chartData, {
  showPercent: true,
});

// Doughnut chart
slide.addChart(pres.charts.DOUGHNUT, chartData, {
  showPercent: true,
});

// Radar chart
slide.addChart(pres.charts.RADAR, chartData);

// Scatter chart
slide.addChart(pres.charts.SCATTER, chartData);
```

## Images

```javascript
// From file path
slide.addImage({ path: "images/chart.png", x: 1, y: 1, w: 5, h: 3 });

// From URL
slide.addImage({ path: "https://example.com/image.jpg", x: 1, y: 1, w: 5, h: 3 });

// From base64 (faster, no file I/O)
slide.addImage({ data: "image/png;base64,iVBORw0KGgo...", x: 1, y: 1, w: 5, h: 3 });
```

### Image Options

```javascript
slide.addImage({
  path: "image.png",
  x: 1, y: 1, w: 5, h: 3,
  rotate: 45,              // 0-359 degrees
  rounding: true,          // Circular crop
  transparency: 50,        // 0-100
  flipH: true,             // Horizontal flip
  flipV: false,            // Vertical flip
  altText: "Description",  // Accessibility
  hyperlink: { url: "https://example.com" }
});
```

### Image Sizing Modes

```javascript
// Contain - fit inside, preserve ratio
{ sizing: { type: 'contain', w: 4, h: 3 } }

// Cover - fill area, preserve ratio (may crop)
{ sizing: { type: 'cover', w: 4, h: 3 } }

// Crop - cut specific portion
{ sizing: { type: 'crop', x: 0.5, y: 0.5, w: 2, h: 2 } }
```

### Calculate Dimensions (preserve aspect ratio)

```javascript
const origWidth = 1978, origHeight = 923, maxHeight = 3.0;
const calcWidth = maxHeight * (origWidth / origHeight);
const centerX = (10 - calcWidth) / 2;
slide.addImage({ path: "image.png", x: centerX, y: 1.2, w: calcWidth, h: maxHeight });
```

### Supported Formats

- **Standard**: PNG, JPG, GIF (animated GIFs work in Microsoft 365)
- **SVG**: Works in modern PowerPoint/Microsoft 365

## Icons

Use react-icons to generate SVG icons, then rasterize to PNG for universal compatibility.

```javascript
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaCheckCircle, FaChartLine } = require("react-icons/fa");

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

// Add to slide
const iconData = await iconToBase64Png(FaCheckCircle, "#4472C4", 256);
slide.addImage({ data: iconData, x: 1, y: 1, w: 0.5, h: 0.5 });
```

**Note**: Use size 256+ for crisp icons. Size controls rasterization resolution, not display size (set by `w` and `h`).

### Icon Libraries

- `react-icons/fa` — Font Awesome
- `react-icons/md` — Material Design
- `react-icons/hi` — Heroicons
- `react-icons/bi` — Bootstrap Icons

## Slide Backgrounds

```javascript
// Solid color
slide.background = { color: "F1F1F1" };

// Color with transparency
slide.background = { color: "FF3399", transparency: 50 };

// Image from URL
slide.background = { path: "https://example.com/bg.jpg" };

// Image from base64
slide.background = { data: "image/png;base64,iVBORw0KGgo..." };
```

## Tables

```javascript
slide.addTable([
  ["Header 1", "Header 2"],
  ["Cell 1", "Cell 2"]
], {
  x: 1, y: 1, w: 8, h: 2,
  border: { pt: 1, color: "999999" },
  colW: [4, 4]
});

// With styled cells
let tableData = [
  [{ text: "Header", options: { fill: { color: "6699CC" }, color: "FFFFFF", bold: true } }, "Cell"],
  [{ text: "Merged", options: { colspan: 2 } }]
];
```

## Slide Masters

```javascript
pres.defineSlideMaster({
  title: "CONTENT_SLIDE",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 1, y: 2, w: 8, h: 2 } } }
  ]
});

let slide = pres.addSlide({ masterName: "CONTENT_SLIDE" });
slide.addText("My Title", { placeholder: "title" });
```

---

## Common Pitfalls

| Issue | Solution |
|-------|----------|
| `#` prefix with hex colors | Remove `#` — use `"FF0000"` not `"#FF0000"` |
| 8-char hex for opacity | Use `opacity` property — `"00000020"` corrupts file |
| Unicode bullets (•) | Use `bullet: true` or `bullet: { characterCode: "2713" }` |
| Missing `breakLine: true` | Add between array items or text runs together |
| Reusing option objects | PptxGenJS mutates objects in-place — create fresh each time |
| `letterSpacing` | Use `charSpacing` instead (letterSpacing silently ignored) |
| Negative shadow offset | Use `angle: 270` with positive offset for upward shadows |
| `lineSpacing` with bullets | Use `paraSpaceAfter` instead (lineSpacing causes excessive gaps) |
| `ROUNDED_RECTANGLE` + accent bar | Use `RECTANGLE` (accent won't cover rounded corners) |
| `margin` not set | Set `margin: 0` for precise alignment with shapes |
| `barDir: "bar"` | Use `barDir: "col"` for vertical columns |
| Bubble chart | Broken in pptxgenjs 4.0.1 — skip or use bar3d |
| Each presentation needs fresh instance | Don't reuse `pptxgen()` objects |

---

## Shadow Helper

```javascript
const makeShadow = () => ({
  type: "outer",
  color: "000000",
  blur: 6,
  offset: 2,
  angle: 135,
  opacity: 0.15
});

// Use fresh object each call
slide.addShape(pres.shapes.RECTANGLE, { shadow: makeShadow(), ... });
slide.addShape(pres.shapes.RECTANGLE, { shadow: makeShadow(), ... });
```

---

## Quick Reference

- **Shapes**: RECTANGLE, OVAL, LINE, ROUNDED_RECTANGLE
- **Charts**: BAR, LINE, PIE, DOUGHNUT, SCATTER, BUBBLE, RADAR
- **Layouts**: LAYOUT_16x9 (10"×5.625"), LAYOUT_16x10, LAYOUT_4x3, LAYOUT_WIDE (13.3"×7.5")
- **Alignment**: "left", "center", "right"
- **Chart data labels**: "outEnd", "inEnd", "center"
- **Legend positions**: "b", "t", "l", "r", "tr"
