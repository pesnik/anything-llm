# PptxGenJS Tutorial for Slide Generation

## Setup

```javascript
const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3" × 7.5"
```

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
// From base64
slide.addImage({
  data: "image/png;base64,iVBORw0KGgo...",
  x: 1, y: 1, w: 5, h: 3
});

// Sizing modes
{ sizing: { type: 'contain', w: 4, h: 3 } }  // Fit inside
{ sizing: { type: 'cover', w: 4, h: 3 } }    // Fill area
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
