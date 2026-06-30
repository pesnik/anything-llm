/**
 * Component Registry — Composable slide components.
 * 
 * Pattern from: SlideForge (40+ composable components)
 * Pattern from: IBM/chuk-mcp-pptx (universal component API)
 * 
 * Each component is a function that returns pptxgenjs operations.
 * Components compose — nest them in layouts to build complex slides.
 */

const P = require("./primitives");
const brand = require("../brand");

// ── Component: Background ──
function Background(slide, { color = P.colors.white, image } = {}) {
  if (image) {
    slide.background = { path: image };
  } else {
    slide.background = { color };
  }
}

// ── Component: Title ──
function Title(slide, { text, y = P.spacing.titleY, style = "slideTitle" }) {
  if (!text) return;
  P.addText(slide, text, {
    x: P.spacing.slideMargin,
    y,
    w: P.spacing.contentWidth,
    h: P.spacing.titleHeight,
    style,
  });
}

// ── Component: Subtitle ──
function Subtitle(slide, { text, y = P.spacing.subtitleY }) {
  if (!text) return;
  P.addText(slide, text, {
    x: P.spacing.slideMargin,
    y,
    w: P.spacing.contentWidth,
    h: 0.4,
    style: "subtitle",
  });
}

// ── Component: LogoIcon ──
function LogoIcon(slide, { position = "top-right" } = {}) {
  P.addLogo(slide, { position });
}

// ── Component: KPI Cards ──
function KPICards(slide, pptx, { metrics, y = 1.8, cardW = 3.0, cardH = 2.5 }) {
  if (!metrics || metrics.length === 0) return;
  
  const count = Math.min(metrics.length, 4);
  const totalW = cardW * count + P.spacing.cardGap * (count - 1);
  const startX = (P.spacing.slideWidth - totalW) / 2;
  
  for (let i = 0; i < count; i++) {
    const m = metrics[i];
    const x = startX + i * (cardW + P.spacing.cardGap);
    
    // Card background
    P.addRect(slide, pptx, { x, y, w: cardW, h: cardH, fill: P.colors.white, line: { color: P.colors.orange, width: 1 } });
    
    // Label (top)
    P.addText(slide, m.label || "", {
      x: x + P.spacing.cardPadding,
      y: y + 0.3,
      w: cardW - P.spacing.cardPadding * 2,
      h: 0.5,
      style: "caption",
      align: "center",
    });
    
    // Big number (center)
    P.addText(slide, m.value || "", {
      x: x + P.spacing.cardPadding,
      y: y + 0.9,
      w: cardW - P.spacing.cardPadding * 2,
      h: 1.0,
      style: "bigNumber",
    });
    
    // Sublabel (bottom)
    if (m.sublabel) {
      P.addText(slide, m.sublabel, {
        x: x + P.spacing.cardPadding,
        y: y + cardH - 0.6,
        w: cardW - P.spacing.cardPadding * 2,
        h: 0.4,
        style: "label",
      });
    }
  }
}

// ── Component: Two Column ──
function TwoColumn(slide, pptx, { leftTitle, left, rightTitle, right, y = 1.5 }) {
  const colW = (P.spacing.contentWidth - 0.5) / 2;
  const leftX = P.spacing.slideMargin;
  const rightX = P.spacing.slideMargin + colW + 0.5;
  
  // Left column
  if (leftTitle) {
    P.addText(slide, leftTitle, {
      x: leftX, y, w: colW, h: 0.5,
      style: "subtitle", bold: true,
    });
  }
  if (left && left.length > 0) {
    P.addBullets(slide, left, {
      x: leftX, y: y + 0.6, w: colW, h: 4.0,
    });
  }
  
  // Divider line
  P.addLine(slide, pptx, {
    x: P.spacing.slideMargin + colW + 0.25,
    y,
    w: 0,
    color: P.colors.divider,
  });
  
  // Right column
  if (rightTitle) {
    P.addText(slide, rightTitle, {
      x: rightX, y, w: colW, h: 0.5,
      style: "subtitle", bold: true,
    });
  }
  if (right && right.length > 0) {
    P.addBullets(slide, right, {
      x: rightX, y: y + 0.6, w: colW, h: 4.0,
    });
  }
}

// ── Component: Timeline ──
function Timeline(slide, pptx, { milestones, y = 1.8 }) {
  if (!milestones || milestones.length === 0) return;
  
  const count = Math.min(milestones.length, 6);
  const colW = 1.8;
  const colGap = 0.2;
  const totalW = colW * count + colGap * (count - 1);
  const startX = (P.spacing.slideWidth - totalW) / 2;
  
  for (let i = 0; i < count; i++) {
    const m = milestones[i];
    const x = startX + i * (colW + colGap);
    
    // Month label (orange circle)
    P.addCircle(slide, pptx, { x, y, size: 0.5, fill: P.colors.orange });
    P.addText(slide, m.label || m.date || "", {
      x, y, w: 0.5, h: 0.5,
      fontSize: 12, bold: true, color: P.colors.white,
      fontFace: P.typography.fonts.body,
      align: "center", valign: "middle",
    });
    
    // Details card
    P.addRect(slide, pptx, {
      x, y: y + 0.6, w: colW, h: 3.5,
      fill: P.colors.lightGray,
    });
    P.addText(slide, m.details || m.desc || "", {
      x: x + 0.15, y: y + 0.75, w: colW - 0.3, h: 3.2,
      style: "body", wrap: true,
    });
  }
}

// ── Component: Image Cards ──
function ImageCards(slide, pptx, { cards, y = 1.8 }) {
  if (!cards || cards.length === 0) return;
  
  const count = Math.min(cards.length, 4);
  const cardW = 3.5;
  const cardH = 3.0;
  const totalW = cardW * count + P.spacing.cardGap * (count - 1);
  const startX = (P.spacing.slideWidth - totalW) / 2;
  
  for (let i = 0; i < count; i++) {
    const card = cards[i];
    const x = startX + i * (cardW + P.spacing.cardGap);
    
    // Card background
    P.addRect(slide, pptx, { x, y, w: cardW, h: cardH, fill: P.colors.white, line: { color: P.colors.orange, width: 1 } });
    
    // Caption
    if (card.caption) {
      P.addText(slide, card.caption, {
        x, y: y + cardH + 0.15, w: cardW, h: 0.6,
        style: "caption", align: "center", wrap: true,
      });
    }
  }
}

// ── Component: Section Break ──
function SectionBreak(slide, { title, subtitle }) {
  Background(slide, { color: P.colors.orange });
  
  if (title) {
    P.addText(slide, title, {
      x: 1.0, y: 2.5, w: 11.333, h: 2.0,
      fontSize: 36, bold: true, color: P.colors.white,
      fontFace: P.typography.fonts.heading,
      align: "center", valign: "middle",
    });
  }
  
  if (subtitle) {
    P.addText(slide, subtitle, {
      x: 1.0, y: 4.5, w: 11.333, h: 0.8,
      fontSize: 14, color: P.colors.white,
      fontFace: P.typography.fonts.body,
      align: "center", valign: "top",
    });
  }
  
  P.addLogo(slide, { position: "center" });
}

// ── Component: Stat Callout (big number + label) ──
function StatCallout(slide, { value, label, x, y, w, valueFontSize = 48, labelFontSize = 12 }) {
  P.addText(slide, value, {
    x, y, w, h: 0.8,
    fontSize: valueFontSize, bold: true, color: P.colors.orange,
    fontFace: P.typography.fonts.heading,
    align: "center", valign: "bottom",
  });
  P.addText(slide, label, {
    x, y: y + 0.8, w, h: 0.4,
    fontSize: labelFontSize, color: P.colors.midGray,
    fontFace: P.typography.fonts.body,
    align: "center", valign: "top",
  });
}

// ── Component: Icon Circle ──
function IconCircle(slide, pptx, { iconData, x, y, size = 0.6, bgColor = P.colors.orange }) {
  P.addCircle(slide, pptx, { x, y, size, fill: bgColor });
  if (iconData) {
    const padding = size * 0.2;
    P.addImage(slide, { data: iconData, x: x + padding, y: y + padding, w: size - padding * 2, h: size - padding * 2 });
  }
}

// ── Component: Quote ──
function Quote(slide, { text, attribution, y = 1.5 }) {
  P.addText(slide, text, {
    x: 1.5, y, w: 10.0, h: 2.0,
    fontSize: 20, italic: true, color: P.colors.nearBlack,
    fontFace: P.typography.fonts.body,
    align: "center", valign: "middle",
  });
  
  if (attribution) {
    P.addText(slide, `— ${attribution}`, {
      x: 1.5, y: y + 2.2, w: 10.0, h: 0.5,
      fontSize: 14, color: P.colors.midGray,
      fontFace: P.typography.fonts.body,
      align: "center", valign: "top",
    });
  }
}

// ── Component: Closing Slide ──
function ClosingSlide(slide, { cta = "Thank You", contact }) {
  Background(slide, { color: P.colors.dark });
  
  P.addText(slide, cta, {
    x: 1.0, y: 2.0, w: 11.333, h: 1.5,
    fontSize: 44, bold: true, color: P.colors.white,
    fontFace: P.typography.fonts.heading,
    align: "center", valign: "middle",
  });
  
  if (contact) {
    P.addText(slide, contact, {
      x: 1.0, y: 3.8, w: 11.333, h: 0.8,
      fontSize: 16, color: P.colors.midGray,
      fontFace: P.typography.fonts.body,
      align: "center", valign: "top",
    });
  }
  
  // Ghost logo via XML injection (handled by builder)
}

// ── Component: Numbered List (two-column with numbers) ──
function NumberedList(slide, pptx, { leftTitle, left, rightTitle, right, y = 1.5 }) {
  const colW = (P.spacing.contentWidth - 0.5) / 2;
  const leftX = P.spacing.slideMargin;
  const rightX = P.spacing.slideMargin + colW + 0.5;
  
  const renderColumn = (title, items, x) => {
    if (title) {
      P.addText(slide, title, {
        x, y, w: colW, h: 0.5,
        fontSize: 14, bold: true, color: P.colors.midGray,
        fontFace: P.typography.fonts.body,
      });
    }
    if (items) {
      items.forEach((item, i) => {
        const num = String(i + 1).padStart(2, "0");
        P.addText(slide, num, {
          x, y: y + 0.6 + i * 0.5, w: 0.5, h: 0.4,
          fontSize: 14, bold: true, color: P.colors.orange,
          fontFace: P.typography.fonts.heading,
        });
        P.addText(slide, item, {
          x: x + 0.55, y: y + 0.6 + i * 0.5, w: colW - 0.6, h: 0.4,
          fontSize: 13, color: P.colors.nearBlack,
          fontFace: P.typography.fonts.body,
        });
      });
    }
  };
  
  renderColumn(leftTitle, left, leftX);
  renderColumn(rightTitle, right, rightX);
}

// ── Component: KPI with Headline (headline + big number + date) ──
function KPIWithHeadline(slide, pptx, { metrics, y = 1.8, cardW = 2.8, cardH = 2.8 }) {
  if (!metrics || metrics.length === 0) return;
  
  const count = Math.min(metrics.length, 4);
  const totalW = cardW * count + P.spacing.cardGap * (count - 1);
  const startX = (P.spacing.slideWidth - totalW) / 2;
  
  for (let i = 0; i < count; i++) {
    const m = metrics[i];
    const x = startX + i * (cardW + P.spacing.cardGap);
    
    // Card background (orange border)
    P.addRect(slide, pptx, { x, y, w: cardW, h: cardH, fill: P.colors.white, line: { color: P.colors.orange, width: 1 } });
    
    // Headline
    P.addText(slide, m.headline || m.label || "", {
      x: x + P.spacing.cardPadding, y: y + 0.2, w: cardW - P.spacing.cardPadding * 2, h: 0.8,
      fontSize: 14, bold: true, color: P.colors.nearBlack,
      fontFace: P.typography.fonts.body, align: "center", valign: "top",
    });
    
    // Big number
    P.addText(slide, m.value || "", {
      x: x + P.spacing.cardPadding, y: y + 1.0, w: cardW - P.spacing.cardPadding * 2, h: 1.0,
      fontSize: 48, bold: true, color: P.colors.orange,
      fontFace: P.typography.fonts.heading, align: "center", valign: "middle",
    });
    
    // Date
    if (m.date) {
      P.addText(slide, m.date, {
        x: x + P.spacing.cardPadding, y: y + cardH - 0.5, w: cardW - P.spacing.cardPadding * 2, h: 0.4,
        fontSize: 12, color: P.colors.orange,
        fontFace: P.typography.fonts.body, align: "center", valign: "bottom",
      });
    }
  }
}

// ── Component: Large Image Placeholder ──
function LargeImagePlaceholder(slide, pptx, { caption, y = 1.5 }) {
  // Large orange-bordered rectangle
  P.addRect(slide, pptx, {
    x: 1.5, y, w: 10.0, h: 4.5,
    fill: P.colors.white, line: { color: P.colors.orange, width: 1 },
  });
  
  // Caption below
  if (caption) {
    P.addBullets(slide, [caption], {
      x: 1.5, y: y + 4.7, w: 10.0, h: 0.8,
    });
  }
}

// ── Component: Dashboard (2x2 mini charts) ──
function Dashboard(slide, pptx, { charts, y = 1.5 }) {
  if (!charts || charts.length === 0) return;
  
  const count = Math.min(charts.length, 4);
  const cellW = (P.spacing.contentWidth - 0.3) / 2;
  const cellH = 2.8;
  
  for (let i = 0; i < count; i++) {
    const c = charts[i];
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = P.spacing.slideMargin + col * (cellW + 0.3);
    const cellY = y + row * (cellH + 0.3);
    
    // Cell background
    P.addRect(slide, pptx, { x, y: cellY, w: cellW, h: cellH, fill: P.colors.white, line: { color: P.colors.divider, width: 0.5 } });
    
    // Chart title
    if (c.title) {
      P.addText(slide, c.title, {
        x: x + 0.2, y: cellY + 0.1, w: cellW - 0.4, h: 0.4,
        fontSize: 11, bold: true, color: P.colors.nearBlack,
        fontFace: P.typography.fonts.body,
      });
    }
    
    // Chart subtitle
    if (c.subtitle) {
      P.addText(slide, c.subtitle, {
        x: x + 0.2, y: cellY + 0.4, w: cellW - 0.4, h: 0.3,
        fontSize: 9, color: P.colors.midGray,
        fontFace: P.typography.fonts.body,
      });
    }
    
    // Mini chart
    if (c.data) {
      P.addChart(slide, pptx, {
        type: c.type || "bar",
        data: c.data,
        x: x + 0.2, y: cellY + 0.7, w: cellW - 0.4, h: cellH - 0.9,
        options: { showLegend: false, showTitle: false },
      });
    }
  }
}

// ── Component: Calendar Timeline (with calendar icons) ──
function CalendarTimeline(slide, pptx, { milestones, y = 1.5 }) {
  if (!milestones || milestones.length === 0) return;
  
  const count = Math.min(milestones.length, 6);
  const colW = 1.7;
  const colGap = 0.15;
  const totalW = colW * count + colGap * (count - 1);
  const startX = (P.spacing.slideWidth - totalW) / 2;
  
  for (let i = 0; i < count; i++) {
    const m = milestones[i];
    const x = startX + i * (colW + colGap);
    
    // Calendar icon (simplified: orange top bar + white box)
    P.addRect(slide, pptx, { x: x + 0.35, y, w: colW - 0.7, h: 0.15, fill: P.colors.orange });
    P.addRect(slide, pptx, { x: x + 0.3, y: y + 0.15, w: colW - 0.6, h: 0.8, fill: P.colors.white, line: { color: P.colors.divider, width: 0.5 } });
    P.addText(slide, m.label || "", {
      x: x + 0.3, y: y + 0.15, w: colW - 0.6, h: 0.8,
      fontSize: 14, bold: true, color: P.colors.nearBlack,
      fontFace: P.typography.fonts.body, align: "center", valign: "middle",
    });
    
    // Details card (gradient gray)
    const grayShade = Math.floor(200 + (i / count) * 40);
    P.addRect(slide, pptx, { x, y: y + 1.1, w: colW, h: 4.5, fill: `D${grayShade.toString(16)}D${grayShade.toString(16)}D${grayShade.toString(16)}` });
    
    // Month + details
    P.addText(slide, m.month || m.label || "", {
      x: x + 0.15, y: y + 1.2, w: colW - 0.3, h: 0.4,
      fontSize: 12, bold: true, color: P.colors.nearBlack,
      fontFace: P.typography.fonts.body,
    });
    P.addText(slide, m.details || m.desc || "", {
      x: x + 0.15, y: y + 1.6, w: colW - 0.3, h: 3.8,
      fontSize: 11, color: P.colors.nearBlack,
      fontFace: P.typography.fonts.body, wrap: true,
    });
  }
}

// ── Component: Table with Checklist ──
function TableWithChecklist(slide, pptx, { headers, rows, notes, y = 1.8 }) {
  if (!rows || rows.length === 0) return;
  
  const tableW = 9.5;
  const tableData = [];
  
  // Headers (orange bg)
  if (headers) {
    tableData.push(headers.map(h => ({
      text: h,
      options: {
        bold: true, fontSize: 12, fontFace: P.typography.fonts.body,
        color: P.colors.white, fill: { color: P.colors.orange },
        align: "left", valign: "middle", margin: [4, 6, 4, 6],
      },
    })));
    // Add empty header for checklist column
    tableData[0].push({ text: "", options: { fill: { color: P.colors.orange } } });
  }
  
  // Rows
  rows.forEach((row, idx) => {
    const rowData = row.map(cell => ({
      text: String(cell),
      options: {
        fontSize: 12, fontFace: P.typography.fonts.body,
        color: P.colors.nearBlack,
        fill: { color: idx % 2 === 1 ? P.colors.lightGray : P.colors.white },
        align: "left", valign: "middle", margin: [4, 6, 4, 6],
      },
    }));
    // Add checkmark
    rowData.push({
      text: row[row.length - 1] === true || row[row.length - 1] === "✓" ? "✓" : "",
      options: {
        fontSize: 18, color: P.colors.orange, bold: true,
        fill: { color: idx % 2 === 1 ? P.colors.lightGray : P.colors.white },
        align: "center", valign: "middle",
      },
    });
    tableData.push(rowData);
  });
  
  const colCount = headers ? headers.length + 1 : rows[0].length + 1;
  const colW = Array(colCount).fill(tableW / colCount);
  colW[colW.length - 1] = 0.8; // Checklist column narrower
  
  slide.addTable(tableData, {
    x: P.spacing.slideMargin, y, w: tableW,
    colW, rowH: 0.45,
    border: { type: "solid", pt: 0.5, color: P.colors.divider },
  });
  
  // Notes
  if (notes) {
    P.addText(slide, notes, {
      x: P.spacing.slideMargin + tableW + 0.3, y, w: 2.5, h: 2.0,
      fontSize: 10, color: P.colors.midGray,
      fontFace: P.typography.fonts.body, wrap: true,
    });
  }
}

// ── Component: Combined Table + Pie Chart ──
function TablePieCombo(slide, pptx, { tables, pieData, pieTitle, y = 1.5 }) {
  const leftW = 7.0;
  const rightX = P.spacing.slideMargin + leftW + 0.5;
  const rightW = P.spacing.contentWidth - leftW - 0.5;
  
  // Left side: tables
  let tableY = y;
  if (tables) {
    tables.forEach(t => {
      if (t.title) {
        P.addText(slide, t.title, {
          x: P.spacing.slideMargin, y: tableY, w: leftW, h: 0.4,
          fontSize: 14, bold: true, color: P.colors.nearBlack,
          fontFace: P.typography.fonts.body,
        });
        tableY += 0.45;
      }
      if (t.headers && t.rows) {
        const tableData = [];
        tableData.push(t.headers.map(h => ({
          text: h, options: {
            bold: true, fontSize: 11, fontFace: P.typography.fonts.body,
            color: P.colors.white, fill: { color: P.colors.orange },
            align: "left", valign: "middle", margin: [3, 6, 3, 6],
          },
        })));
        t.rows.forEach((row, idx) => {
          tableData.push(row.map(cell => ({
            text: String(cell), options: {
              fontSize: 11, fontFace: P.typography.fonts.body,
              color: P.colors.nearBlack,
              fill: { color: idx % 2 === 1 ? P.colors.lightGray : P.colors.white },
              align: "left", valign: "middle", margin: [3, 6, 3, 6],
            },
          })));
        });
        slide.addTable(tableData, {
          x: P.spacing.slideMargin, y: tableY, w: leftW,
          colW: leftW / t.headers.length, rowH: 0.35,
          border: { type: "solid", pt: 0.5, color: P.colors.divider },
        });
        tableY += 0.35 * (tableData.length + 1) + 0.3;
      }
    });
  }
  
  // Right side: pie chart
  if (pieData) {
    if (pieTitle) {
      P.addText(slide, pieTitle, {
        x: rightX, y, w: rightW, h: 0.5,
        fontSize: 11, color: P.colors.nearBlack,
        fontFace: P.typography.fonts.body,
      });
    }
    P.addChart(slide, pptx, {
      type: "pie",
      data: pieData,
      x: rightX, y: y + 0.5, w: rightW, h: 3.5,
      options: { showPercent: true, showLegend: true, legendPos: "b" },
    });
  }
}

// ── Component: Table with Notes ──
function TableWithNotes(slide, pptx, { headers, rows, notes, y = 1.8 }) {
  const tableW = 10.0;
  const tableData = [];
  
  if (headers) {
    tableData.push(headers.map(h => ({
      text: h, options: {
        bold: true, fontSize: 12, fontFace: P.typography.fonts.body,
        color: P.colors.white, fill: { color: P.colors.orange },
        align: "left", valign: "middle", margin: [4, 8, 4, 8],
      },
    })));
  }
  
  if (rows) {
    rows.forEach((row, idx) => {
      tableData.push(row.map(cell => ({
        text: String(cell), options: {
          fontSize: 12, fontFace: P.typography.fonts.body,
          color: P.colors.nearBlack,
          fill: { color: idx % 2 === 1 ? P.colors.lightGray : P.colors.white },
          align: "left", valign: "middle", margin: [4, 8, 4, 8],
        },
      })));
    });
  }
  
  slide.addTable(tableData, {
    x: P.spacing.slideMargin, y, w: tableW,
    colW: tableW / (headers?.length || rows[0]?.length || 1),
    rowH: 0.45,
    border: { type: "solid", pt: 0.5, color: P.colors.divider },
  });
  
  // Notes on right side
  if (notes) {
    P.addText(slide, notes, {
      x: P.spacing.slideMargin + tableW + 0.3, y, w: 1.8, h: 3.0,
      fontSize: 10, color: P.colors.midGray,
      fontFace: P.typography.fonts.body, wrap: true,
    });
  }
}

// ── Component: Dashboard with Notes ──
function DashboardWithNotes(slide, pptx, { charts, notes, footer, y = 1.5 }) {
  // Same as Dashboard but with notes area
  Dashboard(slide, pptx, { charts, y });
  
  if (notes) {
    P.addText(slide, notes, {
      x: 8.5, y: 5.5, w: 4.5, h: 1.5,
      fontSize: 10, color: P.colors.midGray,
      fontFace: P.typography.fonts.body, wrap: true,
    });
  }
  
  if (footer) {
    P.addText(slide, footer, {
      x: P.spacing.slideMargin, y: 7.0, w: P.spacing.contentWidth, h: 0.3,
      fontSize: 8, color: P.colors.midGray,
      fontFace: P.typography.fonts.body,
    });
  }
}

// ── Component: Closing Slide (white bg version) ──
function ClosingSlideWhite(slide, { cta = "Thank You" }) {
  Background(slide, { color: P.colors.white });
  
  P.addText(slide, cta, {
    x: 3.0, y: 2.5, w: 7.0, h: 1.5,
    fontSize: 36, color: P.colors.nearBlack,
    fontFace: P.typography.fonts.body,
    align: "center", valign: "middle",
  });
  
  // Ghost logo (left side, light beige)
  // Handled by ghost-logo XML injection
}

// ── Component Registry ──
const components = {
  Background,
  Title,
  Subtitle,
  LogoIcon,
  KPICards,
  TwoColumn,
  Timeline,
  ImageCards,
  SectionBreak,
  StatCallout,
  IconCircle,
  Quote,
  ClosingSlide,
  NumberedList,
  KPIWithHeadline,
  LargeImagePlaceholder,
  Dashboard,
  CalendarTimeline,
  TableWithChecklist,
  TablePieCombo,
  TableWithNotes,
  DashboardWithNotes,
  ClosingSlideWhite,
};

// ── Component metadata for agent discovery ──
const registry = [
  { name: "Background", use: "Set slide background color or image", props: "color, image" },
  { name: "Title", use: "Slide title (orange, bold, 28pt)", props: "text, y, style" },
  { name: "Subtitle", use: "Slide subtitle (gray, 18pt)", props: "text, y" },
  { name: "LogoIcon", use: "Banglalink logo icon", props: "position" },
  { name: "KPICards", use: "3-4 metric cards with big numbers", props: "metrics, y, cardW, cardH" },
  { name: "TwoColumn", use: "Side-by-side comparison", props: "leftTitle, left, rightTitle, right, y" },
  { name: "Timeline", use: "Horizontal timeline with milestones", props: "milestones, y" },
  { name: "ImageCards", use: "3-4 image/caption cards", props: "cards, y" },
  { name: "SectionBreak", use: "Full orange divider slide", props: "title, subtitle" },
  { name: "StatCallout", use: "Big number + label at any position", props: "value, label, x, y, w" },
  { name: "IconCircle", use: "Colored circle with icon", props: "iconData, x, y, size, bgColor" },
  { name: "Quote", use: "Quote with attribution", props: "text, attribution, y" },
  { name: "ClosingSlide", use: "Thank You slide with ghost logo", props: "cta, contact" },
  { name: "NumberedList", use: "Two-column numbered items (01-09)", props: "leftTitle, left, rightTitle, right, y" },
  { name: "KPIWithHeadline", use: "KPI cards with headline + big number + date", props: "metrics, y, cardW, cardH" },
  { name: "LargeImagePlaceholder", use: "Large orange-bordered rectangle + caption", props: "caption, y" },
  { name: "Dashboard", use: "2x2 grid of mini charts", props: "charts, y" },
  { name: "CalendarTimeline", use: "Timeline with calendar icons", props: "milestones, y" },
  { name: "TableWithChecklist", use: "Table with checkmark column", props: "headers, rows, notes, y" },
  { name: "TablePieCombo", use: "Tables + pie chart side by side", props: "tables, pieData, pieTitle, y" },
  { name: "TableWithNotes", use: "Table with notes on right side", props: "headers, rows, notes, y" },
  { name: "DashboardWithNotes", use: "Dashboard with notes and footer", props: "charts, notes, footer, y" },
  { name: "ClosingSlideWhite", use: "Thank You slide on white bg", props: "cta" },
];

module.exports = {
  ...components,
  registry,
  P, // Re-export primitives for convenience
};
