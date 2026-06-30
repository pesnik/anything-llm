const brand = require("../brand");
const themes = require("./themes");
const shapes = require("./shapes");

function renderTitleSlide(slide, pptx, data) {
  const c = brand.colors;
  const t = brand.titleSlide;

  slide.background = { color: c.primaryOrange };

  slide.addText(data.title || "Presentation", {
    x: t.titleX,
    y: t.titleY,
    w: t.titleW,
    h: t.titleH,
    fontSize: brand.fontSize.heroTitle,
    bold: true,
    color: c.white,
    fontFace: brand.fonts.heading,
    align: "center",
    valign: "bottom",
  });

  shapes.addAccentBar(slide, pptx, 5.167, 2.85, 3.0, c.amber);

  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: t.subtitleX,
      y: t.subtitleY,
      w: t.subtitleW,
      h: t.subtitleH,
      fontSize: brand.fontSize.subtitle,
      color: c.white,
      fontFace: brand.fonts.body,
      align: "center",
      valign: "top",
    });
  }

  shapes.addLogoFull(slide);
}

function renderSectionSlide(slide, pptx, data) {
  const c = brand.colors;
  const sd = brand.sectionDivider;

  slide.background = { color: c.white };

  shapes.addBrandedRect(slide, pptx, 0, 0, sd.orangePanelW, brand.slideSize.height, c.primaryOrange);

  if (data.ghostSwoosh) {
    slide.addShape(pptx.ShapeType.rect, {
      x: sd.orangePanelW - 0.5,
      y: 0,
      w: 1.0,
      h: brand.slideSize.height,
      fill: { color: c.primaryOrange, transparency: 30 },
      line: { color: c.primaryOrange, width: 0 },
    });
  }

  slide.addText(data.title || "Section", {
    x: sd.titleX,
    y: sd.titleY,
    w: sd.titleW,
    h: sd.titleH,
    fontSize: brand.fontSize.sectionTitle,
    bold: true,
    color: c.nearBlack,
    fontFace: brand.fonts.heading,
    align: "left",
    valign: "middle",
  });

  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: sd.titleX,
      y: sd.titleY + 1.2,
      w: sd.titleW,
      h: 0.6,
      fontSize: brand.fontSize.body,
      color: c.midGray,
      fontFace: brand.fonts.body,
      align: "left",
      valign: "top",
    });
  }

  shapes.addLogoIcon(slide, pptx);
}

function renderBulletSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: 0.22,
      w: l.contentW,
      h: 0.55,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      valign: "top",
    });
  }

  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: l.marginX,
      y: 0.78,
      w: l.contentW,
      h: 0.3,
      fontSize: 13,
      color: c.darkGray,
      fontFace: brand.fonts.body,
      valign: "top",
    });
  }

  let contentY = data.subtitle ? 1.15 : 0.9;
  shapes.addBulletList(slide, data.content || [], l.marginX, contentY, l.contentW, 4.2);

  shapes.addLogoIcon(slide, pptx);
}

function renderTwoColumnSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: l.marginY,
      w: l.contentW,
      h: 0.65,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.nearBlack,
      fontFace: brand.fonts.heading,
      valign: "bottom",
    });

    shapes.addAccentBar(slide, pptx, l.marginX, 1.15, 1.5, c.primaryOrange);
  }

  const colY = 1.5;
  const colH = 5.0;
  const colW = (l.contentW - 0.5) / 2;
  const colGap = 0.5;

  if (data.leftContent) {
    shapes.addBulletList(slide, data.leftContent, l.marginX, colY, colW, colH);
  }

  if (data.rightContent) {
    const rightX = l.marginX + colW + colGap;

    if (typeof data.rightContent === "string") {
      slide.addText(data.rightContent, {
        x: rightX,
        y: colY,
        w: colW,
        h: colH,
        fontSize: brand.fontSize.body,
        color: c.nearBlack,
        fontFace: brand.fonts.body,
        valign: "top",
      });
    } else if (Array.isArray(data.rightContent)) {
      shapes.addBulletList(slide, data.rightContent, rightX, colY, colW, colH);
    }
  }

  shapes.addLogoIcon(slide, pptx);
}

function renderTableSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: l.marginY,
      w: l.contentW,
      h: 0.65,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.nearBlack,
      fontFace: brand.fonts.heading,
      valign: "bottom",
    });

    shapes.addAccentBar(slide, pptx, l.marginX, 1.15, 1.5, c.primaryOrange);
  }

  shapes.addBrandedTable(
    slide,
    pptx,
    data.headers,
    data.rows,
    l.marginX,
    1.5,
    l.contentW,
    data.colWidths || null
  );

  shapes.addLogoIcon(slide, pptx);
}

function renderQuoteSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  shapes.addBrandedRect(slide, pptx, l.marginX, 1.5, 0.08, 3.0, c.primaryOrange);

  if (data.quote) {
    slide.addText(`"${data.quote}"`, {
      x: l.marginX + 0.4,
      y: 1.8,
      w: l.contentW - 0.4,
      h: 2.0,
      fontSize: brand.fontSize.subtitle,
      italic: true,
      color: c.nearBlack,
      fontFace: brand.fonts.body,
      align: "left",
      valign: "top",
      wrap: true,
    });
  }

  if (data.attribution) {
    slide.addText(`— ${data.attribution}`, {
      x: l.marginX + 0.4,
      y: 3.8,
      w: l.contentW - 0.4,
      h: 0.5,
      fontSize: brand.fontSize.body,
      color: c.midGray,
      fontFace: brand.fonts.body,
      align: "left",
      valign: "top",
    });
  }

  shapes.addLogoIcon(slide, pptx);
}

function renderClosingSlide(slide, pptx, data) {
  const c = brand.colors;

  slide.background = { color: c.dark };
  slide._isClosing = true;

  slide.addText(data.cta || "Thank You", {
    x: 1.0,
    y: 2.0,
    w: 11.333,
    h: 1.5,
    fontSize: 44,
    bold: true,
    color: c.white,
    fontFace: brand.fonts.heading,
    align: "center",
    valign: "middle",
  });

  if (data.contact) {
    slide.addText(data.contact, {
      x: 1.0,
      y: 3.8,
      w: 11.333,
      h: 0.8,
      fontSize: brand.fontSize.body,
      color: c.white,
      fontFace: brand.fonts.body,
      align: "center",
      valign: "top",
    });
  }

  shapes.addLogoFull(slide);
}

function renderKpiSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;
  const kpi = brand.kpiSlide;

  slide.background = { color: c.white };

  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: l.marginY,
      w: l.contentW,
      h: 0.65,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      valign: "bottom",
    });
  }

  const kpis = data.kpis || [];
  const count = Math.min(kpis.length, 4);
  if (count === 0) return;

  const totalW = kpi.cardW * count + kpi.cardGap * (count - 1);
  const startX = (brand.slideSize.width - totalW) / 2;

  for (let i = 0; i < count; i++) {
    const k = kpis[i];
    const x = startX + i * (kpi.cardW + kpi.cardGap);

    shapes.addBrandedRect(slide, pptx, x, kpi.cardY, kpi.cardW, kpi.cardH, c.white, {
      line: { color: c.swooshOrange, width: 1 },
    });

    slide.addText(k.label || "", {
      x: x + 0.2,
      y: kpi.cardY + 0.3,
      w: kpi.cardW - 0.4,
      h: 0.5,
      fontSize: brand.fontSize.caption,
      color: c.nearBlack,
      fontFace: brand.fonts.body,
      bold: true,
      align: "center",
      valign: "top",
    });

    slide.addText(k.value || "", {
      x: x + 0.2,
      y: kpi.cardY + 0.9,
      w: kpi.cardW - 0.4,
      h: 1.0,
      fontSize: kpi.numberFontSize,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      bold: true,
      align: "center",
      valign: "middle",
    });

    if (k.sublabel) {
      slide.addText(k.sublabel, {
        x: x + 0.2,
        y: kpi.cardY + kpi.cardH - 0.6,
        w: kpi.cardW - 0.4,
        h: 0.4,
        fontSize: brand.fontSize.footnote,
        color: c.midGray,
        fontFace: brand.fonts.body,
        align: "center",
        valign: "bottom",
      });
    }
  }

  shapes.addLogoIcon(slide, pptx);
}

function renderImageCardsSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;
  const ic = brand.imageCards;

  slide.background = { color: c.white };

  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: l.marginY,
      w: l.contentW,
      h: 0.65,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      valign: "bottom",
    });
  }

  const cards = data.cards || [];
  const count = Math.min(cards.length, 4);
  if (count === 0) return;

  const totalW = ic.cardW * count + ic.cardGap * (count - 1);
  const startX = (brand.slideSize.width - totalW) / 2;

  for (let i = 0; i < count; i++) {
    const card = cards[i];
    const x = startX + i * (ic.cardW + ic.cardGap);

    shapes.addBrandedRect(slide, pptx, x, ic.cardY, ic.cardW, ic.cardH, c.white, {
      line: { color: c.swooshOrange, width: 1 },
    });

    if (card.caption) {
      slide.addText(card.caption, {
        x: x,
        y: ic.cardY + ic.cardH + 0.15,
        w: ic.cardW,
        h: ic.captionH,
        fontSize: brand.fontSize.caption,
        color: c.midGray,
        fontFace: brand.fonts.body,
        align: "center",
        valign: "top",
        wrap: true,
      });
    }
  }

  shapes.addLogoIcon(slide, pptx);
}

function renderChartSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: l.marginY,
      w: l.contentW,
      h: 0.65,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      valign: "bottom",
    });
  }

  if (data.chartTitle) {
    slide.addText(data.chartTitle, {
      x: l.marginX,
      y: 1.3,
      w: l.contentW,
      h: 0.4,
      fontSize: brand.fontSize.body,
      color: c.nearBlack,
      fontFace: brand.fonts.body,
      valign: "top",
    });
  }

  const chartType = data.chartType || "bar";
  const chartData = data.chartData || [];

  if (chartData.length > 0) {
    const chartColors = [c.swooshOrange, c.deepOrange, c.amber, c.primaryOrange, c.checkmarkAmber];

    const chartOpts = {
      x: l.marginX,
      y: 1.8,
      w: 9.5,
      h: 5.0,
      showLegend: true,
      legendPos: "r",
      legendFontSize: brand.fontSize.footnote,
      showTitle: false,
      chartColors: chartColors,
      catAxisLabelFontSize: brand.fontSize.footnote,
      valAxisLabelFontSize: brand.fontSize.footnote,
      dataLabelFontSize: brand.fontSize.footnote,
      dataLabelColor: c.nearBlack,
      catAxisOrientation: "minMax",
      valGridLine: { style: "dash", color: c.dividerGray, size: 0.5 },
    };

    if (chartType === "pie" || chartType === "doughnut") {
      chartOpts.showLegend = true;
      chartOpts.legendPos = "r";
      delete chartOpts.valGridLine;
      delete chartOpts.catAxisOrientation;
    }

    if (chartType === "bar") {
      chartOpts.barDir = "bar";
    }

    slide.addChart(pptx.ChartType[chartType] || pptx.ChartType.bar, chartData, chartOpts);
  }

  shapes.addLogoIcon(slide, pptx);
}

function renderTimelineSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;
  const tl = brand.timeline;

  slide.background = { color: c.white };

  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: l.marginY,
      w: l.contentW,
      h: 0.65,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      valign: "bottom",
    });
  }

  const months = data.months || [];
  const count = Math.min(months.length, 6);
  if (count === 0) return;

  const totalW = tl.colW * count + tl.colGap * (count - 1);
  const startX = (brand.slideSize.width - totalW) / 2;

  for (let i = 0; i < count; i++) {
    const m = months[i];
    const x = startX + i * (tl.colW + tl.colGap);

    shapes.addBrandedRect(slide, pptx, x, tl.colY, tl.colW, tl.iconSize, c.swooshOrange);

    slide.addText(m.label || "", {
      x: x,
      y: tl.colY,
      w: tl.colW,
      h: tl.iconSize,
      fontSize: brand.fontSize.caption,
      color: c.white,
      fontFace: brand.fonts.body,
      bold: true,
      align: "center",
      valign: "middle",
    });

    shapes.addBrandedRect(slide, pptx, x, tl.colY + tl.iconSize + 0.1, tl.colW, tl.colH, c.lightGray);

    slide.addText(m.details || "", {
      x: x + 0.15,
      y: tl.colY + tl.iconSize + 0.25,
      w: tl.colW - 0.3,
      h: tl.colH - 0.3,
      fontSize: brand.fontSize.footnote,
      color: c.nearBlack,
      fontFace: brand.fonts.body,
      valign: "top",
      wrap: true,
    });
  }

  shapes.addLogoIcon(slide, pptx);
}

function renderBulletImageSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: l.marginY,
      w: l.contentW,
      h: 0.65,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      valign: "bottom",
    });
  }

  const bulletW = l.contentW * 0.45;
  const imgX = l.marginX + bulletW + 0.5;
  const imgW = l.contentW * 0.5;

  if (data.content) {
    shapes.addBulletList(slide, data.content, l.marginX, 1.5, bulletW, 5.0);
  }

  shapes.addBrandedRect(slide, pptx, imgX, 1.5, imgW, 4.5, c.white, {
    line: { color: c.swooshOrange, width: 1 },
  });

  shapes.addLogoIcon(slide, pptx);
}

function renderSectionBreakSlide(slide, pptx, data) {
  const c = brand.colors;

  slide.background = { color: c.swooshOrange };

  slide.addText(data.title || "Section Break", {
    x: 1.0,
    y: 2.5,
    w: 11.333,
    h: 2.0,
    fontSize: brand.fontSize.sectionTitle,
    bold: true,
    color: c.white,
    fontFace: brand.fonts.heading,
    align: "center",
    valign: "middle",
  });

  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: 1.0,
      y: 4.5,
      w: 11.333,
      h: 0.8,
      fontSize: brand.fontSize.body,
      color: c.white,
      fontFace: brand.fonts.body,
      align: "center",
      valign: "top",
    });
  }

  shapes.addLogoFull(slide);
}

function renderSectionIntroSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: 2.0,
      w: l.contentW,
      h: 1.5,
      fontSize: brand.fontSize.sectionTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      align: "left",
      valign: "middle",
    });
  }

  if (data.subtitle) {
    slide.addText(data.subtitle, {
      x: l.marginX,
      y: 3.8,
      w: l.contentW,
      h: 1.0,
      fontSize: brand.fontSize.body,
      color: c.midGray,
      fontFace: brand.fonts.body,
      align: "left",
      valign: "top",
    });
  }

  shapes.addLogoIcon(slide, pptx);
}

// ── Numbered List Slide ──
function renderNumberedListSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  // Title
  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: 0.22,
      w: l.contentW,
      h: 0.55,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      align: "left",
      valign: "top",
    });
  }

  // Two columns with numbered items
  const colW = (l.contentW - 0.5) / 2;
  const leftX = l.marginX;
  const rightX = l.marginX + colW + 0.5;
  const startY = 0.9;

  const renderColumn = (items, x) => {
    if (!items) return;
    items.forEach((item, i) => {
      const num = String(i + 1).padStart(2, "0");
      slide.addText(num, {
        x,
        y: startY + i * 0.5,
        w: 0.5,
        h: 0.4,
        fontSize: 14,
        bold: true,
        color: c.swooshOrange,
        fontFace: brand.fonts.heading,
      });
      slide.addText(item, {
        x: x + 0.55,
        y: startY + i * 0.5,
        w: colW - 0.6,
        h: 0.4,
        fontSize: 13,
        color: c.nearBlack,
        fontFace: brand.fonts.body,
      });
    });
  };

  renderColumn(data.left, leftX);
  renderColumn(data.right, rightX);

  shapes.addLogoIcon(slide, pptx);
}

// ── KPI with Headline Slide ──
function renderKPIHeadlineSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  // Title
  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: 0.22,
      w: l.contentW,
      h: 0.55,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      align: "left",
      valign: "top",
    });
  }

  // KPI cards
  if (data.metrics && data.metrics.length > 0) {
    const count = Math.min(data.metrics.length, 4);
    const cardW = 2.8;
    const cardH = 2.8;
    const gap = 0.4;
    const totalW = cardW * count + gap * (count - 1);
    const startX = (l.contentW - totalW) / 2 + l.marginX;
    const startY = 0.9;

    for (let i = 0; i < count; i++) {
      const m = data.metrics[i];
      const x = startX + i * (cardW + gap);

      // Card background
      slide.addShape(pptx.ShapeType.rect, {
        x,
        y: startY,
        w: cardW,
        h: cardH,
        fill: { color: c.white },
        line: { color: c.swooshOrange, width: 1 },
      });

      // Headline
      slide.addText(m.headline || m.label || "", {
        x: x + 0.2,
        y: startY + 0.2,
        w: cardW - 0.4,
        h: 0.8,
        fontSize: 14,
        bold: true,
        color: c.nearBlack,
        fontFace: brand.fonts.body,
        align: "center",
        valign: "top",
      });

      // Big number
      slide.addText(m.value || "", {
        x: x + 0.2,
        y: startY + 1.0,
        w: cardW - 0.4,
        h: 1.0,
        fontSize: 48,
        bold: true,
        color: c.swooshOrange,
        fontFace: brand.fonts.heading,
        align: "center",
        valign: "middle",
      });

      // Date
      if (m.date) {
        slide.addText(m.date, {
          x: x + 0.2,
          y: startY + cardH - 0.5,
          w: cardW - 0.4,
          h: 0.4,
          fontSize: 12,
          color: c.swooshOrange,
          fontFace: brand.fonts.body,
          align: "center",
          valign: "bottom",
        });
      }
    }
  }

  shapes.addLogoIcon(slide, pptx);
}

// ── Large Image Placeholder Slide ──
function renderLargeImageSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  // Title
  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: 0.22,
      w: l.contentW,
      h: 0.55,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      align: "left",
      valign: "top",
    });
  }

  // Large orange-bordered rectangle
  slide.addShape(pptx.ShapeType.rect, {
    x: 1.5,
    y: 0.9,
    w: 10.0,
    h: 4.5,
    fill: { color: c.white },
    line: { color: c.swooshOrange, width: 1 },
  });

  // Caption below
  if (data.caption) {
    slide.addText(data.caption, {
      x: 1.5,
      y: 5.6,
      w: 10.0,
      h: 0.8,
      fontSize: 13,
      color: c.nearBlack,
      fontFace: brand.fonts.body,
      wrap: true,
    });
  }

  shapes.addLogoIcon(slide, pptx);
}

// ── Dashboard Slide (2x2 mini charts) ──
function renderDashboardSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  // Title
  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: 0.22,
      w: l.contentW,
      h: 0.55,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      align: "left",
      valign: "top",
    });
  }

  // Dashboard charts
  if (data.charts && data.charts.length > 0) {
    const count = Math.min(data.charts.length, 4);
    const cellW = (l.contentW - 0.3) / 2;
    const cellH = 2.8;
    const startY = 0.9;

    for (let i = 0; i < count; i++) {
      const ch = data.charts[i];
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = l.marginX + col * (cellW + 0.3);
      const y = startY + row * (cellH + 0.3);

      // Cell background
      slide.addShape(pptx.ShapeType.rect, {
        x,
        y,
        w: cellW,
        h: cellH,
        fill: { color: c.white },
        line: { color: c.divider, width: 0.5 },
      });

      // Chart title
      if (ch.title) {
        slide.addText(ch.title, {
          x: x + 0.2,
          y: y + 0.1,
          w: cellW - 0.4,
          h: 0.4,
          fontSize: 11,
          bold: true,
          color: c.nearBlack,
          fontFace: brand.fonts.body,
        });
      }

      // Chart subtitle
      if (ch.subtitle) {
        slide.addText(ch.subtitle, {
          x: x + 0.2,
          y: y + 0.4,
          w: cellW - 0.4,
          h: 0.3,
          fontSize: 9,
          color: c.midGray,
          fontFace: brand.fonts.body,
        });
      }

      // Mini chart
      if (ch.data) {
        const chartType = ch.type || "bar";
        const chartColors = [c.swooshOrange, c.deepOrange, c.amber, c.primaryOrange, c.checkmarkAmber];
        slide.addChart(pptx.ChartType[chartType] || pptx.ChartType.bar, ch.data, {
          x: x + 0.2,
          y: y + 0.7,
          w: cellW - 0.4,
          h: cellH - 0.9,
          showLegend: false,
          showTitle: false,
          chartColors: chartColors,
          catAxisLabelFontSize: 8,
          valAxisLabelFontSize: 8,
          dataLabelFontSize: 8,
        });
      }
    }
  }

  shapes.addLogoIcon(slide, pptx);
}

// ── Calendar Timeline Slide ──
function renderCalendarTimelineSlide(slide, pptx, data) {
  const c = brand.colors;
  const l = brand.contentSlide;

  slide.background = { color: c.white };

  // Title
  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: 0.22,
      w: l.contentW,
      h: 0.55,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.swooshOrange,
      fontFace: brand.fonts.heading,
      align: "left",
      valign: "top",
    });
  }

  // Calendar timeline
  if (data.milestones && data.milestones.length > 0) {
    const count = Math.min(data.milestones.length, 6);
    const colW = 1.7;
    const colGap = 0.15;
    const totalW = colW * count + colGap * (count - 1);
    const startX = (l.contentW - totalW) / 2 + l.marginX;
    const startY = 0.9;

    for (let i = 0; i < count; i++) {
      const m = data.milestones[i];
      const x = startX + i * (colW + colGap);

      // Calendar icon (simplified: orange top bar + white box)
      slide.addShape(pptx.ShapeType.rect, {
        x: x + 0.35,
        y: startY,
        w: colW - 0.7,
        h: 0.15,
        fill: { color: c.swooshOrange },
      });
      slide.addShape(pptx.ShapeType.rect, {
        x: x + 0.3,
        y: startY + 0.15,
        w: colW - 0.6,
        h: 0.8,
        fill: { color: c.white },
        line: { color: c.divider, width: 0.5 },
      });
      slide.addText(m.label || "", {
        x: x + 0.3,
        y: startY + 0.15,
        w: colW - 0.6,
        h: 0.8,
        fontSize: 14,
        bold: true,
        color: c.nearBlack,
        fontFace: brand.fonts.body,
        align: "center",
        valign: "middle",
      });

      // Details card (gradient gray)
      const grayShade = Math.floor(200 + (i / count) * 40);
      const grayHex = grayShade.toString(16).padStart(2, "0");
      slide.addShape(pptx.ShapeType.rect, {
        x,
        y: startY + 1.1,
        w: colW,
        h: 4.5,
        fill: { color: grayHex + grayHex + grayHex },
      });

      // Month + details
      slide.addText(m.month || m.label || "", {
        x: x + 0.15,
        y: startY + 1.2,
        w: colW - 0.3,
        h: 0.4,
        fontSize: 12,
        bold: true,
        color: c.nearBlack,
        fontFace: brand.fonts.body,
      });
      slide.addText(m.details || m.desc || "", {
        x: x + 0.15,
        y: startY + 1.6,
        w: colW - 0.3,
        h: 3.8,
        fontSize: 11,
        color: c.nearBlack,
        fontFace: brand.fonts.body,
        wrap: true,
      });
    }
  }

  shapes.addLogoIcon(slide, pptx);
}

module.exports = {
  renderTitleSlide,
  renderSectionSlide,
  renderBulletSlide,
  renderTwoColumnSlide,
  renderTableSlide,
  renderQuoteSlide,
  renderClosingSlide,
  renderKpiSlide,
  renderImageCardsSlide,
  renderChartSlide,
  renderTimelineSlide,
  renderBulletImageSlide,
  renderSectionBreakSlide,
  renderSectionIntroSlide,
  renderNumberedListSlide,
  renderKPIHeadlineSlide,
  renderLargeImageSlide,
  renderDashboardSlide,
  renderCalendarTimelineSlide,
};
