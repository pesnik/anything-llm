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

  shapes.addBrandedRect(slide, pptx, 0, 0, "100%", l.headerHeight, c.primaryOrange);

  if (data.title) {
    slide.addText(data.title, {
      x: l.marginX,
      y: 0.1,
      w: l.contentW,
      h: 0.6,
      fontSize: brand.fontSize.slideTitle,
      bold: true,
      color: c.white,
      fontFace: brand.fonts.heading,
      valign: "middle",
    });
  }

  let contentY = l.headerHeight + 0.3;
  shapes.addBulletList(slide, data.content || [], l.marginX, contentY, l.contentW, 4.5);

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

  slide.background = { color: c.primaryOrange };
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

module.exports = {
  renderTitleSlide,
  renderSectionSlide,
  renderBulletSlide,
  renderTwoColumnSlide,
  renderTableSlide,
  renderQuoteSlide,
  renderClosingSlide,
};
