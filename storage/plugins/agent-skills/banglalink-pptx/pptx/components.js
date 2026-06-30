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
];

module.exports = {
  ...components,
  registry,
  P, // Re-export primitives for convenience
};
