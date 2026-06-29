const brand = require("../brand");

function hexToRgbObj(hex) {
  const h = hex.replace("#", "");
  return {
    r: parseInt(h.substr(0, 2), 16),
    g: parseInt(h.substr(2, 2), 16),
    b: parseInt(h.substr(4, 2), 16),
  };
}

function buildTheme(slideType, themeMode) {
  const isDark = themeMode === "dark";
  const c = brand.colors;

  return {
    background: isDark ? c.nearBlack : c.white,
    textColor: isDark ? c.white : c.nearBlack,
    accentColor: c.primaryOrange,
    accentColorRGB: hexToRgbObj(c.primaryOrange),
    amberColor: c.amber,
    amberColorRGB: hexToRgbObj(c.amber),
    deepOrangeColor: c.deepOrange,
    midGrayColor: c.midGray,
    lightGrayColor: c.lightGray,
    slideTitleSize: brand.fontSize.slideTitle,
    bodySize: brand.fontSize.body,
    fontBody: brand.fonts.body,
    fontHeading: brand.fonts.heading,
  };
}

function buildTableTheme() {
  return {
    headerFill: brand.colors.amber,
    headerColor: brand.colors.white,
    headerFontSize: 14,
    headerBold: true,
    bodyFill: brand.colors.white,
    bodyAltFill: brand.colors.lightGray,
    bodyColor: brand.colors.nearBlack,
    bodyFontSize: 13,
    borderColor: brand.colors.dividerGray,
  };
}

module.exports = {
  buildTheme,
  buildTableTheme,
  hexToRgbObj,
};
