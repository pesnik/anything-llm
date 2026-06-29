const path = require("path");
const fs = require("fs");

const GENERATED_FILES_DIR = path.join(
  __dirname, "..", "..", "..", "..", "storage", "generated-files"
);

function ensureGeneratedDir() {
  if (!fs.existsSync(GENERATED_FILES_DIR)) {
    fs.mkdirSync(GENERATED_FILES_DIR, { recursive: true });
  }
  return GENERATED_FILES_DIR;
}

function slugify(text) {
  return text
    .toLowerCase()
    .replace(/[^\w\s-]/g, "")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .slice(0, 100);
}

function validateSlideData(data) {
  if (!data) return false;
  const VALID_TYPES = ["title", "section", "bullet", "two-column", "twocolumn", "table", "quote", "closing"];

  if (data.slides && Array.isArray(data.slides)) {
    return data.slides.every((s) => VALID_TYPES.includes(s.type));
  }

  if (data.type && !VALID_TYPES.includes(data.type)) {
    return false;
  }

  return true;
}

function processHtmlContent(html, theme) {
  const HEX_MAP = {
    "0f172a": "171616",
    "1e293b": "1c1c1c",
    "334155": "2d2d2d",
  };

  for (const [oldHex, newHex] of Object.entries(HEX_MAP)) {
    html = html.replaceAll(oldHex, newHex);
  }

  const CSS_OVERRIDE = `
<style id="bl-brand">
.text-blue-100{color:#fff7ed!important}.text-blue-200{color:#fed7aa!important}
.text-blue-300{color:#fdba74!important}.text-blue-400{color:#f56f1f!important}
.text-blue-500{color:#f86c02!important}.text-blue-600{color:#ea580c!important}
.text-blue-700{color:#c2410c!important}
.text-cyan-200{color:#fde68a!important}.text-cyan-300{color:#fcd34d!important}
.text-cyan-400{color:#faa106!important}.text-cyan-500{color:#d97706!important}
.text-cyan-600{color:#b45309!important}
.text-teal-200{color:#fdba74!important}.text-teal-300{color:#fb923c!important}
.text-teal-400{color:#f86c02!important}.text-teal-500{color:#ea580c!important}
.bg-blue-500{background-color:#f86c02!important}
.bg-blue-600{background-color:#ea580c!important}
.bg-cyan-600{background-color:#b45309!important}
.border-blue-300{border-color:#fdba74!important}.border-blue-400{border-color:#f56f1f!important}
.border-blue-500{border-color:#f86c02!important}
.border-cyan-300{border-color:#fcd34d!important}.border-cyan-400{border-color:#faa106!important}
</style>`;

  html = html.replace("</body>", CSS_OVERRIDE + "\n</body>");

  if (theme === "light") {
    const LIGHT_OVERRIDE = `
<style id="bl-theme-light">
body, .ppt-slide { background: #FFFFFF !important; background-image: none !important; }
.text-white, .text-gray-100, .text-gray-200, .text-gray-300 { color: #171616 !important; }
.bg-white { background-color: #F2F2F2 !important; border-color: #E2E2E2 !important; }
</style>`;
    html = html.replace("</body>", LIGHT_OVERRIDE + "\n</body>");
  }

  return html;
}

module.exports = {
  ensureGeneratedDir,
  slugify,
  validateSlideData,
  processHtmlContent,
  GENERATED_FILES_DIR,
};
