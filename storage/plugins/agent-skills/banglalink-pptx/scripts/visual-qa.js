#!/usr/bin/env node
/**
 * Visual QA script for Banglalink PPTX.
 * Converts presentation to images for visual inspection.
 *
 * Usage:
 *   node scripts/visual-qa.js <input.pptx> [output-dir]
 *
 * Requires: LibreOffice (soffice) and Poppler (pdftoppm)
 *   brew install libreoffice poppler
 */

const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const input = process.argv[2];
const outDir = process.argv[3] || path.join(path.dirname(input), ".qa-output");

if (!input) {
  console.error("Usage: node visual-qa.js <input.pptx> [output-dir]");
  process.exit(1);
}

if (!fs.existsSync(input)) {
  console.error(`File not found: ${input}`);
  process.exit(1);
}

// Create output directory
if (!fs.existsSync(outDir)) {
  fs.mkdirSync(outDir, { recursive: true });
}

const pdfName = path.basename(input, path.extname(input)) + ".pdf";
const pdfPath = path.join(outDir, pdfName);
const imgPrefix = path.join(outDir, "slide");

console.log(`\n  Visual QA Pipeline`);
console.log(`  Input: ${input}`);
console.log(`  Output: ${outDir}\n`);

// Step 1: Convert PPTX → PDF via LibreOffice
console.log("  1. Converting PPTX → PDF (LibreOffice)...");
try {
  execSync(`soffice --headless --convert-to pdf --outdir "${outDir}" "${input}"`, {
    stdio: "pipe",
    timeout: 60000,
  });
  if (!fs.existsSync(pdfPath)) {
    console.error("  ✗ PDF conversion failed — LibreOffice may not be installed");
    console.error("    Install: brew install libreoffice");
    process.exit(1);
  }
  console.log("    ✓ PDF created");
} catch (e) {
  console.error("  ✗ LibreOffice conversion failed:", e.message);
  console.error("    Install: brew install libreoffice");
  process.exit(1);
}

// Step 2: Convert PDF → JPEG images via Poppler
console.log("  2. Converting PDF → JPEG images (Poppler)...");
try {
  execSync(`pdftoppm -jpeg -r 150 "${pdfPath}" "${imgPrefix}"`, {
    stdio: "pipe",
    timeout: 60000,
  });
  const images = fs.readdirSync(outDir).filter((f) => f.startsWith("slide") && f.endsWith(".jpg"));
  console.log(`    ✓ ${images.length} slide images created`);
} catch (e) {
  console.error("  ✗ Poppler conversion failed:", e.message);
  console.error("    Install: brew install poppler");
  process.exit(1);
}

// Step 3: Create thumbnail grid
console.log("  3. Creating thumbnail grid...");
const images = fs.readdirSync(outDir)
  .filter((f) => f.startsWith("slide") && f.endsWith(".jpg"))
  .sort((a, b) => {
    const na = parseInt(a.match(/slide-(\d+)/)?.[1] || "0", 10);
    const nb = parseInt(b.match(/slide-(\d+)/)?.[1] || "0", 10);
    return na - nb;
  });

if (images.length > 0) {
  try {
    const { createCanvas, loadImage } = require("canvas");
    const cols = Math.min(3, images.length);
    const rows = Math.ceil(images.length / cols);
    const thumbW = 400;
    const thumbH = 225;
    const padding = 20;

    const gridW = cols * thumbW + (cols + 1) * padding;
    const gridH = rows * (thumbH + 30 + padding) + padding;
    const canvas = createCanvas(gridW, gridH);
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, gridW, gridH);

    for (let i = 0; i < images.length; i++) {
      const row = Math.floor(i / cols);
      const col = i % cols;
      const x = padding + col * (thumbW + padding);
      const y = padding + row * (thumbH + 30 + padding);

      // Label
      ctx.fillStyle = "#000000";
      ctx.font = "14px Arial";
      ctx.textAlign = "center";
      ctx.fillText(images[i], x + thumbW / 2, y + 16);

      // Thumbnail
      loadImage(path.join(outDir, images[i])).then((img) => {
        ctx.drawImage(img, x, y + 24, thumbW, thumbH);
      });
    }

    const gridPath = path.join(outDir, "grid.jpg");
    const buffer = canvas.toBuffer("image/jpeg", { quality: 0.95 });
    fs.writeFileSync(gridPath, buffer);
    console.log(`    ✓ Grid created: ${gridPath}`);
  } catch (e) {
    console.log("    ⚠ canvas not available — skipping grid (install: npm install canvas)");
  }
}

// Step 4: Print QA checklist
console.log(`\n  ── QA Checklist ──`);
console.log(`  Open individual slide images in ${outDir}/ and check for:`);
console.log(`    □ Overlapping elements`);
console.log(`    □ Text overflow or cut-off`);
console.log(`    □ Low-contrast text/icons`);
console.log(`    □ Uneven spacing/gaps`);
console.log(`    □ Missing content`);
console.log(`    □ Brand consistency (orange titles, Arial font)`);
console.log(`    □ No placeholder text remaining`);
console.log(`\n  Slide images:`);
images.forEach((img) => console.log(`    ${path.join(outDir, img)}`));
console.log("");
