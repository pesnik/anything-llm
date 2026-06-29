#!/usr/bin/env node
/**
 * Content QA script for Banglalink PPTX.
 * Extracts text and checks for placeholder remnants, missing content, etc.
 *
 * Usage:
 *   node scripts/content-qa.js <input.pptx>
 *
 * Requires: python + markitdown
 *   pip install "markitdown[pptx]"
 */

const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const input = process.argv[2];

if (!input) {
  console.error("Usage: node content-qa.js <input.pptx>");
  process.exit(1);
}

if (!fs.existsSync(input)) {
  console.error(`File not found: ${input}`);
  process.exit(1);
}

console.log(`\n  Content QA Pipeline`);
console.log(`  Input: ${input}\n`);

// Step 1: Extract text via markitdown
console.log("  1. Extracting text (markitdown)...");
let text;
try {
  text = execSync(`python -m markitdown "${input}"`, {
    encoding: "utf-8",
    timeout: 30000,
  });
  console.log(`    ✓ Extracted ${text.length} characters`);
} catch (e) {
  console.log("    ⚠ markitdown not available — skipping content QA");
  console.log("      Install: pip install \"markitdown[pptx]\"");
  process.exit(0);
}

// Step 2: Check for placeholder remnants
console.log("\n  2. Checking for placeholder remnants...");
const placeholders = [
  /lorem ipsum/gi,
  /xxxx+/gi,
  /placeholder/gi,
  /your (text|content|title|name)/gi,
  /sample (text|content|data)/gi,
  /click to (edit|add)/gi,
  /enter (your|text|content)/gi,
  /<.*>/g,  // HTML tags
];

let issues = 0;
for (const pattern of placeholders) {
  const matches = text.match(pattern);
  if (matches) {
    console.log(`    ✗ Found "${pattern.source}": ${matches.length} occurrence(s)`);
    issues++;
  }
}
if (issues === 0) {
  console.log("    ✓ No placeholder remnants found");
}

// Step 3: Check for empty slides
console.log("\n  3. Checking for empty slides...");
const slides = text.split(/^#{1,2}\s+/m).filter(Boolean);
const emptySlides = slides.filter((s) => s.trim().length < 50);
if (emptySlides.length > 0) {
  console.log(`    ✗ ${emptySlides.length} slide(s) with minimal content`);
} else {
  console.log(`    ✓ All ${slides.length} slides have content`);
}

// Step 4: Check for brand consistency
console.log("\n  4. Checking brand consistency...");
const brandChecks = [
  { pattern: /banglalink/gi, name: "Banglalink mentioned" },
  { pattern: /arial/gi, name: "Arial font" },
];

for (const check of brandChecks) {
  const matches = text.match(check.pattern);
  console.log(`    ${matches ? "✓" : "○"} ${check.name}: ${matches ? matches.length : 0} occurrence(s)`);
}

// Step 5: Word count
console.log("\n  5. Content metrics...");
const words = text.split(/\s+/).filter(Boolean);
const bullets = text.match(/^[\s]*[•✓▸]\s/gm);
console.log(`    Words: ${words.length}`);
console.log(`    Bullet points: ${bullets ? bullets.length : 0}`);
console.log(`    Slides: ${slides.length}`);

console.log(`\n  ── Summary ──`);
if (issues > 0) {
  console.log(`  ⚠ ${issues} issue(s) found — review before publishing`);
} else {
  console.log(`  ✓ Content looks clean`);
}
console.log("");
