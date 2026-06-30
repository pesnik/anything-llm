/**
 * Comprehensive content-verification test for the Banglalink PPTX pipeline.
 * Asserts structural validity, slide content, branding, and ghost logo injection.
 *
 * Usage:  node test-pipeline.js
 *         node test-pipeline.js --save   # keep output file for manual inspection
 */

const path = require("path");
const fs = require("fs");
const JSZip = require("jszip");

const builder = require("./pptx/builder");

// ── Assertion helpers ──────────────────────────────────────────
let passed = 0;
let failed = 0;

function assert(label, condition, detail) {
  if (condition) {
    passed++;
    console.log(`  ✓ ${label}`);
  } else {
    failed++;
    console.log(`  ✗ ${label}  ${detail ? "-- " + detail : ""}`);
  }
}

function naturalSortSlideFiles(files) {
  return [...files].sort((a, b) => {
    const na = parseInt(a.match(/slide(\d+)/)[1], 10);
    const nb = parseInt(b.match(/slide(\d+)/)[1], 10);
    return na - nb;
  });
}

function decodeXmlEntities(str) {
  return str
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&#x201[89];/g, "'");
}

function getSlideTexts(slideXmls) {
  const result = {};
  for (const [file, xml] of Object.entries(slideXmls)) {
    const texts = (xml.match(/<a:t[^>]*>([^<]*)<\/a:t>/g) || [])
      .map((t) => t.replace(/<[^>]*>/g, ""))
      .map(decodeXmlEntities);
    result[file] = texts;
  }
  return result;
}

// ── Sample input (the user's exact LLM call) ──────────────────
const SAMPLE_ARG = {
  topic: "Banglalink Offer & Plan Strategy – Smart Value, Simple Life",
  sections: [
    {
      title: "Executive Summary",
      topic: "Strategic Overview",
      bulletPoints: [
        "GP hides prices behind personalization; BL exposes all prices publicly → Price War invitation",
        "Shift from Volume Leadership to Value Leadership",
        "Pivot from 'More GB for Less Tk' to 'Right GB for Right Life'",
        "Goal: Defend Market Share, Grow ARPU, Build Digital Loyalty Moat",
      ],
    },
    {
      title: "Situation Analysis: Tale of Two Architectures",
      topic: "Competitive Landscape Comparison",
      bulletPoints: [
        "GP: Curated & Segmented (6 clear tiles) | BL: Cluttered & Transactional (30+ SKUs)",
        "GP: Sells Experience & Status | BL: Sells Commodity (MBs/Taka)",
        "GP: High lock-in via GP Star, MyGP App, Postpaid Ecosystem | BL: Low lock-in, transactional recharge",
      ],
    },
    {
      title: "Core Strategy: Smart Value, Simple Life",
      topic: "Strategic Positioning",
      bulletPoints: [
        "12 Core SKUs (Good-Better-Best framework) – Kill the long tail of 30+ packs",
        "BL Star Loyalty Ecosystem – Tiered rewards (Silver/Gold/Platinum/Diamond)",
      ],
    },
    {
      title: "Pillar 1: Radical Portfolio Simplification",
      topic: "Good-Better-Best Framework",
      bulletPoints: [
        "Starter: 1GB/1Day (Tk19), 3GB/3Day (Tk49), 7GB/7Day (Tk99)",
        "Core Monthly (Good/Better/Best): 12GB (Tk299), 25GB (Tk449), 50GB (Tk699)",
      ],
    },
    {
      title: "Pillar 2: BL Star Loyalty Ecosystem",
      topic: "Counter to GP Star",
      bulletPoints: [
        "Silver/Gold/Platinum/Diamond tiers based on tenure, ARPU & app engagement",
      ],
    },
    {
      title: "Pillar 3: Personalization & My Offers Engine",
      topic: "Counter to GP 'My Offers'",
      bulletPoints: [
        "Churn Risk (Low Usage 7d): Push 'Comeback Pack' (5GB @ Tk29) via SMS/App",
        "High Value (Heavy Streamer): Push 'Unlimited YouTube Night Pack' @ Tk49",
      ],
    },
    {
      title: "Pillar 4: Postpaid & Convergence Push",
      topic: "BL Infinity Launch",
      bulletPoints: [
        "Infinity Solo: Tk599/mo – 40GB + Unlimited Voice + Toffee Premium + 5G Ready SIM",
      ],
    },
    {
      title: "Tactical Quick Wins (0-90 Days)",
      topic: "Immediate Execution Plan",
      bulletPoints: [
        "Data Rollover on Core Packs – High perceived value, low actual cost (breakage <15%)",
      ],
    },
    {
      title: "KPI Dashboard & Scorecard",
      topic: "Strategy Health Metrics",
      bulletPoints: [
        "Prepaid ARPU: Tk180 → Tk210 | Postpaid Revenue Mix: <5% → >12%",
      ],
    },
    {
      title: "Risk Mitigation & GP Reaction Plan",
      topic: "Contingency Strategies",
      bulletPoints: [
        "Nuclear Option: Launch 'BL Unlimited Social' (FB/Insta/WA/YouTube/TikTok Unlimited + 10GB @ Tk349)",
      ],
    },
    {
      title: "Roadmap & Next Steps",
      topic: "Implementation Timeline",
      bulletPoints: [
        "Month 9-12: 5G Experience Packs + Full BL Star Ecosystem Rollout",
      ],
    },
  ],
  audience: "Chief Marketing Officer, CEO, Board of Directors – Banglalink Digital",
  outputFileName: "Banglalink_Offer_Plan_Strategy_Smart_Value_Simple_Life",
  theme: "dark",
  enableResearch: false,
};

// ── Helper: replicate handler's slide construction ─────────────
function buildSlideData(section) {
  return [
    {
      type: section.isTitle ? "title" : section.isClosing ? "closing" : "bullet",
      title: section.title || section.topic,
      subtitle: section.subtitle,
      content: section.bulletPoints || ["Content placeholder"],
      cta: section.isClosing ? "Thank You" : undefined,
    },
  ];
}

// ── Main test ──────────────────────────────────────────────────
async function runTest() {
  const saveOutput = process.argv.includes("--save");

  console.log("\n═════════════════════════════════════════════════════════");
  console.log("  Banglalink PPTX Pipeline – Content Verification Test");
  console.log("  " + new Date().toISOString());
  console.log("═════════════════════════════════════════════════════════\n");

  // ── Step 1: Build sections ──────────────────────────
  console.log("▶ 1. SECTION CONSTRUCTION");
  const sections = [];
  for (const section of SAMPLE_ARG.sections) {
    sections.push({ title: section.title || section.topic, slides: buildSlideData(section) });
  }
  assert("11 sections built", sections.length === 11);
  const totalSlides = sections.reduce((sum, s) => sum + (s.slides?.length || 0), 0);
  assert("1 slide per section (no research)", totalSlides === 11);

  // ── Step 2: Render via builder ──────────────────────
  console.log("\n▶ 2. PPTX GENERATION");
  let buffer;
  try {
    buffer = await builder.buildPresentation(sections, { theme: SAMPLE_ARG.theme });
  } catch (e) {
    console.error("BUILD FAILED:", e.message);
    process.exit(1);
  }
  assert("Buffer size > 10 KB", buffer.length > 10000, `got ${buffer.length} bytes`);

  // ── Step 3: Validate ZIP/PPTX structure ─────────────
  console.log("\n▶ 3. PPTX STRUCTURE");
  let zip;
  try {
    zip = await JSZip.loadAsync(buffer);
    assert("Valid ZIP/PPTX archive", true);
  } catch (e) {
    assert("Valid ZIP/PPTX archive", false, e.message);
    process.exit(1);
  }

  const ctXml = await zip.file("[Content_Types].xml").async("string");
  assert(
    "[Content_Types].xml declares pptx MIME",
    ctXml.includes("presentationml.presentation"),
  );

  const slideFiles = naturalSortSlideFiles(
    Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
  );
  assert("11 slide XML files", slideFiles.length === 11, `got ${slideFiles.length}`);

  const hasRel = !!zip.files["ppt/_rels/presentation.xml.rels"];
  assert("Presentation relationships exist", hasRel);

  const slideLayouts = Object.keys(zip.files).filter((n) =>
    n.startsWith("ppt/slideLayouts/"),
  );
  assert("Slide layouts exist", slideLayouts.length > 0);

  // ── Step 4: Read all slide text content (decoded) ──
  console.log("\n▶ 4. SLIDE CONTENT");

  const slideXmls = {};
  for (const sf of slideFiles) {
    slideXmls[sf] = await zip.file(sf).async("string");
  }
  const slideTexts = getSlideTexts(slideXmls);

  // Each section title must appear in decoded slide text
  for (const title of SAMPLE_ARG.sections.map((s) => s.title)) {
    const found = Object.values(slideTexts).some((texts) =>
      texts.some((t) => t.includes(title))
    );
    assert(
      `Section title "${title.substring(0, 40)}..." in a slide`,
      found,
    );
  }

  // Each bullet point prefix (first 25 chars) must appear in decoded slide text
  for (const section of SAMPLE_ARG.sections) {
    for (const bp of section.bulletPoints) {
      const prefix25 = bp.substring(0, 25);
      const found = Object.values(slideTexts).some((texts) =>
        texts.some((t) => t.includes(prefix25))
      );
      assert(
        `Bullet content "${prefix25}..." in slide text`,
        found,
      );
    }
  }

  // ── Step 5: Brand colors in raw slide XMLs ─────────
  console.log("\n▶ 5. BRAND COLORS & FONTS");
  const allXml = Object.values(slideXmls).join(" ");

  // Brand swoosh orange accent on every content slide (header bar)
  assert("Brand swoosh orange #EF6E23 in slides", allXml.includes("EF6E23"));
  // Near-black text color
  assert("Near-black #171616 text", allXml.includes("171616"));
  // White background / text
  assert("White #FFFFFF", allXml.includes("FFFFFF"));
  // Arial font
  assert("Arial typeface", allXml.includes("Arial"));

  // ── Step 6: Ghost logo injection ────────────────────
  console.log("\n▶ 6. GHOST LOGO INJECTION");
  const lastSlideFile = slideFiles[slideFiles.length - 1];
  const lastSlideXml = slideXmls[lastSlideFile];

  // Ghost logo color + alpha
  assert("Ghost logo color F8A01B in last slide", lastSlideXml.includes("F8A01B"));
  assert("Ghost logo 35% alpha (35000)", lastSlideXml.includes("35000"));
  // Ghost logo shape element
  assert("Ghost logo <p:sp> element", lastSlideXml.includes("<p:sp"));

  // Logo icon via r:embed on content slides (not first, not last — all are bullet type)
  for (let i = 0; i < slideFiles.length; i++) {
    const xml = slideXmls[slideFiles[i]];
    if (i === slideFiles.length - 1) continue; // ghost logo on last, no icon
    assert(
      `Logo icon (r:embed) on slide ${i + 1}`,
      xml.includes('r:embed'),
    );
  }

  // ── Step 7: First slide (bullet type, shows first section title) ──
  console.log("\n▶ 7. FIRST SLIDE");
  const firstXml = slideXmls[slideFiles[0]];
  assert("Swoosh orange header bar on first slide", firstXml.includes("EF6E23"));
  assert(
    "First section title on slide 1",
    slideTexts[slideFiles[0]].some((t) => t.includes("Executive Summary")),
  );

  // ── Step 8: Last slide (bullet type, shows last section title) ──
  console.log("\n▶ 8. LAST SLIDE");
  assert("Swoosh orange header bar on last slide", lastSlideXml.includes("EF6E23"));
  assert(
    "Last section title on slide 11",
    slideTexts[lastSlideFile].some((t) => t.includes("Roadmap & Next Steps")),
  );

  // ── Step 9: Explicit slide types (title + closing) ──
  console.log("\n▶ 9. EXPLICIT SLIDE TYPES");
  const explicitSections = [
    { title: "Test Deck", slides: [{ type: "title", title: "Test Deck", subtitle: "For QA" }] },
    { title: "Findings", slides: [{ type: "bullet", title: "Findings", content: ["All tests pass"] }] },
    { title: "End", slides: [{ type: "closing", cta: "Thank You", contact: "QA Team" }] },
  ];
  const explicitBuffer = await builder.buildPresentation(explicitSections, { theme: "dark" });
  const explicitZip = await JSZip.loadAsync(explicitBuffer);
  const explicitSlideFiles = naturalSortSlideFiles(
    Object.keys(explicitZip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
  );

  // Slide 1: title slide
  const e1 = await explicitZip.file(explicitSlideFiles[0]).async("string");
  assert("Title slide has white background", e1.includes("FFFFFF"));
  assert("Title text on title slide", e1.includes("Test Deck"));
  assert("Subtitle on title slide", e1.includes("For QA"));

  // Slide 2: bullet slide
  const e2 = await explicitZip.file(explicitSlideFiles[1]).async("string");
  assert("Bullet slide content", e2.includes("All tests pass"));

  // Slide 3: closing slide
  const e3 = await explicitZip.file(explicitSlideFiles[2]).async("string");
  assert("Closing slide has 'Thank You'", e3.includes("Thank You"));
  assert("Closing slide contact info", e3.includes("QA Team"));
  assert("Closing slide ghost logo F8A01B", e3.includes("F8A01B"));
  assert("Closing slide ghost logo 35% alpha", e3.includes("35000"));

  // ── Step 10: File-save round trip ───────────────────
  console.log("\n▶ 10. FILE SAVE ROUND TRIP");
  const outDir = path.join(__dirname, ".test-output");
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, "test-output.pptx");
  fs.writeFileSync(outPath, Buffer.from(buffer));

  const savedBuf = fs.readFileSync(outPath);
  assert("Saved file re-reads same size", savedBuf.length === buffer.length);

  let savedZip;
  try {
    savedZip = await JSZip.loadAsync(savedBuf);
    assert("Saved file opens as valid ZIP/PPTX", true);
  } catch (e) {
    assert("Saved file opens as valid ZIP/PPTX", false, e.message);
  }

  const savedSlides = naturalSortSlideFiles(
    Object.keys(savedZip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
  );
  assert(
    "Saved file has same slide count",
    savedSlides.length === slideFiles.length,
    `got ${savedSlides.length}`,
  );

  if (!saveOutput) {
    fs.rmSync(outDir, { recursive: true, force: true });
    assert("Temp output cleaned up", true);
  } else {
    console.log(`\n  (kept ${outPath})`);
  }

  // ── Step 11: Diverse layout types ───────────────────
  console.log("\n▶ 11. DIVERSE LAYOUT TYPES");
  const diverseSections = [
    { title: "Opening", slides: [{ type: "title", title: "Strategy Deck", subtitle: "Q4 2025" }] },
    { title: "Overview", slides: [{ type: "section-intro", title: "Strategic Overview", subtitle: "Key highlights" }] },
    { title: "Metrics", slides: [{ type: "bullet", title: "Key Metrics", content: ["ARPU growth", "Market share"] }] },
    { title: "Dashboard", slides: [{ type: "kpi", title: "KPI Dashboard", kpis: [
      { label: "Revenue", value: "Tk8.5B", sublabel: "+12% YoY" },
      { label: "Subscribers", value: "45M", sublabel: "+3% YoY" },
      { label: "ARPU", value: "Tk185", sublabel: "+8% YoY" },
      { label: "Churn", value: "2.1%", sublabel: "-0.5pp" },
    ] }] },
    { title: "Comparison", slides: [{ type: "two-column", title: "GP vs BL", leftContent: ["GP: Curated", "BL: Cluttered"], rightContent: ["GP: Experience", "BL: Commodity"] }] },
    { title: "Timeline", slides: [{ type: "timeline", title: "Implementation Roadmap", months: [
      { label: "Q1", details: "Phase 1 launch" },
      { label: "Q2", details: "Phase 2 rollout" },
      { label: "Q3", details: "Full deployment" },
      { label: "Q4", details: "Optimization" },
    ] }] },
    { title: "Data", slides: [{ type: "table", title: "Financial Summary", headers: ["Metric", "FY2024", "FY2025"], rows: [["Revenue", "8.5B", "9.6B"], ["EBITDA", "40.9%", "41.2%"]] }] },
    { title: "Transition", slides: [{ type: "section-break", title: "Part 2: Execution" }] },
    { title: "Visuals", slides: [{ type: "image-cards", title: "Product Portfolio", cards: [{ caption: "Starter" }, { caption: "Core" }, { caption: "Power" }] }] },
    { title: "Closing", slides: [{ type: "closing", cta: "Thank You", contact: "Strategy Team" }] },
  ];

  const diverseBuffer = await builder.buildPresentation(diverseSections, { theme: "dark" });
  const diverseZip = await JSZip.loadAsync(diverseBuffer);
  const diverseSlideFiles = naturalSortSlideFiles(
    Object.keys(diverseZip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
  );

  assert("10 diverse slides generated", diverseSlideFiles.length === 10, `got ${diverseSlideFiles.length}`);

  // Check KPI slide has metric values
  const kpiXml = await diverseZip.file(diverseSlideFiles[3]).async("string");
  assert("KPI slide has metric value", kpiXml.includes("Tk8.5B"));
  assert("KPI slide has metric label", kpiXml.includes("Revenue"));

  // Check timeline slide has month labels
  const tlXml = await diverseZip.file(diverseSlideFiles[5]).async("string");
  assert("Timeline slide has month label", tlXml.includes("Q1"));

  // Check section-break slide has orange background
  const sbXml = await diverseZip.file(diverseSlideFiles[7]).async("string");
  assert("Section-break slide has swoosh orange", sbXml.includes("EF6E23"));

  // Check closing slide has dark background
  const clXml = await diverseZip.file(diverseSlideFiles[9]).async("string");
  assert("Closing slide has white bg", clXml.includes("FFFFFF"));

  // ── Step 12: Extended layout types ───────────────────
  console.log("\n▶ 12. EXTENDED LAYOUT TYPES");
  const extendedSections = [
    { title: "Opening", slides: [{ type: "title", title: "Extended Test", subtitle: "Component test" }] },
    { title: "Numbered", slides: [{ type: "numbered-list", title: "Key Features",
      left: ["Feature one", "Feature two", "Feature three"],
      right: ["Benefit one", "Benefit two", "Benefit three"] }] },
    { title: "KPI Headline", slides: [{ type: "kpi-headline", title: "KPI Dashboard",
      metrics: [
        { headline: "Revenue", value: "8.5B", date: "FY2025" },
        { headline: "Subscribers", value: "45M", date: "Q4" },
        { headline: "ARPU", value: "Tk185", date: "YoY" },
      ] }] },
    { title: "Large Image", slides: [{ type: "large-image", title: "Product Demo", caption: "Banglalink 5G coverage map" }] },
    { title: "Dashboard", slides: [{ type: "dashboard", title: "Analytics Overview",
      charts: [
        { title: "Revenue Trend", data: [{ name: "Rev", labels: ["Q1","Q2","Q3","Q4"], values: [100,120,110,130] }] },
        { title: "Subscriber Growth", data: [{ name: "Sub", labels: ["Q1","Q2","Q3","Q4"], values: [1000,1100,1200,1300] }] },
      ] }] },
    { title: "Calendar", slides: [{ type: "calendar-timeline", title: "Implementation Plan",
      milestones: [
        { label: "Phase 1", month: "Jan", details: "Discovery and planning" },
        { label: "Phase 2", month: "Mar", details: "Development kickoff" },
        { label: "Phase 3", month: "Jun", details: "Launch" },
      ] }] },
    { title: "Closing", slides: [{ type: "closing", cta: "Thank You", contact: "Extended Test Team" }] },
  ];

  const extendedBuffer = await builder.buildPresentation(extendedSections, { theme: "dark" });
  const extendedZip = await JSZip.loadAsync(extendedBuffer);
  const extendedSlideFiles = naturalSortSlideFiles(
    Object.keys(extendedZip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
  );

  assert("7 extended slides generated", extendedSlideFiles.length === 7, `got ${extendedSlideFiles.length}`);

  // Check numbered-list slide
  const nlXml = await extendedZip.file(extendedSlideFiles[1]).async("string");
  assert("Numbered-list slide has numbered items", nlXml.includes("Feature one"));

  // Check kpi-headline slide
  const khlXml = await extendedZip.file(extendedSlideFiles[2]).async("string");
  assert("KPI-headline slide has headline", khlXml.includes("Revenue"));
  assert("KPI-headline slide has big number", khlXml.includes("8.5B"));

  // Check large-image slide
  const liXml = await extendedZip.file(extendedSlideFiles[3]).async("string");
  assert("Large-image slide has caption", liXml.includes("5G coverage map"));

  // Check dashboard slide
  const dXml = await extendedZip.file(extendedSlideFiles[4]).async("string");
  assert("Dashboard slide has chart titles", dXml.includes("Revenue Trend"));

  // Check calendar-timeline slide
  const calXml = await extendedZip.file(extendedSlideFiles[5]).async("string");
  assert("Calendar-timeline has month labels", calXml.includes("Jan"));

  // ── Summary ─────────────────────────────────────────
  const total = passed + failed;
  console.log("\n═════════════════════════════════════════════════════════");
  console.log(`  ${passed}/${total} assertions passed  ${failed > 0 ? "✗ FAILURES: " + failed : "✓ ALL PASSED"}`);
  console.log(`  File: ${(buffer.length / 1024).toFixed(1)} KB, ${slideFiles.length} slides`);
  console.log("═════════════════════════════════════════════════════════\n");

  process.exit(failed > 0 ? 1 : 0);
}

runTest().catch((e) => {
  console.error("Unhandled error:", e.message);
  process.exit(1);
});
