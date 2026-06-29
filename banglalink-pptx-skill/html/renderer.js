const path = require("path");
const os = require("os");
const fs = require("fs");

const COMMON_CHROME_PATHS = {
  darwin: [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
  ],
  linux: [
    "/usr/bin/google-chrome",
    "/usr/bin/google-chrome-stable",
    "/usr/bin/chromium-browser",
    "/usr/bin/chromium",
    "/snap/bin/chromium",
  ],
  win32: [
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files\\Chromium\\Application\\chrome.exe",
  ],
};

async function detectBrowser() {
  try {
    const playwright = require("playwright");
    const browser = await playwright.chromium.launch({ headless: true });
    return { browser, engine: "playwright" };
  } catch (e) {
    const playwrightMissing = e.code === "MODULE_NOT_FOUND" || e.message?.includes("Cannot find module");
  }

  try {
    const platform = os.platform();
    const chromePaths = COMMON_CHROME_PATHS[platform] || [];

    for (const chromePath of chromePaths) {
      if (fs.existsSync(chromePath)) {
        try {
          const playwright = require("playwright");
          const browser = await playwright.chromium.launch({
            headless: true,
            executablePath: chromePath,
          });
          return { browser, engine: "playwright-system-chrome" };
        } catch (e) {
          continue;
        }
      }
    }
  } catch (e) {}

  try {
    const puppeteer = require("puppeteer-core");
    const platform = os.platform();
    const chromePaths = COMMON_CHROME_PATHS[platform] || [];

    for (const chromePath of chromePaths) {
      if (fs.existsSync(chromePath)) {
        try {
          const browser = await puppeteer.launch({
            headless: true,
            executablePath: chromePath,
          });
          return { browser, engine: "puppeteer" };
        } catch (e) {
          continue;
        }
      }
    }
  } catch (e) {}

  return { browser: null, engine: null };
}

async function renderHTML(html, options) {
  const result = await detectBrowser();
  if (!result.browser) {
    return null;
  }

  try {
    const page = await result.browser.newPage({
      viewport: {
        width: options?.width || 1280,
        height: options?.height || 720,
      },
    });

    await page.setContent(html, { waitUntil: "networkidle", timeout: 15000 }).catch(() =>
      page.waitForLoadState("domcontentloaded", { timeout: 5000 })
    );

    await page.waitForTimeout(900);

    const elements = await page.evaluate(extractElements);
    await result.browser.close();
    return elements;
  } catch (e) {
    try {
      await result.browser.close();
    } catch (_) {}
    return null;
  }
}

const extractElements = `() => {
    const SKIP_TAGS = new Set(['HTML','HEAD','META','LINK','SCRIPT','STYLE','TITLE','BODY']);
    const SKIP_CLS = ['blur-3xl','opacity-5'];

    function parseRGBA(s) {
        if (!s || s === 'transparent') return null;
        const m = s.match(/rgba?\\((\\d+),\\s*(\\d+),\\s*(\\d+)(?:,\\s*([\\d.]+))?\\)/);
        return m ? {r:+m[1],g:+m[2],b:+m[3],a:m[4]!==undefined?+m[4]:1} : null;
    }

    const shapes = [];

    function walkElements(el, depth) {
        if (depth > 9) return;
        if (SKIP_TAGS.has(el.tagName)) return;
        const cls = Array.from(el.classList || []);
        if (cls.some(c => SKIP_CLS.some(s => c.includes(s)))) return;
        const rect = el.getBoundingClientRect();
        if (rect.width < 4 || rect.height < 4) return;
        const st = window.getComputedStyle(el);
        const tag = el.tagName.toLowerCase();
        const isImg = tag === 'img';
        const bgRaw = st.backgroundColor;
        const bg = parseRGBA(bgRaw);
        const hasBg = bg && bg.a > 0.03;
        const bwRaw = parseFloat(st.borderWidth)||0;
        const bcRaw = bwRaw > 0 ? parseRGBA(st.borderColor) : null;
        const hasBorder = bwRaw > 0 && bcRaw && bcRaw.a > 0.03;
        const bgImg = st.backgroundImage;
        const hasGrad = bgImg && bgImg.includes('gradient');

        if (hasBg || hasBorder || hasGrad || isImg) {
            const brRaw = parseFloat(st.borderRadius)||0;
            let boxBg = bg;
            if (!hasBg && hasGrad) {
                boxBg = {r: 28, g: 28, b: 28, a: 1};
            }
            shapes.push({
                type: isImg ? 'image' : 'rect',
                x: rect.left, y: rect.top, w: rect.width, h: rect.height,
                bg: boxBg, br: brRaw,
                border: bcRaw, bw: bwRaw,
                src: isImg ? el.getAttribute('src') : null,
                depth: depth,
            });
        }
        for (const child of el.children) walkElements(child, depth+1);
    }

    const slide = document.querySelector('.ppt-slide');
    if (slide) walkElements(slide, 0);

    const treeWalker = document.createTreeWalker(
        slide || document.body, NodeFilter.SHOW_TEXT, null, false
    );
    let node;
    while(node = treeWalker.nextNode()) {
        const text = node.textContent.trim();
        if(text.length === 0) continue;
        const p = node.parentElement;
        if(SKIP_TAGS.has(p.tagName)) continue;
        const cls = Array.from(p.classList || []);
        if (cls.some(c => SKIP_CLS.some(s => c.includes(s)))) continue;
        if (cls.includes('material-icons')) continue;
        const range = document.createRange();
        range.selectNodeContents(node);
        const rect = range.getBoundingClientRect();
        if(rect.width < 1 || rect.height < 1) continue;
        const st = window.getComputedStyle(p);
        const color = parseRGBA(st.color);
        shapes.push({
            type: 'text',
            x: rect.left, y: rect.top, w: rect.width, h: rect.height,
            text: text, color: color,
            fontSize: parseFloat(st.fontSize),
            bold: parseInt(st.fontWeight) >= 600,
            align: st.textAlign,
            depth: 10 + shapes.length*0.001
        });
    }

    shapes.sort((a,b) => a.depth - b.depth);
    return shapes;
}`;

async function htmlToScreenshot(html, outputPath, options) {
  const result = await detectBrowser();
  if (!result.browser) {
    return null;
  }

  try {
    const page = await result.browser.newPage({
      viewport: {
        width: options?.width || 1280,
        height: options?.height || 720,
      },
    });

    await page.setContent(html, { waitUntil: "networkidle", timeout: 15000 }).catch(() =>
      page.waitForLoadState("domcontentloaded", { timeout: 5000 })
    );

    await page.waitForTimeout(900);

    await page.screenshot({
      path: outputPath,
      clip: { x: 0, y: 0, width: options?.width || 1280, height: options?.height || 720 },
    });

    await result.browser.close();
    return outputPath;
  } catch (e) {
    try {
      await result.browser.close();
    } catch (_) {}
    return null;
  }
}

module.exports = { detectBrowser, renderHTML, htmlToScreenshot };
