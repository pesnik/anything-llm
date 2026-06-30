/**
 * Slide Code Executor — Executes agent-written slide code.
 * 
 * Pattern from: artifact-kit/pptxgenjs-jsx (code execution)
 * 
 * The agent writes JavaScript code that uses primitives and components.
 * This module safely executes that code to generate slides.
 */

const brand = require("../brand");
const primitives = require("./primitives");
const components = require("./components");

/**
 * Execute agent-written slide code.
 * 
 * @param {string} code - JavaScript code string written by the agent
 * @param {object} pptx - PptxGenJS instance
 * @param {object} context - Additional context (data, theme, etc.)
 * @returns {Array} Array of slide operations
 */
function executeSlideCode(code, pptx, context = {}) {
  const slides = [];
  
  // Create sandboxed execution environment
  const sandbox = {
    // Primitives
    P: primitives,
    colors: primitives.colors,
    typography: primitives.typography,
    spacing: primitives.spacing,
    
    // Components
    ...components,
    
    // PptxGenJS instance
    pptx,
    
    // Context data
    ...context,
    
    // Helper: create a new slide
    addSlide: () => {
      const slide = pptx.addSlide();
      slides.push(slide);
      return slide;
    },
    
    // Helper: get current slide
    currentSlide: () => slides[slides.length - 1],
    
    // Helper: raw pptxgenjs access
    raw: primitives.raw,
    
    // Console for debugging
    console: {
      log: (...args) => console.log("[slide-code]", ...args),
      warn: (...args) => console.warn("[slide-code]", ...args),
      error: (...args) => console.error("[slide-code]", ...args),
    },
    
    // Math utilities
    Math,
    parseInt,
    parseFloat,
  };
  
  try {
    // Create function with sandboxed context
    const fn = new Function(...Object.keys(sandbox), code);
    fn(...Object.values(sandbox));
  } catch (e) {
    console.error("[slide-code-executor] Execution error:", e.message);
    console.error("[slide-code-executor] Code:", code.substring(0, 200));
    throw e;
  }
  
  return slides;
}

/**
 * Execute multiple slide code blocks (one per section).
 * 
 * @param {Array} codeBlocks - Array of { title, code } objects
 * @param {object} pptx - PptxGenJS instance
 * @param {object} context - Additional context
 * @returns {Array} Array of all slides
 */
function executeAllSlides(codeBlocks, pptx, context = {}) {
  const allSlides = [];
  
  for (const block of codeBlocks) {
    try {
      const slides = executeSlideCode(block.code, pptx, {
        ...context,
        sectionTitle: block.title,
        sectionData: block.data,
      });
      allSlides.push(...slides);
    } catch (e) {
      console.error(`[slide-code-executor] Failed section "${block.title}":`, e.message);
      // Continue with next section
    }
  }
  
  return allSlides;
}

module.exports = {
  executeSlideCode,
  executeAllSlides,
};
