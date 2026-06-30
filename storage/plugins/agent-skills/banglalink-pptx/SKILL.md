# Banglalink PPTX Skill

## Quick Reference

| Task | Guide |
|------|-------|
| Create from scratch | Use `section-agent.js` with primitives toolkit |
| Read/analyze content | `python -m markitdown presentation.pptx` |
| Visual inspection | `python scripts/visual-qa.js output.pptx` |
| Content QA | `python scripts/content-qa.js output.pptx` |

---

## Design Palettes

Choose colors that match your topic. Use these palettes as inspiration:

| Theme | Primary | Secondary | Accent |
|-------|---------|-----------|--------|
| **Banglalink Orange** | `EF6E23` (swoosh) | `F86C02` (deep) | `FAA106` (amber) |
| **Midnight Executive** | `1E2761` (navy) | `CADCFC` (ice blue) | `FFFFFF` (white) |
| **Forest & Moss** | `2C5F2D` (forest) | `97BC62` (moss) | `F5F5F5` (cream) |
| **Coral Energy** | `F96167` (coral) | `F9E795` (gold) | `2F3C7E` (navy) |
| **Warm Terracotta** | `B85042` (terracotta) | `E7E8D1` (sand) | `A7BEAE` (sage) |
| **Ocean Gradient** | `065A82` (deep blue) | `1C7293` (teal) | `21295C` (midnight) |
| **Charcoal Minimal** | `36454F` (charcoal) | `F2F2F2` (off-white) | `212121` (black) |
| **Teal Trust** | `028090` (teal) | `00A896` (seafoam) | `02C39A` (mint) |
| **Berry & Cream** | `6D2E46` (berry) | `A26769` (dusty rose) | `ECE2D0` (cream) |
| **Sage Calm** | `84B59F` (sage) | `69A297` (eucalyptus) | `50808E` (slate) |
| **Cherry Bold** | `990011` (cherry) | `FCF6F5` (off-white) | `2F3C7E` (navy) |

---

## Typography Pairings

| Header Font | Body Font | Use Case |
|-------------|-----------|----------|
| Arial Black | Arial | Bold, modern presentations |
| Georgia | Calibri | Executive, formal decks |
| Calibri | Calibri Light | Clean, corporate |
| Cambria | Calibri | Traditional, academic |
| Trebuchet MS | Calibri | Tech, startup pitches |
| Impact | Arial | High-impact statements |

**Banglalink default:** Arial for both headings and body (brand requirement).

---

## Layout Variety Rules (CRITICAL)

**NEVER repeat the same layout type for consecutive slides.**

Track which layouts you've used and alternate:
- bullet → kpi → two-column → timeline → chart → image-cards → table → quote
- bullet → section-break → kpi → two-column → timeline → chart → closing

Available layouts:
| Layout | Best For |
|--------|----------|
| bullet | Key points, lists |
| kpi | 3-6 metrics with big numbers |
| two-column | Comparisons, pros/cons |
| timeline | Time-based data, phases |
| image-cards | 3-4 items to showcase |
| chart | Data trends, distributions |
| table | Structured data, financials |
| quote | Testimonials, key insights |
| section-break | Major transitions |
| section-intro | Section openings |
| numbered-list | Step-by-step processes |
| kpi-headline | KPIs with dates |
| large-image | Product demos, visuals |
| dashboard | Multiple data views |
| calendar-timeline | Monthly milestones |

**Aim for 3-4 different layouts per presentation.**

---

## Visual Design Principles

1. **Every slide needs a visual element** — image, chart, icon, or shape. Text-only slides are forgettable.
2. **Color dominance** — 60-70% white, 20% orange accent, 10% dark text.
3. **Leave breathing room** — don't fill every inch with text.
4. **Icons in colored circles** next to section headers add polish.
5. **Dark/light contrast** — dark backgrounds for title + conclusion slides, light for content.

---

## Typography Hierarchy

| Element | Size | Style |
|---------|------|-------|
| Display (hero titles, big numbers) | 48pt | Bold |
| Heading (slide titles) | 28pt | Bold |
| Subhead (subtitles) | 18pt | Regular |
| Body (paragraph text) | 14pt | Regular |
| Caption (labels) | 12pt | Regular |
| Micro (footnotes) | 10pt | Regular |

**NEVER use body size for titles — titles must be 24pt+.**

---

## QA Workflow (Required)

**Assume there are problems. Your job is to find them.**

### Content QA

```bash
python scripts/content-qa.js output.pptx
```

Checks for:
- Missing content
- Placeholder text
- Brand color presence

### Visual QA

**⚠️ USE SUBAGENTS** — even for 2-3 slides. You've been staring at the code and will see what you expect, not what's there. Subagents have fresh eyes.

```bash
python scripts/visual-qa.js output.pptx
```

This creates `slide-01.jpg`, `slide-02.jpg`, etc.

Use this prompt for subagents:

```
Visually inspect these slides. Assume there are issues — find them.

Look for:
- Overlapping elements (text through shapes, lines through words)
- Text overflow or cut off at edges/box boundaries
- Elements too close (< 0.3" gaps) or nearly touching
- Uneven gaps (large empty area in one place, cramped in another)
- Insufficient margin from slide edges (< 0.5")
- Low-contrast text (light text on light background)
- Text boxes too narrow causing excessive wrapping

For each slide, list issues or areas of concern, even if minor.

Read and analyze these images:
1. /path/to/slide-01.jpg (Expected: [brief description])
2. /path/to/slide-02.jpg (Expected: [brief description])

Report ALL issues found, including minor ones.
```

### Verification Loop

1. Generate slides → Convert to images → Inspect
2. **List issues found** (if none found, look again more critically)
3. Fix issues
4. **Re-verify affected slides** — one fix often creates another problem
5. Repeat until a full pass reveals no new issues

**Do not declare success until you've completed at least one fix-and-verify cycle.**

---

## Editing Existing PPTX

When editing an existing presentation:

1. **Analyze template**:
   ```bash
   python scripts/visual-qa.js template.pptx
   python -m markitdown template.pptx
   ```

2. **Unpack**:
   ```bash
   python scripts/office/unpack.py template.pptx unpacked/
   ```

3. **Edit slides**: Each slide is `ppt/slides/slide{N}.xml`

4. **Clean**:
   ```bash
   python scripts/clean.py unpacked/
   ```

5. **Pack**:
   ```bash
   python scripts/office/pack.py unpacked/ output.pptx --original template.pptx
   ```

### Editing Rules

- **Bold all headers**: Use `b="1"` on `<a:rPr>`
- **Never use unicode bullets**: Use proper list formatting
- **Separate multi-item content**: Create separate `<a:p>` elements for each item
- **Smart quotes**: Use XML entities (`&#x201C;` for left double quote)

---

## Common Pitfalls

| Issue | Solution |
|-------|----------|
| `#` prefix with hex colors | Remove `#` — use `"FF0000"` not `"#FF0000"` |
| 8-char hex for opacity | Use `opacity` property instead |
| Unicode bullets (•) | Use `bullet: true` |
| Missing `breakLine: true` | Add between array items |
| Reusing option objects | Create fresh objects each time |
| Text-only slides | Add visual element (chart, cards, icons) |
| Repeating same layout | Alternate layout types |
| Low-contrast text | Check against background color |

---

## Dependencies

- `pip install "markitdown[pptx]"` — text extraction
- `pip install Pillow` — thumbnail grids
- `npm install pptxgenjs` — creating from scratch
- LibreOffice (`soffice`) — PDF conversion
- Poppler (`pdftoppm`) — PDF to images
