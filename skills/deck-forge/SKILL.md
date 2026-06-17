---
name: deck-forge
description: Use this skill when the user wants an end-to-end AI presentation workflow that combines Playwright/browser capture, frontend-designed slide canvases, harness automation, and PPTX export or verification. Trigger for requests like "按 Playwright MCP + frontend-design + harness-anything 做 PPT 插件", "build a deck workflow", "capture website into PPT", or "automate a presentation pipeline".
---

# DeckForge

DeckForge is the orchestrator skill for a presentation workflow built around three lanes:

1. **Capture**: collect screenshots, DOM text, assets, and source URLs with Playwright MCP, the Browser plugin, or local Playwright.
2. **Design**: build each slide as a frontend canvas using `frontend-design` when available; otherwise create HTML/CSS/SVG with the same discipline.
3. **Harness**: write a small task-specific automation wrapper that renders, exports, checks, and iterates until the deck is deliverable.

## Non-Negotiable Intake Gate

For any deck creation or redesign request, the first response must ask intake questions and wait for user confirmation before creating files, generating slides, capturing sources, or editing PPTX. Do not infer missing deck direction from defaults.

Ask and lock:

- page count;
- title;
- audience and use case;
- source material scope, including whether web capture is allowed;
- language: Chinese, English, or bilingual;
- dark, light, or mixed background system;
- desired style: premium atmospheric, clean professional, academic defense, product launch, technical report, or user-specified;
- editability target: full-slide images, editable hybrid, or image-to-editable reconstruction;
- quality reference: user-provided deck/images, or an agreed verbal quality bar;
- exact terms/names that must not be changed;
- whether the user approves the proposed visual direction and slide plan.

After intake, produce a slide plan and visual direction summary. Wait for approval before capture/design/export unless the user explicitly asks for analysis only.

## Visual Quality Bar

Every DeckForge deck should aim for a visibly polished, presentation-ready result. "Good-looking" is not optional; convert it into checkable criteria:

- strong main visual or dominant information structure on each slide;
- clear hierarchy, confident title treatment, and deliberate typography;
- enough visual density for the context, avoiding sparse title-and-bullet pages;
- coherent palette, icon language, spacing, and repeated design motif;
- high-quality screenshots/assets when the deck discusses real software, websites, products, or games;
- formal credibility for academic/report decks, and product-level polish for commercial decks;
- varied layouts across the deck while preserving a consistent design system.

Before exporting PPTX, create or inspect preview images. If a slide looks generic, sparse, poorly aligned, low contrast, or weaker than the agreed quality reference, revise it before asking for final confirmation.

## Default Pipeline

Use this sequence unless the user gives a narrower task:

1. Complete intake and get approval for the slide plan plus visual direction.
2. Create a `deckforge/` workspace under the project:
   - `captures/` for screenshots and source pulls
   - `slides/` for HTML slide canvases
   - `renders/` for PNG previews
   - `pptx/` for generated decks
   - `harness/` for project-specific scripts
3. Capture source material:
   - Prefer a Playwright MCP/browser tool if available.
   - Use local Playwright only when no MCP/browser tool is available.
   - Store screenshots and a short `sources.json`.
4. Design slides:
   - If `frontend-design` is installed, use it for visual direction and component quality.
   - Keep slides as fixed 16:9 canvases, one HTML file per deck or per slide.
   - Avoid relying on PowerPoint layout while designing; PowerPoint is an export target.
5. Build a harness:
   - Use `harness-anything` to create a small adapter for the exact local tools in use.
   - Prefer deterministic file operations over GUI automation.
   - For structured, repeatable decks, use a JSON element-router schema before writing custom PPTX code.
   - Treat Windows WPS/MS Office COM as an optional backend; on macOS, default to `python-pptx` plus LibreOffice.
6. Export:
   - For highest fidelity, screenshot each HTML slide and place the PNGs into PPTX.
   - For editability, recreate selected text/shapes with `python-pptx` only where needed.
7. Verify:
   - Render PPTX to PDF/images.
   - Check page count, nonblank output, screenshot contact sheet, and obvious cropping.
   - Iterate before final response.

## Local CLI

The plugin includes `scripts/deckforge.py` for common filesystem and export tasks:

```bash
python3 ~/plugins/deck-forge-harness/scripts/deckforge.py init --project deckforge
```

```bash
python3 ~/plugins/deck-forge-harness/scripts/deckforge.py html-to-png \
  --html deckforge/slides/index.html \
  --out-dir deckforge/renders
```

```bash
python3 ~/plugins/deck-forge-harness/scripts/deckforge.py images-to-pptx \
  --input-dir deckforge/renders \
  --output deckforge/pptx/deck.pptx
```

```bash
python3 ~/plugins/deck-forge-harness/scripts/deckforge.py render-pptx \
  --pptx deckforge/pptx/deck.pptx \
  --out-dir deckforge/qa
```

```bash
python3 ~/plugins/deck-forge-harness/scripts/deckforge.py doctor
```

```bash
python3 ~/plugins/deck-forge-harness/scripts/deckforge.py schema-to-pptx \
  --schema deckforge/harness/deck-schema.json \
  --output deckforge/pptx/deck.pptx
```

## Decision Rules

- Use browser/Playwright capture for websites, dashboards, app screens, or live references.
- Use frontend slide canvases for original design work.
- Use direct PPTX file editing for assembly and light modifications.
- Use JSON element routing for decks with repeated components or data-driven layouts.
- Use PowerPoint/WPS GUI automation only when the user specifically needs behavior that cannot be done through file operations.
- On Windows, WPS/MS Office COM can be a first-class live Office backend when installed; on macOS, do not promise WPS COM and use LibreOffice only for conversion/render QA.
- If `frontend-design`, Playwright MCP, or a GUI automation harness is unavailable, state the fallback and keep moving with local HTML/Playwright/LibreOffice where possible.
- For PowerPoint deliverables, do not rely only on a contact sheet or LibreOffice render. When PowerPoint, WPS, or another target Office app is available, open the generated PPTX and inspect real app screenshots for text overflow, missing wrapping, clipping, and overlap before delivery.

## Completion Evidence

Final responses should include:

- generated PPTX path;
- locked intake summary or a note that the task was analysis-only;
- capture/source path when relevant;
- render/contact-sheet path;
- real Office app screenshot inspection result when an Office app is available;
- verification commands run;
- unavailable tools or fallback choices.
