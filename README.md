# DeckForge Harness

[中文说明](README-zh.md)

Codex plugin for building polished presentation decks from intake, web/browser capture, frontend slide design, harness automation, and PPTX export.

DeckForge Harness is for making presentation decks with a stronger design process than "generate a few slides and hope they look good." It treats a deck as a workflow: first lock intent and visual direction with the user, then capture evidence or assets, design slides as frontend canvases, export to PPTX, and verify the result with rendered previews.

## Current Recommendation

Use DeckForge when you want Codex to help produce a real presentation deliverable, not just a PPTX file. The plugin is strongest when the deck can benefit from web/app screenshots, frontend-style visual design, or repeatable export/QA steps.

On macOS and Linux, DeckForge intentionally defaults to file-based automation: `python-pptx` for PPTX creation and LibreOffice/Poppler for render checks. WPS/MS Office COM is treated as an optional Windows backend, not a cross-platform promise.

## Good For

- project reports, defense decks, product walkthroughs, research presentations, and technical demos;
- decks that need a polished visual style rather than plain title-and-bullet slides;
- workflows that combine website/app capture, screenshots, generated images, and structured slide content;
- turning frontend slide canvases or JSON element schemas into PPTX;
- repeatable QA where the deck is rendered back to previews before delivery.

## Core Flow

1. Ask and confirm page count, title, audience, language, source scope, style direction, editability target, and quality reference.
2. Capture source material with Playwright, Browser tools, local screenshots, or user-provided files.
3. Design slides as frontend canvases or structured JSON elements.
4. Export the chosen representation to PPTX.
5. Render PPTX previews and inspect for blank pages, cropping, low quality, or layout issues.
6. Iterate before final delivery.

## Style Bar

DeckForge should produce decks that look intentionally designed. It should avoid sparse template pages, weak typography, low-density card grids, and generic bullet slides. Each deck should have a clear visual system, strong hierarchy, legible screenshots/assets, and enough visual structure for the use case.

## Editability Boundary

DeckForge supports multiple output strategies:

- full-slide image decks for maximum visual fidelity;
- editable hybrid decks with selected text/shapes rebuilt in PPTX;
- JSON element-router decks for structured editable layouts.

It does not magically recover native chart data, hidden design layers, or exact source vectors from screenshots. For Windows users, WPS/MS Office COM can be used as a live Office backend when available; for macOS users, WPS COM is not available.

## Install

Run one command:

```bash
curl -fsSL https://raw.githubusercontent.com/JerryLiu-uestc/deck-forge-harness/main/install.sh | bash
```

Or download the repository ZIP from GitHub, unzip it, and run:

```bash
bash install.sh
```

Then restart Codex or refresh plugins, and enable **DeckForge Harness** from the **Personal** marketplace.

The installer downloads the plugin, installs Python dependencies, registers the Codex marketplace entry, and installs LibreOffice/Poppler when a supported package manager is available.

## Use

In Codex, ask:

```text
Use DeckForge to create a polished PPT from my materials.
```

DeckForge will ask for page count, title, audience, language, source scope, style direction, editability target, and quality reference before it starts making slides.

## What It Does

- Captures web/app material with Playwright, Browser tools, or local screenshots.
- Designs slides as frontend canvases before exporting to PPTX.
- Uses small harness adapters for repeatable local workflows.
- Builds PPTX with `python-pptx` and verifies renders with LibreOffice/Poppler.
- Supports a portable JSON element-router path for structured decks.

## Platform Notes

- macOS/Linux: uses `python-pptx` plus LibreOffice for export and QA.
- Windows: can optionally use WPS/MS Office COM when `pywin32` and Office are installed.
- WPS COM is not a macOS capability.

Manual CLI reference: [docs/CLI.md](docs/CLI.md)
