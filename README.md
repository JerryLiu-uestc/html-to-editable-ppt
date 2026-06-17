# HTML to Editable PPT

[中文说明](README-zh.md)

Codex plugin for building polished presentation decks from intake, web/browser capture, frontend slide design, harness automation, and PPTX export.

HTML to Editable PPT is for making presentation decks with a stronger design process than "generate a few slides and hope they look good." It treats a deck as a workflow: first lock intent and visual direction with the user, then capture evidence or assets, design slides as frontend canvases, export to PPTX, and verify the result with rendered previews.

## Current Recommendation

Use HTML to Editable PPT when you want Codex to help produce a real presentation deliverable, not just a PPTX file. The plugin is strongest when the deck can benefit from web/app screenshots, frontend-style visual design, or repeatable export/QA steps.

On macOS and Linux, HTML to Editable PPT intentionally defaults to file-based automation: `python-pptx` for PPTX creation and LibreOffice/Poppler for render checks. WPS/MS Office COM is treated as an optional Windows backend, not a cross-platform promise.

## Good For

- project reports, defense decks, product walkthroughs, research presentations, and technical demos;
- decks that need a polished visual style rather than plain title-and-bullet slides;
- workflows that combine website/app capture, screenshots, generated images, and structured slide content;
- turning frontend slide canvases or JSON element schemas into PPTX;
- repeatable QA where the deck is rendered back to previews before delivery.

## Demo

Generated result, shown as individual slide renders. The full PPTX has 5 slides; the README shows the first 4 for readability.

![Demo slide 1](https://github.com/JerryLiu-uestc/html-to-editable-ppt/raw/main/examples/html-to-editable-ppt-demo/qa/slide-1.png)

![Demo slide 2](https://github.com/JerryLiu-uestc/html-to-editable-ppt/raw/main/examples/html-to-editable-ppt-demo/qa/slide-2.png)

![Demo slide 3](https://github.com/JerryLiu-uestc/html-to-editable-ppt/raw/main/examples/html-to-editable-ppt-demo/qa/slide-3.png)

![Demo slide 4](https://github.com/JerryLiu-uestc/html-to-editable-ppt/raw/main/examples/html-to-editable-ppt-demo/qa/slide-4.png)

The repository includes the editable PPTX and source used for this demo:

- PPTX: [html-to-editable-ppt-demo.pptx](examples/html-to-editable-ppt-demo/html-to-editable-ppt-demo.pptx)
- Source schema: [deck-schema.json](examples/html-to-editable-ppt-demo/deck-schema.json)
- Demo folder: [examples/html-to-editable-ppt-demo](examples/html-to-editable-ppt-demo)

Prompt used for the demo:

```text
Use HTML to Editable PPT to create a 5-slide demo deck for this plugin.
Show the workflow from intake to editable PPTX export and QA.
Use a clean technical style, strong hierarchy, dark/light contrast, and make the result editable in PowerPoint.
After generating the PPTX, render individual slide previews, then open it in PowerPoint and check for text overflow or overlap.
```

## Core Flow

1. Ask and confirm page count, title, audience, language, source scope, style direction, editability target, and quality reference.
2. Capture source material with Playwright, Browser tools, local screenshots, or user-provided files.
3. Design slides as frontend canvases or structured JSON elements.
4. Export the chosen representation to PPTX.
5. Render PPTX previews and inspect for blank pages, cropping, low quality, or layout issues.
6. Iterate before final delivery.

## Style Bar

HTML to Editable PPT should produce decks that look intentionally designed. It should avoid sparse template pages, weak typography, low-density card grids, and generic bullet slides. Each deck should have a clear visual system, strong hierarchy, legible screenshots/assets, and enough visual structure for the use case.

## Editability Boundary

HTML to Editable PPT supports multiple output strategies:

- full-slide image decks for maximum visual fidelity;
- editable hybrid decks with selected text/shapes rebuilt in PPTX;
- JSON element-router decks for structured editable layouts.

It does not magically recover native chart data, hidden design layers, or exact source vectors from screenshots. For Windows users, WPS/MS Office COM can be used as a live Office backend when available; for macOS users, WPS COM is not available.

## Install

Recommended:

1. Download this repository as a ZIP: [html-to-editable-ppt-main.zip](https://github.com/JerryLiu-uestc/html-to-editable-ppt/archive/refs/heads/main.zip).
2. Unzip it locally.
3. In Codex, install or enable the local plugin folder `html-to-editable-ppt`.

Then restart Codex or refresh plugins, and enable **HTML to Editable PPT** from the **Personal** marketplace.

If you want Codex to help install it, send this prompt to your Codex agent:

```text
Please install the Codex plugin from https://github.com/JerryLiu-uestc/html-to-editable-ppt.
Use the safe local-plugin flow: download or clone the repository, place it under ~/plugins/html-to-editable-ppt, register it in my personal Codex marketplace if needed, then run the plugin validator. Do not run remote shell installer commands.
After installation, tell me how to enable HTML to Editable PPT in Codex and whether optional runtime dependencies are missing on my OS.
```

For Claude Code, this repository works as a local workflow package rather than a Codex marketplace plugin. Ask Claude Code to download or clone the repository, read [skills/html-to-editable-ppt/SKILL.md](skills/html-to-editable-ppt/SKILL.md), and use [scripts/html_to_editable_ppt.py](scripts/html_to_editable_ppt.py) for local generation and QA.

Copy this into Claude Code:

```text
Please set up https://github.com/JerryLiu-uestc/html-to-editable-ppt for Claude Code.
Clone or download it into ~/plugins/html-to-editable-ppt, read skills/html-to-editable-ppt/SKILL.md as the workflow guide, and verify scripts/html_to_editable_ppt.py doctor runs.
When I ask for a deck, follow the intake gate, generate editable PPTX output when appropriate, render QA previews, and inspect the deck in PowerPoint/WPS if available.
Do not run remote shell installer commands.
```

Detailed install notes and optional dependency setup: [INSTALL.md](INSTALL.md)

Platform dependency notes:

- macOS: optional render QA needs LibreOffice and Poppler.
- Linux: optional render QA needs LibreOffice and Poppler utilities.
- Windows: optional live Office automation needs WPS or Microsoft Office plus `pywin32`; WPS/MS Office COM is Windows-only.

## Use

In Codex, ask:

```text
Use HTML to Editable PPT to create a polished PPT from my materials.
```

HTML to Editable PPT will ask for page count, title, audience, language, source scope, style direction, editability target, and quality reference before it starts making slides.

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
