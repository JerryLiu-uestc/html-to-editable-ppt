# DeckForge Harness

[中文说明](README-zh.md)

DeckForge Harness is a Codex plugin for agentic presentation workflows.

It is built around the idea:

```text
intake + visual quality bar -> Playwright/browser capture -> frontend slide design -> harness-anything automation -> PPTX export + QA
```

## Skills

- `deck-forge`: orchestrates end-to-end deck workflows.
- `frontend-slide-design`: designs slides as frontend canvases, compatible with `frontend-design` when that skill is available.
- `harness-anything`: creates small task-specific harnesses around local tools, browser flows, PPTX export, LibreOffice verification, PowerPoint, or WPS fallbacks.

DeckForge asks intake questions before creating a deck and treats visual quality as a gate, not a nice-to-have. It should lock page count, title, audience, language, source scope, style direction, editability target, and reference quality before capture or design work starts.

## Local Harness CLI

```bash
python3 scripts/deckforge.py init --project deckforge
```

```bash
python3 scripts/deckforge.py html-to-png \
  --html deckforge/slides/index.html \
  --out-dir deckforge/renders
```

```bash
python3 scripts/deckforge.py images-to-pptx \
  --input-dir deckforge/renders \
  --output deckforge/pptx/deck.pptx
```

```bash
python3 scripts/deckforge.py render-pptx \
  --pptx deckforge/pptx/deck.pptx \
  --out-dir deckforge/qa
```

## Dependencies

- Python 3
- `python-pptx`
- Pillow
- Node.js + Playwright for HTML capture
- LibreOffice `soffice` for PPTX to PDF export
- Poppler `pdftoppm` for slide preview rendering

The plugin uses MCP/browser tooling when available, but keeps deterministic local scripts for the export and verification stages.
