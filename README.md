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

## Quick Start

Clone the plugin into your local Codex plugins folder:

```bash
mkdir -p ~/plugins
git clone https://github.com/JerryLiu-uestc/deck-forge-harness.git ~/plugins/deck-forge-harness
```

Install Python dependencies:

```bash
python3 -m pip install python-pptx pillow
```

Install render dependencies:

```bash
# macOS
brew install --cask libreoffice
brew install poppler

# Ubuntu / Debian
sudo apt-get update
sudo apt-get install -y libreoffice poppler-utils
```

Install Playwright only if you want HTML/browser slide capture:

```bash
npm install -g playwright
playwright install chromium
```

Register the plugin in your personal Codex marketplace:

```bash
mkdir -p ~/.agents/plugins
cat > ~/.agents/plugins/marketplace.json <<'JSON'
{
  "name": "personal",
  "interface": {
    "displayName": "Personal"
  },
  "plugins": [
    {
      "name": "deck-forge-harness",
      "source": {
        "source": "local",
        "path": "./plugins/deck-forge-harness"
      },
      "policy": {
        "installation": "AVAILABLE",
        "authentication": "ON_INSTALL"
      },
      "category": "Productivity"
    }
  ]
}
JSON
```

Restart Codex or refresh plugins, then enable `deck-forge-harness` from your personal marketplace.

Verify the local harness:

```bash
cd ~/plugins/deck-forge-harness
python3 scripts/deckforge.py doctor
```

Expected macOS output should say WPS COM is unavailable and recommend `python-pptx + LibreOffice`.

Run a one-slide smoke test:

```bash
mkdir -p /tmp/deckforge-smoke
cat > /tmp/deckforge-smoke/deck-schema.json <<'JSON'
{
  "canvas": {"w": 1280, "h": 720},
  "slides": [
    {
      "background": "#FFFFFF",
      "elements": [
        {"type": "text", "x": 80, "y": 70, "w": 980, "h": 90, "text": "DeckForge is ready", "fs": 44, "bold": true, "color": "#0F766E"},
        {"type": "card", "x": 80, "y": 190, "w": 460, "h": 180, "fill": "#ECFDF5", "line": "#99F6E4"},
        {"type": "text", "x": 110, "y": 230, "w": 400, "h": 80, "text": "Schema to editable PPTX", "fs": 26, "bold": true, "color": "#134E4A"}
      ]
    }
  ]
}
JSON

python3 scripts/deckforge.py schema-to-pptx \
  --schema /tmp/deckforge-smoke/deck-schema.json \
  --output /tmp/deckforge-smoke/deck.pptx

python3 scripts/deckforge.py render-pptx \
  --pptx /tmp/deckforge-smoke/deck.pptx \
  --out-dir /tmp/deckforge-smoke/qa

python3 scripts/deckforge.py qa \
  --render-dir /tmp/deckforge-smoke/qa \
  --contact-sheet /tmp/deckforge-smoke/qa/contact.png
```

## Local Harness CLI

```bash
python3 scripts/deckforge.py init --project deckforge
```

```bash
python3 scripts/deckforge.py doctor
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
python3 scripts/deckforge.py schema-to-pptx \
  --schema deckforge/harness/deck-schema.json \
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

## Harness-Anything Backend Notes

DeckForge borrows the useful `harness-anything` pattern: expose fragile GUI or document workflows as small, observable CLI adapters. It does not assume every platform can control WPS or PowerPoint directly.

- Windows: WPS/MS Office COM can be used as an optional live Office backend when `pywin32` and Office are installed.
- macOS/Linux: default to `python-pptx` for PPTX file creation and LibreOffice headless for conversion/render QA.
- Structured decks: use JSON element routing (`schema-to-pptx`) before writing one-off PPT generation code.
