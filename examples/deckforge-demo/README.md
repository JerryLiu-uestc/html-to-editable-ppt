# DeckForge Demo

This folder contains a small editable PowerPoint demo generated through the DeckForge schema path.

Files:

- `deck-schema.json` - source schema with editable text, cards, lines, and layout positions.
- `deckforge-demo.pptx` - generated editable PowerPoint deck.
- `qa/contact.png` - rendered contact sheet used for visual QA.

This demo was also opened in Microsoft PowerPoint for macOS and checked with real app screenshots. The PowerPoint check is important because native Office text wrapping can differ from LibreOffice/PDF rendering.

Regenerate the deck:

```bash
python scripts/deckforge.py schema-to-pptx \
  --schema examples/deckforge-demo/deck-schema.json \
  --output examples/deckforge-demo/deckforge-demo.pptx
```

Render and QA:

```bash
python scripts/deckforge.py render-pptx \
  --pptx examples/deckforge-demo/deckforge-demo.pptx \
  --out-dir examples/deckforge-demo/qa

python scripts/deckforge.py qa \
  --render-dir examples/deckforge-demo/qa \
  --contact-sheet examples/deckforge-demo/qa/contact.png
```
