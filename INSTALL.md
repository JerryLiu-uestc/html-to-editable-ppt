# Installation

## Download

Download the clean release package:

```text
https://github.com/JerryLiu-uestc/html-to-editable-ppt/releases/latest/download/html-to-editable-ppt-plugin.zip
```

Or use `git clone` if you prefer keeping the plugin up to date through Git:

```bash
git clone https://github.com/JerryLiu-uestc/html-to-editable-ppt ~/plugins/html-to-editable-ppt
```

The GitHub source ZIP includes examples and development files. Use it only when you want the demo assets or source tree.

## Enable In Codex

1. Download the release package from GitHub Releases.
2. Unzip it.
3. Move the folder to a local plugin location, for example:

```text
~/plugins/html-to-editable-ppt
```

4. Open Codex, refresh plugins, and enable **HTML to Editable PPT** from the local/personal marketplace.

If your Codex setup does not automatically discover `~/plugins`, add this plugin to your personal marketplace entry. The expected plugin path is:

```text
./plugins/html-to-editable-ppt
```

## Ask Codex To Install It

Copy this into Codex:

```text
Please install the Codex plugin from https://github.com/JerryLiu-uestc/html-to-editable-ppt.
Use the safe local-plugin flow: download or clone the repository, place it under ~/plugins/html-to-editable-ppt, register it in my personal Codex marketplace if needed, then run the plugin validator. Do not run remote shell installer commands.
After installation, tell me how to enable HTML to Editable PPT in Codex and whether optional runtime dependencies are missing on my OS.
```

## Use With Claude Code

Claude Code cannot enable a Codex marketplace plugin directly. Use this repository as a local workflow package: download or clone it, ask Claude Code to read `skills/html-to-editable-ppt/SKILL.md`, and use `scripts/html_to_editable_ppt.py` for local PPTX generation and QA.

Copy this into Claude Code:

```text
Please set up https://github.com/JerryLiu-uestc/html-to-editable-ppt for Claude Code.
Clone or download it into ~/plugins/html-to-editable-ppt, read skills/html-to-editable-ppt/SKILL.md as the workflow guide, and verify scripts/html_to_editable_ppt.py doctor runs.
When I ask for a deck, follow the intake gate, generate editable PPTX output when appropriate, render QA previews, and inspect the deck in PowerPoint/WPS if available.
Do not run remote shell installer commands.
```

## Optional Runtime Dependencies By Platform

HTML to Editable PPT can be enabled as a plugin without running a shell installer. Some local harness features need external tools:

- `python-pptx` and Pillow for PPTX generation;
- LibreOffice and Poppler for PPTX render QA;
- Node.js and Playwright for HTML/browser capture.

macOS:

```bash
brew install --cask libreoffice
brew install poppler
python3 -m pip install python-pptx pillow
```

Linux:

```bash
sudo apt-get install libreoffice poppler-utils
python3 -m pip install python-pptx pillow
```

Windows:

```powershell
py -m pip install python-pptx pillow pywin32
```

On Windows, WPS/MS Office COM automation additionally requires WPS Office or Microsoft Office to be installed. On macOS, WPS COM is not available.

If these are missing, HTML to Editable PPT should report the missing capability and use the best available fallback.
