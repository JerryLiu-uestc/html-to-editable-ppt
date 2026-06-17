# DeckForge Harness

[English README](README.md)

DeckForge Harness 是一个用于智能 PPT / 演示文稿工作流的 Codex 插件。

它的核心流程是：

```text
需求确认 + 视觉质量门槛 -> Playwright / 浏览器采集 -> 前端幻灯片设计 -> harness-anything 自动化 -> PPTX 导出 + QA
```

## 技能

- `deck-forge`：总控端到端演示文稿工作流。
- `frontend-slide-design`：把幻灯片设计成 HTML/CSS/React/SVG 前端画布；如果环境里有 `frontend-design`，会与它配合使用。
- `harness-anything`：围绕本地工具、浏览器流程、PPTX 导出、LibreOffice 验证、PowerPoint 或 WPS 兜底流程创建小型自动化 harness。

DeckForge 在开始制作 PPT 前会先询问并锁定需求，并把视觉质量作为硬性门槛，而不是可选优化。它会先确认页数、标题、受众、语言、资料范围、风格方向、可编辑目标和质量参考，再进入采集或设计阶段。

## 快速开始

把插件克隆到本地 Codex 插件目录：

```bash
mkdir -p ~/plugins
git clone https://github.com/JerryLiu-uestc/deck-forge-harness.git ~/plugins/deck-forge-harness
```

安装 Python 依赖：

```bash
python3 -m pip install python-pptx pillow
```

安装渲染依赖：

```bash
# macOS
brew install --cask libreoffice
brew install poppler

# Ubuntu / Debian
sudo apt-get update
sudo apt-get install -y libreoffice poppler-utils
```

如果需要 HTML / 浏览器截图采集，再安装 Playwright：

```bash
npm install -g playwright
playwright install chromium
```

把插件注册到 Codex 个人 marketplace：

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

重启 Codex 或刷新插件列表，然后在个人 marketplace 中启用 `deck-forge-harness`。

验证本地 harness：

```bash
cd ~/plugins/deck-forge-harness
python3 scripts/deckforge.py doctor
```

在 macOS 上，预期输出会提示 WPS COM 不可用，并推荐 `python-pptx + LibreOffice`。

运行一页 PPTX 冒烟测试：

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

## 本地 Harness CLI

初始化 DeckForge 工作区：

```bash
python3 scripts/deckforge.py init --project deckforge
```

检查当前机器可用的 Office / 渲染后端：

```bash
python3 scripts/deckforge.py doctor
```

将 HTML 幻灯片画布渲染为 PNG：

```bash
python3 scripts/deckforge.py html-to-png \
  --html deckforge/slides/index.html \
  --out-dir deckforge/renders
```

将渲染出的图片拼装为 PPTX：

```bash
python3 scripts/deckforge.py images-to-pptx \
  --input-dir deckforge/renders \
  --output deckforge/pptx/deck.pptx
```

用 JSON 元素路由生成可编辑 PPTX：

```bash
python3 scripts/deckforge.py schema-to-pptx \
  --schema deckforge/harness/deck-schema.json \
  --output deckforge/pptx/deck.pptx
```

将 PPTX 渲染为预览图用于检查：

```bash
python3 scripts/deckforge.py render-pptx \
  --pptx deckforge/pptx/deck.pptx \
  --out-dir deckforge/qa
```

## 依赖

- Python 3
- `python-pptx`
- Pillow
- Node.js + Playwright，用于 HTML / 网页截图采集
- LibreOffice `soffice`，用于 PPTX 转 PDF
- Poppler `pdftoppm`，用于生成幻灯片预览图

插件会在可用时使用 MCP / 浏览器工具，同时保留确定性的本地脚本来完成导出和验证阶段。

## Harness-Anything 后端说明

DeckForge 借鉴 `harness-anything` 的关键思想：把脆弱的 GUI 或文档工作流封装成小型、可观测的 CLI 适配器。但它不会假设所有平台都能直接操控 WPS 或 PowerPoint。

- Windows：安装 `pywin32` 和 WPS / Microsoft Office 后，可以把 WPS/MS Office COM 作为可选的真实 Office 后端。
- macOS/Linux：默认使用 `python-pptx` 创建 PPTX，并用 LibreOffice headless 做转换和渲染 QA。
- 结构化 deck：优先使用 JSON 元素路由（`schema-to-pptx`），再考虑写一次性 PPT 生成代码。
