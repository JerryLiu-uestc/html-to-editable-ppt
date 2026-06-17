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

## 本地 Harness CLI

初始化 DeckForge 工作区：

```bash
python3 scripts/deckforge.py init --project deckforge
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
