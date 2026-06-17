# DeckForge Harness

[English README](README.md)

用于制作高质量 PPT / 演示文稿的 Codex 插件。它会先确认需求和视觉标准，再结合网页/浏览器采集、前端幻灯片设计、harness 自动化和 PPTX 导出完成交付。

DeckForge Harness 的目标不是“随便生成几页 PPT”，而是把 PPT 当成一个完整工作流：先和用户确认意图与视觉方向，再采集资料或截图，用前端画布或结构化元素设计页面，导出 PPTX，并通过渲染预览做 QA。

## 当前建议

当你希望 Codex 帮你做一个真正能交付的演示文稿，而不是只生成一个 `.pptx` 文件时，适合使用 DeckForge。它尤其适合需要网页/应用截图、前端式视觉设计、自动导出和检查流程的 PPT。

在 macOS 和 Linux 上，DeckForge 默认走文件级自动化：用 `python-pptx` 创建 PPTX，用 LibreOffice / Poppler 渲染检查。WPS / Microsoft Office COM 只作为 Windows 上的可选后端，不承诺跨平台可用。

## 适合做什么

- 项目汇报、答辩报告、产品演示、研究展示、技术 demo；
- 需要更好视觉风格，而不是普通标题加 bullet 的 PPT；
- 需要组合网页/应用截图、图片素材、生成图和结构化内容的 deck；
- 把前端幻灯片画布或 JSON 元素 schema 转成 PPTX；
- 需要在交付前自动渲染预览并检查的工作流。

## 核心流程

1. 先询问并确认页数、标题、受众、语言、资料范围、风格方向、可编辑目标和质量参考。
2. 使用 Playwright、Browser 工具、本地截图或用户文件采集资料。
3. 用前端画布或结构化 JSON 元素设计幻灯片。
4. 将设计结果导出为 PPTX。
5. 渲染 PPTX 预览图，检查空白页、裁切、低质量和布局问题。
6. 修正后再交付。

## 风格要求

DeckForge 应该产出“看起来经过设计”的 deck。它要避免空洞模板页、弱标题设计、低密度卡片网格和普通 bullet 页。每个 deck 都应该有清晰的视觉系统、明确的信息层级、可读的截图/素材，以及符合场景的视觉结构。

## 可编辑性边界

DeckForge 支持多种输出策略：

- 整页图片型 PPT，优先保证视觉还原；
- 可编辑混合型 PPT，重建部分文字和形状；
- JSON 元素路由型 PPT，适合结构化、可编辑布局。

它不会从截图里自动恢复原生图表数据、隐藏设计图层或精确矢量源文件。Windows 用户可以在环境支持时选择 WPS / MS Office COM 后端；macOS 用户不能使用 WPS COM。

## 安装

运行一个命令：

```bash
curl -fsSL https://raw.githubusercontent.com/JerryLiu-uestc/deck-forge-harness/main/install.sh | bash
```

或者在 GitHub 下载 ZIP，解压后运行：

```bash
bash install.sh
```

然后重启 Codex 或刷新插件列表，在 **Personal** marketplace 里启用 **DeckForge Harness**。

安装脚本会下载插件、安装 Python 依赖、注册 Codex marketplace，并在检测到可用包管理器时安装 LibreOffice / Poppler。

## 使用

在 Codex 里说：

```text
用 DeckForge 根据我的材料做一个好看的 PPT。
```

DeckForge 会先询问页数、标题、受众、语言、资料范围、风格方向、可编辑目标和质量参考，确认后才开始制作。

## 能做什么

- 用 Playwright、Browser 工具或本地截图采集网页/应用素材。
- 先用前端画布设计幻灯片，再导出为 PPTX。
- 用小型 harness 适配器封装可重复的本地流程。
- 用 `python-pptx` 生成 PPTX，并用 LibreOffice / Poppler 渲染检查。
- 支持 JSON 元素路由，适合结构化、数据驱动的 deck。

## 平台说明

- macOS/Linux：默认使用 `python-pptx` + LibreOffice 进行导出和 QA。
- Windows：安装 `pywin32` 和 WPS / Microsoft Office 后，可以选择 WPS/MS Office COM 后端。
- WPS COM 不是 macOS 能力。

手动 CLI 参考：[docs/CLI-zh.md](docs/CLI-zh.md)
