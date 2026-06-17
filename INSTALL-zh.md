# 安装说明

## 下载

从 GitHub Releases 下载干净的 release 包：

```text
https://github.com/JerryLiu-uestc/html-to-editable-ppt/releases/latest/download/html-to-editable-ppt-plugin.zip
```

如果你希望后续用 Git 更新，也可以 clone：

```bash
git clone https://github.com/JerryLiu-uestc/html-to-editable-ppt ~/plugins/html-to-editable-ppt
```

GitHub 源码 ZIP 会包含 examples 和开发文件。只有在你想查看 demo 素材或源码树时才使用源码 ZIP。

## 在 Codex 中启用

1. 在 GitHub Releases 下载 release 包。
2. 解压。
3. 将文件夹移动到本地插件目录，例如：

```text
~/plugins/html-to-editable-ppt
```

4. 打开 Codex，刷新插件列表，并在本地/个人 marketplace 中启用 **HTML to Editable PPT**。

如果你的 Codex 没有自动发现 `~/plugins`，请把插件加入个人 marketplace。期望插件路径是：

```text
./plugins/html-to-editable-ppt
```

## 让 Codex 帮你安装

把下面这段话复制给 Codex：

```text
请帮我安装这个 Codex 插件：https://github.com/JerryLiu-uestc/html-to-editable-ppt。
请使用安全的本地插件安装流程：下载或 clone 仓库，放到 ~/plugins/html-to-editable-ppt，必要时注册到我的个人 Codex marketplace，然后运行插件校验。不要运行远程 shell 安装命令。
安装后请告诉我如何在 Codex 里启用 HTML to Editable PPT，并检查我当前系统是否缺少可选运行依赖。
```

## 在 Claude Code 中使用

Claude Code 不能直接启用 Codex marketplace 插件。你可以把这个仓库当作本地 workflow/skill 包使用：下载或 clone 后，让 Claude Code 阅读 `skills/html-to-editable-ppt/SKILL.md`，并使用 `scripts/html_to_editable_ppt.py` 做本地 PPTX 生成和 QA。

把下面这段话复制给 Claude Code：

```text
请帮我为 Claude Code 配置这个项目：https://github.com/JerryLiu-uestc/html-to-editable-ppt。
请下载或 clone 到 ~/plugins/html-to-editable-ppt，阅读 skills/html-to-editable-ppt/SKILL.md 作为工作流说明，并验证 scripts/html_to_editable_ppt.py doctor 可以运行。
之后我让你做 PPT 时，请遵守需求确认流程，在合适时生成可编辑 PPTX，渲染 QA 预览，并在可用时用 PowerPoint/WPS 打开检查。
不要运行远程 shell 安装命令。
```

## 按平台配置可选运行依赖

HTML to Editable PPT 可以作为插件直接启用，不需要运行 shell 安装脚本。部分本地 harness 功能需要额外工具：

- `python-pptx` 和 Pillow，用于生成 PPTX；
- LibreOffice 和 Poppler，用于渲染 QA；
- Node.js 和 Playwright，用于 HTML / 浏览器采集。

macOS：

```bash
brew install --cask libreoffice
brew install poppler
python3 -m pip install python-pptx pillow
```

Linux：

```bash
sudo apt-get install libreoffice poppler-utils
python3 -m pip install python-pptx pillow
```

Windows：

```powershell
py -m pip install python-pptx pillow pywin32
```

Windows 上使用 WPS/MS Office COM 自动化还需要安装 WPS Office 或 Microsoft Office。macOS 不支持 WPS COM。

如果缺少这些工具，HTML to Editable PPT 应该报告缺失能力，并使用当前可用的最佳兜底路径。
