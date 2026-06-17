#!/usr/bin/env python3
"""Build a clean release ZIP for the plugin.

The source repository includes examples and development files. The release ZIP
intentionally contains only the standard plugin/workflow package contents.
"""

from __future__ import annotations

import argparse
import json
import shutil
import tempfile
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PACKAGE_NAME = "html-to-editable-ppt"
INCLUDE_PATHS = [
    ".codex-plugin/plugin.json",
    "skills",
    "scripts/html_to_editable_ppt.py",
    "docs",
    "README.md",
    "README-zh.md",
    "INSTALL.md",
    "INSTALL-zh.md",
    "LICENSE",
]


def copy_path(src: Path, dst: Path) -> None:
    if src.is_dir():
        shutil.copytree(src, dst)
    else:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)


def iter_files(path: Path) -> list[Path]:
    return sorted(p for p in path.rglob("*") if p.is_file())


def build_zip(output: Path) -> Path:
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp:
        package_root = Path(tmp) / PACKAGE_NAME
        for rel in INCLUDE_PATHS:
            copy_path(ROOT / rel, package_root / rel)

        manifest = json.loads((package_root / ".codex-plugin/plugin.json").read_text(encoding="utf-8"))
        if manifest.get("name") != PACKAGE_NAME:
            raise SystemExit(f"plugin name mismatch: {manifest.get('name')!r}")

        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for file in iter_files(package_root):
                zf.write(file, file.relative_to(Path(tmp)).as_posix())

    return output


def parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Build HTML to Editable PPT release ZIP")
    p.add_argument("--output", default=f"dist/{PACKAGE_NAME}-plugin.zip")
    return p


def main() -> None:
    args = parser().parse_args()
    output = build_zip(Path(args.output))
    print(output)


if __name__ == "__main__":
    main()
