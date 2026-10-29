#!/usr/bin/env python3
"""Build dist/excel-merge-tool.html — one self-contained, fully offline file.

Inlines src/styles.css, the vendored libraries and the app scripts into
src/index.html. Run from the repository root:  python3 build.py
"""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
OUT = ROOT / "dist" / "excel-merge-tool.html"


def read(rel_from_src: str) -> str:
    return (SRC / rel_from_src).resolve().read_text(encoding="utf-8")


def inline_script(match: re.Match) -> str:
    code = read(match.group(1))
    # A literal "</script" inside JS would end the tag early in HTML.
    code = re.sub(r"</(script)", r"<\\/\1", code, flags=re.IGNORECASE)
    return f"<script>\n{code}\n</script>"


def main() -> None:
    html = read("index.html")
    html = re.sub(
        r'<link rel="stylesheet" href="([^"]+)">',
        lambda m: f"<style>\n{read(m.group(1))}\n</style>",
        html,
    )
    html = re.sub(r'<script src="([^"]+)"></script>', inline_script, html)
    if re.search(r'<script src=|<link rel="stylesheet"', html):
        raise SystemExit("Build failed: an external reference was left in the output.")
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(html, encoding="utf-8")
    print(f"Built {OUT.relative_to(ROOT)} ({OUT.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
