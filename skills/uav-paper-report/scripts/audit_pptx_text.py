from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
NS.update({"p": "http://schemas.openxmlformats.org/presentationml/2006/main"})

SUSPICIOUS = [
    "报告口径",
    "讲解重点",
    "应该强调",
    "这页",
    "这一页",
    "这些图说明",
    "可以看到",
    "问题是",
    "代价是",
    "下一步应该",
    "如果结合",
    "不是只在仿真中有效",
    "这里主要",
]

BLANK_PATTERNS = [
    re.compile(r"\n\s*\n"),
    re.compile(r"^\s*[●•\-–]\s*$"),
]

BODY_LINE_BREAK = re.compile(r"[^\n]\n[^\n]")


def iter_slide_text(pptx: Path):
    with zipfile.ZipFile(pptx) as zf:
        names = sorted(
            (name for name in zf.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")),
            key=lambda x: int(re.search(r"slide(\d+)\.xml", x).group(1)),
        )
        for name in names:
            slide_no = int(re.search(r"slide(\d+)\.xml", name).group(1))
            root = ET.fromstring(zf.read(name))
            texts = [node.text or "" for node in root.findall(".//a:t", NS)]
            yield slide_no, "".join(texts)


def iter_slide_paragraphs(pptx: Path):
    with zipfile.ZipFile(pptx) as zf:
        names = sorted(
            (name for name in zf.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")),
            key=lambda x: int(re.search(r"slide(\d+)\.xml", x).group(1)),
        )
        for name in names:
            slide_no = int(re.search(r"slide(\d+)\.xml", name).group(1))
            root = ET.fromstring(zf.read(name))
            for shape in root.findall(".//p:sp", NS):
                name_node = shape.find("./p:nvSpPr/p:cNvPr", NS)
                shape_name = name_node.get("name", "") if name_node is not None else ""
                if shape_name.startswith("MATH_"):
                    continue
                text_body = shape.find("./p:txBody", NS)
                if text_body is None:
                    continue
                paragraphs = text_body.findall(".//a:p", NS)
                texts = ["".join(node.text or "" for node in paragraph.findall(".//a:t", NS)) for paragraph in paragraphs]
                has_text = any(text.strip() for text in texts)
                if not has_text:
                    continue
                for text, paragraph in zip(texts, paragraphs):
                    yield slide_no, text, paragraph


def iter_run_sizes(paragraph) -> list[float]:
    sizes = []
    for rpr in paragraph.findall(".//a:rPr", NS):
        raw = rpr.get("sz")
        if raw and raw.isdigit():
            sizes.append(int(raw) / 100.0)
    return sizes


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit PPTX text for AI/process wording and blank-line risks.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--fail-on-warning", action="store_true")
    args = parser.parse_args()

    if not args.pptx.exists():
        print(f"missing PPTX: {args.pptx}", file=sys.stderr)
        return 2

    warnings: list[str] = []
    for slide_no, text in iter_slide_text(args.pptx):
        for phrase in SUSPICIOUS:
            if phrase in text:
                warnings.append(f"slide {slide_no:02d}: suspicious wording `{phrase}`")
        for pattern in BLANK_PATTERNS:
            if pattern.search(text):
                warnings.append(f"slide {slide_no:02d}: possible blank paragraph or empty bullet")

    paragraphs = list(iter_slide_paragraphs(args.pptx))
    last_slide = max((slide_no for slide_no, _, _ in paragraphs), default=0)
    for slide_no, text, paragraph in paragraphs:
        if not text.strip():
            warnings.append(f"slide {slide_no:02d}: empty paragraph object")
            continue
        stripped = text.strip()
        if stripped in {"●", "•", "-", "–"}:
            warnings.append(f"slide {slide_no:02d}: standalone bullet marker")
        if BODY_LINE_BREAK.search(text) and slide_no not in (1, last_slide):
            warnings.append(f"slide {slide_no:02d}: manual line break inside body text")

        if slide_no in (1, last_slide):
            continue
        sizes = iter_run_sizes(paragraph)
        if not sizes:
            continue
        max_size = max(sizes)
        min_size = min(sizes)
        if stripped.startswith(("●", "•", "–", "-")) and min_size < 16.0:
            warnings.append(f"slide {slide_no:02d}: body bullet font too small ({min_size:.1f} pt)")
        if stripped.startswith(("●", "•", "–", "-")) and max_size > 22.5:
            warnings.append(f"slide {slide_no:02d}: body bullet font too large ({max_size:.1f} pt)")

    if warnings:
        print("Text audit warnings:")
        for warning in warnings:
            print(" -", warning)
    else:
        print("Text audit passed.")

    return 1 if warnings and args.fail_on_warning else 0


if __name__ == "__main__":
    raise SystemExit(main())
