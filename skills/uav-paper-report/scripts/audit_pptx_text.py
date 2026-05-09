from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}

DEFAULT_SLIDE_W = 12191695
DEFAULT_SLIDE_H = 6858000
EMU_TOL = 10000

SUSPICIOUS = [
    "报告口径",
    "从汇报角度",
    "讲解重点",
    "应该强调",
    "这一页",
    "本页",
    "这页",
    "这里主要",
    "可以看到",
    "问题是",
    "代价是",
    "下一步应该",
    "如果继续",
    "最值得借鉴",
    "一句话概括",
    "更像是",
    "key point to explain",
    "should emphasize",
    "future-work advice",
]

BLANK_PATTERNS = [
    re.compile(r"\n\s*\n"),
    re.compile(r"^\s*[●•–\-\s]*$"),
]

BODY_LINE_BREAK = re.compile(r"[^\n]\n[^\n]")
BODY_MARKERS = ("●", "•", "–", "-")
BULLET_LEVELS = {
    "●": "main",
    "•": "secondary",
    "–": "tertiary",
    "-": "tertiary",
}


def bullet_level(text: str) -> str | None:
    stripped = text.lstrip()
    if not stripped:
        return None
    return BULLET_LEVELS.get(stripped[0])


def collect_pptx(paths: list[Path]) -> list[Path]:
    files: list[Path] = []
    for path in paths:
        if path.is_dir():
            files.extend(sorted(path.rglob("*.pptx")))
        elif path.suffix.lower() == ".pptx":
            files.append(path)
    return files


def slide_size(zf: zipfile.ZipFile) -> tuple[int, int]:
    try:
        root = ET.fromstring(zf.read("ppt/presentation.xml"))
    except KeyError:
        return DEFAULT_SLIDE_W, DEFAULT_SLIDE_H
    node = root.find(".//p:sldSz", NS)
    if node is None:
        return DEFAULT_SLIDE_W, DEFAULT_SLIDE_H
    return int(node.get("cx", DEFAULT_SLIDE_W)), int(node.get("cy", DEFAULT_SLIDE_H))


def slide_names(zf: zipfile.ZipFile) -> list[str]:
    return sorted(
        (name for name in zf.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")),
        key=lambda x: int(re.search(r"slide(\d+)\.xml", x).group(1)),
    )


def iter_slide_text(pptx: Path):
    with zipfile.ZipFile(pptx) as zf:
        for name in slide_names(zf):
            slide_no = int(re.search(r"slide(\d+)\.xml", name).group(1))
            root = ET.fromstring(zf.read(name))
            texts = [node.text or "" for node in root.findall(".//a:t", NS)]
            yield slide_no, "".join(texts)


def iter_text_bodies(pptx: Path):
    with zipfile.ZipFile(pptx) as zf:
        for name in slide_names(zf):
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
                paragraphs = text_body.findall("./a:p", NS)
                texts = [
                    "".join(node.text or "" for node in paragraph.findall(".//a:t", NS))
                    for paragraph in paragraphs
                ]
                yield slide_no, shape_name, texts, paragraphs


def iter_shape_bounds(pptx: Path):
    with zipfile.ZipFile(pptx) as zf:
        width, height = slide_size(zf)
        for name in slide_names(zf):
            slide_no = int(re.search(r"slide(\d+)\.xml", name).group(1))
            root = ET.fromstring(zf.read(name))
            nodes = root.findall(".//p:sp", NS) + root.findall(".//p:pic", NS) + root.findall(".//p:graphicFrame", NS)
            for node in nodes:
                name_node = node.find(".//p:cNvPr", NS)
                shape_name = name_node.get("name", "") if name_node is not None else ""
                xfrm = node.find(".//a:xfrm", NS)
                if xfrm is None:
                    continue
                off = xfrm.find("./a:off", NS)
                ext = xfrm.find("./a:ext", NS)
                if off is None or ext is None:
                    continue
                left = int(off.get("x", "0"))
                top = int(off.get("y", "0"))
                w = int(ext.get("cx", "0"))
                h = int(ext.get("cy", "0"))
                yield slide_no, shape_name, left, top, w, h, width, height


def run_sizes(paragraph) -> list[float]:
    sizes = []
    for rpr in paragraph.findall(".//a:rPr", NS):
        raw = rpr.get("sz")
        if raw and raw.isdigit():
            sizes.append(int(raw) / 100.0)
    return sizes


def range_text(values: list[float]) -> str:
    return f"{min(values):.1f}-{max(values):.1f} pt"


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit PPTX text, text bodies, font hierarchy, and shape bounds.")
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument("--fail-on-warning", action="store_true")
    parser.add_argument("--body-min", type=float, default=16.0)
    parser.add_argument("--body-max", type=float, default=22.0)
    parser.add_argument(
        "--strict-body-hierarchy",
        action="store_true",
        help="Require fixed bullet-size bands and small per-level font drift.",
    )
    parser.add_argument("--main-min", type=float, default=17.6)
    parser.add_argument("--main-max", type=float, default=18.8)
    parser.add_argument("--secondary-min", type=float, default=16.4)
    parser.add_argument("--secondary-max", type=float, default=17.6)
    parser.add_argument("--tertiary-min", type=float, default=16.0)
    parser.add_argument("--tertiary-max", type=float, default=16.9)
    parser.add_argument("--max-level-spread", type=float, default=0.8)
    parser.add_argument("--max-run-spread", type=float, default=1.0)
    args = parser.parse_args()

    pptx_files = collect_pptx(args.paths)
    if not pptx_files:
        print("no PPTX files found", file=sys.stderr)
        return 2

    warnings: list[str] = []
    for pptx in pptx_files:
        prefix = pptx.name
        level_sizes: dict[str, list[float]] = {"main": [], "secondary": [], "tertiary": []}
        level_ranges = {
            "main": (args.main_min, args.main_max),
            "secondary": (args.secondary_min, args.secondary_max),
            "tertiary": (args.tertiary_min, args.tertiary_max),
        }
        for slide_no, text in iter_slide_text(pptx):
            for phrase in SUSPICIOUS:
                if phrase.lower() in text.lower():
                    warnings.append(f"{prefix} slide {slide_no:02d}: suspicious wording `{phrase}`")
            for pattern in BLANK_PATTERNS:
                if pattern.search(text):
                    warnings.append(f"{prefix} slide {slide_no:02d}: possible blank paragraph or empty bullet")

        paragraphs_for_last = []
        for slide_no, _, texts, paragraphs in iter_text_bodies(pptx):
            if any(text.strip() for text in texts):
                paragraphs_for_last.extend((slide_no, text, paragraph) for text, paragraph in zip(texts, paragraphs))
        last_slide = max((slide_no for slide_no, _, _ in paragraphs_for_last), default=0)

        for slide_no, shape_name, texts, paragraphs in iter_text_bodies(pptx):
            if texts and not any(text.strip() for text in texts):
                warnings.append(f"{prefix} slide {slide_no:02d}: empty text body on `{shape_name}`")
            if texts and any(text.strip() for text in texts) and any(not text.strip() for text in texts):
                warnings.append(f"{prefix} slide {slide_no:02d}: mixed empty paragraph in `{shape_name}`")
            for text, paragraph in zip(texts, paragraphs):
                if "\n" in text and slide_no not in (1, last_slide):
                    warnings.append(f"{prefix} slide {slide_no:02d}: manual newline in `{shape_name}`")
                stripped = text.strip()
                if not stripped:
                    continue
                if stripped in {"●", "•", "–", "-"}:
                    warnings.append(f"{prefix} slide {slide_no:02d}: standalone bullet marker")
                if BODY_LINE_BREAK.search(text) and slide_no not in (1, last_slide):
                    warnings.append(f"{prefix} slide {slide_no:02d}: manual line break inside body text")
                level = bullet_level(stripped)
                if level is not None:
                    sizes = run_sizes(paragraph)
                    if not sizes:
                        continue
                    min_size = min(sizes)
                    max_size = max(sizes)
                    level_sizes[level].append(max_size)
                    if min_size < args.body_min:
                        warnings.append(f"{prefix} slide {slide_no:02d}: body bullet font too small ({min_size:.1f} pt)")
                    if max_size > args.body_max:
                        warnings.append(f"{prefix} slide {slide_no:02d}: body bullet font too large ({max_size:.1f} pt)")
                    if args.strict_body_hierarchy:
                        expected_min, expected_max = level_ranges[level]
                        if min_size < expected_min or max_size > expected_max:
                            warnings.append(
                                f"{prefix} slide {slide_no:02d}: {level} bullet font outside fixed hierarchy "
                                f"({range_text(sizes)}, expected {expected_min:.1f}-{expected_max:.1f} pt)"
                            )
                        if max_size - min_size > args.max_run_spread:
                            warnings.append(
                                f"{prefix} slide {slide_no:02d}: mixed font sizes inside one bullet "
                                f"({range_text(sizes)})"
                            )

        for slide_no, shape_name, left, top, w, h, slide_w, slide_h in iter_shape_bounds(pptx):
            if (
                left < -EMU_TOL
                or top < -EMU_TOL
                or left + w > slide_w + EMU_TOL
                or top + h > slide_h + EMU_TOL
            ):
                warnings.append(f"{prefix} slide {slide_no:02d}: shape `{shape_name}` exceeds slide bounds")

        if args.strict_body_hierarchy:
            for level, values in level_sizes.items():
                if len(values) < 2:
                    continue
                spread = max(values) - min(values)
                if spread > args.max_level_spread:
                    warnings.append(
                        f"{prefix}: {level} bullet font hierarchy drift is too large "
                        f"({range_text(values)}, spread {spread:.1f} pt)"
                    )

    if warnings:
        print("Text/layout audit warnings:")
        for warning in warnings:
            print(" -", warning)
    else:
        print("Text/layout audit passed.")

    return 1 if warnings and args.fail_on_warning else 0


if __name__ == "__main__":
    raise SystemExit(main())
