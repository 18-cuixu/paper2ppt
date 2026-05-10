from __future__ import annotations

import argparse
import io
import math
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

DEFAULT_SLIDE_W = 12191695
DEFAULT_SLIDE_H = 6858000
EMU_TOL = 10000
EMU_PER_INCH = 914400
PT_PER_INCH = 72

SUSPICIOUS = [
    "报告口径",
    "从汇报角度",
    "汇报重点",
    "讲解重点",
    "应该强调",
    "表格强调",
    "收益项强调",
    "方法强调",
    "这一页",
    "本页",
    "这页",
    "这里主要",
    "可以看到",
    "问题是",
    "代价是",
    "下一步应该",
    "如果继续",
    "后续工作",
    "适合借鉴",
    "借鉴到",
    "最值得借鉴",
    "一句话概括",
    "更像是",
    "key point to explain",
    "should emphasize",
    "future-work advice",
    "模板压力",
    "回归检查",
    "回归对象",
    "通过标准",
    "用于测试",
    "该页检查",
    "生成器应",
    "PPTX 审计",
    "template smoke",
    "multi-template regression",
    "scaffold",
    "QA",
]

BLANK_PATTERNS = [
    re.compile(r"\n\s*\n"),
    re.compile(r"^\s*[●•–\-\s]*$"),
]

BODY_LINE_BREAK = re.compile(r"[^\n]\n[^\n]")
CODE_STYLE_MATH = re.compile(
    r"(?:[A-Za-zΑ-Ωα-ωϕϑϖϱ][A-Za-z0-9Α-Ωα-ωϕϑϖϱ]*|[πλΔηψθγℓϕ])_(?:\{[^}\s]+\}|[A-Za-z0-9Α-Ωα-ωϕϑϖϱ]+)"
)
BODY_MARKERS = ("●", "•", "–", "-")
BULLET_LEVELS = {
    "●": "main",
    "•": "secondary",
    "–": "tertiary",
    "-": "tertiary",
}
CONTENT_SHAPE_NAMES = ("MATH_", "RULE", "METRIC_BOX_")
INTENTIONAL_BREAK_SHAPE_NAMES = ("DIAG_",)

PROFILE_PRESETS = {
    "compact": {
        "body_min": 16.0,
        "plain_body_min": 14.0,
        "table_min": 11.6,
        "metric_value_min": 14.8,
        "metric_label_min": 11.4,
        "metric_note_min": 14.0,
        "diagram_label_min": 13.8,
        "main_min": 17.6,
        "main_max": 18.8,
        "secondary_min": 16.4,
        "secondary_max": 17.6,
        "tertiary_min": 16.0,
        "tertiary_max": 16.9,
        "max_level_spread": 0.8,
        "max_run_spread": 1.0,
        "max_formula_width_factor": 1.05,
    },
    "dense-visual": {
        "body_min": 15.8,
        "plain_body_min": 14.0,
        "table_min": 11.6,
        "metric_value_min": 14.8,
        "metric_label_min": 11.4,
        "metric_note_min": 14.0,
        "diagram_label_min": 13.8,
        "main_min": 17.6,
        "main_max": 19.1,
        "secondary_min": 15.8,
        "secondary_max": 18.0,
        "tertiary_min": 15.8,
        "tertiary_max": 17.1,
        "max_level_spread": 2.0,
        "max_run_spread": 1.2,
        "max_formula_width_factor": 1.15,
    },
    "classic-large": {
        "body_min": 16.0,
        "plain_body_min": 14.5,
        "table_min": 12.0,
        "metric_value_min": 15.4,
        "metric_label_min": 11.8,
        "metric_note_min": 14.5,
        "diagram_label_min": 13.8,
        "main_min": 18.0,
        "main_max": 21.2,
        "secondary_min": 16.8,
        "secondary_max": 21.0,
        "tertiary_min": 16.0,
        "tertiary_max": 20.2,
        "max_level_spread": 4.4,
        "max_run_spread": 1.4,
        "max_formula_width_factor": 1.12,
    },
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


def iter_text_paragraphs(pptx: Path):
    with zipfile.ZipFile(pptx) as zf:
        for name in slide_names(zf):
            slide_no = int(re.search(r"slide(\d+)\.xml", name).group(1))
            root = ET.fromstring(zf.read(name))
            containers = root.findall(".//p:sp", NS) + root.findall(".//p:graphicFrame", NS)
            for idx, container in enumerate(containers):
                name_node = container.find(".//p:cNvPr", NS)
                shape_name = name_node.get("name", f"shape-{idx}") if name_node is not None else f"shape-{idx}"
                for paragraph in container.findall(".//a:p", NS):
                    text = "".join(node.text or "" for node in paragraph.findall(".//a:t", NS))
                    yield slide_no, shape_name, text


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


def iter_text_boxes(pptx: Path):
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
                xfrm = shape.find(".//a:xfrm", NS)
                if text_body is None or xfrm is None:
                    continue
                off = xfrm.find("./a:off", NS)
                ext = xfrm.find("./a:ext", NS)
                if off is None or ext is None:
                    continue
                paragraphs = text_body.findall("./a:p", NS)
                texts = [
                    "".join(node.text or "" for node in paragraph.findall(".//a:t", NS))
                    for paragraph in paragraphs
                ]
                yield (
                    slide_no,
                    shape_name,
                    int(off.get("x", "0")) / EMU_PER_INCH,
                    int(off.get("y", "0")) / EMU_PER_INCH,
                    int(ext.get("cx", "0")) / EMU_PER_INCH,
                    int(ext.get("cy", "0")) / EMU_PER_INCH,
                    texts,
                    paragraphs,
                )


def iter_math_text_bodies(pptx: Path):
    with zipfile.ZipFile(pptx) as zf:
        for name in slide_names(zf):
            slide_no = int(re.search(r"slide(\d+)\.xml", name).group(1))
            root = ET.fromstring(zf.read(name))
            for shape in root.findall(".//p:sp", NS):
                name_node = shape.find("./p:nvSpPr/p:cNvPr", NS)
                shape_name = name_node.get("name", "") if name_node is not None else ""
                if not shape_name.startswith("MATH_"):
                    continue
                text_body = shape.find("./p:txBody", NS)
                if text_body is None:
                    continue
                xfrm = shape.find(".//a:xfrm", NS)
                width = None
                if xfrm is not None:
                    ext = xfrm.find("./a:ext", NS)
                    if ext is not None:
                        width = int(ext.get("cx", "0")) / EMU_PER_INCH
                paragraphs = text_body.findall("./a:p", NS)
                texts = [
                    "".join(node.text or "" for node in paragraph.findall(".//a:t", NS))
                    for paragraph in paragraphs
                ]
                yield slide_no, shape_name, texts, paragraphs, width


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


def iter_semantic_table_layouts(pptx: Path):
    semantic_headers = {
        ("符号", "含义"): 0.58,
        ("变量", "含义"): 0.58,
        ("对象", "进入方式"): 0.56,
        ("对象", "约束"): 0.56,
    }
    with zipfile.ZipFile(pptx) as zf:
        for name in slide_names(zf):
            slide_no = int(re.search(r"slide(\d+)\.xml", name).group(1))
            root = ET.fromstring(zf.read(name))
            for frame in root.findall(".//p:graphicFrame", NS):
                name_node = frame.find(".//p:cNvPr", NS)
                shape_name = name_node.get("name", "") if name_node is not None else ""
                table = frame.find(".//a:tbl", NS)
                if table is None:
                    continue
                rows = table.findall("./a:tr", NS)
                if not rows:
                    continue
                header_cells = rows[0].findall("./a:tc", NS)
                headers = tuple("".join(t.text or "" for t in cell.findall(".//a:t", NS)).strip() for cell in header_cells)
                min_second_ratio = semantic_headers.get(headers)
                if min_second_ratio is None:
                    continue
                widths = []
                for col in table.findall("./a:tblGrid/a:gridCol", NS):
                    raw = col.get("w")
                    if raw and raw.isdigit():
                        widths.append(int(raw))
                if len(widths) != len(headers) or sum(widths) <= 0:
                    continue
                ratios = [w / sum(widths) for w in widths]
                yield slide_no, shape_name, headers, ratios, min_second_ratio


def iter_table_paragraphs(pptx: Path):
    with zipfile.ZipFile(pptx) as zf:
        for name in slide_names(zf):
            slide_no = int(re.search(r"slide(\d+)\.xml", name).group(1))
            root = ET.fromstring(zf.read(name))
            for frame in root.findall(".//p:graphicFrame", NS):
                name_node = frame.find(".//p:cNvPr", NS)
                shape_name = name_node.get("name", "") if name_node is not None else ""
                table = frame.find(".//a:tbl", NS)
                if table is None:
                    continue
                for row_idx, row in enumerate(table.findall("./a:tr", NS), start=1):
                    for col_idx, cell in enumerate(row.findall("./a:tc", NS), start=1):
                        for paragraph in cell.findall(".//a:p", NS):
                            text = "".join(node.text or "" for node in paragraph.findall(".//a:t", NS))
                            yield slide_no, shape_name, row_idx, col_idx, text, paragraph


def resolve_relationship_target(source: str, target: str) -> str:
    base = Path(source).parent
    parts: list[str] = []
    for part in (base / target).as_posix().split("/"):
        if part in ("", "."):
            continue
        if part == "..":
            if parts:
                parts.pop()
            continue
        parts.append(part)
    return "/".join(parts)


def slide_relationships(zf: zipfile.ZipFile, slide_xml: str) -> dict[str, str]:
    rels_path = f"{Path(slide_xml).parent}/_rels/{Path(slide_xml).name}.rels"
    try:
        root = ET.fromstring(zf.read(rels_path))
    except KeyError:
        return {}
    rels: dict[str, str] = {}
    for rel in root.findall("./rel:Relationship", NS):
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            rels[rid] = resolve_relationship_target(slide_xml, target)
    return rels


def has_source_crop(pic) -> bool:
    src = pic.find(".//a:blipFill/a:srcRect", NS)
    if src is None:
        return False
    return any(src.get(key) not in (None, "0") for key in ("l", "r", "t", "b"))


def iter_picture_aspects(pptx: Path):
    with zipfile.ZipFile(pptx) as zf:
        slide_xmls = slide_names(zf)
        last_slide = len(slide_xmls)
        for name in slide_xmls:
            slide_no = int(re.search(r"slide(\d+)\.xml", name).group(1))
            root = ET.fromstring(zf.read(name))
            rels = slide_relationships(zf, name)
            for pic in root.findall(".//p:pic", NS):
                name_node = pic.find(".//p:cNvPr", NS)
                shape_name = name_node.get("name", "") if name_node is not None else ""
                xfrm = pic.find(".//a:xfrm", NS)
                if xfrm is None:
                    continue
                off = xfrm.find("./a:off", NS)
                ext = xfrm.find("./a:ext", NS)
                if off is None or ext is None:
                    continue
                top = int(off.get("y", "0")) / EMU_PER_INCH
                width = int(ext.get("cx", "0")) / EMU_PER_INCH
                height = int(ext.get("cy", "0")) / EMU_PER_INCH
                if slide_no in (1, last_slide) or top < 1.0 or width <= 0 or height <= 0:
                    continue
                if has_source_crop(pic):
                    continue
                blip = pic.find(".//a:blip", NS)
                rid = blip.get(f"{{{NS['r']}}}embed") if blip is not None else None
                target = rels.get(rid or "")
                if not target or target not in zf.namelist():
                    continue
                try:
                    with Image.open(io.BytesIO(zf.read(target))) as image:
                        image_ratio = image.width / image.height
                except Exception:
                    continue
                shape_ratio = width / height
                yield slide_no, shape_name, shape_ratio, image_ratio


def iter_content_bounds(pptx: Path):
    with zipfile.ZipFile(pptx) as zf:
        slide_w, slide_h = slide_size(zf)
        for name in slide_names(zf):
            slide_no = int(re.search(r"slide(\d+)\.xml", name).group(1))
            root = ET.fromstring(zf.read(name))
            nodes = root.findall(".//p:sp", NS) + root.findall(".//p:pic", NS) + root.findall(".//p:graphicFrame", NS)
            for idx, node in enumerate(nodes):
                name_node = node.find(".//p:cNvPr", NS)
                shape_name = name_node.get("name", f"shape-{idx}") if name_node is not None else f"shape-{idx}"
                text = "".join(t.text or "" for t in node.findall(".//a:t", NS)).strip()
                is_rule = shape_name.startswith("RULE")
                is_metric_box = shape_name.startswith("METRIC_BOX_")
                is_content = bool(text) or node.tag.endswith("pic") or node.tag.endswith("graphicFrame") or is_rule or is_metric_box
                if not is_content and not shape_name.startswith(CONTENT_SHAPE_NAMES):
                    continue
                xfrm = node.find(".//a:xfrm", NS)
                if xfrm is None:
                    continue
                off = xfrm.find("./a:off", NS)
                ext = xfrm.find("./a:ext", NS)
                if off is None or ext is None:
                    continue
                left = int(off.get("x", "0")) / EMU_PER_INCH
                top = int(off.get("y", "0")) / EMU_PER_INCH
                right = left + int(ext.get("cx", "0")) / EMU_PER_INCH
                bottom = top + int(ext.get("cy", "0")) / EMU_PER_INCH
                yield slide_no, shape_name, left, top, right, bottom, slide_w / EMU_PER_INCH, slide_h / EMU_PER_INCH


def run_sizes(paragraph) -> list[float]:
    sizes = []
    for rpr in paragraph.findall(".//a:rPr", NS):
        raw = rpr.get("sz")
        if raw and raw.isdigit():
            sizes.append(int(raw) / 100.0)
    return sizes


def shape_text_role(shape_name: str) -> str | None:
    if shape_name.startswith("METRIC_VALUE_"):
        return "metric_value"
    if shape_name.startswith("METRIC_LABEL_"):
        return "metric_label"
    if shape_name.startswith("METRIC_NOTE_"):
        return "metric_note"
    if shape_name.startswith("DIAG_"):
        return "diagram_label"
    if shape_name.startswith(("BODY_", "TEXT_", "NOTE_", "CALLOUT_")):
        return "plain_body"
    return None


def paragraph_spacing_pt(paragraph) -> tuple[float, float]:
    ppr = paragraph.find("./a:pPr", NS)
    if ppr is None:
        return 0.0, 0.0

    def read_spacing(tag: str) -> float:
        node = ppr.find(f"./a:{tag}/a:spcPts", NS)
        if node is not None:
            raw = node.get("val")
            if raw and raw.lstrip("-").isdigit():
                return int(raw) / 100.0
        return 0.0

    return read_spacing("spcBef"), read_spacing("spcAft")


def has_manual_break(paragraph) -> bool:
    return paragraph.find(".//a:br", NS) is not None


def allows_manual_break(shape_name: str) -> bool:
    return shape_name.startswith(INTENTIONAL_BREAK_SHAPE_NAMES)


def estimated_text_width_pt(text: str, size: float) -> float:
    units = 0.0
    for ch in re.sub(r"\s+", " ", text).strip():
        if "\u4e00" <= ch <= "\u9fff":
            units += 0.95
        elif ch.isspace():
            units += 0.28
        elif ch in ",.;:()[]{}+-=≤≥≈→⇒⇔∑∫‖·/":
            units += 0.36
        elif ch in "ilI|":
            units += 0.26
        elif ch in "MWQ@":
            units += 0.78
        else:
            units += 0.50
    return units * size


def paragraph_line_count(text: str, sizes: list[float], width_in: float) -> tuple[int, bool]:
    cleaned = re.sub(r"\s+", " ", text).strip()
    if not cleaned:
        return 0, False
    size = max(sizes) if sizes else 17.0
    available_pt = max(12.0, (width_in - 0.10) * PT_PER_INCH)
    estimated_lines = max(1, math.ceil(estimated_text_width_pt(cleaned, size) / available_pt))
    tokens = re.findall(r"[A-Za-z0-9][A-Za-z0-9+\-/_.]*|[^\s]", cleaned)
    has_orphan = any(estimated_text_width_pt(token, size) > available_pt * 0.82 for token in tokens)
    return estimated_lines, has_orphan


def text_width_overflow(text: str, sizes: list[float], width_in: float, factor: float = 0.96) -> bool:
    cleaned = re.sub(r"\s+", " ", text).strip()
    if not cleaned:
        return False
    size = max(sizes) if sizes else 12.0
    available_pt = max(12.0, (width_in - 0.10) * PT_PER_INCH)
    return estimated_text_width_pt(cleaned, size) > available_pt * factor


def diagram_label_overflow(text: str, sizes: list[float], width_in: float, height_in: float) -> bool:
    cleaned = re.sub(r"\s+", " ", text).strip()
    if not cleaned:
        return False
    size = max(sizes) if sizes else 12.0
    line_count, orphan_risk = paragraph_line_count(cleaned, sizes, width_in)
    if orphan_risk:
        return True
    if line_count <= 1:
        # Diagram nodes often contain compact math or mixed Chinese/English labels.
        # The width estimator is intentionally conservative, so keep a small
        # tolerance for single-line labels while still failing true orphan tokens.
        return text_width_overflow(cleaned, sizes, width_in, factor=1.06)
    if line_count > 2:
        return True
    estimated_height = line_count * size * 1.10 / PT_PER_INCH + 0.08
    return estimated_height > max(0.05, height_in * 0.88)


def estimated_text_height_in(texts: list[str], paragraphs, width_in: float) -> tuple[float, int, bool]:
    total_pt = 0.0
    max_lines = 0
    orphan_risk = False
    for text, paragraph in zip(texts, paragraphs):
        stripped = text.strip()
        if not stripped:
            continue
        sizes = run_sizes(paragraph)
        size = max(sizes) if sizes else 17.0
        line_count, has_orphan = paragraph_line_count(stripped, sizes, width_in)
        space_before, space_after = paragraph_spacing_pt(paragraph)
        total_pt += space_before + line_count * size * 1.10 + min(space_after, 3.0)
        max_lines = max(max_lines, line_count)
        orphan_risk = orphan_risk or has_orphan
    # PowerPoint text boxes include top/bottom inner margins even when the XML
    # shape looks tall enough; keep a conservative rendered slack allowance.
    return total_pt / PT_PER_INCH + 0.08, max_lines, orphan_risk


def range_text(values: list[float]) -> str:
    return f"{min(values):.1f}-{max(values):.1f} pt"


def overlap_area(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    return max(0.0, min(a[2], b[2]) - max(a[0], b[0])) * max(0.0, min(a[3], b[3]) - max(a[1], b[1]))


def vertical_gap(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    if a[3] <= b[1]:
        return b[1] - a[3]
    if b[3] <= a[1]:
        return a[1] - b[3]
    return 0.0


def horizontal_overlap(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    return max(0.0, min(a[2], b[2]) - max(a[0], b[0]))


def should_check_overlap(name: str, box: tuple[float, float, float, float], slide_h: float) -> bool:
    left, top, right, bottom = box
    if top < 1.0 or bottom > slide_h - 0.42:
        return False
    if right - left < 0.05 or bottom - top < 0.05:
        return False
    if name.startswith(("RULE", "METRIC_BOX_")):
        return True
    lowered = name.lower()
    if "connector" in lowered or "rectangle" in lowered and not name.startswith("MATH_"):
        return False
    return True


def allowed_overlap(name_a: str, name_b: str) -> bool:
    if name_a.startswith("METRIC_BOX_") and name_b.startswith(("METRIC_VALUE_", "METRIC_LABEL_")):
        return True
    if name_a.startswith("METRIC_BOX_") and name_b.startswith("METRIC_NOTE_"):
        return True
    if name_b.startswith("METRIC_BOX_") and name_a.startswith(("METRIC_VALUE_", "METRIC_LABEL_")):
        return True
    if name_b.startswith("METRIC_BOX_") and name_a.startswith("METRIC_NOTE_"):
        return True
    return False


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit PPTX text, text bodies, font hierarchy, and shape bounds.")
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument("--fail-on-warning", action="store_true")
    parser.add_argument("--body-min", type=float, default=16.0)
    parser.add_argument("--body-max", type=float, default=22.0)
    parser.add_argument("--plain-body-min", type=float, default=14.0)
    parser.add_argument("--table-min", type=float, default=11.6)
    parser.add_argument("--metric-value-min", type=float, default=14.8)
    parser.add_argument("--metric-label-min", type=float, default=11.4)
    parser.add_argument("--metric-note-min", type=float, default=14.0)
    parser.add_argument("--diagram-label-min", type=float, default=13.8)
    parser.add_argument(
        "--strict-body-hierarchy",
        action="store_true",
        help="Require fixed bullet-size bands and small per-level font drift.",
    )
    parser.add_argument(
        "--profile",
        choices=sorted(PROFILE_PRESETS),
        default=None,
        help="Use a template-aware body hierarchy preset before applying explicit CLI overrides.",
    )
    parser.add_argument("--main-min", type=float, default=17.6)
    parser.add_argument("--main-max", type=float, default=18.8)
    parser.add_argument("--secondary-min", type=float, default=16.4)
    parser.add_argument("--secondary-max", type=float, default=17.6)
    parser.add_argument("--tertiary-min", type=float, default=16.0)
    parser.add_argument("--tertiary-max", type=float, default=16.9)
    parser.add_argument("--max-level-spread", type=float, default=0.8)
    parser.add_argument("--max-run-spread", type=float, default=1.0)
    parser.add_argument("--max-body-space-after", type=float, default=3.0)
    parser.add_argument("--max-body-space-before", type=float, default=0.5)
    parser.add_argument("--max-formula-width-factor", type=float, default=1.05)
    parser.add_argument("--max-text-height-fill", type=float, default=0.94)
    parser.add_argument("--max-body-lines", type=int, default=3)
    parser.add_argument("--min-wide-text-width", type=float, default=5.4)
    parser.add_argument(
        "--max-picture-ratio-drift",
        type=float,
        default=0.08,
        help="Warn when an uncropped content picture's displayed aspect ratio differs from its source image.",
    )
    args = parser.parse_args()

    if args.profile:
        for key, value in PROFILE_PRESETS[args.profile].items():
            setattr(args, key, value)

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

        for slide_no, shape_name, text in iter_text_paragraphs(pptx):
            for match in CODE_STYLE_MATH.finditer(text):
                warnings.append(
                    f"{prefix} slide {slide_no:02d}: code-style math label `{match.group(0)}` in `{shape_name}`"
                )

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
                if "\n" in text and slide_no not in (1, last_slide) and not allows_manual_break(shape_name):
                    warnings.append(f"{prefix} slide {slide_no:02d}: manual newline in `{shape_name}`")
                if has_manual_break(paragraph) and slide_no not in (1, last_slide) and not allows_manual_break(shape_name):
                    warnings.append(f"{prefix} slide {slide_no:02d}: manual line break element in `{shape_name}`")
                stripped = text.strip()
                if not stripped:
                    continue
                space_before, space_after = paragraph_spacing_pt(paragraph)
                if space_before > args.max_body_space_before:
                    warnings.append(
                        f"{prefix} slide {slide_no:02d}: body paragraph has excessive space_before "
                        f"({space_before:.1f} pt) in `{shape_name}`"
                    )
                if space_after > args.max_body_space_after:
                    warnings.append(
                        f"{prefix} slide {slide_no:02d}: body paragraph has excessive space_after "
                        f"({space_after:.1f} pt) in `{shape_name}`"
                    )
                if stripped in {"●", "•", "–", "-"}:
                    warnings.append(f"{prefix} slide {slide_no:02d}: standalone bullet marker")
                if BODY_LINE_BREAK.search(text) and slide_no not in (1, last_slide):
                    warnings.append(f"{prefix} slide {slide_no:02d}: manual line break inside body text")
                level = bullet_level(stripped)
                role = shape_text_role(shape_name)
                if role is not None:
                    sizes = run_sizes(paragraph)
                    if sizes:
                        min_size = min(sizes)
                        threshold = getattr(args, f"{role}_min")
                        if min_size < threshold:
                            role_label = role.replace("_", " ")
                            warnings.append(
                                f"{prefix} slide {slide_no:02d}: {role_label} font too small "
                                f"({min_size:.1f} pt, expected at least {threshold:.1f} pt) in `{shape_name}`"
                            )
                    if args.strict_body_hierarchy and role == "plain_body" and level is None:
                        warnings.append(
                            f"{prefix} slide {slide_no:02d}: body paragraph lacks bullet marker in `{shape_name}`"
                        )
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

        if args.strict_body_hierarchy:
            for slide_no, shape_name, row_idx, col_idx, text, paragraph in iter_table_paragraphs(pptx):
                stripped = text.strip()
                if not stripped:
                    continue
                sizes = run_sizes(paragraph)
                if not sizes:
                    continue
                min_size = min(sizes)
                if min_size < args.table_min:
                    warnings.append(
                        f"{prefix} slide {slide_no:02d}: table text font too small "
                        f"({min_size:.1f} pt, expected at least {args.table_min:.1f} pt) "
                        f"in `{shape_name}` cell R{row_idx}C{col_idx}"
                    )

        text_overflow_boxes: dict[tuple[int, str], tuple[float, tuple[float, float, float, float]]] = {}
        if args.strict_body_hierarchy:
            for slide_no, shape_name, left, top, width_in, height_in, texts, paragraphs in iter_text_boxes(pptx):
                role = shape_text_role(shape_name)
                if role is None or slide_no in (1, last_slide):
                    continue
                if role not in {"plain_body", "metric_note"}:
                    continue
                if not any(text.strip() for text in texts):
                    continue
                estimated_height, max_lines, orphan_risk = estimated_text_height_in(texts, paragraphs, width_in)
                fill_ratio = estimated_height / max(height_in, 0.01)
                text_overflow_boxes[(slide_no, shape_name)] = (
                    estimated_height,
                    (left, top, left + width_in, max(top + height_in, top + estimated_height)),
                )
                if fill_ratio > args.max_text_height_fill:
                    warnings.append(
                        f"{prefix} slide {slide_no:02d}: text has too little rendered slack in `{shape_name}` "
                        f"({estimated_height:.2f} in estimated for {height_in:.2f} in box)"
                    )
                if (
                    role == "plain_body"
                    and width_in < args.min_wide_text_width
                    and max_lines > args.max_body_lines
                ):
                    warnings.append(
                        f"{prefix} slide {slide_no:02d}: narrow body box causes avoidable wrapping in `{shape_name}` "
                        f"({width_in:.2f} in wide, {max_lines} estimated lines)"
                    )
                if orphan_risk:
                    warnings.append(
                        f"{prefix} slide {slide_no:02d}: long token may create orphan/wrapped line in `{shape_name}`"
                    )

        if args.strict_body_hierarchy:
            for slide_no, shape_name, left, top, width_in, height_in, texts, paragraphs in iter_text_boxes(pptx):
                role = shape_text_role(shape_name)
                if role != "diagram_label" or slide_no in (1, last_slide):
                    continue
                for text, paragraph in zip(texts, paragraphs):
                    stripped = text.strip()
                    if not stripped:
                        continue
                    sizes = run_sizes(paragraph)
                    if not sizes:
                        continue
                    min_size = min(sizes)
                    if min_size < args.diagram_label_min:
                        warnings.append(
                            f"{prefix} slide {slide_no:02d}: diagram label font too small "
                            f"({min_size:.1f} pt, expected at least {args.diagram_label_min:.1f} pt) in `{shape_name}`"
                        )
                    if diagram_label_overflow(stripped, sizes, width_in, height_in):
                        warnings.append(
                            f"{prefix} slide {slide_no:02d}: diagram label may overflow its node in `{shape_name}`"
                        )

        if args.strict_body_hierarchy:
            for slide_no, shape_name, texts, paragraphs, width_in in iter_math_text_bodies(pptx):
                for text, paragraph in zip(texts, paragraphs):
                    if has_manual_break(paragraph):
                        warnings.append(f"{prefix} slide {slide_no:02d}: manual line break element in `{shape_name}`")
                    if "\n" in text:
                        warnings.append(f"{prefix} slide {slide_no:02d}: manual newline in `{shape_name}`")
                    stripped = text.strip()
                    if not stripped or width_in is None or not shape_name.startswith("MATH_BODY"):
                        continue
                    sizes = run_sizes(paragraph)
                    size = max(sizes) if sizes else 19.0
                    estimated = estimated_text_width_pt(stripped, size)
                    available = width_in * PT_PER_INCH
                    if estimated > available * args.max_formula_width_factor:
                        warnings.append(
                            f"{prefix} slide {slide_no:02d}: formula row may wrap or overflow in `{shape_name}` "
                            f"({estimated:.0f} pt estimated > {available:.0f} pt available)"
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
            for slide_no, shape_name, headers, ratios, min_second_ratio in iter_semantic_table_layouts(pptx):
                if len(ratios) >= 2 and ratios[1] < min_second_ratio:
                    warnings.append(
                        f"{prefix} slide {slide_no:02d}: semantic table `{shape_name}` has narrow meaning column "
                        f"({headers[1]} {ratios[1]:.0%}, expected at least {min_second_ratio:.0%})"
                    )

        if args.strict_body_hierarchy:
            slide_items: dict[int, list[tuple[str, tuple[float, float, float, float]]]] = {}
            for slide_no, shape_name, left, top, right, bottom, slide_w, slide_h in iter_content_bounds(pptx):
                box = (left, top, right, bottom)
                overflow = text_overflow_boxes.get((slide_no, shape_name))
                if overflow is not None:
                    box = overflow[1]
                if should_check_overlap(shape_name, box, slide_h):
                    slide_items.setdefault(slide_no, []).append((shape_name, box))
            for slide_no, items in slide_items.items():
                for i in range(len(items)):
                    name_a, box_a = items[i]
                    for j in range(i + 1, len(items)):
                        name_b, box_b = items[j]
                        if allowed_overlap(name_a, name_b):
                            continue
                        if name_a.startswith("RULE") or name_b.startswith("RULE"):
                            gap = vertical_gap(box_a, box_b)
                            h_overlap = horizontal_overlap(box_a, box_b)
                            if gap < 0.12 and h_overlap > 0.35:
                                warnings.append(
                                    f"{prefix} slide {slide_no:02d}: text or content too close to rule "
                                    f"`{name_a}` and `{name_b}`"
                                )
                        area = overlap_area(box_a, box_b)
                        if area <= 0.015:
                            continue
                        smaller = min(
                            max(0.001, (box_a[2] - box_a[0]) * (box_a[3] - box_a[1])),
                            max(0.001, (box_b[2] - box_b[0]) * (box_b[3] - box_b[1])),
                        )
                        if area / smaller > 0.08:
                            warnings.append(
                                f"{prefix} slide {slide_no:02d}: content shapes overlap "
                                f"`{name_a}` and `{name_b}`"
                            )

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

            for slide_no, shape_name, shape_ratio, image_ratio in iter_picture_aspects(pptx):
                drift = abs(math.log(shape_ratio / image_ratio))
                if drift > args.max_picture_ratio_drift:
                    warnings.append(
                        f"{prefix} slide {slide_no:02d}: picture `{shape_name}` aspect ratio drift "
                        f"({shape_ratio:.2f} displayed vs {image_ratio:.2f} source)"
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
