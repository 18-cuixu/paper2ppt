from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

import run_template_matrix as matrix_runner


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_PAPERS = ROOT / "assets" / "template-profiles" / "stress-papers.json"
DEFAULT_TEMPLATES = ROOT / "assets" / "template-profiles" / "template-smoke.local.example.json"
AUDIT = ROOT / "scripts" / "audit_pptx_text.py"
RENDER = ROOT / "scripts" / "render_pptx_previews.py"
SCAN = ROOT / "scripts" / "scan_rendered_slides.py"

EMU_PER_INCH = 914400
FONT = "Times New Roman"
BLACK = RGBColor(25, 25, 25)
GRAY = RGBColor(96, 105, 116)
LIGHT_GRAY = RGBColor(243, 245, 248)
MID_GRAY = RGBColor(210, 215, 222)
RED = RGBColor(205, 32, 44)
WHITE = RGBColor(255, 255, 255)

PROFILE_SIZE = {
    "compact": {"main": 18.2, "secondary": 16.8, "tertiary": 16.2, "formula": 19.0, "table": 11.8},
    "dense-visual": {"main": 18.2, "secondary": 16.8, "tertiary": 16.2, "formula": 19.0, "table": 11.8},
    "classic-large": {"main": 19.2, "secondary": 18.4, "tertiary": 17.0, "formula": 20.0, "table": 12.4},
}

THEMES = {
    "blue": RGBColor(0, 121, 192),
    "cyan": RGBColor(0, 145, 166),
    "green": RGBColor(35, 145, 95),
    "red": RGBColor(172, 38, 47),
    "purple": RGBColor(92, 78, 152),
    "dark": RGBColor(25, 41, 58),
}


def expand_path(raw: str, *, template_root: Path | None) -> Path:
    text = os.path.expandvars(raw)
    if template_root is not None:
        text = text.replace("${PPTAGENT_TEMPLATE_ROOT}", str(template_root))
    return Path(text).expanduser().resolve()


def rgb_tuple(color: RGBColor) -> tuple[int, int, int]:
    return color[0], color[1], color[2]


def brightness(color: RGBColor) -> float:
    r, g, b = rgb_tuple(color)
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255.0


def theme_color(template: Path, theme: str | None) -> RGBColor:
    configured = THEMES.get((theme or "").lower())
    if configured is not None:
        return configured

    colors: Counter[tuple[int, int, int]] = Counter()
    try:
        prs = Presentation(str(template))
        for slide in list(prs.slides)[:4]:
            for shape in slide.shapes:
                try:
                    fill = shape.fill
                    if fill.type != 1:
                        continue
                    rgb = fill.fore_color.rgb
                    if rgb is None:
                        continue
                    r, g, b = rgb_tuple(rgb)
                    if max(r, g, b) > 238 or max(r, g, b) < 35:
                        continue
                    if abs(r - g) + abs(g - b) + abs(r - b) < 42:
                        continue
                    colors[(r, g, b)] += 1
                except Exception:
                    continue
    except Exception:
        pass
    if not colors:
        return THEMES["blue"]
    r, g, b = colors.most_common(1)[0][0]
    return RGBColor(r, g, b)


def clear_slides(prs: Presentation) -> None:
    sld_id_list = prs.slides._sldIdLst  # noqa: SLF001 - python-pptx has no public slide-delete API.
    for sld_id in list(sld_id_list):
        prs.part.drop_rel(sld_id.rId)
        sld_id_list.remove(sld_id)


def set_typeface(rpr, font: str = FONT) -> None:
    for tag in ("latin", "ea", "cs"):
        node = rpr.find(qn(f"a:{tag}"))
        if node is None:
            node = rpr.makeelement(qn(f"a:{tag}"))
            rpr.append(node)
        node.set("typeface", font)


def set_run(run, size: float, *, bold: bool = False, color: RGBColor = BLACK) -> None:
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    set_typeface(run._r.get_or_add_rPr(), FONT)  # noqa: SLF001


def add_textbox(slide, left: float, top: float, width: float, height: float, *, name: str):
    shape = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    shape.name = name
    shape.text_frame.margin_left = Inches(0.03)
    shape.text_frame.margin_right = Inches(0.03)
    shape.text_frame.margin_top = Inches(0.02)
    shape.text_frame.margin_bottom = Inches(0.02)
    shape.text_frame.word_wrap = True
    return shape


def clean_inline_text(text: str, *, context: str) -> str:
    if "\n" in text or "\r" in text:
        raise ValueError(f"manual newline is not allowed in {context}")
    cleaned = " ".join(text.split())
    if not cleaned:
        raise ValueError(f"empty text is not allowed in {context}")
    return cleaned


def clean_formula_text(text: str) -> str:
    if "\n" in text or "\r" in text:
        raise ValueError("manual newline is not allowed in formula runs")
    cleaned = re.sub(r"[ \t]+", " ", text)
    if not cleaned.strip():
        raise ValueError("empty formula run is not allowed")
    return cleaned


def add_plain(slide, text: str, left: float, top: float, width: float, height: float, size: float, *,
              name: str, bold: bool = False, color: RGBColor = BLACK, align=PP_ALIGN.LEFT) -> None:
    shape = add_textbox(slide, left, top, width, height, name=name)
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(0)
    p.line_spacing = 0.95
    run = p.add_run()
    run.text = clean_inline_text(text, context=name)
    set_run(run, size, bold=bold, color=color)


def set_baseline(run, value: int) -> None:
    run._r.get_or_add_rPr().set("baseline", str(value))  # noqa: SLF001


def add_math_run(paragraph, piece: str | tuple[str, str], size: float) -> None:
    if isinstance(piece, tuple):
        text, mode = piece
    else:
        text, mode = piece, ""
    text = clean_formula_text(text)
    run = paragraph.add_run()
    run.text = text
    run_size = size * 0.72 if mode in {"sub", "sup"} else size
    set_run(run, run_size, color=BLACK)
    if mode == "sub":
        set_baseline(run, -25000)
    elif mode == "sup":
        set_baseline(run, 30000)


def add_bullets(slide, lines: list[str], left: float, top: float, width: float, height: float, sizes: dict[str, float], *,
                name: str, space_after: float = 1.2) -> None:
    emphasis_terms = (
        "安全",
        "时间最优",
        "运行时",
        "CBF",
        "LQR",
        "执行器",
        "多模态",
        "最小间距",
        "覆盖率",
        "平滑",
        "视觉惯性",
        "事件相机",
        "体素",
        "通信",
        "多智能体",
        "优先级目标",
        "deterministic",
    )

    def add_emphasis_runs(paragraph, text: str, size: float, *, base_bold: bool) -> None:
        cursor = 0
        while cursor < len(text):
            hit = None
            hit_pos = len(text)
            for term in emphasis_terms:
                pos = text.find(term, cursor)
                if pos != -1 and pos < hit_pos:
                    hit = term
                    hit_pos = pos
            if hit is None:
                run = paragraph.add_run()
                run.text = text[cursor:]
                set_run(run, size, bold=base_bold, color=BLACK)
                break
            if hit_pos > cursor:
                run = paragraph.add_run()
                run.text = text[cursor:hit_pos]
                set_run(run, size, bold=base_bold, color=BLACK)
            run = paragraph.add_run()
            run.text = hit
            set_run(run, size, bold=True, color=RED)
            cursor = hit_pos + len(hit)

    shape = add_textbox(slide, left, top, width, height, name=name)
    tf = shape.text_frame
    tf.clear()
    if not lines:
        raise ValueError(f"empty bullet block is not allowed in {name}")
    for idx, raw in enumerate(lines):
        line = clean_inline_text(raw, context=name)
        marker = line.lstrip()[0]
        if marker == "●":
            size = sizes["main"]
            indent = Inches(0.28)
            first = Inches(-0.18)
        elif marker == "•":
            size = sizes["secondary"]
            indent = Inches(0.43)
            first = Inches(-0.16)
        else:
            size = sizes["tertiary"]
            indent = Inches(0.60)
            first = Inches(-0.14)
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = 0.93
        p.left_margin = indent
        p.first_line_indent = first
        bold = marker == "●" and len(line) < 34
        add_emphasis_runs(p, line, size, base_bold=bold)


def strip_empty_text_bodies(prs: Presentation) -> int:
    removed = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            element = shape._element  # noqa: SLF001
            tx_body = element.find(qn("p:txBody"))
            if tx_body is None:
                continue
            texts = [node.text or "" for node in tx_body.findall(".//" + qn("a:t"))]
            if not any(text.strip() for text in texts):
                element.remove(tx_body)
                removed += 1
    return removed


def add_rect(slide, left: float, top: float, width: float, height: float, color: RGBColor, *, line: RGBColor | None = None):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    return shape


def add_rule(slide, left: float, top: float, width: float, accent: RGBColor) -> None:
    shape = add_rect(slide, left, top, width, 0.018, accent)
    shape.name = "RULE"


def add_header(slide, w: float, h: float, accent: RGBColor, part: str, title: str) -> None:
    # Some public templates keep sample lab names, footers, or date fields on the
    # master. Cover the master canvas before adding our own report layout.
    add_rect(slide, 0.0, 0.0, w, h, WHITE)
    add_rect(slide, 0.0, 0.0, w, 0.48, LIGHT_GRAY)
    add_rect(slide, 0.0, 0.48, w, 0.04, accent)
    add_plain(slide, part, 0.46, 0.12, 1.30, 0.24, 11.0, name="HEADER_PART", bold=True, color=accent)
    add_plain(slide, title, 1.55, 0.12, max(2.0, w - 2.1), 0.24, 11.0, name="HEADER_TITLE", color=GRAY)


def add_section(slide, w: float, y: float, number: str, title: str, accent: RGBColor) -> None:
    add_plain(slide, number, 0.46, y, 0.78, 0.36, 18.5, name="SECTION_NO", bold=True, color=accent)
    add_plain(slide, title, 1.24, y + 0.03, max(2.5, w - 1.74), 0.34, 18.5, name="SECTION_TITLE", bold=True, color=BLACK)


def fit_picture(slide, path: Path, left: float, top: float, width: float, height: float, *, name: str):
    with Image.open(path) as image:
        ratio = image.width / image.height
    box_ratio = width / height
    if ratio > box_ratio:
        draw_w = width
        draw_h = width / ratio
        x = left
        y = top + (height - draw_h) / 2
    else:
        draw_h = height
        draw_w = height * ratio
        x = left + (width - draw_w) / 2
        y = top
    pic = slide.shapes.add_picture(str(path), Inches(x), Inches(y), width=Inches(draw_w), height=Inches(draw_h))
    pic.name = name
    return pic


def draw_assets(out_dir: Path, accent: RGBColor) -> dict[str, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    red_ch, green_ch, blue_ch = rgb_tuple(accent)
    accent_rgb = (red_ch, green_ch, blue_ch)
    muted = (min(255, red_ch + 62), min(255, green_ch + 62), min(255, blue_ch + 62))
    paths: dict[str, Path] = {}

    wide = Image.new("RGB", (1100, 430), (250, 252, 255))
    d = ImageDraw.Draw(wide)
    d.rectangle((55, 52, 1045, 364), outline=(190, 198, 210), width=3)
    for x in range(120, 1000, 125):
        d.line((x, 65, x, 352), fill=(226, 230, 236), width=1)
    for y in range(100, 340, 58):
        d.line((70, y, 1030, y), fill=(226, 230, 236), width=1)
    pts = [(90, 330), (210, 270), (342, 285), (480, 195), (620, 170), (760, 116), (918, 126), (1010, 84)]
    d.line(pts, fill=accent_rgb, width=7, joint="curve")
    for x, y in pts[1:-1]:
        d.ellipse((x - 10, y - 10, x + 10, y + 10), fill=muted, outline=accent_rgb, width=2)
    for rect in [(270, 92, 370, 170), (694, 232, 840, 314)]:
        d.rounded_rectangle(rect, radius=12, fill=(255, 228, 230), outline=(205, 32, 44), width=2)
    d.text((75, 24), "trajectory candidates and safe primitive selection", fill=(45, 50, 60))
    path = out_dir / "wide-trajectory.png"
    wide.save(path)
    paths["wide"] = path

    square = Image.new("RGB", (760, 560), (250, 252, 255))
    d = ImageDraw.Draw(square)
    boxes = [
        (65, 80, 230, 165, "primitive library"),
        (300, 80, 465, 165, "collision check"),
        (535, 80, 700, 165, "cost select"),
        (185, 300, 350, 385, "trajectory"),
        (420, 300, 585, 385, "controller"),
    ]
    for x0, y0, x1, y1, text in boxes:
        d.rounded_rectangle((x0, y0, x1, y1), radius=15, fill=(238, 246, 255), outline=accent_rgb, width=3)
        d.text((x0 + 18, y0 + 30), text, fill=(34, 40, 48))
    for start, end in [((230, 122), (300, 122)), ((465, 122), (535, 122)), ((617, 165), (502, 300)), ((382, 122), (267, 300)), ((350, 342), (420, 342))]:
        d.line((*start, *end), fill=accent_rgb, width=4)
    d.text((60, 25), "method pipeline and closed-loop execution", fill=(45, 50, 60))
    path = out_dir / "square-pipeline.png"
    square.save(path)
    paths["square"] = path

    tall = Image.new("RGB", (530, 760), (250, 252, 255))
    d = ImageDraw.Draw(tall)
    d.rounded_rectangle((70, 55, 460, 700), radius=28, outline=(190, 198, 210), width=3, fill=(255, 255, 255))
    y = 120
    for idx, label in enumerate(["RRT* path", "LQR tracking", "safety filter", "safe input"]):
        color = accent_rgb if idx < 3 else (205, 32, 44)
        d.ellipse((90, y - 14, 118, y + 14), fill=color)
        d.line((104, y + 14, 104, y + 100), fill=(190, 198, 210), width=3)
        d.rounded_rectangle((150, y - 32, 420, y + 32), radius=12, fill=(242, 246, 250), outline=color, width=2)
        d.text((170, y - 10), label, fill=(32, 38, 46))
        y += 140
    d.text((80, 24), "runtime safety assurance flow", fill=(45, 50, 60))
    path = out_dir / "tall-runtime.png"
    tall.save(path)
    paths["tall"] = path
    return paths


def add_table(slide, rows: list[list[str]], left: float, top: float, width: float, height: float, size: float, accent: RGBColor) -> None:
    graphic = slide.shapes.add_table(len(rows), len(rows[0]), Inches(left), Inches(top), Inches(width), Inches(height))
    graphic.name = "TABLE_RESULTS"
    tbl = graphic.table
    for ri, row in enumerate(rows):
        for ci, value in enumerate(row):
            cell = tbl.cell(ri, ci)
            value = clean_inline_text(value, context="table cell")
            cell.text = value
            cell.margin_left = Inches(0.02)
            cell.margin_right = Inches(0.02)
            cell.margin_top = Inches(0.01)
            cell.margin_bottom = Inches(0.01)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT_GRAY if ri == 0 else WHITE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.space_after = Pt(0)
            for run in p.runs:
                value_is_key = value in {"低", "运行时输入滤波", "时间最优候选", "实机实验"}
                set_run(run, size, bold=ri == 0 or ci == 0 or value_is_key, color=RED if value_is_key else BLACK)


FORMULA_SPECS: dict[str, list[list[str | tuple[str, str]]]] = {
    "primitive-planner-2025": [
        ["p(t) = p", ("0", "sub"), " + v", ("0", "sub"), "t + 1/2 a t", ("2", "sup")],
        ["J(π) = T(π) + λ", ("c", "sub"), " C(π) + λ", ("s", "sub"), " S(π)"],
        ["π", ("*", "sup"), " = arg min", ("π∈Pₛ", "sub"), " J(π)"],
    ],
    "safety-assurance-2025": [
        ["u", ("s", "sub"), " = arg min", ("u", "sub"), " ||u - u", ("n", "sub"), "||", ("2", "sup")],
        ["h(x) ≥ 0,    ḣ(x) + α h(x) ≥ 0"],
        ["ẋ = f(x) + g(x)u", ("s", "sub")],
    ],
    "quad-lcd-2025": [
        ["u = clip(π", ("θ", "sub"), "(x,r), u", ("min", "sub"), ", u", ("max", "sub"), ")"],
        ["L = L", ("track", "sub"), " + λ", ("u", "sub"), "L", ("sat", "sub"), " + λ", ("s", "sub"), "L", ("smooth", "sub")],
        ["e", ("t", "sub"), " = ||p", ("t", "sub"), " - p", ("ref", "sub"), "||", ("2", "sub"), ",    Δu", ("t", "sub"), " = ||u", ("t", "sub"), " - u", ("t-1", "sub"), "||", ("2", "sub")],
    ],
    "mc-swarm-2025": [
        ["m", ("i", "sub"), " = arg max", ("m", "sub"), " q", ("m", "sub"), "(x", ("i", "sub"), ", X", ("N", "sub"), ", g)"],
        ["d", ("ij", "sub"), " = ||p", ("i", "sub"), " - p", ("j", "sub"), "||", ("2", "sub"), " ≥ d", ("min", "sub")],
        ["u", ("i", "sub"), " = u", ("i", "sub"), ("m", "sup"), " + k", ("s", "sub"), " ∇h", ("i", "sub"), "(x)"],
    ],
    "smooth-coverage-2025": [
        ["min ", "Σ", ("i", "sub"), " (L", ("i", "sub"), " + λ", ("k", "sub"), "K", ("i", "sub"), " + λ", ("b", "sub"), "B", ("i", "sub"), ")"],
        ["s.t.  ∪", ("i", "sub"), " C", ("i", "sub"), " = Ω,    C", ("i", "sub"), " ∩ C", ("j", "sub"), " = ∅"],
        ["||p", ("i", "sub"), "(t) - p", ("j", "sub"), "(t)||", ("2", "sub"), " ≥ d", ("min", "sub")],
    ],
    "downfacing-vio-2025": [
        ["T", ("k+1", "sub"), " = T", ("k", "sub"), " exp(ξ", ("k", "sub"), " Δt)"],
        ["r", ("i", "sub"), " = z", ("i", "sub"), " - π(T", ("cw", "sub"), " P", ("i", "sub"), ")"],
        ["E = Σ", ("i", "sub"), " ||r", ("i", "sub"), "||", ("2", "sup"), " + λ", ("b", "sub"), " ||b||", ("2", "sup")],
    ],
    "voxel-esvio-2025": [
        ["p", ("v", "sub"), " = TopK(p | v, s(p))"],
        ["E", ("map", "sub"), " = Σ", ("v", "sub"), " w", ("v", "sub"), " ||e", ("v", "sub"), "||", ("2", "sup")],
        ["x", ("*", "sup"), " = arg min", ("x", "sub"), " (E", ("imu", "sub"), " + E", ("event", "sub"), " + E", ("stereo", "sub"), ")"],
    ],
    "persistent-monitoring-2025": [
        ["a", ("i", "sub"), ("*", "sup"), " = arg max", ("a", "sub"), " Q", ("i", "sub"), "(o", ("i", "sub"), ", m", ("i", "sub"), ", a)"],
        ["R", ("ij", "sub"), "(t) ≥ R", ("min", "sub"), ",    ||p", ("i", "sub"), " - p", ("j", "sub"), "|| ≤ d", ("comm", "sub")],
        ["J = Σ", ("l", "sub"), " γ", ("l", "sub"), " q", ("l", "sub"), "(t) - λ", ("c", "sub"), " C", ("i", "sub"), "(t)"],
    ],
    "terrain-aware-uav-2025": [
        ["z", ("k", "sub"), " = h(x", ("k", "sub"), ", y", ("k", "sub"), ") + d", ("safe", "sub")],
        ["v", ("k", "sub"), " = Π(K, R", ("k", "sub"), ", t", ("k", "sub"), ", X)"],
        ["J = Σ", ("k", "sub"), " (1 - cov(v", ("k", "sub"), ")) + λ", ("h", "sub"), "|Δz", ("k", "sub"), "| + λ", ("ψ", "sub"), "|Δψ", ("k", "sub"), "|"],
    ],
    "zoom-ptz-coverage-2025": [
        ["GSD", ("k", "sub"), " = H", ("k", "sub"), " p / f", ("k", "sub")],
        ["f", ("k", "sub"), ("*", "sup"), " = clip(H", ("k", "sub"), " p / GSD", ("0", "sub"), ", f", ("min", "sub"), ", f", ("max", "sub"), ")"],
        ["J = T + λ", ("L", "sub"), " L + λ", ("z", "sub"), " Σ", ("k", "sub"), " |Δz", ("k", "sub"), "|"],
    ],
    "diff-physics-flight-2025": [
        ["a", ("t", "sub"), " = π", ("θ", "sub"), "(D", ("t", "sub"), ", v", ("t", "sub"), ")"],
        ["v", ("t+1", "sub"), " = v", ("t", "sub"), " + a", ("t", "sub"), " Δt"],
        ["L = Σ", ("t", "sub"), " ℓ", ("obs", "sub"), "(p", ("t", "sub"), ") + λ", ("u", "sub"), " ||a", ("t", "sub"), "||", ("2", "sup")],
    ],
    "shape-adaptive-quad-2025": [
        ["x", ("k+1", "sub"), " = f(x", ("k", "sub"), ", u", ("k", "sub"), ", η", ("k", "sub"), ")"],
        ["J = Σ", ("k", "sub"), " ||p", ("k", "sub"), " - p", ("ref", "sub"), "||", ("2", "sup"), " + λ", ("η", "sub"), " ||Δη", ("k", "sub"), "||", ("2", "sup")],
        ["η", ("*", "sup"), " = arg min", ("η", "sub"), " J", ("path", "sub"), "(x, η)"],
    ],
}


def formula_rows_for(paper: dict[str, Any]) -> list[list[str | tuple[str, str]]]:
    return FORMULA_SPECS.get(paper["id"], [[row] for row in paper["formula_rows"]])


def add_formula_rows(slide, rows: list[list[str | tuple[str, str]]], left: float, top: float, width: float, row_h: float, size: float, accent: RGBColor) -> None:
    for idx, row in enumerate(rows):
        y = top + idx * row_h
        add_plain(slide, f"({idx + 1})", left, y + 0.04, 0.50, 0.32, size - 2.2, name=f"MATH_LABEL_{idx}", color=GRAY, align=PP_ALIGN.RIGHT)
        shape = add_textbox(slide, left + 0.68, y, width - 0.68, 0.46, name=f"MATH_BODY_{idx}")
        shape.text_frame.word_wrap = False
        shape.text_frame.margin_left = Inches(0)
        shape.text_frame.margin_right = Inches(0)
        shape.text_frame.margin_top = Inches(0)
        shape.text_frame.margin_bottom = Inches(0)
        p = shape.text_frame.paragraphs[0]
        p.space_after = Pt(0)
        p.line_spacing = 1.0
        for piece in row:
            add_math_run(p, piece, size)
    add_rule(slide, left + 0.62, top + len(rows) * row_h + 0.08, width - 0.62, accent)


def metric_row(slide, items: list[list[str]], left: float, top: float, width: float, accent: RGBColor, size: float) -> None:
    gap = 0.12
    box_w = (width - gap * (len(items) - 1)) / len(items)
    for idx, item in enumerate(items):
        x = left + idx * (box_w + gap)
        box = add_rect(slide, x, top, box_w, 0.64, RGBColor(248, 250, 253), line=MID_GRAY)
        box.name = f"METRIC_BOX_{idx}"
        add_plain(slide, item[0], x + 0.08, top + 0.08, box_w - 0.16, 0.22, size, name=f"METRIC_VALUE_{idx}", bold=True, color=accent, align=PP_ALIGN.CENTER)
        add_plain(slide, item[1], x + 0.08, top + 0.36, box_w - 0.16, 0.18, 9.8, name=f"METRIC_LABEL_{idx}", color=GRAY, align=PP_ALIGN.CENTER)


def blank_layout(prs: Presentation):
    for layout in prs.slide_layouts:
        if "blank" in layout.name.lower() or "空白" in layout.name:
            return layout
    return prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)]


def scrub_placeholders(slide) -> None:
    banned = ("单击此处", "Click to", "‹#›")
    for shape in list(slide.shapes):
        if not getattr(shape, "is_placeholder", False):
            continue
        text = getattr(shape, "text", "")
        if not text or any(token.lower() in text.lower() for token in banned):
            slide.shapes._spTree.remove(shape._element)  # noqa: SLF001


def add_blank_slide(prs: Presentation):
    slide = prs.slides.add_slide(blank_layout(prs))
    scrub_placeholders(slide)
    return slide


def cover_slide(prs: Presentation, paper: dict[str, Any], accent: RGBColor, template_id: str, sizes: dict[str, float]) -> None:
    w = prs.slide_width / EMU_PER_INCH
    h = prs.slide_height / EMU_PER_INCH
    slide = add_blank_slide(prs)
    add_rect(slide, 0, 0, w, h, RGBColor(246, 249, 252))
    band_h = h * 0.40
    band_y = h * 0.27
    add_rect(slide, 0, band_y, w, band_h, accent)
    title_color = WHITE if brightness(accent) < 0.72 else BLACK
    add_plain(slide, paper["title"], w * 0.08, band_y + band_h * 0.24, w * 0.84, band_h * 0.34, min(28, sizes["main"] + 8.8), name="COVER_TITLE", bold=True, color=title_color, align=PP_ALIGN.CENTER)
    add_plain(slide, f"{paper['topic']} / {paper['year']}", w * 0.10, band_y + band_h * 0.62, w * 0.80, 0.34, sizes["secondary"], name="COVER_SUBTITLE", bold=True, color=title_color, align=PP_ALIGN.CENTER)
    add_plain(slide, "中文学术论文汇报", w * 0.32, h - 0.72, w * 0.36, 0.28, 13.0, name="COVER_FOOTER", color=accent, align=PP_ALIGN.CENTER)


def content_position(prs: Presentation, paper: dict[str, Any], accent: RGBColor, sizes: dict[str, float]) -> None:
    w = prs.slide_width / EMU_PER_INCH
    h = prs.slide_height / EMU_PER_INCH
    slide = add_blank_slide(prs)
    add_header(slide, w, h, accent, "Part. 01", "论文定位")
    add_section(slide, w, 0.78, "1.1", "研究问题与核心思路", accent)
    body_top = 1.35
    if w < 11.0:
        add_bullets(slide, ["● " + paper["problem"], "● " + paper["main_claim"]] + paper["bullets"][:2], 0.55, body_top, w - 1.1, 2.58, sizes, name="BODY_POSITION", space_after=0.78)
        add_table(slide, paper["comparison_rows_narrow"], 0.68, 3.92, w - 1.36, 1.08, 10.7, accent)
        add_bullets(slide, paper["position_notes"], 0.72, 5.18, w - 1.44, 0.64, sizes, name="BODY_POSITION_NOTE", space_after=0.0)
        metric_row(slide, paper["metrics"], 0.65, h - 1.16, w - 1.30, accent, 15.2)
    else:
        add_bullets(slide, ["● " + paper["problem"], "● " + paper["main_claim"]] + paper["bullets"][:2], 0.70, body_top, w * 0.54, 3.12, sizes, name="BODY_POSITION", space_after=0.82)
        metric_row(slide, paper["metrics"], w * 0.61, body_top + 0.12, w * 0.33, accent, 16.0)
        add_bullets(slide, paper["position_notes"], w * 0.61, body_top + 1.80, w * 0.33, 1.04, sizes, name="BODY_POSITION_SIDE", space_after=0.25)
        add_rule(slide, 0.76, 4.48, w - 1.52, accent)
        add_table(slide, paper["comparison_rows_wide"], 0.92, 4.72, w - 1.84, 1.28, 10.8, accent)
        add_bullets(slide, paper["bullets"][2:3], 0.94, 6.20, w - 1.88, 0.44, sizes, name="BODY_POSITION_BOTTOM", space_after=0.0)


def method_overview(prs: Presentation, paper: dict[str, Any], assets: dict[str, Path], accent: RGBColor, sizes: dict[str, float]) -> None:
    w = prs.slide_width / EMU_PER_INCH
    h = prs.slide_height / EMU_PER_INCH
    slide = add_blank_slide(prs)
    add_header(slide, w, h, accent, "Part. 02", "研究方法")
    add_section(slide, w, 0.78, "2.1", "方法流程", accent)
    if w < 11.0:
        fit_picture(slide, assets["square"], 0.72, 1.34, 4.22, 3.12, name="FIG_PIPELINE")
        add_bullets(slide, paper["bullets"][:2], 5.12, 1.34, w - 5.66, 1.82, sizes, name="BODY_METHOD_RIGHT", space_after=0.75)
        add_table(
            slide,
            [["阶段", "输出"], ["候选库", "基元"], ["碰撞检测", "安全集"], ["代价选择", "轨迹"]],
            5.18,
            3.44,
            w - 5.80,
            0.94,
            10.4,
            accent,
        )
        add_rule(slide, 0.68, 4.72, w - 1.36, accent)
        add_bullets(slide, paper["method_notes_narrow"], 0.72, 4.92, w - 1.44, 0.78, sizes, name="BODY_METHOD_BOTTOM", space_after=0.35)
        metric_row(slide, [["流程", "候选生成"], ["约束", "安全筛选"], ["输出", "连续轨迹"]], 0.82, 6.04, w - 1.64, accent, 14.8)
    else:
        fit_picture(slide, assets["square"], 0.76, 1.52, w * 0.47, 4.15, name="FIG_PIPELINE")
        add_bullets(slide, paper["bullets"] + paper["method_notes_wide"], w * 0.54, 1.52, w * 0.38, 3.98, sizes, name="BODY_METHOD_RIGHT", space_after=0.9)
        add_table(
            slide,
            paper["method_rows"],
            0.84,
            4.68,
            w - 1.68,
            0.86,
            10.8,
            accent,
        )
        add_rule(slide, 0.76, 5.86, w - 1.52, accent)
        add_bullets(slide, paper["method_bottom_notes"], 0.82, 6.04, w - 1.64, 0.52, sizes, name="BODY_METHOD_NOTE", space_after=0.0)


def formula_slide(prs: Presentation, paper: dict[str, Any], accent: RGBColor, sizes: dict[str, float]) -> None:
    w = prs.slide_width / EMU_PER_INCH
    h = prs.slide_height / EMU_PER_INCH
    slide = add_blank_slide(prs)
    add_header(slide, w, h, accent, "Part. 02", "研究方法")
    add_section(slide, w, 0.78, "2.2", "约束与目标函数", accent)
    formula_rows = formula_rows_for(paper)
    if w < 11.0:
        add_formula_rows(slide, formula_rows, 0.74, 1.38, w - 1.45, 0.58, sizes["formula"], accent)
        rows = [["符号", "含义"]] + paper["terms"][:4]
        add_table(slide, rows, 0.70, 3.35, w - 1.40, 1.92, sizes["table"], accent)
        add_bullets(slide, paper["formula_notes_narrow"] + paper["formula_notes_wide"][1:2], 0.72, 5.48, w - 1.44, 0.88, sizes, name="BODY_FORMULA_NOTE", space_after=0.15)
    else:
        add_formula_rows(slide, formula_rows, 0.86, 1.44, w * 0.54, 0.62, sizes["formula"], accent)
        rows = [["符号", "含义"]] + paper["terms"][:4]
        add_table(slide, rows, w * 0.61, 1.42, w * 0.33, 2.50, sizes["table"], accent)
        add_rule(slide, 0.86, 4.22, w - 1.72, accent)
        add_bullets(slide, paper["formula_notes_wide"], 0.86, 4.46, w * 0.52, 1.10, sizes, name="BODY_FORMULA_EXPLAIN", space_after=0.48)
        add_table(slide, [["对象", "进入方式"], ["目标", "代价函数"], ["约束", "可行集合"], ["输出", "控制/轨迹"]], w * 0.61, 4.40, w * 0.33, 1.26, sizes["table"], accent)
        metric_row(slide, [["目标", "优化量"], ["约束", "安全集"], ["动力学", "状态传播"]], 0.98, h - 1.02, w - 1.96, accent, 15.4)


def evidence_slide(prs: Presentation, paper: dict[str, Any], assets: dict[str, Path], accent: RGBColor, sizes: dict[str, float]) -> None:
    w = prs.slide_width / EMU_PER_INCH
    h = prs.slide_height / EMU_PER_INCH
    slide = add_blank_slide(prs)
    add_header(slide, w, h, accent, "Part. 03", "实验与结果")
    add_section(slide, w, 0.78, "3.1", "宽图与表格排版", accent)
    fit_picture(slide, assets["wide"], 0.72, 1.34, w - 1.44, 2.50, name="FIG_TRAJECTORY")
    rows = [
        ["维度", "基线", "本文方法"],
        ["在线负担", "中等", paper["metrics"][0][1]],
        ["轨迹性质", "可行候选", paper["metrics"][1][1]],
        ["验证范围", "仿真", paper["metrics"][2][1]],
    ]
    if w < 11.0:
        add_table(slide, rows, 0.70, 4.10, w - 1.40, 1.40, sizes["table"], accent)
        add_bullets(slide, paper["evidence_notes_narrow"], 0.70, 5.74, w - 1.40, 0.72, sizes, name="BODY_EVIDENCE", space_after=0.0)
    else:
        add_table(slide, rows, 0.80, 4.16, w * 0.48, 1.36, sizes["table"], accent)
        add_bullets(slide, paper["evidence_notes_wide"], w * 0.56, 4.14, w * 0.36, 1.28, sizes, name="BODY_EVIDENCE", space_after=0.7)
        metric_row(slide, [["轨迹", "候选收敛"], ["对比", "性能差异"], ["结论", "可执行性"]], 0.94, h - 1.08, w - 1.88, accent, 15.4)


def narrow_figure_slide(prs: Presentation, paper: dict[str, Any], assets: dict[str, Path], accent: RGBColor, sizes: dict[str, float]) -> None:
    w = prs.slide_width / EMU_PER_INCH
    h = prs.slide_height / EMU_PER_INCH
    slide = add_blank_slide(prs)
    add_header(slide, w, h, accent, "Part. 03", "实验与结果")
    add_section(slide, w, 0.78, "3.2", "纵向流程与术语解释", accent)
    if w < 11.0:
        fit_picture(slide, assets["tall"], 0.70, 1.42, 3.35, 4.85, name="FIG_TALL_FLOW")
        add_bullets(slide, paper["tall_figure_notes_narrow"], 4.32, 1.48, w - 4.85, 3.48, sizes, name="BODY_NARROW", space_after=0.8)
        add_table(slide, [["变量", "含义"]] + paper["terms"][:2], 4.34, 5.18, w - 4.90, 0.92, 9.6, accent)
    else:
        fit_picture(slide, assets["tall"], 0.92, 1.42, w * 0.28, 4.80, name="FIG_TALL_FLOW")
        add_bullets(slide, paper["tall_figure_notes_wide"], w * 0.39, 1.52, w * 0.52, 2.82, sizes, name="BODY_NARROW", space_after=0.9)
        add_rule(slide, w * 0.39, 4.70, w * 0.52, accent)
        add_bullets(slide, paper["tall_figure_bottom_notes"], w * 0.39, 4.92, w * 0.52, 0.70, sizes, name="BODY_NARROW_NOTE", space_after=0.0)
        metric_row(slide, [["输入", "观测/状态"], ["约束", "安全/通信"], ["输出", "轨迹/位姿"]], w * 0.39, h - 1.08, w * 0.52, accent, 14.8)


def summary_slide(prs: Presentation, paper: dict[str, Any], accent: RGBColor, sizes: dict[str, float]) -> None:
    w = prs.slide_width / EMU_PER_INCH
    h = prs.slide_height / EMU_PER_INCH
    slide = add_blank_slide(prs)
    add_header(slide, w, h, accent, "Part. 04", "总结")
    add_section(slide, w, 0.78, "4.1", "报告结论", accent)
    summary_lines = [
        "● " + paper["main_claim"],
        "• " + paper["summary_method"],
        "• " + paper["summary_result"],
        "• 方法边界主要来自约束建模、候选覆盖范围和真实执行误差。",
    ]
    if w < 11.0:
        summary_lines.extend(paper["summary_extra"])
    add_bullets(slide, summary_lines, 0.76, 1.48, w - 1.52, 3.18 if w < 11.0 else 1.86, sizes, name="BODY_SUMMARY", space_after=0.48)
    if w < 11.0:
        add_table(slide, [["对象", "约束"], ["轨迹", "连续性"], ["控制", "输入边界"], ["实验", "可执行性"]], 0.88, 3.64, w - 1.76, 1.05, 10.7, accent)
        metric_row(slide, [["问题", "规划约束"], ["方法", "在线选择"], ["结果", "飞行验证"]], 0.88, h - 1.84, w - 1.76, accent, 15.2)
    else:
        add_table(
            slide,
            paper["summary_rows"],
            0.88,
            3.42,
            w - 1.76,
            1.34,
            10.8,
            accent,
        )
        add_bullets(slide, paper["summary_bridge"], 0.96, 5.02, w - 1.92, 0.46, sizes, name="BODY_SUMMARY_BRIDGE", space_after=0.0)
        metric_row(slide, [["问题", "任务约束"], ["方法", "闭环求解"], ["结果", "实验证据"]], 0.88, h - 1.74, w - 1.76, accent, 15.8)


def thanks_slide(prs: Presentation, accent: RGBColor, template_id: str) -> None:
    w = prs.slide_width / EMU_PER_INCH
    h = prs.slide_height / EMU_PER_INCH
    slide = add_blank_slide(prs)
    add_rect(slide, 0, 0, w, h, RGBColor(246, 249, 252))
    band_h = h * 0.40
    band_y = h * 0.27
    add_rect(slide, 0, band_y, w, band_h, accent)
    title_color = WHITE if brightness(accent) < 0.72 else BLACK
    add_plain(slide, "谢谢！", w * 0.10, band_y + band_h * 0.33, w * 0.80, 0.64, 42.0, name="THANKS_TITLE", bold=True, color=title_color, align=PP_ALIGN.CENTER)
    add_plain(slide, "学术论文汇报", w * 0.36, h - 0.72, w * 0.28, 0.24, 12.0, name="THANKS_FOOTER", color=accent, align=PP_ALIGN.CENTER)


def build_deck(template: dict[str, Any], paper: dict[str, Any], out_dir: Path, *, template_root: Path | None) -> tuple[Path, str]:
    template_path = expand_path(template["path"], template_root=template_root)
    if not template_path.exists():
        raise FileNotFoundError(template_path)
    template_id = template["id"]
    profile = template.get("audit_profile", "compact")
    sizes = PROFILE_SIZE[profile]
    accent = theme_color(template_path, template.get("theme"))

    prs = Presentation(str(template_path))
    clear_slides(prs)
    assets = draw_assets(out_dir / "assets" / f"{template_id}-{paper['id']}", accent)
    cover_slide(prs, paper, accent, template_id, sizes)
    content_position(prs, paper, accent, sizes)
    method_overview(prs, paper, assets, accent, sizes)
    formula_slide(prs, paper, accent, sizes)
    evidence_slide(prs, paper, assets, accent, sizes)
    narrow_figure_slide(prs, paper, assets, accent, sizes)
    summary_slide(prs, paper, accent, sizes)
    thanks_slide(prs, accent, template_id)
    strip_empty_text_bodies(prs)

    deck_dir = out_dir / f"{template_id}-{paper['id']}"
    deck_dir.mkdir(parents=True, exist_ok=True)
    deck_path = deck_dir / "input.pptx"
    prs.save(deck_path)
    return deck_path, profile


def run(cmd: list[str], *, timeout: int) -> subprocess.CompletedProcess[str]:
    print("+", " ".join(cmd), flush=True)
    return subprocess.run(cmd, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=timeout)


def render_case(pptx: Path, case_dir: Path, *, timeout: int) -> int:
    office = matrix_runner.soffice_path()
    if not office:
        print("LibreOffice not found; smoke render skipped.")
        return 2
    try:
        pdf = matrix_runner.export_pdf(pptx, case_dir, office, timeout=timeout)
    except Exception as exc:
        print(f"LibreOffice export failed: {exc}")
        return 1
    py = matrix_runner.script_python()
    png_dir = case_dir / "png"
    preview = case_dir / "preview-grid.png"
    proc = run([py, str(RENDER), str(pdf), "--out-dir", str(png_dir), "--preview-grid", str(preview)], timeout=120)
    print(proc.stdout)
    if proc.returncode != 0:
        return proc.returncode
    proc = run([
        py,
        str(SCAN),
        str(png_dir),
        "--fail-on-warning",
        "--ignore-edge-slides",
        "--blank-warn",
        "0.76",
        "--min-band-fraction",
        "0.14",
    ], timeout=60)
    print(proc.stdout)
    return proc.returncode


def load_templates(path: Path, *, template_root: Path | None) -> list[dict[str, Any]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    templates = data.get("templates", [])
    resolved = []
    for template in templates:
        try:
            template_path = expand_path(template["path"], template_root=template_root)
        except Exception:
            continue
        if template_path.exists():
            resolved.append(template)
        else:
            print(f"skip missing template: {template.get('id', template.get('path'))} -> {template_path}")
    return resolved


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate small paper-report decks across local PPTX templates.")
    parser.add_argument("--papers", type=Path, default=DEFAULT_PAPERS)
    parser.add_argument("--templates", type=Path, default=DEFAULT_TEMPLATES)
    parser.add_argument("--template-root", type=Path, help="Root containing PPTAgent-style template subfolders.")
    parser.add_argument("--out-dir", type=Path, default=ROOT / "out" / "template-smoke")
    parser.add_argument("--case-limit", type=int, default=0, help="Limit generated cases for quick debugging.")
    parser.add_argument("--skip-render", action="store_true")
    parser.add_argument("--keep-going", action="store_true")
    parser.add_argument("--export-timeout", type=int, default=90)
    args = parser.parse_args()

    papers = json.loads(args.papers.read_text(encoding="utf-8")).get("papers", [])
    templates = load_templates(args.templates, template_root=args.template_root)
    if not papers:
        print("no stress papers configured", file=sys.stderr)
        return 2
    if not templates:
        print("no readable templates configured", file=sys.stderr)
        return 2

    py = matrix_runner.script_python()
    failures: list[str] = []
    count = 0
    args.out_dir.mkdir(parents=True, exist_ok=True)
    for template in templates:
        for paper in papers:
            if args.case_limit and count >= args.case_limit:
                break
            case_id = f"{template['id']}-{paper['id']}"
            print(f"\n== smoke {case_id} ==", flush=True)
            try:
                pptx, profile = build_deck(template, paper, args.out_dir, template_root=args.template_root)
            except Exception as exc:
                print(f"build failed: {exc}")
                failures.append(f"{case_id}: build failed")
                if not args.keep_going:
                    break
                continue
            proc = run([
                py,
                str(AUDIT),
                str(pptx),
                "--strict-body-hierarchy",
                "--profile",
                profile,
                "--fail-on-warning",
            ], timeout=60)
            print(proc.stdout)
            if proc.returncode != 0:
                failures.append(f"{case_id}: PPTX audit failed")
                if not args.keep_going:
                    break
                continue
            if not args.skip_render:
                rc = render_case(pptx, pptx.parent, timeout=args.export_timeout)
                if rc != 0:
                    failures.append(f"{case_id}: render/scan failed")
                    if not args.keep_going:
                        break
            count += 1
        if args.case_limit and count >= args.case_limit:
            break
        if failures and not args.keep_going:
            break

    if failures:
        print("\nTemplate smoke failed:")
        for failure in failures:
            print(" -", failure)
        return 1
    print(f"\nTemplate smoke passed ({count} case(s)).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
