from __future__ import annotations

from pathlib import Path

import fitz
from PIL import Image, ImageChops, ImageDraw


ROOT = Path(__file__).resolve().parent
WORK = ROOT / "rl-privileged-test"
PDF = ROOT / "test-papers" / "rl_privileged_quadrotor.pdf"
PAGES = WORK / "paper-pages"
CROPS = WORK / "clean-crops"


def render_pages(doc: fitz.Document, zoom: float = 2.5) -> None:
    PAGES.mkdir(parents=True, exist_ok=True)
    for old in PAGES.glob("page-*.png"):
        old.unlink()
    for index, page in enumerate(doc, 1):
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        pix.save(PAGES / f"page-{index:02d}.png")


def trim_white(img: Image.Image, pad: int = 8) -> Image.Image:
    rgb = img.convert("RGB")
    bg = Image.new("RGB", rgb.size, "white")
    bbox = ImageChops.difference(rgb, bg).getbbox()
    if not bbox:
        return rgb
    left, top, right, bottom = bbox
    return rgb.crop((
        max(0, left - pad),
        max(0, top - pad),
        min(rgb.width, right + pad),
        min(rgb.height, bottom + pad),
    ))


def crop_pdf_rect(
    doc: fitz.Document,
    page_no: int,
    rect: tuple[float, float, float, float],
    name: str,
    *,
    zoom: float = 3.0,
    trim: bool = True,
    pad: int = 8,
) -> None:
    CROPS.mkdir(parents=True, exist_ok=True)
    page = doc[page_no - 1]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=fitz.Rect(*rect), alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    if trim:
        img = trim_white(img, pad=pad)
    img.save(CROPS / name)


def make_preview_grid(folder: Path, pattern: str, out: Path, *, cols: int = 4) -> None:
    paths = sorted(folder.glob(pattern))
    thumbs: list[Image.Image] = []
    for path in paths:
        img = Image.open(path).convert("RGB")
        img.thumbnail((320, 210))
        canvas = Image.new("RGB", (340, 245), "white")
        canvas.paste(img, ((340 - img.width) // 2, 28))
        ImageDraw.Draw(canvas).text((10, 8), path.name[:48], fill=(0, 0, 0))
        thumbs.append(canvas)
    if not thumbs:
        return
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * 340, rows * 245), "white")
    for i, thumb in enumerate(thumbs):
        sheet.paste(thumb, ((i % cols) * 340, (i // cols) * 245))
    out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out)


def main() -> None:
    doc = fitz.open(PDF)
    WORK.mkdir(parents=True, exist_ok=True)
    render_pages(doc)
    CROPS.mkdir(parents=True, exist_ok=True)
    for old in CROPS.glob("*.png"):
        old.unlink()

    # Hand-tuned crops from the 8-page arXiv paper. Full captions/body text are excluded.
    crop_pdf_rect(doc, 1, (300, 150, 575, 300), "fig1_system_overview.png", pad=4)
    crop_pdf_rect(doc, 2, (50, 58, 305, 185), "fig2_diff_dynamics.png", pad=6)
    crop_pdf_rect(doc, 3, (50, 50, 274, 203), "fig3_network_architecture.png", pad=4)
    crop_pdf_rect(doc, 4, (50, 48, 300, 190), "fig4_toa_map_paths.png", pad=6)
    crop_pdf_rect(doc, 5, (50, 220, 300, 455), "fig6_attitude_control.png", pad=6)
    crop_pdf_rect(doc, 6, (50, 45, 574, 205), "fig7_success_modes.png", pad=4)
    crop_pdf_rect(doc, 6, (50, 286, 574, 458), "fig8_traj_comparison.png", pad=6)
    crop_pdf_rect(doc, 7, (50, 52, 574, 295), "fig9_gravity_randomization.png", pad=4)
    crop_pdf_rect(doc, 7, (306, 420, 574, 630), "table3_hardware_trials.png", pad=6)
    crop_pdf_rect(doc, 8, (50, 50, 300, 255), "fig10_forest_navigation.png", pad=6)
    crop_pdf_rect(doc, 8, (50, 412, 300, 548), "fig11_night_flight.png", pad=4)

    make_preview_grid(PAGES, "page-*.png", WORK / "paper-pages-preview-grid.png")
    make_preview_grid(CROPS, "*.png", WORK / "clean-crops-preview-grid.png")
    print("pages", len(list(PAGES.glob("page-*.png"))))
    print("crops", len(list(CROPS.glob("*.png"))))
    print("crop_preview_grid", WORK / "clean-crops-preview-grid.png")


if __name__ == "__main__":
    main()
