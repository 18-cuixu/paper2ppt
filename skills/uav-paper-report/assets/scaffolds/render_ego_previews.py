from __future__ import annotations

from pathlib import Path

import fitz
from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parent
PDF = ROOT / "rendered-ego" / "uav-ego-planner-report.pdf"
OUT = ROOT / "rendered-ego-check" / "png"
PREVIEW_GRID = ROOT / "rendered-ego-check" / "preview-grid.png"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for old in OUT.glob("slide-*.png"):
        old.unlink()
    doc = fitz.open(PDF)
    paths = []
    for index, page in enumerate(doc, 1):
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        path = OUT / f"slide-{index:02d}.png"
        pix.save(path)
        paths.append(path)

    cols = 4
    thumbs = []
    for path in paths:
        img = Image.open(path).convert("RGB")
        img.thumbnail((360, 203))
        canvas = Image.new("RGB", (380, 235), "white")
        canvas.paste(img, ((380 - img.width) // 2, 20))
        ImageDraw.Draw(canvas).text((12, 6), path.stem, fill=(0, 0, 0))
        thumbs.append(canvas)

    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * 380, rows * 235), "white")
    for i, thumb in enumerate(thumbs):
        sheet.paste(thumb, ((i % cols) * 380, (i // cols) * 235))
    sheet.save(PREVIEW_GRID)
    print("pages", len(paths))
    print("preview_grid", PREVIEW_GRID)


if __name__ == "__main__":
    main()
