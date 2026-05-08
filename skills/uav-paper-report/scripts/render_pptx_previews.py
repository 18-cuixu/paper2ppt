from __future__ import annotations

import argparse
from pathlib import Path

import fitz
from PIL import Image, ImageDraw


def render(pdf: Path, out_dir: Path, preview_grid: Path, *, scale: float = 2.0, cols: int = 4) -> int:
    out_dir.mkdir(parents=True, exist_ok=True)
    preview_grid.parent.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("slide-*.png"):
        old.unlink()

    doc = fitz.open(pdf)
    paths: list[Path] = []
    for index, page in enumerate(doc, 1):
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        path = out_dir / f"slide-{index:02d}.png"
        pix.save(path)
        paths.append(path)

    thumbs = []
    for path in paths:
        img = Image.open(path).convert("RGB")
        img.thumbnail((360, 203))
        canvas = Image.new("RGB", (380, 235), "white")
        canvas.paste(img, ((380 - img.width) // 2, 20))
        ImageDraw.Draw(canvas).text((12, 6), path.stem, fill=(0, 0, 0))
        thumbs.append(canvas)

    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * 380, max(rows, 1) * 235), "white")
    for i, thumb in enumerate(thumbs):
        sheet.paste(thumb, ((i % cols) * 380, (i // cols) * 235))
    sheet.save(preview_grid)
    return len(paths)


def main() -> int:
    parser = argparse.ArgumentParser(description="Render a PPT-exported PDF to slide PNGs and a preview grid.")
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--preview-grid", "--contact", dest="preview_grid", type=Path, required=True)
    parser.add_argument("--scale", type=float, default=2.0)
    args = parser.parse_args()

    count = render(args.pdf, args.out_dir, args.preview_grid, scale=args.scale)
    print("pages", count)
    print("preview_grid", args.preview_grid)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
