from __future__ import annotations

from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parent
PAGES = ROOT / "multi-quad-cbf-test" / "pdf-pages"
OUT = ROOT / "multi-quad-cbf-test" / "clean-crops"


CROPS = {
    "fig1_real_gap.png": ("page-01.png", (650, 145, 980, 350)),
    "fig2_system_convention.png": ("page-03.png", (105, 100, 560, 340)),
    "fig3_polytope.png": ("page-04.png", (95, 105, 555, 355)),
    "fig4_sim_envs.png": ("page-05.png", (610, 95, 1115, 390)),
    "fig5_block_diagram.png": ("page-05.png", (610, 445, 1120, 620)),
    "fig6_tracking.png": ("page-06.png", (180, 90, 1075, 570)),
    "table2_success.png": ("page-06.png", (105, 675, 590, 820)),
}


def trim_white(im: Image.Image, pad: int = 8, threshold: int = 250) -> Image.Image:
    gray = im.convert("L")
    mask = gray.point(lambda p: 0 if p > threshold else 255)
    bbox = mask.getbbox()
    if not bbox:
        return im
    left, top, right, bottom = bbox
    left = max(0, left - pad)
    top = max(0, top - pad)
    right = min(im.width, right + pad)
    bottom = min(im.height, bottom + pad)
    return im.crop((left, top, right, bottom))


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for name, (page, box) in CROPS.items():
        im = Image.open(PAGES / page).convert("RGB")
        crop = im.crop(box)
        crop = trim_white(crop)
        crop.save(OUT / name)
        print(name, crop.size)


if __name__ == "__main__":
    main()
