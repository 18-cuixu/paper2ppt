from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageChops, ImageStat


def nonwhite_bbox(img: Image.Image, threshold: int = 248):
    gray = img.convert("L")
    mask = gray.point(lambda p: 255 if p < threshold else 0)
    return mask.getbbox()


def white_fraction(img: Image.Image, threshold: int = 248) -> float:
    gray = img.convert("L")
    hist = gray.histogram()
    white = sum(hist[threshold:])
    total = img.width * img.height
    return white / total if total else 1.0


def edge_density(img: Image.Image) -> float:
    gray = img.convert("L")
    shifted = ImageChops.offset(gray, 1, 0)
    diff = ImageChops.difference(gray, shifted)
    stat = ImageStat.Stat(diff)
    return stat.mean[0] / 255.0


def large_internal_whitespace(img: Image.Image, threshold: int = 248) -> list[str]:
    gray = img.convert("L")
    width, height = gray.size
    # Ignore the fixed header/footer margins and inspect the content body.
    top = int(height * 0.16)
    bottom = int(height * 0.92)
    if bottom <= top:
        return []

    row_blank: list[float] = []
    for y in range(top, bottom):
        row = gray.crop((0, y, width, y + 1))
        hist = row.histogram()
        row_blank.append(sum(hist[threshold:]) / width)

    warnings: list[str] = []
    run_start = None
    min_band = int(height * 0.105)
    for idx, blank in enumerate(row_blank + [0.0]):
        if blank > 0.985 and run_start is None:
            run_start = idx
        elif (blank <= 0.985 or idx == len(row_blank)) and run_start is not None:
            run_len = idx - run_start
            y0 = top + run_start
            y1 = top + idx
            if run_len >= min_band and y0 > height * 0.22 and y1 < height * 0.88:
                warnings.append(f"large internal horizontal whitespace band {y0 / height:.0%}-{y1 / height:.0%}")
            run_start = None
    return warnings


def scan(path: Path, blank_warn: float, edge_warn: float) -> list[str]:
    img = Image.open(path).convert("RGB")
    warnings: list[str] = []
    wf = white_fraction(img)
    ed = edge_density(img)
    bbox = nonwhite_bbox(img)
    if wf > blank_warn:
        warnings.append(f"{path.name}: high white/blank fraction {wf:.2%}")
    if ed > edge_warn:
        warnings.append(f"{path.name}: high edge density {ed:.3f}; inspect for crowding/overlap")
    if bbox is None:
        warnings.append(f"{path.name}: appears blank")
    else:
        left, top, right, bottom = bbox
        margin_warnings = []
        if top > img.height * 0.22:
            margin_warnings.append("large top blank area")
        if bottom < img.height * 0.78:
            margin_warnings.append("large bottom blank area")
        if margin_warnings:
            warnings.append(f"{path.name}: {', '.join(margin_warnings)}")
    for warning in large_internal_whitespace(img):
        warnings.append(f"{path.name}: {warning}")
    return warnings


def main() -> int:
    parser = argparse.ArgumentParser(description="Heuristic scan for blank/crowded rendered slide PNGs.")
    parser.add_argument("png_dir", type=Path)
    parser.add_argument("--blank-warn", type=float, default=0.80)
    parser.add_argument("--edge-warn", type=float, default=0.075)
    parser.add_argument("--fail-on-warning", action="store_true")
    args = parser.parse_args()

    paths = sorted(args.png_dir.glob("slide-*.png"))
    if not paths:
        print(f"no slide PNGs found in {args.png_dir}")
        return 2

    warnings: list[str] = []
    for path in paths:
        warnings.extend(scan(path, args.blank_warn, args.edge_warn))

    if warnings:
        print("Rendered slide scan warnings:")
        for warning in warnings:
            print(" -", warning)
    else:
        print("Rendered slide scan passed.")

    return 1 if warnings and args.fail_on_warning else 0


if __name__ == "__main__":
    raise SystemExit(main())
