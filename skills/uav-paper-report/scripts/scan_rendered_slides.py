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


def body_content_bbox(img: Image.Image, threshold: int = 248):
    width, height = img.size
    top = int(height * 0.16)
    bottom = int(height * 0.92)
    if bottom <= top:
        return None
    body = img.crop((0, top, width, bottom)).convert("L")
    mask = body.point(lambda p: 255 if p < threshold else 0)
    bbox = mask.getbbox()
    if bbox is None:
        return None
    left, y0, right, y1 = bbox
    return left, top + y0, right, top + y1


def edge_density(img: Image.Image) -> float:
    gray = img.convert("L")
    shifted = ImageChops.offset(gray, 1, 0)
    diff = ImageChops.difference(gray, shifted)
    stat = ImageStat.Stat(diff)
    return stat.mean[0] / 255.0


def large_internal_whitespace(img: Image.Image, threshold: int = 248, min_band_fraction: float = 0.18) -> list[str]:
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
    min_band = int(height * min_band_fraction)
    for idx, blank in enumerate(row_blank + [0.0]):
        if blank > 0.995 and run_start is None:
            run_start = idx
        elif (blank <= 0.995 or idx == len(row_blank)) and run_start is not None:
            run_len = idx - run_start
            y0 = top + run_start
            y1 = top + idx
            if run_len >= min_band and y0 > height * 0.22 and y1 < height * 0.88:
                warnings.append(f"large internal horizontal whitespace band {y0 / height:.0%}-{y1 / height:.0%}")
            run_start = None
    return warnings


def large_empty_quadrants(img: Image.Image, threshold: int = 248) -> list[str]:
    gray = img.convert("L")
    width, height = gray.size
    top = int(height * 0.20)
    bottom = int(height * 0.88)
    left = int(width * 0.06)
    right = int(width * 0.94)
    if bottom <= top or right <= left:
        return []

    mid_x = (left + right) // 2
    mid_y = (top + bottom) // 2
    regions = [
        ("upper-left", (left, top, mid_x, mid_y)),
        ("upper-right", (mid_x, top, right, mid_y)),
        ("lower-left", (left, mid_y, mid_x, bottom)),
        ("lower-right", (mid_x, mid_y, right, bottom)),
    ]
    warnings: list[str] = []
    for label, box in regions:
        crop = gray.crop(box)
        if crop.width < width * 0.30 or crop.height < height * 0.24:
            continue
        hist = crop.histogram()
        blank = sum(hist[threshold:]) / (crop.width * crop.height)
        if blank > 0.985:
            warnings.append(f"large empty {label} body region")
    return warnings


def dense_text_bands(img: Image.Image, threshold: int = 248) -> list[str]:
    gray = img.convert("L")
    width, height = gray.size
    top = int(height * 0.16)
    bottom = int(height * 0.92)
    if bottom <= top:
        return []

    row_ink: list[float] = []
    for y in range(top, bottom):
        row = gray.crop((0, y, width, y + 1))
        hist = row.histogram()
        ink = sum(hist[:threshold]) / width
        row_ink.append(ink)

    warnings: list[str] = []
    run_start = None
    # A long row band with many dark pixels usually means text/image collision,
    # over-crowded tables, or a figure squeezed into prose.
    for idx, ink in enumerate(row_ink + [0.0]):
        if ink > 0.18 and run_start is None:
            run_start = idx
        elif (ink <= 0.18 or idx == len(row_ink)) and run_start is not None:
            run_len = idx - run_start
            if run_len >= int(height * 0.055):
                y0 = top + run_start
                y1 = top + idx
                warnings.append(f"dense rendered content band {y0 / height:.0%}-{y1 / height:.0%}; inspect for overlap")
            run_start = None
    return warnings


def scan(path: Path, blank_warn: float, edge_warn: float, dense_band: bool, min_band_fraction: float) -> list[str]:
    img = Image.open(path).convert("RGB")
    warnings: list[str] = []
    wf = white_fraction(img)
    ed = edge_density(img)
    bbox = nonwhite_bbox(img)
    body_bbox = body_content_bbox(img)
    sparse_body = True
    if body_bbox is not None:
        b_left, b_top, b_right, b_bottom = body_bbox
        body_h = int(img.height * 0.92) - int(img.height * 0.16)
        body_span = (b_bottom - b_top) / body_h if body_h > 0 else 0.0
        body_width = (b_right - b_left) / img.width if img.width else 0.0
        sparse_body = body_span < 0.46 or body_width < 0.52
    if wf > blank_warn and sparse_body:
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
    for warning in large_internal_whitespace(img, min_band_fraction=min_band_fraction):
        warnings.append(f"{path.name}: {warning}")
    for warning in large_empty_quadrants(img):
        warnings.append(f"{path.name}: {warning}")
    if dense_band:
        for warning in dense_text_bands(img):
            warnings.append(f"{path.name}: {warning}")
    return warnings


def main() -> int:
    parser = argparse.ArgumentParser(description="Heuristic scan for blank/crowded rendered slide PNGs.")
    parser.add_argument("png_dir", type=Path)
    parser.add_argument("--blank-warn", type=float, default=0.80)
    parser.add_argument("--edge-warn", type=float, default=0.075)
    parser.add_argument("--dense-band", action="store_true", help="Warn on long high-ink rendered bands.")
    parser.add_argument(
        "--min-band-fraction",
        type=float,
        default=0.18,
        help="Minimum slide-height fraction for an internal blank band warning.",
    )
    parser.add_argument(
        "--ignore-edge-slides",
        action="store_true",
        help="Skip the first and last rendered slides, normally cover/thanks pages.",
    )
    parser.add_argument("--fail-on-warning", action="store_true")
    args = parser.parse_args()

    paths = sorted(args.png_dir.glob("slide-*.png"))
    if not paths:
        print(f"no slide PNGs found in {args.png_dir}")
        return 2

    warnings: list[str] = []
    for index, path in enumerate(paths):
        if args.ignore_edge_slides and index in (0, len(paths) - 1):
            continue
        warnings.extend(scan(path, args.blank_warn, args.edge_warn, args.dense_band, args.min_band_fraction))

    if warnings:
        print("Rendered slide scan warnings:")
        for warning in warnings:
            print(" -", warning)
    else:
        print("Rendered slide scan passed.")

    return 1 if warnings and args.fail_on_warning else 0


if __name__ == "__main__":
    raise SystemExit(main())
