from __future__ import annotations

import argparse
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


PROFILE_TARGETS = {
    "compact": {"main": 18.2, "secondary": 16.8, "tertiary": 16.2},
    "dense-visual": {"main": 18.2, "secondary": 16.8, "tertiary": 16.2},
    "classic-large": {"main": 19.2, "secondary": 18.4, "tertiary": 17.0},
}

BULLET_LEVELS = {
    "●": "main",
    "•": "secondary",
    "–": "tertiary",
    "-": "tertiary",
}


def slide_names(zf: zipfile.ZipFile) -> list[str]:
    return sorted(
        (name for name in zf.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")),
        key=lambda x: int(Path(x).stem.replace("slide", "")),
    )


def bullet_level(text: str) -> str | None:
    stripped = text.lstrip()
    if not stripped:
        return None
    return BULLET_LEVELS.get(stripped[0])


def set_typeface(rpr: ET.Element, font: str) -> None:
    for tag in ("latin", "ea", "cs"):
        qname = f"{{{NS['a']}}}{tag}"
        node = rpr.find(f"./a:{tag}", NS)
        if node is None:
            node = ET.SubElement(rpr, qname)
        node.set("typeface", font)


def normalize_run(run: ET.Element, size_pt: float, font: str) -> bool:
    rpr = run.find("./a:rPr", NS)
    if rpr is None:
        rpr = ET.Element(f"{{{NS['a']}}}rPr")
        run.insert(0, rpr)
    wanted = str(int(round(size_pt * 100)))
    changed = rpr.get("sz") != wanted
    rpr.set("sz", wanted)
    before = ET.tostring(rpr, encoding="unicode")
    set_typeface(rpr, font)
    return changed or before != ET.tostring(rpr, encoding="unicode")


def normalize_paragraph(paragraph: ET.Element, level: str, size_pt: float, font: str) -> bool:
    changed = False
    for run in paragraph.findall("./a:r", NS):
        if normalize_run(run, size_pt, font):
            changed = True
    end = paragraph.find("./a:endParaRPr", NS)
    if end is not None:
        wanted = str(int(round(size_pt * 100)))
        if end.get("sz") != wanted:
            end.set("sz", wanted)
            changed = True
        before = ET.tostring(end, encoding="unicode")
        set_typeface(end, font)
        changed = changed or before != ET.tostring(end, encoding="unicode")
    return changed


def repair_slide(xml: bytes, targets: dict[str, float], *, font: str, fix_newlines: bool) -> tuple[bytes, int]:
    root = ET.fromstring(xml)
    changes = 0
    for shape in root.findall(".//p:sp", NS):
        text_body = shape.find("./p:txBody", NS)
        if text_body is None:
            continue
        paragraphs = text_body.findall("./a:p", NS)
        texts = [
            "".join(node.text or "" for node in paragraph.findall(".//a:t", NS))
            for paragraph in paragraphs
        ]
        if paragraphs and not any(text.strip() for text in texts):
            shape.remove(text_body)
            changes += 1
            continue
        for paragraph, text in zip(paragraphs, texts):
            if fix_newlines:
                for node in paragraph.findall(".//a:t", NS):
                    if node.text and "\n" in node.text:
                        node.text = " ".join(node.text.split())
                        changes += 1
            level = bullet_level(text)
            if level and normalize_paragraph(paragraph, level, targets[level], font):
                changes += 1
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changes


def repair_pptx(src: Path, dst: Path, *, profile: str, font: str, fix_newlines: bool) -> int:
    targets = PROFILE_TARGETS[profile]
    dst.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = dst.with_name(f"{dst.stem}.repairing{dst.suffix}")
    if tmp_path.exists():
        tmp_path.unlink()
    total_changes = 0
    try:
        with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            slide_set = set(slide_names(zin))
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in slide_set:
                    data, changes = repair_slide(data, targets, font=font, fix_newlines=fix_newlines)
                    total_changes += changes
                zout.writestr(item, data)
        tmp_path.replace(dst)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
    return total_changes


def main() -> int:
    parser = argparse.ArgumentParser(description="Repair common generated PPTX layout hygiene issues.")
    parser.add_argument("pptx", nargs="+", type=Path)
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--in-place", action="store_true")
    parser.add_argument("--profile", choices=sorted(PROFILE_TARGETS), default="compact")
    parser.add_argument("--font", default="Times New Roman")
    parser.add_argument("--fix-newlines", action="store_true")
    args = parser.parse_args()

    if not args.in_place and not args.out_dir:
        parser.error("use --in-place or --out-dir")
    for path in args.pptx:
        dst = path if args.in_place else args.out_dir / path.name
        changes = repair_pptx(path, dst, profile=args.profile, font=args.font, fix_newlines=args.fix_newlines)
        print(f"{dst}: repaired {changes} item(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
