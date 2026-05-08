from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

for prefix, uri in {
    "cp": CP_NS,
    "dc": DC_NS,
    "dcterms": DCTERMS_NS,
    "xsi": XSI_NS,
    "ep": EP_NS,
    "vt": VT_NS,
}.items():
    ET.register_namespace(prefix, uri)


REMOVED_PART_PREFIXES = (
    "ppt/notesSlides/",
    "ppt/notesMasters/",
    "ppt/comments/",
    "ppt/threadedComments/",
    "ppt/persons/",
)
REMOVED_PARTS = {
    "docProps/custom.xml",
    "docProps/thumbnail.jpeg",
    "docProps/thumbnail.jpg",
    "docProps/thumbnail.png",
    "ppt/commentAuthors.xml",
    "ppt/authors.xml",
    "ppt/people.xml",
}
REMOVED_REL_MARKERS = (
    "/notesSlide",
    "/notesMaster",
    "/comments",
    "/commentAuthors",
    "/people",
    "/persons",
    "/custom-properties",
    "/metadata/thumbnail",
)
DEFAULT_REPLACEMENTS = {
    "：<real name>": "：报告人",
    ": <real name>": ": Presenter",
}
DEFAULT_FORBIDDEN = (
    "<real name>",
    "<machine name>",
    "<user name>",
)


def collect_pptx(paths: list[Path]) -> list[Path]:
    files: list[Path] = []
    for path in paths:
        if path.is_dir():
            files.extend(sorted(path.rglob("*.pptx")))
        elif path.suffix.lower() == ".pptx":
            files.append(path)
    return files


def parse_replacements(items: list[str]) -> dict[str, str]:
    replacements = dict(DEFAULT_REPLACEMENTS)
    for item in items:
        if "=" not in item:
            raise ValueError(f"replacement must use OLD=NEW: {item}")
        old, new = item.split("=", 1)
        replacements[old] = new
    return replacements


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def target_is_removed(target: str) -> bool:
    normalized = target.replace("\\", "/")
    return (
        normalized.startswith("../notesSlides/")
        or normalized.startswith("../notesMasters/")
        or normalized.startswith("../comments/")
        or normalized.startswith("../threadedComments/")
        or normalized.startswith("../persons/")
        or normalized in {"../commentAuthors.xml", "../authors.xml", "../people.xml"}
        or normalized.startswith("notesSlides/")
        or normalized.startswith("notesMasters/")
        or normalized.startswith("comments/")
        or normalized.startswith("threadedComments/")
        or normalized.startswith("persons/")
        or normalized in {"commentAuthors.xml", "authors.xml", "people.xml"}
        or normalized == "docProps/custom.xml"
        or normalized in {"docProps/thumbnail.jpeg", "docProps/thumbnail.jpg", "docProps/thumbnail.png"}
    )


def should_remove_part(name: str) -> bool:
    return name in REMOVED_PARTS or any(name.startswith(prefix) for prefix in REMOVED_PART_PREFIXES)


def sanitize_relationships(data: bytes) -> bytes:
    root = ET.fromstring(data)
    changed = False
    for rel in list(root):
        typ = rel.get("Type", "")
        target = rel.get("Target", "")
        if any(marker in typ for marker in REMOVED_REL_MARKERS) or target_is_removed(target):
            root.remove(rel)
            changed = True
    return xml_bytes(root) if changed else data


def sanitize_content_types(data: bytes) -> bytes:
    root = ET.fromstring(data)
    changed = False
    for item in list(root):
        part = item.get("PartName", "").lstrip("/")
        content_type = item.get("ContentType", "")
        if should_remove_part(part) or any(marker.strip("/") in content_type for marker in REMOVED_REL_MARKERS):
            root.remove(item)
            changed = True
    return xml_bytes(root) if changed else data


def set_text(parent: ET.Element, tag: str, value: str) -> None:
    child = parent.find(tag)
    if child is None:
        child = ET.SubElement(parent, tag)
    child.text = value


def sanitize_core(data: bytes, title: str, creator: str) -> bytes:
    root = ET.fromstring(data)
    set_text(root, f"{{{DC_NS}}}title", title)
    set_text(root, f"{{{DC_NS}}}subject", "")
    set_text(root, f"{{{DC_NS}}}creator", creator)
    set_text(root, f"{{{CP_NS}}}keywords", "")
    set_text(root, f"{{{CP_NS}}}lastModifiedBy", creator)
    set_text(root, f"{{{CP_NS}}}revision", "1")
    created = root.find(f"{{{DCTERMS_NS}}}created")
    if created is not None:
        created.text = "2026-01-01T00:00:00Z"
        created.set(f"{{{XSI_NS}}}type", "dcterms:W3CDTF")
    modified = root.find(f"{{{DCTERMS_NS}}}modified")
    if modified is not None:
        modified.text = "2026-01-01T00:00:00Z"
        modified.set(f"{{{XSI_NS}}}type", "dcterms:W3CDTF")
    return xml_bytes(root)


def sanitize_app(data: bytes, creator: str) -> bytes:
    root = ET.fromstring(data)
    for tag, value in {
        f"{{{EP_NS}}}Company": "",
        f"{{{EP_NS}}}Manager": "",
        f"{{{EP_NS}}}Application": "Microsoft PowerPoint",
        f"{{{EP_NS}}}AppVersion": "16.0000",
    }.items():
        node = root.find(tag)
        if node is None:
            node = ET.SubElement(root, tag)
        node.text = value
    total_time = root.find(f"{{{EP_NS}}}TotalTime")
    if total_time is not None:
        total_time.text = "0"
    for tag in ("Notes", "HiddenSlides", "MMClips"):
        node = root.find(f"{{{EP_NS}}}{tag}")
        if node is not None:
            node.text = "0"
    return xml_bytes(root)


def replace_visible_text(data: bytes, replacements: dict[str, str]) -> bytes:
    text = data.decode("utf-8", errors="ignore")
    original = text
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text.encode("utf-8") if text != original else data


def sanitize_pptx(path: Path, replacements: dict[str, str], title: str, creator: str, *, in_place: bool) -> Path:
    if not path.exists():
        raise FileNotFoundError(path)

    out_path = path if in_place else path.with_name(f"{path.stem}-sanitized{path.suffix}")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp_path = Path(tmp.name)

    try:
        with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                name = info.filename
                if should_remove_part(name):
                    continue
                data = zin.read(name)
                if name.endswith(".rels"):
                    data = sanitize_relationships(data)
                elif name == "[Content_Types].xml":
                    data = sanitize_content_types(data)
                elif name == "docProps/core.xml":
                    data = sanitize_core(data, title, creator)
                elif name == "docProps/app.xml":
                    data = sanitize_app(data, creator)
                elif name.endswith(".xml"):
                    data = replace_visible_text(data, replacements)
                zout.writestr(info, data)
        shutil.move(str(tmp_path), str(out_path))
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
    return out_path


def scan_for_privacy(path: Path, forbidden: list[str]) -> list[str]:
    warnings: list[str] = []
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if should_remove_part(name):
                warnings.append(f"{path.name}: remaining private/comment/notes part `{name}`")
            if not name.endswith((".xml", ".rels")):
                continue
            try:
                text = zf.read(name).decode("utf-8")
            except UnicodeDecodeError:
                continue
            for token in forbidden:
                if token and re.search(re.escape(token), text, flags=re.IGNORECASE):
                    warnings.append(f"{path.name}: forbidden token `{token}` in `{name}`")
    return warnings


def main() -> int:
    parser = argparse.ArgumentParser(description="Remove personal metadata, notes, comments, and known names from PPTX files.")
    parser.add_argument("paths", nargs="+", type=Path, help="PPTX files or directories containing PPTX files")
    parser.add_argument("--in-place", action="store_true", help="Rewrite the input files instead of writing *-sanitized.pptx")
    parser.add_argument("--replace", action="append", default=[], help="Visible XML replacement in OLD=NEW form")
    parser.add_argument("--forbid", action="append", default=[], help="Forbidden token to check after sanitizing")
    parser.add_argument("--title", default="Paper Report")
    parser.add_argument("--creator", default="paper2ppt")
    parser.add_argument("--check-only", action="store_true", help="Only run the privacy scan")
    parser.add_argument("--fail-on-warning", action="store_true")
    args = parser.parse_args()

    pptx_files = collect_pptx(args.paths)
    if not pptx_files:
        print("no PPTX files found", file=sys.stderr)
        return 2

    replacements = parse_replacements(args.replace)
    forbidden = list(DEFAULT_FORBIDDEN) + args.forbid

    all_warnings: list[str] = []
    for pptx in pptx_files:
        target = pptx
        if not args.check_only:
            target = sanitize_pptx(pptx, replacements, args.title, args.creator, in_place=args.in_place)
            print(f"sanitized: {target}")
        warnings = scan_for_privacy(target, forbidden)
        all_warnings.extend(warnings)

    if all_warnings:
        print("Privacy warnings:")
        for warning in all_warnings:
            print(" -", warning)
    else:
        print("Privacy scan passed.")

    return 1 if all_warnings and args.fail_on_warning else 0


if __name__ == "__main__":
    raise SystemExit(main())
