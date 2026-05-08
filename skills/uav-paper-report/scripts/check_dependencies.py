from __future__ import annotations

import importlib.util
import shutil
import sys
from pathlib import Path


PYTHON_PACKAGES = {
    "pptx": "python-pptx",
    "PIL": "Pillow",
    "fitz": "PyMuPDF",
}

WINDOWS_SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.com"


def main() -> int:
    missing: list[str] = []
    for module, package in PYTHON_PACKAGES.items():
        if importlib.util.find_spec(module) is None:
            missing.append(package)

    soffice = shutil.which("soffice") or shutil.which("soffice.com")
    if not soffice and not Path(WINDOWS_SOFFICE).exists():
        missing.append("LibreOffice")

    if missing:
        print("Missing dependencies:")
        for item in missing:
            print(" -", item)
        print("Install Python packages with: python -m pip install python-pptx Pillow PyMuPDF")
        print("Install LibreOffice or set the soffice.com path before PDF export.")
        return 1

    print("Dependencies available.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
