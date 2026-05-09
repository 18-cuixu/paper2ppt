from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_MATRIX = ROOT / "assets" / "template-profiles" / "matrix.json"
AUDIT = ROOT / "scripts" / "audit_pptx_text.py"
RENDER = ROOT / "scripts" / "render_pptx_previews.py"
SCAN = ROOT / "scripts" / "scan_rendered_slides.py"

WINDOWS_SOFFICE = Path(r"C:\Program Files\LibreOffice\program\soffice.com")
WINDOWS_SOFFICE_EXE = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")


def resolve_case_path(matrix_path: Path, raw: str) -> Path:
    path = Path(raw)
    if not path.is_absolute():
        path = (matrix_path.parent / path).resolve()
    return path


def workspace_root() -> Path:
    # skills/uav-paper-report/scripts -> repository root by default.
    return ROOT.parents[2]


def resolve_optional_path(matrix_path: Path, raw: str | None) -> Path | None:
    if not raw:
        return None
    path = resolve_case_path(matrix_path, raw)
    if path.exists():
        return path
    alt = (workspace_root() / raw).resolve()
    return alt if alt.exists() else path


def soffice_path() -> str | None:
    found = shutil.which("soffice") or shutil.which("soffice.com")
    if found:
        return found
    if WINDOWS_SOFFICE_EXE.exists():
        return str(WINDOWS_SOFFICE_EXE)
    if WINDOWS_SOFFICE.exists():
        return str(WINDOWS_SOFFICE)
    return None


def run(
    cmd: list[str],
    *,
    cwd: Path | None = None,
    timeout: int | None = None,
) -> subprocess.CompletedProcess[str]:
    print("+", " ".join(cmd), flush=True)
    try:
        return subprocess.run(
            cmd,
            cwd=cwd,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:
        output = exc.stdout if isinstance(exc.stdout, str) else ""
        return subprocess.CompletedProcess(cmd, 124, output + f"\nTIMEOUT after {timeout}s\n")


def export_pdf(pptx: Path, out_dir: Path, office: str, *, timeout: int) -> Path:
    pptx = pptx.resolve()
    out_dir = out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    staged = out_dir / "input.pptx"
    shutil.copy2(pptx, staged)
    profile = out_dir / "lo-profile"
    profile.mkdir(parents=True, exist_ok=True)
    cmd = [
        office,
        "--headless",
        f"-env:UserInstallation=file:///{profile.as_posix()}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(staged),
    ]
    proc = run(cmd, timeout=timeout)
    print(proc.stdout)
    if proc.returncode != 0:
        raise RuntimeError(f"LibreOffice export failed for {pptx.name}")
    pdf = out_dir / f"{staged.stem}.pdf"
    if not pdf.exists():
        matches = sorted(out_dir.glob("*.pdf"))
        if not matches:
            raise RuntimeError(f"LibreOffice did not create a PDF for {pptx.name}")
        pdf = matches[-1]
    return pdf


def main() -> int:
    parser = argparse.ArgumentParser(description="Run PPTX audit/render/scan across template-profile examples.")
    parser.add_argument("--matrix", type=Path, default=DEFAULT_MATRIX)
    parser.add_argument("--out-dir", type=Path, default=ROOT / "out" / "template-matrix")
    parser.add_argument("--case", action="append", dest="case_ids", help="Run only the named case id.")
    parser.add_argument("--skip-render", action="store_true", help="Only run PPTX audits.")
    parser.add_argument("--keep-going", action="store_true", help="Continue after a failed case.")
    parser.add_argument("--export-timeout", type=int, default=90, help="LibreOffice export timeout per case in seconds.")
    args = parser.parse_args()

    matrix_path = args.matrix.resolve()
    data = json.loads(matrix_path.read_text(encoding="utf-8"))
    profiles = data["profiles"]
    cases = data["cases"]
    if args.case_ids:
        requested = set(args.case_ids)
        cases = [case for case in cases if case["id"] in requested]
        missing = requested - {case["id"] for case in cases}
        if missing:
            print(f"unknown case id(s): {', '.join(sorted(missing))}", file=sys.stderr)
            return 2

    office = None if args.skip_render else soffice_path()
    if not args.skip_render and not office:
        print("LibreOffice not found; rerun with --skip-render or install LibreOffice.", file=sys.stderr)
        return 2

    failures: list[str] = []
    args.out_dir.mkdir(parents=True, exist_ok=True)

    for case in cases:
        case_id = case["id"]
        pptx = resolve_case_path(matrix_path, case["pptx"])
        profile = profiles[case["template_profile"]]["audit_profile"]
        print(f"\n== {case_id} [{case['template_profile']} -> {profile}] ==", flush=True)
        if not pptx.exists():
            message = f"{case_id}: missing PPTX {pptx}"
            print(message)
            failures.append(message)
            if not args.keep_going:
                break
            continue

        proc = run([
            sys.executable,
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

        if args.skip_render:
            continue

        case_dir = args.out_dir / case_id
        pdf = None
        legacy_pdf = resolve_optional_path(matrix_path, case.get("pdf"))
        if case.get("legacy_pdf_baseline") and (not legacy_pdf or not legacy_pdf.exists()):
            print("legacy rendered PDF baseline is not bundled; running PPTX audit only for this reference case.")
            continue
        if case.get("legacy_pdf_baseline") and legacy_pdf and legacy_pdf.exists():
            print(f"using legacy rendered PDF baseline: {legacy_pdf}", flush=True)
            pdf = legacy_pdf
        else:
            try:
                pdf = export_pdf(pptx, case_dir, office or "soffice", timeout=args.export_timeout)
            except RuntimeError as exc:
                print(exc)
                failures.append(f"{case_id}: LibreOffice export failed")
                if not args.keep_going:
                    break
                continue
        png_dir = case_dir / "png"
        preview = case_dir / "preview-grid.png"
        proc = run([
            sys.executable,
            str(RENDER),
            str(pdf),
            "--out-dir",
            str(png_dir),
            "--preview-grid",
            str(preview),
        ], timeout=120)
        print(proc.stdout)
        if proc.returncode != 0:
            failures.append(f"{case_id}: PNG render failed")
            if not args.keep_going:
                break
            continue

        scan_cmd = [sys.executable, str(SCAN), str(png_dir), "--fail-on-warning"]
        if "scan_blank_warn" in case:
            scan_cmd.extend(["--blank-warn", str(case["scan_blank_warn"])])
        proc = run(scan_cmd, timeout=60)
        print(proc.stdout)
        if proc.returncode != 0:
            if case.get("legacy_pdf_baseline"):
                print(f"{case_id}: legacy rendered baseline has scan warnings; keeping as non-blocking reference.")
            else:
                failures.append(f"{case_id}: rendered scan failed")
            if not args.keep_going:
                break

    if failures:
        print("\nTemplate matrix failed:")
        for failure in failures:
            print(" -", failure)
        return 1

    print("\nTemplate matrix passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
