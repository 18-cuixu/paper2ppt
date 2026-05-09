from __future__ import annotations

import argparse
import json
import os
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
CODEX_RUNTIME_PYTHON = Path.home() / ".cache" / "codex-runtimes" / "codex-primary-runtime" / "dependencies" / "python" / "python.exe"


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


def resolve_baseline_pdf(case_id: str, baseline_dir: Path | None) -> Path | None:
    if baseline_dir is None:
        return None
    candidate = (baseline_dir / case_id / "input.pdf").resolve()
    return candidate if candidate.exists() else None


def soffice_path() -> str | None:
    configured = os.environ.get("PAPER2PPT_SOFFICE")
    if configured:
        return configured
    found = shutil.which("soffice.com") or shutil.which("soffice")
    if found:
        return found
    if WINDOWS_SOFFICE.exists():
        return str(WINDOWS_SOFFICE)
    if WINDOWS_SOFFICE_EXE.exists():
        return str(WINDOWS_SOFFICE_EXE)
    return None


def script_python() -> str:
    configured = os.environ.get("PAPER2PPT_PYTHON")
    if configured:
        return configured
    if CODEX_RUNTIME_PYTHON.exists():
        return str(CODEX_RUNTIME_PYTHON)
    return sys.executable


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


def assert_readable_file(path: Path, label: str) -> None:
    try:
        with path.open("rb") as handle:
            handle.read(4)
    except OSError as exc:
        raise RuntimeError(f"{label} is not readable: {path} ({exc})") from exc


def canonicalize_pptx_for_export(src: Path, dst: Path) -> str:
    """Save through python-pptx before LibreOffice export when possible.

    Some PPTX files pass zip/XML checks and open with python-pptx, but LibreOffice
    refuses them with "source file could not be loaded" until the package is
    reserialized. Keep the source untouched and use a canonical staged copy.
    """
    try:
        from pptx import Presentation

        prs = Presentation(str(src))
        prs.save(dst)
        return "python-pptx canonicalized"
    except Exception:
        shutil.copy2(src, dst)
        return "raw copied"


def make_export_stage(out_dir: Path) -> Path:
    # LibreOffice on Windows is sensitive to deep source/profile paths. Keep the
    # staged PPTX and user profile directly under the workspace root, then copy
    # the exported PDF back to the case output directory.
    del out_dir
    stage_parent = workspace_root()
    stage_root = stage_parent / f"_paper2ppt_lo_stage_{os.getpid()}"
    counter = 0
    while stage_root.exists():
        counter += 1
        stage_root = stage_parent / f"_paper2ppt_lo_stage_{os.getpid()}_{counter}"
    stage_root.mkdir(parents=True, exist_ok=False)
    return stage_root


def export_pdf(pptx: Path, out_dir: Path, office: str, *, timeout: int) -> Path:
    pptx = pptx.resolve()
    assert_readable_file(pptx, "source PPTX")
    out_dir = out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    stage_root = make_export_stage(out_dir)
    staged = stage_root / "input.pptx"
    staged_pdf = stage_root / "input.pdf"
    profile = stage_root / "profile"
    try:
        profile.mkdir(parents=True, exist_ok=True)
        mode = canonicalize_pptx_for_export(pptx, staged)
        print(f"staged PPTX for LibreOffice export: {mode}", flush=True)
        assert_readable_file(staged, "staged PPTX")
        cmd = [
            office,
            "--headless",
            f"-env:UserInstallation=file:///{profile.as_posix()}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(stage_root),
            str(staged),
        ]
        proc = run(cmd, timeout=timeout)
        print(proc.stdout)
        if proc.returncode != 0:
            raise RuntimeError(f"LibreOffice export failed for {pptx.name}")
        if not staged_pdf.exists():
            matches = sorted(stage_root.glob("*.pdf"))
            if not matches:
                raise RuntimeError(f"LibreOffice did not create a PDF for {pptx.name}")
            staged_pdf = matches[-1]
        final_pdf = out_dir / "input.pdf"
        shutil.copy2(staged_pdf, final_pdf)
        return final_pdf
    finally:
        shutil.rmtree(stage_root, ignore_errors=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Run PPTX audit/render/scan across template-profile examples.")
    parser.add_argument("--matrix", type=Path, default=DEFAULT_MATRIX)
    parser.add_argument("--out-dir", type=Path, default=ROOT / "out" / "template-matrix")
    parser.add_argument("--case", action="append", dest="case_ids", help="Run only the named case id.")
    parser.add_argument("--skip-render", action="store_true", help="Only run PPTX audits.")
    parser.add_argument(
        "--baseline-dir",
        type=Path,
        help=(
            "Use existing per-case input.pdf files from this directory for render/scan. "
            "This is an explicit fallback for environments where LibreOffice export is unavailable."
        ),
    )
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
    if not args.skip_render and not office and args.baseline_dir is None:
        print("LibreOffice not found; rerun with --skip-render, provide --baseline-dir, or install LibreOffice.", file=sys.stderr)
        return 2

    failures: list[str] = []
    args.out_dir.mkdir(parents=True, exist_ok=True)
    py = script_python()

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
        try:
            assert_readable_file(pptx, "case PPTX")
        except RuntimeError as exc:
            message = f"{case_id}: {exc}"
            print(message)
            failures.append(message)
            if not args.keep_going:
                break
            continue

        proc = run([
            py,
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
        baseline_pdf = resolve_baseline_pdf(case_id, args.baseline_dir)
        if case.get("legacy_pdf_baseline") and (not legacy_pdf or not legacy_pdf.exists()):
            print("legacy rendered PDF baseline is not bundled; running PPTX audit only for this reference case.")
            continue
        if case.get("legacy_pdf_baseline") and legacy_pdf and legacy_pdf.exists():
            print(f"using legacy rendered PDF baseline: {legacy_pdf}", flush=True)
            pdf = legacy_pdf
        else:
            if office:
                try:
                    pdf = export_pdf(pptx, case_dir, office, timeout=args.export_timeout)
                except RuntimeError as exc:
                    print(exc)
                    if baseline_pdf:
                        print(f"falling back to explicit rendered PDF baseline: {baseline_pdf}", flush=True)
                        pdf = baseline_pdf
                    else:
                        failures.append(f"{case_id}: LibreOffice export failed")
                        if not args.keep_going:
                            break
                        continue
            elif baseline_pdf:
                print(f"LibreOffice not found; using explicit rendered PDF baseline: {baseline_pdf}", flush=True)
                pdf = baseline_pdf
            else:
                failures.append(f"{case_id}: LibreOffice unavailable and no rendered PDF baseline")
                if not args.keep_going:
                    break
                continue
        png_dir = case_dir / "png"
        preview = case_dir / "preview-grid.png"
        proc = run([
            py,
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

        scan_cmd = [py, str(SCAN), str(png_dir), "--fail-on-warning"]
        if case.get("ignore_edge_slides", True):
            scan_cmd.append("--ignore-edge-slides")
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
