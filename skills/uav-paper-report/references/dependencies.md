# Dependencies

The accepted workflow uses:

- Python 3.10+.
- `python-pptx` for editable PPTX generation.
- `Pillow` for image sizing and preview grids.
- `PyMuPDF` (`fitz`) for PDF-to-PNG rendering.
- LibreOffice `soffice.com` for headless PPTX-to-PDF export.

Example Windows paths:

- Python venv: `C:\path\to\project\.venv\Scripts\python.exe`
- LibreOffice: `C:\Program Files\LibreOffice\program\soffice.com`
- User project root: `C:\path\to\project`

If missing, install the dependency rather than skipping QA. The deck is not final until it has been exported and rendered for visual inspection.

Suggested checks:

```powershell
where.exe python
where.exe soffice
python -c "import pptx, PIL, fitz; print('ok')"
```

Use a separate LibreOffice user profile per run, for example:

```powershell
& 'C:\Program Files\LibreOffice\program\soffice.com' `
  '-env:UserInstallation=file:///C:/path/to/work/lo-profile' `
  --headless --convert-to pdf --outdir 'C:\path\to\rendered' 'C:\path\to\deck.pptx'
```
