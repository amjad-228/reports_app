# PPTX Generator Backend (FastAPI)

Run locally:

```bash
python -m venv .venv
. .venv/Scripts/activate  # Windows PowerShell: .venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 8000 --reload
```

Env vars (optional):
- PPTX_TEMPLATE_PATH: absolute path to `report_template.pptx`. If unset, service will use `../public/templates/report_template.pptx` relative to repo root.
- LIBREOFFICE_PATH: full path to `soffice` executable if not on PATH.

Endpoints:
- GET /health
- POST /generate-pptx → returns a PPTX file (binary) based on payload fields
- POST /generate-pdf → returns a PDF generated server-side from PPTX (LibreOffice headless)

CORS is open for development; restrict in production.

PDF conversion (no Microsoft PowerPoint required):
- Install LibreOffice on the server (Windows/macOS/Linux)
- Ensure `soffice` is available on PATH or provide `LIBREOFFICE_PATH`
