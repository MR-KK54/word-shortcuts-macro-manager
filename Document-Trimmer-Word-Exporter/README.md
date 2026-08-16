# Word & PDF Page Exporter Pro

High-fidelity, **MS Word engine** page-range trimming for Word documents and PDFs.
Upload a document, specify page ranges (e.g. `1-4,5-6,7-10`), and get exact
target files whose layout, headers, footers, styles, sections, tables and
formatting match the source — with all trailing breaks removed from the last
page of every split.

Frontend: `index.html` + `static/` (PWA).
Backend: `server.py` (Flask) + `engine/`.

---

## The trimming engine (MS Word, not LibreOffice)

The trimming engine is built around **Microsoft Word** — the authoritative layout
engine — and never uses LibreOffice for splitting:

| Situation                      | Engine used                                                        |
| ------------------------------ | ------------------------------------------------------------------ |
| Windows + MS Word installed    | **MS Word COM** (`engine/word_com.py`) — Word's own pagination & format conversion. Exact fidelity. |
| Linux / Render (no Word)       | **Word's recorded pagination** (`engine/word_markers.py`) — reads the `<w:lastRenderedPageBreak/>` markers MS Word embeds when it saves a document, merged with explicit page/section breaks. Reproduces Word's pagination. |
| PDF source files               | PyMuPDF (`engine/pdf_split.py`).                                   |

### What the engine guarantees
- **100% source match** — every split is a clone of the source Word package
  (`styles.xml`, `numbering.xml`, `theme`, headers, footers, images, media)
  with only `word/document.xml` rewritten. Nothing is regenerated, so every
  layout detail (headers, footers, styles, alignments, indentation, fonts,
  tables, sections) is preserved byte-for-byte at the XML level.
- **Exact page ranges** — explicit page breaks are inserted at the source page
  boundaries so each output page mirrors the source page.
- **Breaks removed from the last page** — for `1-4,5-6,7-10` the last pages
  `4`, `6`, `10` have all trailing breaks (empty paragraphs, page/line break
  runs, section breaks) stripped.
- **Section-aware headers/footers** — each split keeps the correct header/footer
  of the section its pages belong to (verified with `1-4 | 5-6 | 7-10` across
  three sections with different headers).

---

## API (what the frontend calls)

| Method | Endpoint | Purpose |
| ------ | -------- | ------- |
| POST | `/api/upload` | Upload `.docx/.doc/.docm/.dotx/.dotm/.rtf/.pdf` |
| POST | `/api/inspect` | Page / section count of a file |
| POST | `/api/naming-preview` | Filename template preview |
| POST | `/api/export` | Start an export job (files + range spec + format + template) |
| GET  | `/api/job/<id>` | Poll job progress/logs/outputs |
| POST | `/api/job/<id>/cancel` | Cancel a job |
| GET  | `/api/download/<job_id>/<name>` | Download an output file |
| GET  | `/api/preview/<name>` | Render a page as PNG (via Word COM on Windows) |
| GET  | `/api/output-preview/<job_id>/<name>` | Preview an output page |
| POST | `/api/clear-storage` | Clear uploaded files / jobs from the server |

Range spec examples: `1-end`, `1-4,5-6,7-10`, `1, 5, 8-10`, `3-end`, `even`,
`odd`, `all-individual`.

Output formats: same as source, `docx`, `pdf`, `doc`, `rtf`, `docm`.

---

## Run locally (Windows + MS Word — full fidelity)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt   # includes pywin32 on Windows
python server.py                   # http://localhost:5000
```

Requirements: Windows with Microsoft Word installed (for the Word COM engine),
Python 3.10+.

---

## Deploy on Render via GitHub

> Note: Render runs **Linux containers**, where MS Word cannot run. On Render the
> trimming uses Word's recorded pagination markers (see above), which reproduces
> Word's layout for Word-authored documents. For byte-perfect MS Word rendering,
> run on Windows (e.g. this machine) where Word COM is used automatically.

1. Push this repository to GitHub.
2. In Render: **New → Blueprint** and select the repository
   (`render.yaml` is provided), or **New → Web Service**:
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn server:app --workers 1 --threads 8 --timeout 600`
   - Health check path: `/`
3. Done — the service is live and the frontend talks to the API automatically.

The `render.yaml` Blueprint is included:

```yaml
services:
  - type: web
    name: word-exporter
    runtime: python
    plan: free
    buildCommand: pip install -r requirements.txt
    startCommand: gunicorn server:app --workers 1 --threads 8 --timeout 600
    healthCheckPath: /
```
