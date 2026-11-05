# FileShift — Local File Converter

**100% free, runs locally, no cloud needed.**

## Conversions (60 paths)

| From → | PDF | DOCX | PPTX | XLSX | CSV | MD | TXT | HTML | PNG | JPG | WEBP |
|--------|:---:|:----:|:----:|:----:|:---:|:--:|:---:|:----:|:---:|:---:|:----:|
| **PDF**  | — | ✅ |   |   |   | ✅ | ✅ | ✅ | ✅ | ✅ |   |
| **DOCX** | ✅ | — |   |   |   | ✅ | ✅ | ✅ |   |   |   |
| **PPTX** | ✅ |   | — |   |   | ✅ | ✅ |   |   |   |   |
| **XLSX** | ✅ |   |   | — | ✅ | ✅ |   | ✅ |   |   |   |
| **CSV**  | ✅ |   |   | ✅ | — | ✅ |   | ✅ |   |   |   |
| **MD**   | ✅ | ✅ | ✅ |   |   | — | ✅ | ✅ |   |   |   |
| **TXT**  | ✅ | ✅ |   |   |   | ✅ | — | ✅ |   |   |   |
| **HTML** | ✅ | ✅ |   |   |   | ✅ | ✅ | — |   |   |   |
| **PNG**  | ✅ |   |   |   |   |   | 🔍 |   | — | ✅ | ✅ |
| **JPG**  | ✅ |   |   |   |   |   | 🔍 |   | ✅ | — | ✅ |
| **WEBP** | ✅ |   |   |   |   |   | 🔍 |   | ✅ | ✅ | — |
| **DOC** *(old Word)*       | ✅ | ✅ |   |   |   | ✅ | ✅ | ✅ |   |   |   |
| **PPT** *(old PowerPoint)* | ✅ |   | ✅ |   |   | ✅ | ✅ |   |   |   |   |
| **XLS** *(old Excel)*      | ✅ |   |   | ✅ | ✅ | ✅ |   | ✅ |   |   |   |

🔍 = uses OCR. Scanned PDFs (no text layer) are detected automatically: **PDF → TXT / MD / DOCX** runs OCR for you.

Notes:
- Multi-page PDF → PNG/JPG, multi-sheet XLSX → CSV and split results are delivered as a `.zip`.
- MD → PPTX: the first `#` heading becomes the title slide; every `#` / `##` heading starts a new slide.

## PDF tools

| Tool | Endpoint | What it does |
|------|----------|--------------|
| **Merge** | `POST /merge` | Combine PDFs and images (png/jpg/webp) into one PDF, in the order sent |
| **Split** | `POST /tools/split` | `mode=extract` selected pages → 1 PDF · `delete` remove pages · `ranges` one PDF per range · `every` chunks of N pages · `each` one PDF per page. Pages look like `1-3, 5, 8-` |
| **Compress** | `POST /tools/compress` | `level=low` (lossless) · `medium` (≈130 dpi, good for email) · `high` (≈80 dpi, smallest). Returns `X-Original-Size` / `X-Output-Size`. If the file can't get smaller you get it back unchanged |
| **Protect** | `POST /tools/protect` | AES-256 password. Optional `allow_print`, `allow_copy`, `allow_edit` (`false` to restrict) and `owner_password` |
| **Unlock** | `POST /tools/unlock` | Remove the password from a PDF you have the password for |
| **OCR** | `POST /tools/ocr` | Scanned PDF or image → searchable PDF (`output=pdf`) or text (`output=txt`). `language=eng` (or `eng+hin`), `force=true` to re-OCR pages that already have text |

Other endpoints: `GET /tools` (machine-readable catalogue of every tool and its parameters),
`GET /ocr/languages`, `GET /capabilities`, `GET /formats`.

**Try everything right now:** open **http://localhost:8080/docs** — an interactive page where you can upload a file and run any tool.

Passwords are sent only to your own local server, used in memory and never stored or logged.

## Setup (one time)

```bash
pip install -r requirements.txt
```

Optional free programs that unlock extra features:

| Program | Needed for | Install |
|---------|-----------|---------|
| **LibreOffice** | pixel-accurate PPTX → PDF, and reading old `.doc` / `.ppt` / `.xls` | libreoffice.org |
| **Tesseract OCR** | all OCR features | Windows: github.com/UB-Mannheim/tesseract/wiki · macOS: `brew install tesseract` · Linux: `sudo apt install tesseract-ocr` |

For other OCR languages, select them in the Tesseract installer (Windows) or install the language pack
(Linux: `sudo apt install tesseract-ocr-hin`, etc.). Restart FileShift after installing anything.

If a program is missing, FileShift tells you exactly what to install — everything else keeps working.
The **Engine Status** panel and `GET /capabilities` show what was detected.

## Run

```bash
python server.py
```

Then open **http://localhost:8080** in your browser (use `localhost` — never `0.0.0.0`).

| Env var | Default | Purpose |
|---------|---------|---------|
| `PORT`  | `8080`  | Port to listen on |
| `HOST`  | `127.0.0.1` | Set to `0.0.0.0` only if you want other devices on your network to reach it |

## The app

A sidebar with one page per feature:

| Page | What it does |
|------|--------------|
| **Home** | Overview, usage stats and shortcuts to every tool |
| **Convert** | Drop a file, pick a format (only the formats that file can become are shown), download |
| **Merge · Split · Compress · OCR · Protect · Unlock** | One page each, with options, progress bar and a result card |
| **History** | Everything processed in this browser (names and sizes only, stored in the browser's local storage — no database). *Download again* works until you close the tab |
| **System** | Which optional engines were detected (LibreOffice, Tesseract, WeasyPrint…) and how to install the missing ones |

Everything runs on your machine, nothing is uploaded anywhere; temp files are deleted after each download.

## Libraries used (all free/open source)

| Library | Purpose |
|---------|---------|
| `pymupdf` | PDF text/images, OCR bridge, split, merge, compress, encrypt |
| `weasyprint` | HTML/MD/spreadsheet → PDF rendering |
| `python-docx` | Read/write DOCX |
| `python-pptx` | Read PPTX, create PPTX from Markdown |
| `openpyxl` | Read/write XLSX |
| `Pillow` | Image conversion, image downsampling |
| `pdf2docx` | PDF → DOCX layout conversion |
| `beautifulsoup4` + `markdown` | HTML / Markdown parsing |
| `fastapi` + `uvicorn` | Local web server |
| LibreOffice *(optional)* | PPTX/DOCX → PDF, legacy Office input |
| Tesseract *(optional)* | OCR engine |