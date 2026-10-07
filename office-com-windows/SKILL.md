---
name: office-com-windows
description: Read, convert, export and visually verify Word/PowerPoint files on this Windows machine through Office COM automation. Use when a task needs to read a .doc/.docx/.ppt/.pptx, convert .doc to .docx, export a document to PDF, render slides to images, or QA the layout of a finished Word file and there is no reliable LibreOffice.
metadata:
  short-description: Office COM 读写文档、导出 PDF 与渲染校验
---

# Office COM automation on Windows

This machine has Microsoft Office 16 (Word and PowerPoint) but **no LibreOffice**,
so the DOCX/PPTX reference renderers that depend on LibreOffice do not work here.
Everything goes through Office COM instead. The scripts in `scripts/` encode the
workarounds that took significant time to discover; read
`references/environment.md` before improvising.

## Non-obvious rules that break work if ignored

1. **Word cannot write inside the user profile.** `SaveAs2` and
   `ExportAsFixedFormat` hang forever when the destination is under
   `C:\Users\<user>\...` (Documents, Desktop, AppData, `%TEMP%`). Export to a
   directory outside the profile - `C:\CodexTmp` works - then copy the result
   into the workspace with ordinary file tools. Writing the same file with
   python-docx or PowerShell is fine; only Word's own writer hangs.
2. **Launch Word with `/a` before attaching.** A third-party add-in (OfficePLUS in
   this environment) makes `New-Object -ComObject Word.Application` succeed but
   every later document operation hang. `Start-Process WINWORD.EXE /a`, then
   `[Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")`.
   Diagnostics usually show no visible dialog, just an invisible
   `HardwareMonitorWindow`.
3. **Never close a presentation the user already has open.** PowerPoint is
   single-instance, so `Presentations.Open` on the same path attaches to the
   user's window and `Close()` would close their file. Work on a copy, opened with
   `Open(path, ReadOnly, Untitled, WithWindow=False)`.
4. **Keep `.ps1` files ASCII-only.** Windows PowerShell 5.1 reads script files
   with the ANSI code page; Chinese string literals become mojibake and can break
   parsing. Pass paths and non-ASCII strings through environment variables.
5. **Read a `.doc` through Flat OPC, not SaveAs.** `$doc.WordOpenXML` returns the
   whole document as XML without touching Word's writer; `flat2docx.py` rebuilds a
   normal `.docx` that python-docx can open.
6. **Kill stray `WINWORD.EXE` processes between runs.** A Word instance left over
   from a hung call makes the next attach fail with "could not attach".

## Scripts

All Word/PowerPoint scripts read their inputs from environment variables and are
launched as `Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File", <script>`
(or run directly with `powershell -NoProfile -File`).

| Script | Purpose |
| --- | --- |
| `scripts/word_read.ps1` | Read a `.doc`/`.docx` via Word: page/paragraph/table counts, styles, automatic list numbering, full text, or Flat OPC export (`WORD_MODE=flatopc`) |
| `scripts/flat2docx.py` | Rebuild a real `.docx` package from Flat OPC XML |
| `scripts/word_export_pdf.ps1` | Export a document to PDF **outside the profile**, and report the page count |
| `scripts/pdf2png.py` | Rasterise the PDF into `page-NN.png` with pypdfium2 |
| `scripts/docx_inspect.ps1` | Inspect a finished `.docx`: pages, words, list numbering of chosen headings, image sizes/pages, items past a horizontal limit |
| `scripts/pptx_dump.ps1` | Dump slide text/tables/notes and optionally export every slide as PNG |
| `scripts/qa_layout.py` | Report per-page bottom whitespace and captions orphaned onto the next page |

Typical read path for a `.doc` template:

```powershell
$env:WORD_SRC = "C:\path\template.doc"; $env:WORD_OUT = "C:\work\template_flat.xml"; $env:WORD_MODE = "flatopc"
Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","<skill>\scripts\word_read.ps1" -WindowStyle Hidden
# then
python <skill>\scripts\flat2docx.py C:\work\template_flat.xml C:\work\template.docx
```

## Verification loop for any finished Word file

1. `word_export_pdf.ps1` to `C:\CodexTmp\<name>.pdf` (never into the profile).
2. `pdf2png.py` to get page images, then **open every page image and look at it**
   - no clipping, overlap, missing glyphs, or images past the margin.
3. `qa_layout.py` on the same PDF for bottom-whitespace and orphan-caption problems.
4. `docx_inspect.ps1` to confirm page count, automatic numbering, and image sizes.

Layout defects that only show up in the rendered pages - a figure separated from
its caption, a section pushed to the next page, a table split in half - are the
reason this loop is mandatory rather than optional.

## Environment facts worth knowing

- Bundled Python (`...\codex-primary-runtime\dependencies\python\python.exe`) has
  `python-docx`, `python-pptx`, `Pillow`, `lxml`, `numpy`, `pdfplumber`,
  `pypdfium2`. It has **no** matplotlib and no `olefile`.
- Bundled Node has `sharp` but **not** `mathjax-full`, so LaTeX-to-image rendering
  is not available; use native Word equations instead (see the
  `lab-report-first-four-sections` skill).
- Poppler binaries live under
  `...\codex-primary-runtime\dependencies\native\poppler\Library\bin` when a tool
  needs `pdftoppm`.
- Word settings that a task may need to change temporarily (add-in load behaviour,
  default printer) must be restored when the task ends.
