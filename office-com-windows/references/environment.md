# Environment notes and how the Word hang was diagnosed

Read this when something behaves unexpectedly; the short version lives in
`SKILL.md`.

## What is installed

| Tool | State |
| --- | --- |
| Microsoft Office 16 (Word, PowerPoint, Excel) | installed at `C:\Program Files\Microsoft Office\root\Office16\` |
| LibreOffice (`soffice.exe`) | **not installed** - so the `documents` skill's `render_docx.py` cannot run |
| WPS Office | not installed |
| Bundled Python | `C:\Users\<user>\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe` |
| Bundled Node | `...\dependencies\node\bin\node.exe`, packages in `...\node\node_modules` |
| Poppler | `...\dependencies\native\poppler\Library\bin` (`pdftoppm.exe`, `pdfinfo.exe`) |

## The Word "hang on save" problem

Symptoms observed while building a lab report:

- `Word.Application` creates fine (~2 s) and `Documents.Open` returns fine.
- `doc.Content.Text = "..."` fine; `doc.WordOpenXML` fine; `doc.Close(0)` fine.
- `doc.SaveAs2(...)`, `doc.SaveAs(...)` (any format: docx, rtf, txt) and
  `doc.ExportAsFixedFormat(...)` never return.
- Not caused by add-ins (`/a` still hangs), not by `DisplayAlerts`, not by
  `AutoSaveOn`, not by the default printer, not by the target volume.
- PowerPoint performs the same kind of automation perfectly (`Slides.Export`
  wrote 18 PNGs), which proves the COM/automation plumbing is healthy.

**Root cause (empirical): the destination path.** Writing to
`C:\Users\<user>\...` (Documents, Desktop, AppData, `%TEMP%`) hangs; writing to
`C:\CodexTmp\...` completes in about a second. A security/AV product
(Lenovo Anti-Virus powered by Huorong was registered on this machine) is the
likely filter, but the practical rule is simply: **let Word write outside the
user profile**.

Consequences:

- Never call Word's writer for a deliverable located in the workspace; build the
  file with python-docx instead.
- For rendering/QA, export the PDF to `C:\CodexTmp` and convert it there.
- For `.doc` to `.docx`, use Flat OPC (`Document.WordOpenXML` + `flat2docx.py`)
  rather than SaveAs.

## The add-in that blocks automation

`HKCU\Software\Microsoft\Office\Word\Addins\MSOfficePLUS` (LoadBehavior 3) loads a
WebView-based add-in into Word. With it loaded, `Documents.Add()` followed by any
document operation hangs, and window enumeration shows no dialog - only an
invisible `HardwareMonitorWindow` plus `MsoWorkPane title=OfficePLUS`.

`/a` (Word started with no add-ins and no global templates) removes the problem.
If a task ever needs the add-in disabled through the registry instead, set
`LoadBehavior` to 0 and **restore it to 3 afterwards**.

## PowerShell 5.1 and non-ASCII scripts

`powershell.exe -File script.ps1` reads the script using the ANSI code page on
this Chinese-locale machine. A UTF-8 file containing Chinese characters is
decoded incorrectly, which garbles literals and can break parsing (a document
path became `C:\Users\...\瀹為獙鎶ュ憡妯＄増(1).doc` and Word reported
"file not found").

Rules used by every script here:

- keep `.ps1` files ASCII-only;
- pass paths, heading texts and other non-ASCII values in through environment
  variables;
- use `$PSScriptRoot` instead of hard-coding a directory that contains the user
  name.

Python 3 scripts may contain Chinese because Python reads source files as UTF-8;
call `sys.stdout.reconfigure(encoding="utf-8", errors="replace")` so console
output does not raise `UnicodeEncodeError` on a GBK console.

## Copying files with non-ASCII names

Inline PowerShell commands passed through the shell handle Chinese paths
correctly. Prefer copying an input to an ASCII-named file inside the working
directory first (for example `work\template_src.doc`) so later scripts and
external tools never have to deal with the original name.

## Clean-up expectations

- Restore any registry value, default printer, or Office option that was changed
  for diagnosis.
- `C:\CodexTmp` is deliberately outside the workspace, so workspace-scoped delete
  policies may refuse to remove it; tell the user it exists and can be deleted.
