# Export a Word document to PDF through Word COM. ASCII-only on purpose.
#
# WHY THE OUTPUT PATH MATTERS
#   On this machine Word hangs forever inside SaveAs2 / ExportAsFixedFormat when
#   the destination lives under the user profile (Documents, Desktop, AppData,
#   %TEMP%). Writing to a directory outside the profile - for example C:\CodexTmp
#   - works normally. Always export to such a directory, then copy the file into
#   the workspace with ordinary file tools.
#
# Inputs (environment variables):
#   WORD_SRC   document to export                                (required)
#   WORD_OUT   target .pdf path, OUTSIDE the user profile        (required)
#
# The script also prints the page count, which is handy for layout checks.

$ErrorActionPreference = "Stop"

$src = $env:WORD_SRC
$out = $env:WORD_OUT
if (-not $src -or -not $out) {
  Write-Error "WORD_SRC and WORD_OUT must be set"
  exit 1
}
$home_dir = $env:USERPROFILE
if ($out.StartsWith($home_dir, [System.StringComparison]::OrdinalIgnoreCase)) {
  Write-Warning "WORD_OUT is inside the user profile; Word may hang. Prefer a path such as C:\CodexTmp\out.pdf"
}

$exe = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
if (-not (Test-Path $exe)) {
  $exe = (Get-Command WINWORD.EXE -ErrorAction SilentlyContinue).Source
}
if (-not $exe) { Write-Error "WINWORD.EXE not found"; exit 1 }

Start-Process -FilePath $exe -ArgumentList "/a" -WindowStyle Hidden | Out-Null
$word = $null
for ($i = 0; $i -lt 60; $i++) {
  Start-Sleep -Milliseconds 500
  try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    if ($word) { break }
  } catch { }
}
if (-not $word) { Write-Error "could not attach to Word"; exit 1 }
$word.Visible = $false
$word.DisplayAlerts = 0

$doc = $word.Documents.Open($src, $false, $true, $false)
$doc.Repaginate()
$pages = $doc.ComputeStatistics(2)
if (Test-Path $out) { Remove-Item -LiteralPath $out -Force }
$doc.ExportAsFixedFormat($out, 17)
$ok = Test-Path $out
try { $doc.Close(0) } catch { }
try { $word.Quit() } catch { }
Write-Output ("pages=" + $pages + " pdf_exists=" + $ok + " -> " + $out)
if (-not $ok) { exit 1 }
