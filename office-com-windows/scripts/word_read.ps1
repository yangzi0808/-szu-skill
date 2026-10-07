# Read a Word document (.doc / .docx) through Word COM without modifying it.
#
# IMPORTANT: this file must stay ASCII-only. Windows PowerShell 5.1 reads .ps1
# files with the ANSI code page, so non-ASCII literals turn into mojibake and
# can even break parsing. Pass paths and non-ASCII strings through environment
# variables instead of embedding them here.
#
# Inputs (environment variables):
#   WORD_SRC     full path of the document to read            (required)
#   WORD_OUT     path of the report/text file to write        (required)
#   WORD_MODE    "structure" (default) | "flatopc" | "text"
#   WORD_HEADS   optional "|"-separated list of heading texts to locate
#
# Output: WORD_OUT. For "flatopc" the file is Flat OPC XML that
# scripts/flat2docx.py can rebuild into a real .docx package.

$ErrorActionPreference = "Stop"

$src = $env:WORD_SRC
$out = $env:WORD_OUT
$mode = $env:WORD_MODE
if (-not $mode) { $mode = "structure" }
if (-not $src -or -not $out) {
  Write-Error "WORD_SRC and WORD_OUT must be set"
  exit 1
}

$exe = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
if (-not (Test-Path $exe)) {
  $exe = (Get-Command WINWORD.EXE -ErrorAction SilentlyContinue).Source
}
if (-not $exe) { Write-Error "WINWORD.EXE not found"; exit 1 }

# Launch Word without add-ins, then attach. A third-party add-in (for example
# OfficePLUS) can block COM automation forever; /a avoids loading it.
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

if ($mode -eq "flatopc") {
  [System.IO.File]::WriteAllText($out, $doc.WordOpenXML, [System.Text.Encoding]::UTF8)
} else {
  $lines = New-Object System.Collections.ArrayList
  [void]$lines.Add("pages=" + $doc.ComputeStatistics(2))
  [void]$lines.Add("paragraphs=" + $doc.Paragraphs.Count)
  [void]$lines.Add("tables=" + $doc.Tables.Count)
  [void]$lines.Add("inlineShapes=" + $doc.InlineShapes.Count)
  [void]$lines.Add("sections=" + $doc.Sections.Count)
  foreach ($s in $doc.Sections) {
    $ps = $s.PageSetup
    [void]$lines.Add(("page W={0} H={1} top={2} bottom={3} left={4} right={5}" -f `
      $ps.PageWidth, $ps.PageHeight, $ps.TopMargin, $ps.BottomMargin, `
      $ps.LeftMargin, $ps.RightMargin))
  }
  [void]$lines.Add("")
  [void]$lines.Add("=== PARAGRAPHS ===")
  $n = $doc.Paragraphs.Count
  for ($i = 1; $i -le $n; $i++) {
    $p = $doc.Paragraphs.Item($i)
    $t = $p.Range.Text -replace "`r", "" -replace "`a", "" -replace "`n", ""
    $style = ""
    try { $style = $p.Style.NameLocal } catch { }
    $num = ""
    try { $num = $p.Range.ListFormat.ListString } catch { }
    $font = ""
    try { $font = ("{0}/{1}" -f $p.Range.Font.Name, $p.Range.Font.Size) } catch { }
    [void]$lines.Add(("[{0}] style={1} num='{2}' font={3} :: {4}" -f $i, $style, $num, $font, $t))
  }
  [void]$lines.Add("")
  [void]$lines.Add("=== TABLES ===")
  for ($ti = 1; $ti -le $doc.Tables.Count; $ti++) {
    $tb = $doc.Tables.Item($ti)
    [void]$lines.Add(("--- table {0}: rows={1} cols={2}" -f $ti, $tb.Rows.Count, $tb.Columns.Count))
    for ($ri = 1; $ri -le $tb.Rows.Count; $ri++) {
      $cells = @()
      for ($ci = 1; $ci -le $tb.Columns.Count; $ci++) {
        try { $tx = $tb.Cell($ri, $ci).Range.Text -replace "`r", "" -replace "`a", "" }
        catch { $tx = "<merged>" }
        $cells += $tx
      }
      [void]$lines.Add(("R{0}: {1}" -f $ri, ($cells -join " || ")))
    }
  }
  if ($env:WORD_HEADS) {
    [void]$lines.Add("")
    [void]$lines.Add("=== HEADING NUMBERING ===")
    $heads = $env:WORD_HEADS -split [regex]::Escape("|")
    for ($i = 1; $i -le $n; $i++) {
      $p = $doc.Paragraphs.Item($i)
      $t = ($p.Range.Text -replace "`r", "" -replace "`a", "").Trim()
      if ($heads -contains $t) {
        $num = ""
        try { $num = $p.Range.ListFormat.ListString } catch { }
        [void]$lines.Add(("idx={0} num='{1}' text='{2}' page={3}" -f `
          $i, $num, $t, $p.Range.Information(3)))
      }
    }
  }
  if ($mode -eq "text") {
    [void]$lines.Add("")
    [void]$lines.Add("=== FULL TEXT ===")
    [void]$lines.Add($doc.Content.Text)
  }
  $lines -join "`r`n" | Out-File -LiteralPath $out -Encoding UTF8
}

try { $doc.Close(0) } catch { }
try { $word.Quit() } catch { }
Write-Output ("ok " + $mode + " -> " + $out)
