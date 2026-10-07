# Structural inspection of a finished .docx through Word COM: page count, list
# numbering of chosen headings, per-image size/page, and any paragraph that
# drifts past a safe horizontal position. ASCII-only on purpose.
#
# Inputs (environment variables):
#   DOCX_PATH     document to inspect                                (required)
#   DOCX_OUT      text report to write                               (required)
#   DOCX_HEADS    optional "|"-separated heading texts to check      (optional)
#                 (useful to confirm automatic list numbering, e.g. a template
#                  that renders section 2 as the Chinese numeral for two)
#   DOCX_HLIMIT   horizontal position warning threshold in points    (default 440)

$ErrorActionPreference = "Stop"

$src = $env:DOCX_PATH
$out = $env:DOCX_OUT
if (-not $src -or -not $out) {
  Write-Error "DOCX_PATH and DOCX_OUT must be set"
  exit 1
}
$hlimit = 440
if ($env:DOCX_HLIMIT) { $hlimit = [double]$env:DOCX_HLIMIT }

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
$lines = New-Object System.Collections.ArrayList
[void]$lines.Add("pages=" + $doc.ComputeStatistics(2))
[void]$lines.Add("words=" + $doc.ComputeStatistics(0))
[void]$lines.Add("paragraphs=" + $doc.Paragraphs.Count)
[void]$lines.Add("tables=" + $doc.Tables.Count)
[void]$lines.Add("inlineShapes=" + $doc.InlineShapes.Count)
[void]$lines.Add("sections=" + $doc.Sections.Count)

if ($env:DOCX_HEADS) {
  [void]$lines.Add("")
  [void]$lines.Add("=== heading numbering ===")
  $heads = $env:DOCX_HEADS -split [regex]::Escape("|")
  for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
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

[void]$lines.Add("")
[void]$lines.Add("=== items past " + $hlimit + " pt horizontally ===")
$wide = 0
for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
  $p = $doc.Paragraphs.Item($i)
  try {
    $h = $p.Range.Information(5)
    if ($h -gt $hlimit) {
      $t = ($p.Range.Text -replace "`r", "" -replace "`a", "")
      if ($t.Length -gt 24) { $t = $t.Substring(0, 24) }
      [void]$lines.Add(("idx={0} page={1} hpos={2} text='{3}'" -f `
        $i, $p.Range.Information(3), [math]::Round($h, 1), $t))
      $wide++
    }
  } catch { }
}
[void]$lines.Add("wide items=" + $wide)

[void]$lines.Add("")
[void]$lines.Add("=== images ===")
for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
  $sh = $doc.InlineShapes.Item($i)
  [void]$lines.Add(("img {0} w={1} h={2} page={3}" -f $i, `
    [math]::Round($sh.Width, 1), [math]::Round($sh.Height, 1), `
    $sh.Range.Information(3)))
}

$lines -join "`r`n" | Out-File -LiteralPath $out -Encoding UTF8
try { $doc.Close(0) } catch { }
try { $word.Quit() } catch { }
Write-Output ("ok -> " + $out)
