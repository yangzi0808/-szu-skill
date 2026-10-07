# Dump text/tables/notes from a PowerPoint deck (.ppt or .pptx) through COM and
# optionally export every slide as PNG. ASCII-only on purpose.
#
# Inputs (environment variables):
#   PPT_SRC        deck to read                                        (required)
#   PPT_OUT        text report to write                                (required)
#   PPT_PNG_DIR    directory for slide-NN.png when exporting           (optional)
#   PPT_PNG_W      slide PNG width  in px                              (default 2560)
#   PPT_PNG_H      slide PNG height in px                              (default 1920)
#
# Two safety rules:
#   * Open a COPY of the deck, not the file the user has open. PowerPoint is
#     single-instance, so opening the same path again would attach to the user's
#     window and closing it would close their file.
#   * Open read-only and without a window: Presentations.Open(path, ReadOnly,
#     Untitled, WithWindow).

$ErrorActionPreference = "Stop"

$src = $env:PPT_SRC
$out = $env:PPT_OUT
if (-not $src -or -not $out) {
  Write-Error "PPT_SRC and PPT_OUT must be set"
  exit 1
}
$pngDir = $env:PPT_PNG_DIR
$pngW = 2560
$pngH = 1920
if ($env:PPT_PNG_W) { $pngW = [int]$env:PPT_PNG_W }
if ($env:PPT_PNG_H) { $pngH = [int]$env:PPT_PNG_H }

$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($src, $true, $false, $false)
$lines = New-Object System.Collections.ArrayList
[void]$lines.Add("slides=" + $pres.Slides.Count)

$i = 0
foreach ($s in $pres.Slides) {
  $i++
  [void]$lines.Add("--- slide $i ---")
  foreach ($sh in $s.Shapes) {
    if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) {
      $t = $sh.TextFrame.TextRange.Text -replace "`r", " / " -replace "`n", " / "
      if ($t.Trim().Length -gt 0) {
        [void]$lines.Add(("text[{0}] " -f $sh.Name) + $t.Trim())
      }
    }
    if ($sh.HasTable -eq -1) {
      foreach ($row in $sh.Table.Rows) {
        $cells = @()
        foreach ($c in $row.Cells) { $cells += $c.Shape.TextFrame.TextRange.Text }
        [void]$lines.Add("table: " + ($cells -join " | "))
      }
    }
  }
  try {
    $np = $s.NotesPage
    if ($np -and $np.Shapes.Count -ge 2) {
      $nt = $np.Shapes.Placeholders(2).TextFrame.TextRange.Text
      if ($nt -and $nt.Trim().Length -gt 0) { [void]$lines.Add("notes: " + $nt.Trim()) }
    }
  } catch { }
  if ($pngDir) {
    if (-not (Test-Path $pngDir)) { New-Item -ItemType Directory -Path $pngDir -Force | Out-Null }
    $f = Join-Path $pngDir ("slide-{0:d2}.png" -f $i)
    $s.Export($f, "PNG", $pngW, $pngH)
  }
}

$lines -join "`r`n" | Out-File -LiteralPath $out -Encoding UTF8
$pres.Close()
try { $app.Quit() } catch { }
Write-Output ("ok slides=" + $i + " -> " + $out)
