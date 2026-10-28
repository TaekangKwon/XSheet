# ISSUE-93: the pixels Excel painted for every shape of ref-geoms.xlsx (U3: arcTo angles, shape outlines) and for flipped shapes with text (U4: text of flipped shapes).
# Part 1: opens ref-geoms.xlsx READ-ONLY (never saved) and exports every shape (all 225: AutoShapeType 1..183 presets, adjusted, rotated/flipped, freeforms, ...) to PNG.
# Part 2: builds ref-flips.xlsx (written into $Dir): ellipse / diamond / triangle / rightArrow / homePlate x (none, flipH, flipV, flipH+flipV, flipV+rot30),
#         each as a text shape (black text FLAG, top-left anchored, 1pt black outline, no fill) and as a twin without text (the outline only), then exports them to PNG.
# Export of one shape: CopyPicture -> paste into a chart (background magenta) -> Chart.Export PNG (same method as make-ref-colors.ps1; see its notes:
#   Excel must be Visible or the paste is empty; the pasted picture has a WHITE background rectangle; the picture sits at the top-left of the chart; Chart.Export writes indexed PNGs).
# Excel screen at 125% (120 dpi): 1 pt = 1.667 px in the PNGs (the compare script checks this from the picture size).
# SIDE EFFECTS: uses the Windows CLIPBOARD (CopyPicture) - do not copy anything else while it runs; Excel is shown (Visible) and driven by COM; nothing else is touched.
# Writes only into $Dir: geoms-png\*.png, flips-png\*.png, ref-flips.xlsx, flips.com.json (COM values of the flip shapes), export.json (what was exported / notes). PNGs are NOT committed.
# Run (do NOT run on CI; needs Excel 16, screen zoom 125%; PowerShell 5.1, file saved as UTF-8 with BOM):
#   powershell -ExecutionPolicy Bypass -File make-ref-geom-pngs.ps1 -Dir <abs out dir> [-Geoms <ref-geoms.xlsx>] [-Only name1,name2] [-Isolate]
# -Isolate: while a shape is exported, the other shapes within 80 pt of it are hidden (in memory only). Use it (with -Only) if an overlay shows a neighbour's ink in the Excel image
#   (CopyPicture of a shape should draw that shape only; arcs and callouts have frames larger than their pictures, so this is a precaution).
# After the run: node compare-ref-geoms.mjs <out dir>   (renders ref-geoms.xlsx / ref-flips.xlsx in InterHub, compares, writes ref-geoms.compare.json + ref-flips.compare.json).
param([string]$Dir, [string]$Geoms = '', [string[]]$Only = @(), [switch]$Isolate)
$ErrorActionPreference = 'Stop'
# relative -Dir is resolved against the PowerShell location (not the process cwd), then made absolute
if (-not $Dir) { throw '-Dir is required (absolute output directory)' }
if (-not [IO.Path]::IsPathRooted($Dir)) { $Dir = Join-Path (Get-Location).Path $Dir }
$Dir = [IO.Path]::GetFullPath($Dir)
if (-not $Geoms) { $Geoms = Join-Path $PSScriptRoot 'ref-geoms.xlsx' }
if (-not [IO.Path]::IsPathRooted($Geoms)) { $Geoms = Join-Path (Get-Location).Path $Geoms }
$Geoms = [IO.Path]::GetFullPath($Geoms)
if (-not (Test-Path -LiteralPath $Geoms)) { throw "ref-geoms.xlsx not found: $Geoms" }
New-Item -ItemType Directory -Force -Path $Dir | Out-Null
$geomPng = Join-Path $Dir 'geoms-png'
$flipPng = Join-Path $Dir 'flips-png'
New-Item -ItemType Directory -Force -Path $geomPng | Out-Null
New-Item -ItemType Directory -Force -Path $flipPng | Out-Null
# -Only name1,name2 may also arrive as a single 'name1,name2' string
$onlySet = @{}
foreach ($o in $Only) { foreach ($n in ($o -split ',')) { if ($n.Trim()) { $onlySet[$n.Trim()] = $true } } }

# chart background (magenta). Excel RGB int = R + G*256 + B*65536
$BGINT = 255 + 0 * 256 + 255 * 65536
$notes = New-Object System.Collections.ArrayList
$exported = New-Object System.Collections.ArrayList
$script:pasteNoted = $false
$script:Isolate = [bool]$Isolate
$script:ws = $null
$script:wsw = $null
$script:pngDir = $null

function Try-Do($label, [scriptblock]$sb) { try { & $sb } catch { [void]$script:notes.Add("$label failed: $($_.Exception.Message)") } }
function Safe($sb) { try { & $sb } catch { $null } }

# Export one shape to PNG: CopyPicture -> paste into a magenta chart (on the work sheet) -> Chart.Export. Returns $true when something was pasted.
function Export-Png($s, [string]$name) {
  # the picture of a rotated shape is its rotated bounding box (never larger than the diagonal), so size the chart by the longer side
  $side = [math]::Ceiling([math]::Max([double]$s.Width, [double]$s.Height) * 1.5) + 60
  $co = $script:wsw.ChartObjects().Add(0, 0, $side, $side)
  $ok = $false
  try {
    $ch = $co.Chart
    $ch.ChartArea.Format.Fill.Visible = -1
    $ch.ChartArea.Format.Fill.Solid()
    $ch.ChartArea.Format.Fill.ForeColor.RGB = $script:BGINT
    $ch.ChartArea.Format.Line.Visible = 0
    # CopyPicture(xlScreen=1, xlBitmap=2) -> Paste into the chart. Chart.Paste can paste nothing without an error, so check Shapes.Count
    foreach ($strategy in @('activate', 'select')) {
      $script:ws.Activate()
      $s.CopyPicture(1, 2)
      Start-Sleep -Milliseconds 150
      $script:wsw.Activate(); $co.Activate()
      if ($strategy -eq 'select') { $ch.ChartArea.Select() | Out-Null }
      $ch.Paste()
      if ($ch.Shapes.Count -gt 0) {
        $ok = $true
        if (-not $script:pasteNoted) { $script:pasteNoted = $true; [void]$script:notes.Add("paste strategy: $strategy") }
        break
      }
    }
    if (-not $ok) { [void]$script:notes.Add("paste failed (all strategies): $name") }
    [void]$ch.Export((Join-Path $script:pngDir ($name + '.png')), 'PNG')
  } finally { $co.Delete() | Out-Null }
  $ok
}

function Add-Export([string]$set, $s) {
  $name = $s.Name
  $hidden = @()
  if ($script:Isolate) {
    $l = $s.Left - 80; $t = $s.Top - 80; $r = $s.Left + $s.Width + 80; $b = $s.Top + $s.Height + 80
    foreach ($o in $script:ws.Shapes) {
      if ($o.Name -eq $name -or -not $o.Visible) { continue }
      if ($o.Left + $o.Width -lt $l -or $o.Left -gt $r -or $o.Top + $o.Height -lt $t -or $o.Top -gt $b) { continue }
      $o.Visible = 0; $hidden += $o.Name
    }
  }
  try {
    Try-Do "export $name" {
      $ok = Export-Png $s $name
      [void]$script:exported.Add([ordered]@{ set = $set; name = $name; pasted = $ok; width = $s.Width; height = $s.Height; isolated = $hidden.Count })
    }
  } finally { foreach ($n in $hidden) { $script:ws.Shapes.Item($n).Visible = -1 } }
}

$xl = New-Object -ComObject Excel.Application
# Visible: CopyPicture + Chart.Paste silently pasted nothing while Excel was hidden
$xl.Visible = $true; $xl.DisplayAlerts = $false
try {
  # ========== Part 1: every shape of ref-geoms.xlsx (read-only; a work sheet is added in memory and never saved) ==========
  $wb = $xl.Workbooks.Open($Geoms, 0, $true)
  $script:ws = $wb.Worksheets.Item(1)
  $script:wsw = $wb.Worksheets.Add([Type]::Missing, $script:ws)
  $script:wsw.Name = 'work'
  $script:ws.Activate()
  Try-Do 'zoom' { $xl.ActiveWindow.Zoom = 100 }
  $script:pngDir = $geomPng
  $names = @(); foreach ($sh in $script:ws.Shapes) { $names += $sh.Name }
  foreach ($n in $names) {
    if ($onlySet.Count -gt 0 -and -not $onlySet.ContainsKey($n)) { continue }
    Add-Export 'geoms' ($script:ws.Shapes.Item($n))
  }
  $wb.Close($false)

  # ========== Part 2: flipped shapes with text (U4) -> ref-flips.xlsx ==========
  $wb = $xl.Workbooks.Add()
  $wb.Styles.Item('Normal').Font.Name = 'Calibri'
  $script:ws = $wb.Worksheets.Item(1)
  $script:ws.Name = 'flips'
  $script:ws.Cells.Item(1, 1).Value2 = 'ref-flips'
  $kinds = @(@('ellipse', 9), @('diamond', 4), @('triangle', 7), @('rightArrow', 33), @('homePlate', 51))
  $variants = @('none', 'flipH', 'flipV', 'flipHV', 'flipVRot30')
  $W = 110; $H = 80
  $recs = New-Object System.Collections.ArrayList
  function New-Variant($kind, $type, $variant, $col, $row, [bool]$withText) {
    $x = 20 + 150 * $col; $y = 40 + 130 * $row
    $nm = $(if ($withText) { "u4_${kind}_$variant" } else { "u4o_${kind}_$variant" })
    $s = $script:ws.Shapes.AddShape($type, $x, $y, $W, $H); $s.Name = $nm
    $s.Fill.Visible = 0; $s.Line.Weight = 1; $s.Line.ForeColor.RGB = 0
    if ($withText) {
      $tf = $s.TextFrame2
      $tf.WordWrap = 0; $tf.VerticalAnchor = 1
      $tf.TextRange.Text = 'FLAG'
      $tf.TextRange.Font.Name = 'Calibri'; $tf.TextRange.Font.Size = 18
      $tf.TextRange.Font.Fill.ForeColor.RGB = 0
      $tf.TextRange.ParagraphFormat.Alignment = 1
    }
    if ($variant -eq 'flipH' -or $variant -eq 'flipHV') { $s.Flip(0) }
    if ($variant -eq 'flipV' -or $variant -eq 'flipHV' -or $variant -eq 'flipVRot30') { $s.Flip(1) }
    if ($variant -eq 'flipVRot30') { $s.Rotation = 30 }
    [void]$script:recs.Add([ordered]@{ name = $nm; kind = $kind; variant = $variant; twin = (-not $withText); autoShapeType = $type
        Left = $s.Left; Top = $s.Top; Width = $s.Width; Height = $s.Height; Rotation = $s.Rotation
        HorizontalFlip = $s.HorizontalFlip; VerticalFlip = $s.VerticalFlip; HasText = (Safe { $s.TextFrame2.HasText }) })
  }
  for ($r = 0; $r -lt $kinds.Count; $r++) {
    for ($c = 0; $c -lt $variants.Count; $c++) { New-Variant $kinds[$r][0] $kinds[$r][1] $variants[$c] $c $r $true }
  }
  for ($r = 0; $r -lt $kinds.Count; $r++) {
    for ($c = 0; $c -lt $variants.Count; $c++) { New-Variant $kinds[$r][0] $kinds[$r][1] $variants[$c] $c ($r + $kinds.Count + 1) $false }
  }
  $wb.SaveAs((Join-Path $Dir 'ref-flips.xlsx'), 51)
  # the work sheet is added AFTER saving (the saved file has only the 'flips' sheet) and the workbook is closed without saving again
  $script:wsw = $wb.Worksheets.Add([Type]::Missing, $script:ws)
  $script:wsw.Name = 'work'
  $script:ws.Activate()
  Try-Do 'zoom' { $xl.ActiveWindow.Zoom = 100 }
  $script:pngDir = $flipPng
  foreach ($rec in $recs) {
    if ($onlySet.Count -gt 0 -and -not $onlySet.ContainsKey($rec.name)) { continue }
    Add-Export 'flips' ($script:ws.Shapes.Item($rec.name))
  }
  $wb.Close($false)

  # ========== record ==========
  $utf8 = New-Object Text.UTF8Encoding($false)
  $flipJson = [ordered]@{ excelVersion = $xl.Version; shapes = $recs } | ConvertTo-Json -Depth 6
  [IO.File]::WriteAllText((Join-Path $Dir 'flips.com.json'), $flipJson, $utf8)
  $expJson = [ordered]@{ excelVersion = $xl.Version; geoms = $Geoms; notes = $notes; shapes = $exported } | ConvertTo-Json -Depth 6
  [IO.File]::WriteAllText((Join-Path $Dir 'export.json'), $expJson, $utf8)
  "ok: $($exported.Count) PNGs; notes=$($notes.Count)"
  $notes
} finally { try { $xl.Quit() } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
