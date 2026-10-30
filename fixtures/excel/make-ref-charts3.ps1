# ISSUE-110: the third COM run for G1 charts (ADR-0031 section 16 "next COM"): ref-charts3.xlsx (Excel 16 COM) and what Excel itself reports for every chart in it, plus PNG exports.
# make-ref-charts.ps1 / make-ref-charts2.ps1 and their ref-charts*.xlsx / *.expected.json are NOT touched (1st and 2nd runs stay as they are). Helper functions below are copies of make-ref-charts2.ps1
# (one addition: ChartArea.RoundedCorners is recorded).
# Writes into $Dir (nothing else is touched):
#   stage create : ref-charts3.xlsx, excel3.json (COM values, pt as returned)
#   stage png    : png\<chart name>.png and png\png-index.ref-charts3.json (sizes, expected places, and flags for make-ref-charts-png.mjs: Corner / Ticks / LabelBand).
#                  With -Ref2Xlsx <ref-charts2.xlsx> the c7sw_* charts of the 2nd file are exported too (read-only; png\png-index.ref-charts2-c7.json) so that the label geometry
#                  (rotation / wrapping) of the 36 existing scenes can be read from the pictures. Excel is made Visible for this stage (Chart.Export does not work hidden).
#                  The clipboard is never used. The pictures are NOT committed (keep -Dir outside the repository).
# Run (do NOT run on CI; needs Excel 16; screen zoom 125% like make-ref-charts.ps1 - COM values are pt and stay pt, converting to px is the test's job):
#   powershell -ExecutionPolicy Bypass -File make-ref-charts3.ps1 -Dir <out dir> [-Png] [-Ref2Xlsx <ref-charts2.xlsx>] [-SkipCreate] [-Only create,png]
#   (PowerShell 5.1, file saved as UTF-8 with BOM). Then: node make-ref-charts3.extract.mjs <out dir>; node make-ref-charts-png.mjs <out dir> --out ref-charts3-png.json;
#   copy ref-charts3.xlsx, ref-charts3.expected.json and ref-charts3-png.json to __fixtures__/excel.
#   Avoid reserved device names (COM1-9, NUL, ...) for output file names: that is why the JSON is excel3.json (not com3.json).
# Stages (default = create, plus png when -Png is given; -SkipCreate removes create; -Only <list> says exactly which stages run):
#   create   builds ref-charts3.xlsx (below) and records every chart.
#   png      opens <Dir>\ref-charts3.xlsx (must exist: made by an earlier create) and exports PNGs of all its charts (+ the c7sw_* of -Ref2Xlsx if given).
#            A chart outside the visible window is exported as an empty file: scroll to its TopLeftCell first, and retry (300 ms wait, up to 4 times) while the file is empty.
# Charts (each is ChartObject.Name = key of the expected values; Shapes.AddChart2(-1, type, ...) = Excel's current default style). Height is 216 pt except where noted.
#   c1sw_*    C1 (KNOWN 3, ISSUE-110 item 7): horizontal bars x chart width sweep. barStacked_neg / barStacked_negbig / barClustered_mil7 at 200..600 pt step 20 (21 each),
#             barClustered_neg / barClustered_mil8 / barStacked100_allneg at 200..600 pt step 40 (11 each). Read: MinimumScale / MaximumScale / MajorUnit, plot inside width.     96
#   tick_*    KNOWN 1 (item 8): the value axis line none/visible x MinorTickMark none/in/out/cross x minor gridlines no/yes (column, MajorTickMark out) = 16;
#             MajorTickMark none with line none and no minor gridlines x in/out/cross = 3; horizontal bar, line none, no minor gridlines x in/out/cross = 3.           22
#             Read: MinorUnit / tick marks (COM) and the tick marks in the PNG (png-reader: valTicks).
#   c7w_*     ISSUE-107: column, 10 categories, name length 8 / 12 / 20 / 20 with spaces x chart width 240 300 360 480 600 720 pt. Read: plot inside height, CategoryAxis.Pos.Height,
#             and in the PNG the label geometry (labels: cluster count and angle).                                                                                      30
#   c4dark_*  ISSUE-109: ChartStyle 41..48 on a clustered column (std, legend on) + 42 / 45 on a line with markers and a pie. Read: chart area / plot area fills, fonts, series fills.     12
#   round_*   ISSUE-109: ChartArea.RoundedCorners true / false x small / std / large, the chart area filled #FFF2CC with a 1.5 pt blue line. Read: the PNG corner (corner.radiusPx). 6
# => 166 charts. Every record = Rec-Chart (same fields as make-ref-charts2.ps1, + ChartArea.RoundedCorners) + Ticks; the category axis band is CategoryAxis.Pos.Height, the plot inside height is PlotArea.Inside.Height.
param([string]$Dir, [string]$Ref2Xlsx = '', [switch]$SkipCreate, [switch]$Png, [string[]]$Only = @())
$ErrorActionPreference = 'Stop'
# relative paths are resolved against the PowerShell location (not the process cwd), then made absolute
if (-not [IO.Path]::IsPathRooted($Dir)) { $Dir = Join-Path (Get-Location).Path $Dir }
$Dir = [IO.Path]::GetFullPath($Dir)
if ($Ref2Xlsx -ne '') {
  if (-not [IO.Path]::IsPathRooted($Ref2Xlsx)) { $Ref2Xlsx = Join-Path (Get-Location).Path $Ref2Xlsx }
  $Ref2Xlsx = [IO.Path]::GetFullPath($Ref2Xlsx)
  if (-not (Test-Path $Ref2Xlsx)) { throw "-Ref2Xlsx not found: $Ref2Xlsx" }
}
# which stages run
$stages = New-Object System.Collections.ArrayList
if ($Only.Count -gt 0) {
  foreach ($o in $Only) { foreach ($p in ($o -split ',')) { $t = $p.Trim().ToLower(); if ($t -ne '') { [void]$stages.Add($t) } } }
  foreach ($t in $stages) { if ($t -notin @('create', 'png')) { throw "unknown stage '$t' (create / png)" } }
} else {
  if (-not $SkipCreate) { [void]$stages.Add('create') }
  if ($Png) { [void]$stages.Add('png') }
}
if ($stages.Count -eq 0) { throw 'nothing to do: give -Png, or do not use -SkipCreate' }
$refPath = Join-Path $Dir 'ref-charts3.xlsx'
if (('png' -in $stages) -and ('create' -notin $stages) -and -not (Test-Path $refPath)) { throw "stage png needs $refPath (run stage create first)" }
New-Item -ItemType Directory -Force -Path $Dir | Out-Null
$pngDir = Join-Path $Dir 'png'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$notes = New-Object System.Collections.ArrayList
$records = New-Object System.Collections.ArrayList
$script:dataSets = @{}
$script:dataRow = 1
$script:place = @{}
$script:wb = $null
$script:wsData = $null

function Safe($sb) { try { & $sb } catch { $null } }
function Try-Do($label, [scriptblock]$sb) { try { & $sb } catch { [void]$script:notes.Add("$label failed: $($_.Exception.Message)") } }
# Excel colour integer (BGR) -> 'RRGGBB'. Negative values (xlColorIndexAutomatic -4105 etc.) are returned as they are
function Hex-Of($v) {
  if ($null -eq $v) { return $null }
  try { $v = [int]$v } catch { return $null }
  if ($v -lt 0) { return $v }
  '{0:X2}{1:X2}{2:X2}' -f ($v -band 255), (($v -shr 8) -band 255), (($v -shr 16) -band 255)
}
# Left / Top / Width / Height of a chart element (pt)
function Pos4($o) { [ordered]@{ Left = Safe { $o.Left }; Top = Safe { $o.Top }; Width = Safe { $o.Width }; Height = Safe { $o.Height } } }
# ChartFormat.Fill (Format.Fill.ForeColor.RGB is the colour Excel resolved, ObjectThemeColor / Brightness say which theme slot)
function Fill-Info($f) {
  [ordered]@{ Visible = Safe { $f.Visible }; Type = Safe { $f.Type }; RGB = Hex-Of (Safe { $f.ForeColor.RGB }); Theme = Safe { $f.ForeColor.ObjectThemeColor }
    Brightness = Safe { $f.ForeColor.Brightness }; Transparency = Safe { $f.Transparency } }
}
# ChartFormat.Line
function Line-Info($l) {
  [ordered]@{ Visible = Safe { $l.Visible }; Weight = Safe { $l.Weight }; RGB = Hex-Of (Safe { $l.ForeColor.RGB }); Theme = Safe { $l.ForeColor.ObjectThemeColor }
    Brightness = Safe { $l.ForeColor.Brightness }; DashStyle = Safe { $l.DashStyle }; Transparency = Safe { $l.Transparency } }
}
# TextFrame2.TextRange.Font (new object model: Format.TextFrame2.TextRange)
function Font2($tr) {
  [ordered]@{ Size = Safe { $tr.Font.Size }; Bold = Safe { $tr.Font.Bold }; Name = Safe { $tr.Font.Name }; RGB = Hex-Of (Safe { $tr.Font.Fill.ForeColor.RGB })
    Theme = Safe { $tr.Font.Fill.ForeColor.ObjectThemeColor }; Brightness = Safe { $tr.Font.Fill.ForeColor.Brightness } }
}
# ChartFont (old object model: .Font; Color is a BGR integer)
function Font1($f) { [ordered]@{ Size = Safe { $f.Size }; Bold = Safe { $f.Bold }; Name = Safe { $f.Name }; RGB = Hex-Of (Safe { $f.Color }) } }

# One axis: scale, label area (Left/Top/Width/Height of the axis object includes the tick labels), tick labels, gridlines, line, title
function Axis-Info($ch, $kind) {
  $ax = Safe { $ch.Axes($kind) }
  if ($null -eq $ax) { return $null }
  $o = [ordered]@{}
  $o.Pos = Pos4 $ax
  $o.MinimumScale = Safe { $ax.MinimumScale }; $o.MaximumScale = Safe { $ax.MaximumScale }
  $o.MajorUnit = Safe { $ax.MajorUnit }; $o.MinorUnit = Safe { $ax.MinorUnit }
  $o.MinimumScaleIsAuto = Safe { $ax.MinimumScaleIsAuto }; $o.MaximumScaleIsAuto = Safe { $ax.MaximumScaleIsAuto }
  $o.MajorUnitIsAuto = Safe { $ax.MajorUnitIsAuto }; $o.MinorUnitIsAuto = Safe { $ax.MinorUnitIsAuto }
  $o.ScaleType = Safe { $ax.ScaleType }; $o.ReversePlotOrder = Safe { $ax.ReversePlotOrder }; $o.Crosses = Safe { $ax.Crosses }
  $o.AxisBetweenCategories = Safe { $ax.AxisBetweenCategories }
  $o.TickLabelSpacing = Safe { $ax.TickLabelSpacing }; $o.TickLabelSpacingIsAuto = Safe { $ax.TickLabelSpacingIsAuto }
  $o.TickMarkSpacing = Safe { $ax.TickMarkSpacing }
  $o.TickLabelPosition = Safe { $ax.TickLabelPosition }; $o.MajorTickMark = Safe { $ax.MajorTickMark }; $o.MinorTickMark = Safe { $ax.MinorTickMark }
  $o.TickLabels = [ordered]@{ NumberFormat = Safe { $ax.TickLabels.NumberFormat }; NumberFormatLinked = Safe { $ax.TickLabels.NumberFormatLinked }
    Orientation = Safe { $ax.TickLabels.Orientation }; Offset = Safe { $ax.TickLabels.Offset }; Font = Font1 (Safe { $ax.TickLabels.Font }) }
  $o.Line = Line-Info (Safe { $ax.Format.Line })
  $o.HasMajorGridlines = Safe { $ax.HasMajorGridlines }; $o.HasMinorGridlines = Safe { $ax.HasMinorGridlines }
  if ($o.HasMajorGridlines) { $o.MajorGridlines = Line-Info (Safe { $ax.MajorGridlines.Format.Line }) }
  if ($o.HasMinorGridlines) { $o.MinorGridlines = Line-Info (Safe { $ax.MinorGridlines.Format.Line }) }
  $o.HasTitle = Safe { $ax.HasTitle }
  if ($o.HasTitle) { $o.Title = [ordered]@{ Text = Safe { $ax.AxisTitle.Text }; Pos = Pos4 (Safe { $ax.AxisTitle }); Orientation = Safe { $ax.AxisTitle.Orientation }
    Font = Font2 (Safe { $ax.AxisTitle.Format.TextFrame2.TextRange }) } }
  $o
}

# One series: name, values, fill / line / marker, and (pie, doughnut, varyColors, -Points) every point
function Series-Info($s, $withPoints) {
  $o = [ordered]@{}
  $o.Name = Safe { $s.Name }
  $o['Values'] = Safe { @($s.Values) }
  $o.ChartType = Safe { $s.ChartType }; $o.AxisGroup = Safe { $s.AxisGroup }; $o.PlotOrder = Safe { $s.PlotOrder }
  $o.Fill = Fill-Info (Safe { $s.Format.Fill }); $o.Line = Line-Info (Safe { $s.Format.Line })
  $o.MarkerStyle = Safe { $s.MarkerStyle }; $o.MarkerSize = Safe { $s.MarkerSize }
  $o.MarkerBackgroundColor = Hex-Of (Safe { $s.MarkerBackgroundColor }); $o.MarkerForegroundColor = Hex-Of (Safe { $s.MarkerForegroundColor })
  $o.Smooth = Safe { $s.Smooth }; $o.InvertIfNegative = Safe { $s.InvertIfNegative }; $o.Explosion = Safe { $s.Explosion }
  if ($withPoints) {
    $n = Safe { $s.Points().Count }
    $pts = New-Object System.Collections.ArrayList
    for ($j = 1; $j -le $n; $j++) {
      $p = Safe { $s.Points($j) }
      [void]$pts.Add([ordered]@{ Fill = Fill-Info (Safe { $p.Format.Fill }); Line = Line-Info (Safe { $p.Format.Line }); MarkerStyle = Safe { $p.MarkerStyle }
        MarkerSize = Safe { $p.MarkerSize }; MarkerBackgroundColor = Hex-Of (Safe { $p.MarkerBackgroundColor }); Explosion = Safe { $p.Explosion } })
    }
    $o.Points = $pts
  }
  $o
}

# Everything ADR-0031 needs from one ChartObject. $meta = what the script knew when it built the chart (group / type / data / size ...)
function Rec-Chart($co, $meta, [bool]$withPoints) {
  $ch = $co.Chart
  $o = [ordered]@{}
  $o.Name = $co.Name
  foreach ($k in $meta.Keys) { $o[$k] = $meta[$k] }
  $o.ChartType = Safe { $ch.ChartType }; $o.ChartStyle = Safe { $ch.ChartStyle }; $o.ChartColor = Safe { $ch.ChartColor }
  $o.ChartObject = Pos4 $co
  $o.ChartArea = [ordered]@{ Pos = Pos4 (Safe { $ch.ChartArea }); Fill = Fill-Info (Safe { $ch.ChartArea.Format.Fill }); Line = Line-Info (Safe { $ch.ChartArea.Format.Line })
    Font = Font1 (Safe { $ch.ChartArea.Font }); RoundedCorners = Safe { $ch.ChartArea.RoundedCorners } }
  $o.PlotArea = [ordered]@{ Pos = Pos4 (Safe { $ch.PlotArea })
    Inside = [ordered]@{ Left = Safe { $ch.PlotArea.InsideLeft }; Top = Safe { $ch.PlotArea.InsideTop }; Width = Safe { $ch.PlotArea.InsideWidth }; Height = Safe { $ch.PlotArea.InsideHeight } }
    Fill = Fill-Info (Safe { $ch.PlotArea.Format.Fill }); Line = Line-Info (Safe { $ch.PlotArea.Format.Line }); Position = Safe { $ch.PlotArea.Position } }
  $o.HasTitle = Safe { $ch.HasTitle }
  if ($o.HasTitle) {
    $o.Title = [ordered]@{ Text = Safe { $ch.ChartTitle.Text }; Pos = Pos4 (Safe { $ch.ChartTitle }); IncludeInLayout = Safe { $ch.ChartTitle.IncludeInLayout }
      Position = Safe { $ch.ChartTitle.Position }; Font = Font2 (Safe { $ch.ChartTitle.Format.TextFrame2.TextRange }) }
  }
  $o.HasLegend = Safe { $ch.HasLegend }
  if ($o.HasLegend) {
    $lg = Safe { $ch.Legend }
    $entries = New-Object System.Collections.ArrayList
    $n = Safe { $lg.LegendEntries().Count }
    for ($i = 1; $i -le $n; $i++) {
      $e = Safe { $lg.LegendEntries($i) }
      [void]$entries.Add([ordered]@{ Pos = Pos4 $e; KeyPos = Pos4 (Safe { $e.LegendKey }); KeyFillRGB = Hex-Of (Safe { $e.LegendKey.Format.Fill.ForeColor.RGB })
        KeyLineRGB = Hex-Of (Safe { $e.LegendKey.Format.Line.ForeColor.RGB }); KeyMarkerStyle = Safe { $e.LegendKey.MarkerStyle } })
    }
    $o.Legend = [ordered]@{ Position = Safe { $lg.Position }; IncludeInLayout = Safe { $lg.IncludeInLayout }; Pos = Pos4 $lg; Font = Font1 (Safe { $lg.Font })
      Fill = Fill-Info (Safe { $lg.Format.Fill }); Line = Line-Info (Safe { $lg.Format.Line }); Entries = $entries }
  }
  $o.CategoryAxis = Axis-Info $ch 1
  $o.ValueAxis = Axis-Info $ch 2
  $g = Safe { $ch.ChartGroups(1) }
  $o.Group = [ordered]@{ VaryByCategories = Safe { $g.VaryByCategories }; GapWidth = Safe { $g.GapWidth }; Overlap = Safe { $g.Overlap }
    DoughnutHoleSize = Safe { $g.DoughnutHoleSize }; FirstSliceAngle = Safe { $g.FirstSliceAngle }; HasSeriesLines = Safe { $g.HasSeriesLines } }
  $o.SeriesCount = Safe { $ch.SeriesCollection().Count }
  $ser = New-Object System.Collections.ArrayList
  $isRound = ($o.ChartType -eq 5) -or ($o.ChartType -eq -4120)
  $pts = $withPoints -or $isRound -or ($o.Group.VaryByCategories -eq $true)
  for ($i = 1; $i -le $o.SeriesCount; $i++) { [void]$ser.Add((Series-Info (Safe { $ch.SeriesCollection($i) }) $pts)) }
  $o.Series = $ser
  # ISSUE-105 (C12 / C13): the tick marks and markers in one place, so a check does not have to walk the axes and the series
  $cat = $o.CategoryAxis; $val = $o.ValueAxis
  $o.Ticks = [ordered]@{ CatMajor = $(if ($cat) { $cat.MajorTickMark } else { $null }); CatMinor = $(if ($cat) { $cat.MinorTickMark } else { $null })
    ValMajor = $(if ($val) { $val.MajorTickMark } else { $null }); ValMinor = $(if ($val) { $val.MinorTickMark } else { $null })
    MarkerStyles = @(foreach ($s in $ser) { $s.MarkerStyle }); MarkerSizes = @(foreach ($s in $ser) { $s.MarkerSize }) }
  $o
}

# ---- data on the sheet 'Data': A = categories, B.. = series (header row = names). One block per data set ----
function Put-Data($key, $cats, $matrix) {
  $r0 = $script:dataRow
  $ws = $script:wsData
  $ws.Cells($r0, 1).Value2 = $key
  for ($s = 0; $s -lt $matrix.Count; $s++) { $ws.Cells($r0, 2 + $s).Value2 = ('S' + ($s + 1)) }
  for ($c = 0; $c -lt $cats.Count; $c++) {
    $ws.Cells($r0 + 1 + $c, 1).Value2 = $cats[$c]
    for ($s = 0; $s -lt $matrix.Count; $s++) { $ws.Cells($r0 + 1 + $c, 2 + $s).Value2 = [double]$matrix[$s][$c] }
  }
  $script:dataSets[$key] = @{ Row = $r0; Cats = $cats.Count; Series = $matrix.Count }
  $script:dataRow = $r0 + $cats.Count + 3
}
# categories + the first $nser series of a data set
function Get-Range($key, [int]$nser) {
  $d = $script:dataSets[$key]
  , $script:wsData.Range($script:wsData.Cells($d.Row, 1), $script:wsData.Cells($d.Row + $d.Cats, 1 + $nser))
}
function Add-ChartSheet($name) {
  $ws = $script:wb.Worksheets.Add([Type]::Missing, $script:wb.Worksheets.Item($script:wb.Worksheets.Count))
  $ws.Name = $name
  $script:place[$name] = @{ X = 10; Y = 10; RowH = 0 }
  $ws
}
# next free spot on the chart sheet: charts side by side up to 2600 pt wide, then a new row
function Next-Pos($ws, $w, $h) {
  $p = $script:place[$ws.Name]
  if (($p.X + $w) -gt 2600) { $p.X = 10; $p.Y = $p.Y + $p.RowH + 10; $p.RowH = 0 }
  $pos = , @($p.X, $p.Y)
  $p.X = $p.X + $w + 10
  if ($h -gt $p.RowH) { $p.RowH = $h }
  $pos
}
# Create one chart, run $tweak (optional) and record it. $nser = how many series of the data set; title = 'Title' unless $setTitle is $false
function New-Chart($ws, $name, [int]$xlType, $dataKey, [int]$nser, $w, $h, $meta, [scriptblock]$tweak, [bool]$withPoints = $false, [bool]$setTitle = $true) {
  $p = Next-Pos $ws $w $h
  $ws.Activate()
  $shape = $ws.Shapes.AddChart2(-1, $xlType, $p[0], $p[1], $w, $h)
  $shape.Name = $name
  $ch = $shape.Chart
  # AddChart2 may have taken a selection as its data: start from no series, then give exactly the data we mean (xlColumns = 2)
  while ($ch.SeriesCollection().Count -gt 0) { [void]$ch.SeriesCollection(1).Delete() }
  $ch.SetSourceData((Get-Range $dataKey $nser), 2)
  $ch.ChartType = $xlType
  if ($setTitle) { $ch.HasTitle = $true; $ch.ChartTitle.Text = 'Title' }
  if ($null -ne $tweak) { Try-Do "tweak $name" $tweak }
  $co = $ws.ChartObjects($name)
  $m = [ordered]@{ Sheet = $ws.Name; Data = $dataKey; SeriesAsked = $nser; WidthAsked = $w; HeightAsked = $h }
  foreach ($k in $meta.Keys) { $m[$k] = $meta[$k] }
  [void]$script:records.Add((Rec-Chart $co $m $withPoints))
}

# category names for the C7 sweep: 'C01abcdefg...' (no spaces) or 'C01 alpha beta ...' (spaces: Excel can wrap), cut to $len characters
function Cat-Name([int]$i, [int]$len, [bool]$spaced) {
  $head = 'C' + $i.ToString('00')
  if ($spaced) { $s = $head + ' alpha beta gamma delta epsilon' } else { $s = $head + 'abcdefghijklmnopqrstuvwxyz' }
  $s.Substring(0, $len)
}
# delete legend entry $i (1-based, in the order Excel lists them)
function Del-Entry($ch, [int]$i) { [void]$ch.Legend.LegendEntries($i).Delete() }

# ---- PNG: every chart of the workbook at $path whose name matches $filter -> png\<name>.png, and an index (sizes, expected places, flags) for make-ref-charts-png.mjs ----
# Flags: Corner (round_*: read the corner of the chart area), Ticks (tick_*: read the value axis tick marks), LabelBand (c7*: CategoryAxis.Pos, read the label geometry)
function Export-Pngs([string]$path, [string]$filter, [string]$tag) {
  New-Item -ItemType Directory -Force -Path $pngDir | Out-Null
  $idx = New-Object System.Collections.ArrayList
  $pnotes = New-Object System.Collections.ArrayList
  $pwb = $xl.Workbooks.Open($path, 0, $true)
  try {
    foreach ($pws in $pwb.Worksheets) {
      $pws.Activate()
      foreach ($co in $pws.ChartObjects()) {
        $n = $co.Name
        if ($n -notmatch $filter) { continue }
        try {
          $ch = $co.Chart
          $file = Join-Path $pngDir ($n + '.png')
          # a chart outside the visible window is exported as an empty file: scroll it into view first, then retry while the file is empty
          $tl = $co.TopLeftCell
          $xl.ActiveWindow.ScrollRow = $tl.Row
          $xl.ActiveWindow.ScrollColumn = $tl.Column
          $ok = $false
          for ($try = 0; $try -lt 4; $try++) {
            $ok = $ch.Export($file, 'PNG')
            if ((Test-Path $file) -and ((Get-Item $file).Length -gt 0)) { break }
            Start-Sleep -Milliseconds 300
          }
          if (-not ((Test-Path $file) -and ((Get-Item $file).Length -gt 0))) { $ok = $false }
          $ct = Safe { $ch.ChartType }
          $round = ($ct -eq 5) -or ($ct -eq -4120)
          # colours that must not be mistaken for a gridline / axis line: series fills, series lines, point fills
          $ex = New-Object System.Collections.ArrayList
          $sc = Safe { $ch.SeriesCollection().Count }
          for ($i = 1; $i -le $sc; $i++) {
            $s = Safe { $ch.SeriesCollection($i) }
            foreach ($v in @((Hex-Of (Safe { $s.Format.Fill.ForeColor.RGB })), (Hex-Of (Safe { $s.Format.Line.ForeColor.RGB })))) { if ($v -is [string]) { [void]$ex.Add($v) } }
            if ($round) {
              $np = Safe { $s.Points().Count }
              for ($j = 1; $j -le $np; $j++) { $v = Hex-Of (Safe { $s.Points($j).Format.Fill.ForeColor.RGB }); if ($v -is [string]) { [void]$ex.Add($v) } }
            }
          }
          $va = Safe { $ch.Axes(2) }
          $ca = Safe { $ch.Axes(1) }
          $band = $null
          if (($n -match '^c7') -and $ca) { $band = Pos4 $ca }
          [void]$idx.Add([ordered]@{ Name = $n; Source = $tag; File = ($n + '.png'); Exported = $ok; ChartType = $ct; ChartStyle = Safe { $ch.ChartStyle }
            WidthPt = Safe { $co.Width }; HeightPt = Safe { $co.Height }
            PlotInside = [ordered]@{ Left = Safe { $ch.PlotArea.InsideLeft }; Top = Safe { $ch.PlotArea.InsideTop }; Width = Safe { $ch.PlotArea.InsideWidth }; Height = Safe { $ch.PlotArea.InsideHeight } }
            Horizontal = ($ct -in @(57, 58, 59)); HasAxes = (-not $round)
            ValueAxis = $(if ($va) { [ordered]@{ Min = Safe { $va.MinimumScale }; Max = Safe { $va.MaximumScale }; Major = Safe { $va.MajorUnit }; Minor = Safe { $va.MinorUnit }
              HasMajorGrid = Safe { $va.HasMajorGridlines }; HasMinorGrid = Safe { $va.HasMinorGridlines }; Crosses = Safe { $va.Crosses }
              MajorTickMark = Safe { $va.MajorTickMark }; MinorTickMark = Safe { $va.MinorTickMark }; LineVisible = Safe { $va.Format.Line.Visible } } } else { $null })
            CategoryAxisCrosses = $(if ($ca) { Safe { $ca.Crosses } } else { $null })
            Corner = ($n -like 'round_*'); Ticks = ($n -like 'tick_*'); LabelBand = $band
            ExcludeRGB = @($ex | Select-Object -Unique) })
        } catch { [void]$pnotes.Add("png $n failed: $($_.Exception.Message)") }
      }
    }
  } finally { $pwb.Close($false) }
  $out = [ordered]@{ excelVersion = $xl.Version; source = $tag; note = 'WidthPt / HeightPt = ChartObject (pt). PlotInside = COM values (pt, origin 4pt inside the chart area). Colours RRGGBB.'
    notes = @($pnotes); charts = @($idx) }
  $json = ConvertTo-Json -InputObject $out -Depth 8
  [IO.File]::WriteAllText((Join-Path $pngDir "png-index.$tag.json"), $json, (New-Object Text.UTF8Encoding($false)))
  "ok: $($idx.Count) png ($tag); notes=$($pnotes.Count)"
}

try {
  if ('create' -in $stages) {
    $script:wb = $xl.Workbooks.Add()
    $script:wsData = $script:wb.Worksheets.Item(1); $script:wsData.Name = 'Data'
    $cats5 = @('Cat1', 'Cat2', 'Cat3', 'Cat4', 'Cat5')
    # the same data as make-ref-charts2.ps1 (pos / neg / negbig / allneg / mil7 / mil8)
    Put-Data 'pos' $cats5 @(@(12, 19, 7, 15, 9), @(8, 14, 11, 6, 17), @(5, 9, 13, 10, 4))
    Put-Data 'neg' $cats5 @(@(12, -5, 7, -15, 9), @(-8, 14, -11, 6, -3), @(5, -9, 13, -10, 4))
    Put-Data 'negbig' $cats5 @(@(12000, -5000, 7000, -15000, 9000), @(-8000, 14000, -11000, 6000, -3000), @(5000, -9000, 13000, -10000, 4000))
    Put-Data 'allneg' $cats5 @(@(-12, -5, -7, -15, -9), @(-8, -14, -11, -6, -3), @(-5, -9, -13, -10, -4))
    Put-Data 'mil7' $cats5 @(@(125000, 480000, 310000, 950000, 670000), @(83000, 210000, 540000, 400000, 760000), @(155000, 98000, 260000, 330000, 120000))
    Put-Data 'mil8' $cats5 @(@(1250000, 4800000, 3100000, 9500000, 6700000), @(830000, 2100000, 5400000, 4000000, 7600000), @(1550000, 980000, 2600000, 3300000, 1200000))
    # c7w data: w_10_<length>[s]  (10 categories; s = with spaces; same naming rule and values as the c7sw sweep of make-ref-charts2.ps1)
    foreach ($len in @(8, 12, 20)) {
      foreach ($sp in @($false, $true)) {
        if ($sp -and $len -lt 12) { continue }
        $cs = @(); $vs = @()
        for ($c = 1; $c -le 10; $c++) { $cs += (Cat-Name $c $len $sp); $vs += ((($c * 7) % 13) + 3) }
        Put-Data ("w_10_{0}{1}" -f $len, $(if ($sp) { 's' } else { '' })) $cs @(, $vs)
      }
    }

    $size = [ordered]@{ small = @(216, 144); std = @(360, 216); large = @(720, 432) }
    function Ex($name, $xlType, $dataKey, $nser, $sizeKey, $tweak, $meta, $withPoints = $false, $setTitle = $true) {
      if ($sizeKey -is [string]) { $wh = $size[$sizeKey] } else { $wh = $sizeKey }
      New-Chart $script:ws $name $xlType $dataKey $nser $wh[0] $wh[1] $meta $tweak $withPoints $setTitle
    }

    # ========== c1sw (C1 KNOWN 3): horizontal bars x chart width sweep ==========
    $script:ws = Add-ChartSheet 'c1sw'
    $w20 = @(); for ($w = 200; $w -le 600; $w += 20) { $w20 += $w }
    $w40 = @(); for ($w = 200; $w -le 600; $w += 40) { $w40 += $w }
    $c1sw = @(
      @('barStacked_neg', 58, 'neg', $w20), @('barStacked_negbig', 58, 'negbig', $w20), @('barClustered_mil7', 57, 'mil7', $w20),
      @('barClustered_neg', 57, 'neg', $w40), @('barClustered_mil8', 57, 'mil8', $w40), @('barStacked100_allneg', 59, 'allneg', $w40))
    foreach ($def in $c1sw) {
      foreach ($w in $def[3]) {
        Ex ("c1sw_{0}_{1}" -f $def[0], $w) $def[1] $def[2] 3 @($w, 216) $null @{ Kind = 'c1sw'; Check = 'C1'; Type = $def[0]; DataSet = $def[2]; ChartWidth = $w }
      }
    }

    # ========== tick (KNOWN 1): value axis line none / visible x MinorTickMark x minor gridlines ==========
    $script:ws = Add-ChartSheet 'tick'
    $marks = [ordered]@{ none = -4142; 'in' = 2; out = 3; cross = 4 }
    $tickTweak = {
      $a = $ch.Axes(2)
      $a.MajorTickMark = $script:tmaj; $a.MinorTickMark = $script:tmin; $a.HasMinorGridlines = $script:tgrid
      if ($script:tline) { $a.Format.Line.Visible = -1; $a.Format.Line.Weight = 0.75 } else { $a.Format.Line.Visible = 0 }
    }
    foreach ($lineOn in @($false, $true)) {
      foreach ($mk in $marks.Keys) {
        foreach ($grid in @($false, $true)) {
          $script:tmaj = 3; $script:tmin = $marks[$mk]; $script:tgrid = $grid; $script:tline = $lineOn
          $nm = "tick_col_{0}_{1}_{2}_std" -f $(if ($lineOn) { 'lineOn' } else { 'lineNone' }), $mk, $(if ($grid) { 'grid' } else { 'nogrid' })
          Ex $nm 51 'pos' 3 'std' $tickTweak @{ Kind = 'tick'; Check = 'MinorTick'; Type = 'colClustered'; LineVisibleAsked = $lineOn; MinorMarkAsked = $mk; MinorGridAsked = $grid; MajorMarkAsked = 'out' }
        }
      }
    }
    foreach ($mk in @('in', 'out', 'cross')) {
      $script:tmaj = -4142; $script:tmin = $marks[$mk]; $script:tgrid = $false; $script:tline = $false
      Ex ("tick_col_lineNone_{0}_nogrid_majorNone_std" -f $mk) 51 'pos' 3 'std' $tickTweak @{ Kind = 'tick'; Check = 'MinorTick'; Type = 'colClustered'; LineVisibleAsked = $false; MinorMarkAsked = $mk; MinorGridAsked = $false; MajorMarkAsked = 'none' }
    }
    foreach ($mk in @('in', 'out', 'cross')) {
      $script:tmaj = 3; $script:tmin = $marks[$mk]; $script:tgrid = $false; $script:tline = $false
      Ex ("tick_bar_lineNone_{0}_nogrid_std" -f $mk) 57 'pos' 3 'std' $tickTweak @{ Kind = 'tick'; Check = 'MinorTick'; Type = 'barClustered'; LineVisibleAsked = $false; MinorMarkAsked = $mk; MinorGridAsked = $false; MajorMarkAsked = 'out' }
    }

    # ========== c7w (ISSUE-107): category axis labels x chart width (10 categories, name length 8 / 12 / 20 / 20 with spaces) ==========
    $script:ws = Add-ChartSheet 'c7w'
    foreach ($len in @(8, 12, 20)) {
      foreach ($sp in @($false, $true)) {
        if ($sp -and $len -lt 12) { continue }
        $suf = $(if ($sp) { 's' } else { '' })
        foreach ($w in @(240, 300, 360, 480, 600, 720)) {
          Ex ("c7w_col_{0}{1}_{2}" -f $len, $suf, $w) 51 ("w_10_{0}{1}" -f $len, $suf) 1 @($w, 216) $null @{ Kind = 'c7w'; Check = 'C7'; Type = 'colClustered'; Count = 10; NameLen = $len; Spaces = $sp; ChartWidth = $w }
        }
      }
    }

    # ========== c4dark (ISSUE-109): ChartStyle 41..48 ==========
    $script:ws = Add-ChartSheet 'c4dark'
    foreach ($sn in 41..48) {
      $script:sn = $sn
      Ex ("c4dark_{0}_colClustered_std" -f $sn) 51 'pos' 3 'std' { $ch.ChartStyle = [int]$script:sn; $ch.HasLegend = $true } @{ Kind = 'c4dark'; Check = 'C4'; Type = 'colClustered'; StyleAsked = $sn }
    }
    foreach ($sn in @(42, 45)) {
      $script:sn = $sn
      Ex ("c4dark_{0}_lineMarkers_std" -f $sn) 65 'pos' 3 'std' { $ch.ChartStyle = [int]$script:sn; $ch.HasLegend = $true } @{ Kind = 'c4dark'; Check = 'C4'; Type = 'lineMarkers'; StyleAsked = $sn }
      Ex ("c4dark_{0}_pie_std" -f $sn) 5 'pos' 1 'std' { $ch.ChartStyle = [int]$script:sn; $ch.HasLegend = $true } @{ Kind = 'c4dark'; Check = 'C4'; Type = 'pie'; StyleAsked = $sn } $true
    }

    # ========== round (ISSUE-109): ChartArea.RoundedCorners x size; the chart area is yellow with a blue 1.5 pt line so that the corner is easy to read in the PNG ==========
    $script:ws = Add-ChartSheet 'round'
    foreach ($rc in @($true, $false)) {
      foreach ($sk in $size.Keys) {
        $script:rc = $rc
        Ex ("round_{0}_{1}" -f $(if ($rc) { 'on' } else { 'off' }), $sk) 51 'pos' 3 $sk {
          $ch.ChartArea.RoundedCorners = $script:rc
          $f = $ch.ChartArea.Format.Fill; $f.Visible = -1; $f.Solid(); $f.ForeColor.RGB = 13431551
          $l = $ch.ChartArea.Format.Line; $l.Visible = -1; $l.ForeColor.RGB = 16711680; $l.Weight = 1.5
        } @{ Kind = 'round'; Check = 'C4'; Type = 'colClustered'; Size = $sk; RoundedAsked = $rc }
      }
    }

    $script:wb.SaveAs($refPath, 51); $script:wb.Close($false); $script:wb = $null
    $out = [ordered]@{ excelVersion = $xl.Version; unit = 'pt'; note = 'COM values as returned (pt). Screen zoom 125%: 0.6 pt = 1 px. Colours are RRGGBB (negative = Excel colour index constant).'
      notes = $notes; charts = $records }
    $json = $out | ConvertTo-Json -Depth 12
    [IO.File]::WriteAllText((Join-Path $Dir 'excel3.json'), $json, (New-Object Text.UTF8Encoding($false)))
    "ok: $($records.Count) charts; notes=$($notes.Count)"
  }

  # ========== PNG (Excel must be Visible for Chart.Export) ==========
  if ('png' -in $stages) {
    $xl.Visible = $true
    if (Test-Path $refPath) { Export-Pngs $refPath '.*' 'ref-charts3' }
    if ($Ref2Xlsx -ne '') { Export-Pngs $Ref2Xlsx '^c7sw_' 'ref-charts2-c7' }
  }
} finally {
  try { if ($null -ne $script:wb) { $script:wb.Close($false) } } catch {}
  try { $xl.Quit() } catch {}
  [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
