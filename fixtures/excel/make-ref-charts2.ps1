# ISSUE-105: the second COM run for G1 charts (ADR-0031 section 16 "next COM"): ref-charts2.xlsx (Excel 16 COM) and what Excel itself reports for every chart in it,
# plus the C12 / C13 / C4(c:style) charts of ref-charts-default.xlsx (made by make-ref-charts-default.ts) and PNG exports.
# make-ref-charts.ps1 / ref-charts.xlsx / ref-charts.expected.json are NOT touched (the first run stays as it is). Helper functions below are copies of make-ref-charts.ps1.
# Writes into $Dir (nothing else is touched):
#   stage create : ref-charts2.xlsx, excel2.json (COM values, pt as returned)
#   stage default: excel2.default.json (every chart of -DefaultXlsx, read-only; each record has a 'Ticks' block = MajorTickMark / MinorTickMark / MarkerStyle / MarkerSize)
#   stage png    : png\<chart name>.png and png\png-index.ref-charts2.json / png\png-index.default.json (sizes and expected places for make-ref-charts-png.mjs). Excel is made Visible
#                  for this stage (Chart.Export does not work hidden). The clipboard is never used. The pictures are NOT committed (keep -Dir outside the repository).
# Run (do NOT run on CI; needs Excel 16; screen zoom 125% like make-ref-charts.ps1 - COM values are pt and stay pt, converting to px is the test's job):
#   powershell -ExecutionPolicy Bypass -File make-ref-charts2.ps1 -Dir <out dir> [-DefaultXlsx <ref-charts-default.xlsx>] [-SkipCreate] [-Png] [-Only create,default,png]
#   (PowerShell 5.1, file saved as UTF-8 with BOM). Then: node make-ref-charts2.extract.mjs <out dir>; node make-ref-charts-png.mjs <out dir>;
#   copy ref-charts2.xlsx, ref-charts2.expected.json, ref-charts-png.json and (regenerated) ref-charts-default.xlsx to __fixtures__/excel.
# Stages (default = create, plus default when -DefaultXlsx is given, plus png when -Png is given; -SkipCreate removes create; -Only <list> says exactly which stages run):
#   create   builds ref-charts2.xlsx (below) and records every chart.
#   default  opens -DefaultXlsx read-only and records every chart (C12 / C13 / C4 / the first 19 again).
#   png      opens <Dir>\ref-charts2.xlsx (must exist: made by an earlier create) and -DefaultXlsx (if given) read-only and exports PNGs:
#            ref-charts2.xlsx only the scenes c3vary* / c4style* / c5corner* / c6row13* / c7sw* / c11_*, the default file all of its charts.
# Charts (each is ChartObject.Name = key of the expected values; Shapes.AddChart2(-1, type, ...) = Excel's current default style). Sizes (pt): small 216x144, std 360x216, large 720x432.
#   c1neg    C1  : horizontal bars (clustered / stacked) with negative values, data neg (+ negbig = x1000), 3 sizes                      12
#   c2allneg C2  : 100% stacked column / bar, all values negative, 3 sizes                                                              6
#   c3vary*  C3,C6: pie / doughnut with VaryByCategories = false (and a true baseline), legend entries (count, boxes)                      8
#   c1mil    C1  : tick labels of 7 and 8 digits: clustered bar x 3 sizes, clustered column at std                                        8
#   c4style  C4  : ChartStyle 1 2 3 10 26 42 (column, std) + style 3 on line-with-markers and pie                                        8
#   c7sw     C7  : category axis sweep (5/10/15/20/30 categories x name length 4/8/12/20, no spaces; 12 and 20 also with spaces), column std;
#                  plus 6 horizontal bars (15 and 30 categories x length 8/20)                                                         36
#   c5corner C5  : legend at the corner (xlLegendPositionCorner 2): column x 3 sizes, bar / line / pie at std                           6
#   c6row13  C6  : 13 legend entries in a row (pie / doughnut x legend top / bottom x widths 200 280 360 480 720 pt, height 216)        20
#   minor_*  MinorUnit: minor gridlines / minor tick marks / explicit MinorUnit; MinorUnit and MajorUnit are recorded                    11
#   c11_*    C11 : LegendEntries(i).Delete and Series.PlotOrder changes; the remaining entries (positions, key colours) are recorded    12
# => 127 charts. Every record = Rec-Chart (same fields as make-ref-charts.ps1) + Ticks; the category axis band is CategoryAxis.Pos.Height, the plot inside height is PlotArea.Inside.Height.
param([string]$Dir, [string]$DefaultXlsx = '', [switch]$SkipCreate, [switch]$Png, [string[]]$Only = @())
$ErrorActionPreference = 'Stop'
# relative paths are resolved against the PowerShell location (not the process cwd), then made absolute
if (-not [IO.Path]::IsPathRooted($Dir)) { $Dir = Join-Path (Get-Location).Path $Dir }
$Dir = [IO.Path]::GetFullPath($Dir)
if ($DefaultXlsx -ne '') {
  if (-not [IO.Path]::IsPathRooted($DefaultXlsx)) { $DefaultXlsx = Join-Path (Get-Location).Path $DefaultXlsx }
  $DefaultXlsx = [IO.Path]::GetFullPath($DefaultXlsx)
}
# which stages run
$stages = New-Object System.Collections.ArrayList
if ($Only.Count -gt 0) {
  foreach ($o in $Only) { foreach ($p in ($o -split ',')) { $t = $p.Trim().ToLower(); if ($t -ne '') { [void]$stages.Add($t) } } }
  foreach ($t in $stages) { if ($t -notin @('create', 'default', 'png')) { throw "unknown stage '$t' (create / default / png)" } }
} else {
  if (-not $SkipCreate) { [void]$stages.Add('create') }
  if ($DefaultXlsx -ne '') { [void]$stages.Add('default') }
  if ($Png) { [void]$stages.Add('png') }
}
if ($stages.Count -eq 0) { throw 'nothing to do: give -DefaultXlsx / -Png, or do not use -SkipCreate' }
if (('default' -in $stages) -and $DefaultXlsx -eq '') { throw 'stage default needs -DefaultXlsx' }
$refPath = Join-Path $Dir 'ref-charts2.xlsx'
if (('png' -in $stages) -and ('create' -notin $stages) -and ($DefaultXlsx -eq '') -and -not (Test-Path $refPath)) { throw "stage png needs $refPath (run stage create first) or -DefaultXlsx" }
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
    Font = Font1 (Safe { $ch.ChartArea.Font }) }
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

# ---- PNG: every chart that matches $filter -> png\<name>.png, and an index (sizes, expected places) for make-ref-charts-png.mjs ----
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
          [void]$idx.Add([ordered]@{ Name = $n; Source = $tag; File = ($n + '.png'); Exported = $ok; ChartType = $ct; ChartStyle = Safe { $ch.ChartStyle }
            WidthPt = Safe { $co.Width }; HeightPt = Safe { $co.Height }
            PlotInside = [ordered]@{ Left = Safe { $ch.PlotArea.InsideLeft }; Top = Safe { $ch.PlotArea.InsideTop }; Width = Safe { $ch.PlotArea.InsideWidth }; Height = Safe { $ch.PlotArea.InsideHeight } }
            Horizontal = ($ct -in @(57, 58, 59)); HasAxes = (-not $round)
            ValueAxis = $(if ($va) { [ordered]@{ Min = Safe { $va.MinimumScale }; Max = Safe { $va.MaximumScale }; Major = Safe { $va.MajorUnit }; Minor = Safe { $va.MinorUnit }
              HasMajorGrid = Safe { $va.HasMajorGridlines }; HasMinorGrid = Safe { $va.HasMinorGridlines }; Crosses = Safe { $va.Crosses } } } else { $null })
            CategoryAxisCrosses = $(if ($ca) { Safe { $ca.Crosses } } else { $null })
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
    # the same data as make-ref-charts.ps1 (pos / neg / big), plus negbig (neg x1000), allneg (C2), mil7 / mil8 (7 and 8 digit tick labels), 13 series / 13 points
    Put-Data 'pos' $cats5 @(@(12, 19, 7, 15, 9), @(8, 14, 11, 6, 17), @(5, 9, 13, 10, 4))
    Put-Data 'neg' $cats5 @(@(12, -5, 7, -15, 9), @(-8, 14, -11, 6, -3), @(5, -9, 13, -10, 4))
    Put-Data 'dec' $cats5 @(@(0.12, 0.35, 0.27, 0.48, 0.31), @(0.08, 0.22, 0.41, 0.16, 0.37), @(0.05, 0.18, 0.09, 0.26, 0.14))
    Put-Data 'big' $cats5 @(@(12500, 48000, 31000, 95000, 67000), @(8300, 21000, 54000, 40000, 76000), @(15500, 9800, 26000, 33000, 12000))
    Put-Data 'negbig' $cats5 @(@(12000, -5000, 7000, -15000, 9000), @(-8000, 14000, -11000, 6000, -3000), @(5000, -9000, 13000, -10000, 4000))
    Put-Data 'allneg' $cats5 @(@(-12, -5, -7, -15, -9), @(-8, -14, -11, -6, -3), @(-5, -9, -13, -10, -4))
    Put-Data 'mil7' $cats5 @(@(125000, 480000, 310000, 950000, 670000), @(83000, 210000, 540000, 400000, 760000), @(155000, 98000, 260000, 330000, 120000))
    Put-Data 'mil8' $cats5 @(@(1250000, 4800000, 3100000, 9500000, 6700000), @(830000, 2100000, 5400000, 4000000, 7600000), @(1550000, 980000, 2600000, 3300000, 1200000))
    $m13 = @(); for ($s = 0; $s -lt 13; $s++) { $row = @(); for ($c = 0; $c -lt 5; $c++) { $row += ((($s + 1) * 7 + ($c + 1) * 3) % 9) + 2 }; $m13 += , $row }
    Put-Data 'ser13' $cats5 $m13
    $c13 = @(); for ($c = 1; $c -le 13; $c++) { $c13 += ('P' + $c) }
    Put-Data 'pt13' $c13 @(, @(5, 9, 14, 7, 11, 6, 13, 8, 10, 12, 4, 15, 9))
    # C7 sweep data: sw_<count>_<length>[s]  (s = with spaces; only for the lengths 12 and 20)
    foreach ($n in @(5, 10, 15, 20, 30)) {
      foreach ($len in @(4, 8, 12, 20)) {
        foreach ($sp in @($false, $true)) {
          if ($sp -and $len -lt 12) { continue }
          $cs = @(); $vs = @()
          for ($c = 1; $c -le $n; $c++) { $cs += (Cat-Name $c $len $sp); $vs += ((($c * 7) % 13) + 3) }
          Put-Data ("sw_{0}_{1}{2}" -f $n, $len, $(if ($sp) { 's' } else { '' })) $cs @(, $vs)
        }
      }
    }

    $types = [ordered]@{ colClustered = 51; colStacked = 52; colStacked100 = 53; barClustered = 57; barStacked = 58; barStacked100 = 59; line = 4; lineMarkers = 65; pie = 5; doughnut = -4120 }
    $size = [ordered]@{ small = @(216, 144); std = @(360, 216); large = @(720, 432) }
    function Ex($name, $xlType, $dataKey, $nser, $sizeKey, $tweak, $meta, $withPoints = $false, $setTitle = $true) {
      if ($sizeKey -is [string]) { $wh = $size[$sizeKey] } else { $wh = $sizeKey }
      New-Chart $script:ws $name $xlType $dataKey $nser $wh[0] $wh[1] $meta $tweak $withPoints $setTitle
    }

    # ========== C1: horizontal bars (clustered / stacked) with negative values, 3 sizes (data neg and negbig) ==========
    $script:ws = Add-ChartSheet 'c1'
    foreach ($tk in @('barClustered', 'barStacked')) {
      foreach ($dk in @('neg', 'negbig')) {
        foreach ($sk in $size.Keys) {
          Ex ("c1neg_{0}_{1}_{2}" -f $tk, $dk, $sk) $types[$tk] $dk 3 $sk $null @{ Kind = 'c1neg'; Check = 'C1'; Type = $tk; DataSet = $dk; Size = $sk }
        }
      }
    }
    # C1: tick labels of 7 / 8 digits (horizontal bars: the minimum gap of the horizontal axis depends on the label width; a column chart for the vertical side)
    foreach ($dk in @('mil7', 'mil8')) {
      foreach ($sk in $size.Keys) {
        Ex ("c1mil_barClustered_{0}_{1}" -f $dk, $sk) 57 $dk 3 $sk $null @{ Kind = 'c1mil'; Check = 'C1'; Type = 'barClustered'; DataSet = $dk; Size = $sk }
      }
      Ex ("c1mil_colClustered_{0}_std" -f $dk) 51 $dk 3 'std' $null @{ Kind = 'c1mil'; Check = 'C1'; Type = 'colClustered'; DataSet = $dk; Size = 'std' }
    }

    # ========== C2: 100% stacked, all values negative (column and bar), 3 sizes ==========
    $script:ws = Add-ChartSheet 'c2'
    foreach ($tk in @('colStacked100', 'barStacked100')) {
      foreach ($sk in $size.Keys) {
        Ex ("c2allneg_{0}_{1}" -f $tk, $sk) $types[$tk] 'allneg' 3 $sk $null @{ Kind = 'c2allneg'; Check = 'C2'; Type = $tk; DataSet = 'allneg'; Size = $sk }
      }
    }

    # ========== C3 / C6: pie / doughnut with varyColors = 0 (and true as the baseline): legend entries ==========
    $script:ws = Add-ChartSheet 'c3'
    $v0 = @(
      @('c3vary0_pie_1ser_r', 5, 'pos', 1, -4152, $false, 'r'), @('c3vary0_pie_1ser_b', 5, 'pos', 1, -4107, $false, 'b'), @('c3vary0_pie_pt13_r', 5, 'pt13', 1, -4152, $false, 'r'),
      @('c3vary0_doughnut_1ser_r', -4120, 'pos', 1, -4152, $false, 'r'), @('c3vary0_doughnut_3ser_r', -4120, 'pos', 3, -4152, $false, 'r'), @('c3vary0_doughnut_3ser_b', -4120, 'pos', 3, -4107, $false, 'b'),
      @('c3vary1_pie_1ser_b', 5, 'pos', 1, -4107, $true, 'b'), @('c3vary1_doughnut_3ser_r', -4120, 'pos', 3, -4152, $true, 'r'))
    foreach ($v in $v0) {
      $script:vv = $v
      Ex $v[0] $v[1] $v[2] $v[3] 'std' { $ch.ChartGroups(1).VaryByCategories = $script:vv[5]; $ch.HasLegend = $true; $ch.Legend.Position = $script:vv[4] } @{ Kind = 'c3vary'; Check = 'C3, C6'; Type = $(if ($v[1] -eq 5) { 'pie' } else { 'doughnut' }); VaryAsked = $v[5]; LegendPosAsked = $v[6] } $true
    }

    # ========== C4: ChartStyle (1 2 3 10 26 42 on a column chart; 3 on a line with markers and on a pie). Read the colours of the area / text / gridlines / axis lines ==========
    $script:ws = Add-ChartSheet 'c4'
    foreach ($sn in @(1, 2, 3, 10, 26, 42)) {
      $script:sn = $sn
      Ex ("c4style_{0}_colClustered_std" -f $sn) 51 'pos' 3 'std' { $ch.ChartStyle = [int]$script:sn; $ch.HasLegend = $true } @{ Kind = 'c4style'; Check = 'C4'; Type = 'colClustered'; StyleAsked = $sn }
    }
    $script:sn = 3
    Ex 'c4style_3_lineMarkers_std' 65 'pos' 3 'std' { $ch.ChartStyle = 3; $ch.HasLegend = $true } @{ Kind = 'c4style'; Check = 'C4'; Type = 'lineMarkers'; StyleAsked = 3 }
    Ex 'c4style_3_pie_std' 5 'pos' 1 'std' { $ch.ChartStyle = 3; $ch.HasLegend = $true } @{ Kind = 'c4style'; Check = 'C4'; Type = 'pie'; StyleAsked = 3 } $true

    # ========== C7: category axis sweep (rotation / wrapping): column std = count x length (no spaces; 12 and 20 also with spaces), 6 horizontal bars ==========
    $script:ws = Add-ChartSheet 'c7'
    foreach ($n in @(5, 10, 15, 20, 30)) {
      foreach ($len in @(4, 8, 12, 20)) {
        foreach ($sp in @($false, $true)) {
          if ($sp -and $len -lt 12) { continue }
          $suf = $(if ($sp) { 's' } else { '' })
          Ex ("c7sw_col_{0}_{1}{2}" -f $n, $len, $suf) 51 ("sw_{0}_{1}{2}" -f $n, $len, $suf) 1 'std' $null @{ Kind = 'c7sw'; Check = 'C7'; Type = 'colClustered'; Count = $n; NameLen = $len; Spaces = $sp }
        }
      }
    }
    foreach ($n in @(15, 30)) {
      foreach ($len in @(8, 20)) {
        Ex ("c7sw_bar_{0}_{1}" -f $n, $len) 57 ("sw_{0}_{1}" -f $n, $len) 1 'std' $null @{ Kind = 'c7sw'; Check = 'C7'; Type = 'barClustered'; Count = $n; NameLen = $len; Spaces = $false }
      }
      Ex ("c7sw_bar_{0}_20s" -f $n) 57 ("sw_{0}_20s" -f $n) 1 'std' $null @{ Kind = 'c7sw'; Check = 'C7'; Type = 'barClustered'; Count = $n; NameLen = 20; Spaces = $true }
    }

    # ========== C5: legend at the corner (xlLegendPositionCorner = 2): where the plot goes ==========
    $script:ws = Add-ChartSheet 'c5'
    foreach ($sk in $size.Keys) {
      Ex ("c5corner_colClustered_{0}" -f $sk) 51 'pos' 3 $sk { $ch.HasLegend = $true; $ch.Legend.Position = 2 } @{ Kind = 'c5corner'; Check = 'C5'; Type = 'colClustered'; Size = $sk; LegendPosAsked = 'tr' }
    }
    Ex 'c5corner_barClustered_std' 57 'pos' 3 'std' { $ch.HasLegend = $true; $ch.Legend.Position = 2 } @{ Kind = 'c5corner'; Check = 'C5'; Type = 'barClustered'; Size = 'std'; LegendPosAsked = 'tr' }
    Ex 'c5corner_lineMarkers_std' 65 'pos' 3 'std' { $ch.HasLegend = $true; $ch.Legend.Position = 2 } @{ Kind = 'c5corner'; Check = 'C5'; Type = 'lineMarkers'; Size = 'std'; LegendPosAsked = 'tr' }
    Ex 'c5corner_pie_std' 5 'pos' 1 'std' { $ch.HasLegend = $true; $ch.Legend.Position = 2 } @{ Kind = 'c5corner'; Check = 'C5'; Type = 'pie'; Size = 'std'; LegendPosAsked = 'tr' } $true

    # ========== C6: 13 legend entries in a row (pie / doughnut, legend top / bottom, widths): the bounding box of the entries ==========
    $script:ws = Add-ChartSheet 'c6'
    foreach ($tk in @('pie', 'doughnut')) {
      foreach ($lk in @('t', 'b')) {
        foreach ($wd in @(200, 280, 360, 480, 720)) {
          $script:lp = $(if ($lk -eq 't') { -4160 } else { -4107 })
          Ex ("c6row13_{0}_{1}_{2}" -f $tk, $lk, $wd) $types[$tk] 'pt13' 1 @($wd, 216) { $ch.HasLegend = $true; $ch.Legend.Position = $script:lp } @{ Kind = 'c6row13'; Check = 'C6'; Type = $tk; LegendPosAsked = $lk } $true
        }
      }
    }

    # ========== MinorUnit: minor gridlines / minor tick marks / explicit MinorUnit (MinorUnit and MajorUnit of the value axis are recorded) ==========
    $script:ws = Add-ChartSheet 'minor'
    $minorGrid = { $a = $ch.Axes(2); $a.HasMinorGridlines = $true; $a.MinorTickMark = 3 }
    foreach ($sk in $size.Keys) { Ex ("minor_colClustered_pos_{0}" -f $sk) 51 'pos' 3 $sk $minorGrid @{ Kind = 'minor'; Check = 'MinorUnit'; Type = 'colClustered'; DataSet = 'pos'; Size = $sk } }
    foreach ($dk in @('big', 'dec', 'neg')) { Ex ("minor_colClustered_{0}_std" -f $dk) 51 $dk 3 'std' $minorGrid @{ Kind = 'minor'; Check = 'MinorUnit'; Type = 'colClustered'; DataSet = $dk; Size = 'std' } }
    Ex 'minor_barClustered_pos_std' 57 'pos' 3 'std' $minorGrid @{ Kind = 'minor'; Check = 'MinorUnit'; Type = 'barClustered'; DataSet = 'pos'; Size = 'std' }
    Ex 'minor_lineMarkers_pos_std' 65 'pos' 3 'std' $minorGrid @{ Kind = 'minor'; Check = 'MinorUnit'; Type = 'lineMarkers'; DataSet = 'pos'; Size = 'std' }
    Ex 'minor_colClustered_pos_tickOnly_std' 51 'pos' 3 'std' { $a = $ch.Axes(2); $a.MinorTickMark = 3 } @{ Kind = 'minor'; Check = 'MinorUnit'; Type = 'colClustered'; DataSet = 'pos'; Size = 'std'; Note = 'minor tick marks only, no minor gridlines' }
    Ex 'minor_colClustered_pos_major4_minor1_std' 51 'pos' 3 'std' { $a = $ch.Axes(2); $a.MajorUnit = 4; $a.MinorUnit = 1; $a.HasMinorGridlines = $true } @{ Kind = 'minor'; Check = 'MinorUnit'; Type = 'colClustered'; DataSet = 'pos'; Size = 'std'; Note = 'explicit MajorUnit 4 / MinorUnit 1' }
    Ex 'minor_colClustered_pos_major5_minor2p5_std' 51 'pos' 3 'std' { $a = $ch.Axes(2); $a.MajorUnit = 5; $a.MinorUnit = 2.5; $a.HasMinorGridlines = $true } @{ Kind = 'minor'; Check = 'MinorUnit'; Type = 'colClustered'; DataSet = 'pos'; Size = 'std'; Note = 'explicit MajorUnit 5 / MinorUnit 2.5' }

    # ========== C11: LegendEntries(i).Delete and Series.PlotOrder (legend at the right). The remaining entries (positions, key colours) say which series each idx points at ==========
    $script:ws = Add-ChartSheet 'c11'
    $c11 = { $ch.HasLegend = $true; $ch.Legend.Position = -4152 }
    Ex 'c11_col3_del2' 51 'pos' 3 'std' { & $c11; Del-Entry $ch 2 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'delete entry 2' }
    Ex 'c11_col3_del1' 51 'pos' 3 'std' { & $c11; Del-Entry $ch 1 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'delete entry 1' }
    Ex 'c11_col3_del3' 51 'pos' 3 'std' { & $c11; Del-Entry $ch 3 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'delete entry 3' }
    Ex 'c11_col3_swap' 51 'pos' 3 'std' { & $c11; $ch.SeriesCollection(1).PlotOrder = 3 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'S1.PlotOrder = 3' }
    Ex 'c11_col3_swap_del1' 51 'pos' 3 'std' { & $c11; $ch.SeriesCollection(1).PlotOrder = 3; Del-Entry $ch 1 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'S1.PlotOrder = 3; delete entry 1' }
    Ex 'c11_col3_swap_del3' 51 'pos' 3 'std' { & $c11; $ch.SeriesCollection(1).PlotOrder = 3; Del-Entry $ch 3 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'S1.PlotOrder = 3; delete entry 3' }
    Ex 'c11_col3_del2_swap' 51 'pos' 3 'std' { & $c11; Del-Entry $ch 2; $ch.SeriesCollection(1).PlotOrder = 3 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'delete entry 2; S1.PlotOrder = 3' }
    Ex 'c11_line3_swap_del2' 65 'pos' 3 'std' { & $c11; $ch.SeriesCollection(1).PlotOrder = 3; Del-Entry $ch 2 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'S1.PlotOrder = 3; delete entry 2' }
    Ex 'c11_stack3_swap_del2' 52 'pos' 3 'std' { & $c11; $ch.SeriesCollection(1).PlotOrder = 3; Del-Entry $ch 2 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'S1.PlotOrder = 3; delete entry 2' }
    Ex 'c11_bar3_swap_del1' 57 'pos' 3 'std' { & $c11; $ch.SeriesCollection(1).PlotOrder = 3; Del-Entry $ch 1 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'S1.PlotOrder = 3; delete entry 1' }
    Ex 'c11_col4_swap_del2_del3' 51 'ser13' 4 'std' { & $c11; $ch.SeriesCollection(1).PlotOrder = 4; Del-Entry $ch 3; Del-Entry $ch 2 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'S1.PlotOrder = 4; delete entry 3, then entry 2' }
    Ex 'c11_pie5_del3' 5 'pos' 1 'std' { & $c11; Del-Entry $ch 3 } @{ Kind = 'c11'; Check = 'C11'; Ops = 'delete entry 3 (point 3)' } $true

    $script:wb.SaveAs($refPath, 51); $script:wb.Close($false); $script:wb = $null
    $out = [ordered]@{ excelVersion = $xl.Version; unit = 'pt'; note = 'COM values as returned (pt). Screen zoom 125%: 0.6 pt = 1 px. Colours are RRGGBB (negative = Excel colour index constant).'
      notes = $notes; charts = $records }
    $json = $out | ConvertTo-Json -Depth 12
    [IO.File]::WriteAllText((Join-Path $Dir 'excel2.json'), $json, (New-Object Text.UTF8Encoding($false)))
    "ok: $($records.Count) charts; notes=$($notes.Count)"
  }

  # ========== C12 / C13 / C4 (c:style): the charts of make-ref-charts-default.ts, opened read-only ==========
  if ('default' -in $stages) {
    $defRecords = New-Object System.Collections.ArrayList
    $dnotes = New-Object System.Collections.ArrayList
    $dwb = $xl.Workbooks.Open($DefaultXlsx, 0, $true)
    try {
      foreach ($dws in $dwb.Worksheets) {
        foreach ($co in $dws.ChartObjects()) {
          try {
            $meta = [ordered]@{ Kind = 'default'; Sheet = $dws.Name; Check = $(if ($co.Name -like 'def2_tick*') { 'C12' } elseif ($co.Name -like 'def2_line*') { 'C13' } else { 'C4' }) }
            [void]$defRecords.Add((Rec-Chart $co $meta $false))
          } catch { [void]$dnotes.Add("default $($co.Name) failed: $($_.Exception.Message)") }
        }
      }
    } finally { $dwb.Close($false); $dwb = $null }
    $dout = [ordered]@{ excelVersion = $xl.Version; unit = 'pt'; source = [IO.Path]::GetFileName($DefaultXlsx); notes = $dnotes; charts = $defRecords }
    $djson = $dout | ConvertTo-Json -Depth 12
    [IO.File]::WriteAllText((Join-Path $Dir 'excel2.default.json'), $djson, (New-Object Text.UTF8Encoding($false)))
    "ok: $($defRecords.Count) default charts; notes=$($dnotes.Count)"
  }

  # ========== PNG (Excel must be Visible for Chart.Export) ==========
  if ('png' -in $stages) {
    $xl.Visible = $true
    if (Test-Path $refPath) { Export-Pngs $refPath '^(c3vary|c4style|c5corner|c6row13|c7sw|c11_)' 'ref-charts2' }
    if ($DefaultXlsx -ne '') { Export-Pngs $DefaultXlsx '.*' 'default' }
  }
} finally {
  try { if ($null -ne $script:wb) { $script:wb.Close($false) } } catch {}
  try { $xl.Quit() } catch {}
  [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
