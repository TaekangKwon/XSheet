# ISSUE-104: ref-charts.xlsx (Excel 16 COM) and what Excel itself reports for every G1 chart (ADR-0031 section 12, items C1..C10).
# Writes into $Dir (nothing else is touched): ref-charts.xlsx, com.json (COM values, pt as returned), com.default.json (only with -DefaultXlsx),
# png\*.png (only with -Png; the pictures are for a later visual comparison, not committed).
# Run (do NOT run on CI; needs Excel 16; screen zoom 125% like make-ref-colors.ps1 - COM values are pt and stay pt, converting to px is the test's job):
#   powershell -ExecutionPolicy Bypass -File make-ref-charts.ps1 -Dir <out dir> [-DefaultXlsx <ref-charts-default.xlsx>] [-SkipCreate] [-Png]
#   (PowerShell 5.1, file saved as UTF-8 with BOM). Then: node make-ref-charts.extract.mjs <out dir>; copy ref-charts.xlsx and ref-charts.expected.json to __fixtures__/excel.
# Switches:
#   -DefaultXlsx <path>  also open that file (made by make-ref-charts-default.ts: charts with no c:spPr / c:txPr / c:style) read-only and record the same values -> com.default.json (C4 / C10 / C3)
#   -SkipCreate          only the -DefaultXlsx part (no ref-charts.xlsx)
#   -Png                 also Chart.Export every chart as png\<name>.png (default: off). The clipboard is never used (no CopyPicture / Paste).
# Charts (each is ChartObject.Name = key of the expected values; Shapes.AddChart2(-1, type, ...) = Excel's current default style):
#   G1 types (10): colClustered 51, colStacked 52, colStacked100 53, barClustered 57, barStacked 58, barStacked100 59, line 4, lineMarkers 65, pie 5, doughnut -4120.
#   Data (4, 3 series x 5 categories): pos (positives only), neg (mixed signs), dec (decimals below 1), big (tens of thousands).
#   Sizes (pt): small 216x144, std 360x216, large 720x432.
# Why not the full 10 x 4 x 3 = 120 grid: the size only changes the automatic axis range (C1) and layout (C5) -- one data set with a small range (pos) and one with a
# large range (big) show both extremes; neg and dec are about the data (min/max rules, number format), so they stay at std. Pie / doughnut have no axis: pos and big
# at 3 sizes is enough for C5 / C9, neg (negative slices) and dec at std. => per axis type: pos 3 + big 3 + neg 1 + dec 1 = 8 (x8 types = 64), pie/doughnut 8 each (16): 80 charts.
#   Extras (sheet 'extra', std size unless named): 7 and 13 series (C3), 13 points in pie / doughnut (C3), varyColors on/off with 1 series (C3), legend at r/l/t/b/corner/none
#   (C5, C6; also with line markers and pie), manual layout of title / legend / plot (C5), no title / auto title with 1 and 2 series / overlay (C10, C5),
#   axis titles (C5), 30 categories and long category names at small and std (C7: Excel does not say the rotation or the thinning through COM except TickLabelSpacing; the PNG does).
param([string]$Dir, [string]$DefaultXlsx = '', [switch]$SkipCreate, [switch]$Png)
$ErrorActionPreference = 'Stop'
# relative paths are resolved against the PowerShell location (not the process cwd), then made absolute
if (-not [IO.Path]::IsPathRooted($Dir)) { $Dir = Join-Path (Get-Location).Path $Dir }
$Dir = [IO.Path]::GetFullPath($Dir)
if ($DefaultXlsx -ne '') {
  if (-not [IO.Path]::IsPathRooted($DefaultXlsx)) { $DefaultXlsx = Join-Path (Get-Location).Path $DefaultXlsx }
  $DefaultXlsx = [IO.Path]::GetFullPath($DefaultXlsx)
}
if ($SkipCreate -and $DefaultXlsx -eq '') { throw '-SkipCreate needs -DefaultXlsx' }
New-Item -ItemType Directory -Force -Path $Dir | Out-Null
$pngDir = Join-Path $Dir 'png'
if ($Png) { New-Item -ItemType Directory -Force -Path $pngDir | Out-Null }

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
# Create one chart, run $tweak (optional) and record it. $nser = how many series of the data set; title = 'Title' unless $title is $false
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
  if ($Png) { Try-Do "png $name" { [void]$ch.Export((Join-Path $pngDir ($name + '.png')), 'PNG') } }
  $m = [ordered]@{ Sheet = $ws.Name; Data = $dataKey; SeriesAsked = $nser; WidthAsked = $w; HeightAsked = $h }
  foreach ($k in $meta.Keys) { $m[$k] = $meta[$k] }
  [void]$script:records.Add((Rec-Chart $co $m $withPoints))
}

try {
  if (-not $SkipCreate) {
    $script:wb = $xl.Workbooks.Add()
    $script:wsData = $script:wb.Worksheets.Item(1); $script:wsData.Name = 'Data'
    $cats5 = @('Cat1', 'Cat2', 'Cat3', 'Cat4', 'Cat5')
    Put-Data 'pos' $cats5 @(@(12, 19, 7, 15, 9), @(8, 14, 11, 6, 17), @(5, 9, 13, 10, 4))
    Put-Data 'neg' $cats5 @(@(12, -5, 7, -15, 9), @(-8, 14, -11, 6, -3), @(5, -9, 13, -10, 4))
    Put-Data 'dec' $cats5 @(@(0.12, 0.35, 0.27, 0.48, 0.31), @(0.08, 0.22, 0.41, 0.16, 0.37), @(0.05, 0.18, 0.09, 0.26, 0.14))
    Put-Data 'big' $cats5 @(@(12500, 48000, 31000, 95000, 67000), @(8300, 21000, 54000, 40000, 76000), @(15500, 9800, 26000, 33000, 12000))
    $m13 = @(); for ($s = 0; $s -lt 13; $s++) { $row = @(); for ($c = 0; $c -lt 5; $c++) { $row += ((($s + 1) * 7 + ($c + 1) * 3) % 9) + 2 }; $m13 += , $row }
    Put-Data 'ser13' $cats5 $m13
    $c13 = @(); for ($c = 1; $c -le 13; $c++) { $c13 += ('P' + $c) }
    Put-Data 'pt13' $c13 @(, @(5, 9, 14, 7, 11, 6, 13, 8, 10, 12, 4, 15, 9))
    $c30 = @(); $v30 = @(); for ($c = 1; $c -le 30; $c++) { $c30 += ('Item' + $c.ToString('00')); $v30 += ((($c * 7) % 13) + 3) }
    Put-Data 'cat30' $c30 @(, $v30)
    $cLong = @(); $vLong = @(); for ($c = 1; $c -le 12; $c++) { $cLong += ('Long category name ' + $c.ToString('00')); $vLong += ((($c * 5) % 11) + 4) }
    Put-Data 'catlong' $cLong @(, $vLong)

    $types = [ordered]@{ colClustered = 51; colStacked = 52; colStacked100 = 53; barClustered = 57; barStacked = 58; barStacked100 = 59; line = 4; lineMarkers = 65; pie = 5; doughnut = -4120 }
    $size = [ordered]@{ small = @(216, 144); std = @(360, 216); large = @(720, 432) }
    $legendPos = [ordered]@{ r = -4152; l = -4131; t = -4160; b = -4107; tr = 2 }

    # ========== grid: sheets pos / big (3 sizes), neg / dec (std only) ==========
    foreach ($dk in @('pos', 'big', 'neg', 'dec')) {
      $ws = Add-ChartSheet $dk
      foreach ($tk in $types.Keys) {
        foreach ($sk in $size.Keys) {
          if ($dk -in @('neg', 'dec') -and $sk -ne 'std') { continue }
          $meta = [ordered]@{ Kind = 'grid'; Type = $tk; Size = $sk }
          New-Chart $ws ("{0}_{1}_{2}" -f $tk, $dk, $sk) $types[$tk] $dk 3 $size[$sk][0] $size[$sk][1] $meta $null
        }
      }
    }

    # ========== extras ==========
    $ws = Add-ChartSheet 'extra'
    $std = $size['std']
    function Ex($name, $xlType, $dataKey, $nser, $sizeKey, $tweak, $meta, $withPoints = $false, $setTitle = $true) {
      $m = [ordered]@{ Kind = 'extra' }
      foreach ($k in $meta.Keys) { $m[$k] = $meta[$k] }
      New-Chart $ws $name $xlType $dataKey $nser $size[$sizeKey][0] $size[$sizeKey][1] $m $tweak $withPoints $setTitle
    }
    # C3: 7 and 13 series (colours of the 7th series and later), 13 points in pie / doughnut, varyColors with one series
    Ex 'ser7_colClustered_std' 51 'ser13' 7 'std' $null @{ Check = 'C3' }
    Ex 'ser13_colClustered_std' 51 'ser13' 13 'std' $null @{ Check = 'C3' }
    Ex 'ser13_line_std' 4 'ser13' 13 'std' $null @{ Check = 'C3' }
    Ex 'ser13_lineMarkers_std' 65 'ser13' 13 'std' $null @{ Check = 'C3, C8' }
    Ex 'ser13_doughnut_std' -4120 'ser13' 13 'std' $null @{ Check = 'C3, C9' }
    Ex 'pt13_pie_std' 5 'pt13' 1 'std' $null @{ Check = 'C3' }
    Ex 'pt13_doughnut_std' -4120 'pt13' 1 'std' $null @{ Check = 'C3, C9' }
    Ex 'vary_colClustered_1ser_off_std' 51 'pos' 1 'std' { $ch.ChartGroups(1).VaryByCategories = $false } @{ Check = 'C3' } $true
    Ex 'vary_colClustered_1ser_on_std' 51 'pos' 1 'std' { $ch.ChartGroups(1).VaryByCategories = $true } @{ Check = 'C3' } $true
    Ex 'vary_lineMarkers_1ser_on_std' 65 'pos' 1 'std' { $ch.ChartGroups(1).VaryByCategories = $true } @{ Check = 'C3, C8' } $true
    # C5, C6: legend position (5) and none; legend of line markers and of pie
    foreach ($lk in $legendPos.Keys) {
      $lp = $legendPos[$lk]
      Ex ("legend_{0}_colClustered_std" -f $lk) 51 'pos' 3 'std' { $ch.HasLegend = $true; $ch.Legend.Position = $lp } @{ Check = 'C5, C6'; LegendPosAsked = $lk }
    }
    Ex 'legend_none_colClustered_std' 51 'pos' 3 'std' { $ch.HasLegend = $false } @{ Check = 'C5'; LegendPosAsked = 'none' }
    Ex 'legend_r_lineMarkers_std' 65 'pos' 3 'std' { $ch.HasLegend = $true; $ch.Legend.Position = -4152 } @{ Check = 'C5, C6, C8'; LegendPosAsked = 'r' }
    Ex 'legend_b_lineMarkers_std' 65 'pos' 3 'std' { $ch.HasLegend = $true; $ch.Legend.Position = -4107 } @{ Check = 'C5, C6, C8'; LegendPosAsked = 'b' }
    Ex 'legend_r_pie_std' 5 'pos' 1 'std' { $ch.HasLegend = $true; $ch.Legend.Position = -4152 } @{ Check = 'C5, C6'; LegendPosAsked = 'r' } $true
    Ex 'legend_b_pie_std' 5 'pos' 1 'std' { $ch.HasLegend = $true; $ch.Legend.Position = -4107 } @{ Check = 'C5, C6'; LegendPosAsked = 'b' } $true
    # C5: manual layout (what is asked is in meta; what Excel kept is in the recorded values)
    Ex 'manual_colClustered_std' 51 'pos' 3 'std' {
      $ch.HasLegend = $true; $ch.Legend.Position = -4152
      $ch.ChartTitle.Left = 40; $ch.ChartTitle.Top = 6
      $ch.PlotArea.InsideLeft = 60; $ch.PlotArea.InsideTop = 50; $ch.PlotArea.InsideWidth = 200; $ch.PlotArea.InsideHeight = 110
      $ch.Legend.Left = 280; $ch.Legend.Top = 60; $ch.Legend.Width = 70; $ch.Legend.Height = 70
    } @{ Check = 'C5'; Asked = @{ TitleLeft = 40; TitleTop = 6; InsideLeft = 60; InsideTop = 50; InsideWidth = 200; InsideHeight = 110; LegendLeft = 280; LegendTop = 60; LegendWidth = 70; LegendHeight = 70 } }
    Ex 'manual_pie_std' 5 'pos' 1 'std' {
      $ch.HasLegend = $true; $ch.Legend.Position = -4152
      $ch.PlotArea.InsideLeft = 40; $ch.PlotArea.InsideTop = 40; $ch.PlotArea.InsideWidth = 150; $ch.PlotArea.InsideHeight = 150
    } @{ Check = 'C5, C9'; Asked = @{ InsideLeft = 40; InsideTop = 40; InsideWidth = 150; InsideHeight = 150 } } $true
    # C10, C5: no title, automatic title (1 series / 2 series: the text Excel puts), title and legend not taking space (overlay)
    Ex 'title_none_colClustered_std' 51 'pos' 3 'std' { $ch.HasTitle = $false } @{ Check = 'C10, C5' } $false $false
    Ex 'title_auto_1ser_colClustered_std' 51 'pos' 1 'std' $null @{ Check = 'C10' } $false $false
    Ex 'title_auto_2ser_colClustered_std' 51 'pos' 2 'std' $null @{ Check = 'C10' } $false $false
    Ex 'overlay_colClustered_std' 51 'pos' 3 'std' { $ch.HasLegend = $true; $ch.ChartTitle.IncludeInLayout = $false; $ch.Legend.IncludeInLayout = $false } @{ Check = 'C5' }
    # C5: axis titles (the strip they take)
    Ex 'axtitle_colClustered_std' 51 'pos' 3 'std' {
      $ch.Axes(1).HasTitle = $true; $ch.Axes(1).AxisTitle.Text = 'Category axis'; $ch.Axes(2).HasTitle = $true; $ch.Axes(2).AxisTitle.Text = 'Value axis'
    } @{ Check = 'C5' }
    Ex 'axtitle_barClustered_std' 57 'pos' 3 'std' {
      $ch.Axes(1).HasTitle = $true; $ch.Axes(1).AxisTitle.Text = 'Category axis'; $ch.Axes(2).HasTitle = $true; $ch.Axes(2).AxisTitle.Text = 'Value axis'
    } @{ Check = 'C5' }
    # C7: many categories and long names, small and std (thinning / rotation: read TickLabelSpacing and the PNG)
    Ex 'cat30_colClustered_small' 51 'cat30' 1 'small' $null @{ Check = 'C7' }
    Ex 'cat30_colClustered_std' 51 'cat30' 1 'std' $null @{ Check = 'C7' }
    Ex 'catlong_colClustered_small' 51 'catlong' 1 'small' $null @{ Check = 'C7' }
    Ex 'catlong_colClustered_std' 51 'catlong' 1 'std' $null @{ Check = 'C7' }
    Ex 'catlong_barClustered_small' 57 'catlong' 1 'small' $null @{ Check = 'C7' }

    $script:wb.SaveAs((Join-Path $Dir 'ref-charts.xlsx'), 51); $script:wb.Close($false); $script:wb = $null
    $out = [ordered]@{ excelVersion = $xl.Version; unit = 'pt'; note = 'COM values as returned (pt). Screen zoom 125%: 0.6 pt = 1 px. Colours are RRGGBB (negative = Excel colour index constant).'
      notes = $notes; charts = $records }
    $json = $out | ConvertTo-Json -Depth 12
    [IO.File]::WriteAllText((Join-Path $Dir 'com.json'), $json, (New-Object Text.UTF8Encoding($false)))
    "ok: $($records.Count) charts; notes=$($notes.Count)"
  }

  # ========== C4: the charts with no c:spPr / c:txPr / c:style (made by make-ref-charts-default.ts), opened read-only ==========
  if ($DefaultXlsx -ne '') {
    $defRecords = New-Object System.Collections.ArrayList
    $dnotes = New-Object System.Collections.ArrayList
    $dwb = $xl.Workbooks.Open($DefaultXlsx, 0, $true)
    try {
      foreach ($dws in $dwb.Worksheets) {
        foreach ($co in $dws.ChartObjects()) {
          try {
            $meta = [ordered]@{ Kind = 'default'; Sheet = $dws.Name; Check = 'C4' }
            [void]$defRecords.Add((Rec-Chart $co $meta $false))
            if ($Png) { [void]$co.Chart.Export((Join-Path $pngDir ($co.Name + '.png')), 'PNG') }
          } catch { [void]$dnotes.Add("default $($co.Name) failed: $($_.Exception.Message)") }
        }
      }
    } finally { $dwb.Close($false); $dwb = $null }
    $dout = [ordered]@{ excelVersion = $xl.Version; unit = 'pt'; source = [IO.Path]::GetFileName($DefaultXlsx); notes = $dnotes; charts = $defRecords }
    $djson = $dout | ConvertTo-Json -Depth 12
    [IO.File]::WriteAllText((Join-Path $Dir 'com.default.json'), $djson, (New-Object Text.UTF8Encoding($false)))
    "ok: $($defRecords.Count) default charts; notes=$($dnotes.Count)"
  }
} finally {
  try { if ($null -ne $script:wb) { $script:wb.Close($false) } } catch {}
  try { $xl.Quit() } catch {}
  [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
