# showcase.xlsx (Excel 16 COM): one small, good-looking workbook for the README screenshot and the example page.
# It is not a reference for tests (no expected values); it only shows what XSheet draws in one screen:
# a merged title, a header with fill and white bold text, borders, number formats (thousands, %), a red negative,
# a total row, a clustered column chart, a callout shape and a second sheet tab.
# Only what XSheet draws today is used: no cell alignment, wrapping or conditional formatting (README "Not yet drawn"),
# so the screenshot looks the same as the workbook in Excel.
# Run (do NOT run on CI; needs Excel 16):
#   powershell -ExecutionPolicy Bypass -File make-showcase.ps1 -Dir <out dir>
#   Then copy <out dir>\showcase.xlsx to fixtures/excel/.
param([string]$Dir)
$ErrorActionPreference = 'Stop'
if (-not [IO.Path]::IsPathRooted($Dir)) { $Dir = Join-Path (Get-Location).Path $Dir }
$Dir = [IO.Path]::GetFullPath($Dir)
New-Item -ItemType Directory -Force -Path $Dir | Out-Null
$out = Join-Path $Dir 'showcase.xlsx'

# Excel colours are BGR integers
function Rgb($r, $g, $b) { $r + 256 * $g + 65536 * $b }
$navy = Rgb 31 56 100
$band = Rgb 237 242 249
$grey = Rgb 89 89 89
$line = Rgb 191 191 191

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $ws = $wb.Worksheets.Item(1); $ws.Name = 'Sales'
  $notes = $wb.Worksheets.Item(2); $notes.Name = 'Notes'
  $xl.ActiveWindow.DisplayGridlines = $true

  # title and subtitle
  $ws.Range('A1:G1').Merge()
  $ws.Range('A1').Value2 = '2026 Sales by Region'
  $ws.Range('A1').Font.Size = 18; $ws.Range('A1').Font.Bold = $true; $ws.Range('A1').Font.Color = $navy
  $ws.Rows.Item(1).RowHeight = 30
  $ws.Range('A2:G2').Merge()
  $ws.Range('A2').Value2 = 'Figures in USD thousands'
  $ws.Range('A2').Font.Italic = $true; $ws.Range('A2').Font.Color = $grey

  # header
  $head = @('Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Total', 'Growth')
  for ($c = 0; $c -lt $head.Count; $c++) { $ws.Cells(4, $c + 1).Value2 = $head[$c] }
  $h = $ws.Range('A4:G4')
  $h.Interior.Color = $navy; $h.Font.Color = Rgb 255 255 255; $h.Font.Bold = $true
  $ws.Rows.Item(4).RowHeight = 22

  # data (Q1..Q4), total and growth (Q4 against Q1) are formulas so Excel saves their values
  $rows = @(
    @('North', 1240, 1315, 1402, 1488),
    @('South', 980, 1012, 955, 1034),
    @('East', 1530, 1488, 1396, 1342),
    @('West', 860, 934, 1021, 1107),
    @('Online', 620, 744, 829, 918)
  )
  for ($i = 0; $i -lt $rows.Count; $i++) {
    $r = 5 + $i
    for ($c = 0; $c -lt 5; $c++) { if ($c -eq 0) { $ws.Cells($r, 1).Value2 = [string]$rows[$i][0] } else { $ws.Cells($r, $c + 1).Value2 = [double]$rows[$i][$c] } }
    $ws.Cells($r, 6).Formula = "=SUM(B${r}:E${r})"
    $ws.Cells($r, 7).Formula = "=E${r}/B${r}-1"
    if ($i % 2 -eq 1) { $ws.Range("A${r}:G${r}").Interior.Color = $band }
  }
  $ws.Cells(10, 1).Value2 = 'Total'
  foreach ($col in 'B', 'C', 'D', 'E', 'F') { $ws.Range("${col}10").Formula = "=SUM(${col}5:${col}9)" }
  $ws.Range('G10').Formula = '=E10/B10-1'
  $t = $ws.Range('A10:G10'); $t.Font.Bold = $true
  $t.Borders.Item(8).LineStyle = -4119; $t.Borders.Item(8).Color = $navy   # xlEdgeTop, xlDouble
  $t.Borders.Item(9).LineStyle = 1; $t.Borders.Item(9).Weight = -4138; $t.Borders.Item(9).Color = $navy   # xlEdgeBottom, medium

  $ws.Range('B5:F10').NumberFormat = '#,##0'
  # a colour in the format code ([Red]) is refused by a Japanese Excel through COM, so negative growth is coloured by the font instead
  $ws.Range('G5:G10').NumberFormat = '0.0%'
  foreach ($r in 5..10) { if ($ws.Cells($r, 7).Value2 -lt 0) { $ws.Cells($r, 7).Font.Color = Rgb 192 0 0 } }
  $body = $ws.Range('A5:G9')
  foreach ($b in 9, 11, 12) { $body.Borders.Item($b).LineStyle = 1; $body.Borders.Item($b).Color = $line }   # bottom, inside vertical, inside horizontal
  $ws.Range('F4:F10').Font.Bold = $true

  $ws.Columns.Item(1).ColumnWidth = 12
  foreach ($c in 2..6) { $ws.Columns.Item($c).ColumnWidth = 9 }
  $ws.Columns.Item(7).ColumnWidth = 9
  foreach ($r in 5..10) { $ws.Rows.Item($r).RowHeight = 20 }

  # clustered column chart of Q1..Q4 per region, right of the table
  $left = $ws.Range('I1').Left; $top = $ws.Range('I1').Top + 4
  $co = $ws.Shapes.AddChart2(-1, 51, $left, $top, 400, 236)
  $co.Name = 'salesChart'
  $ch = $co.Chart
  $ch.SetSourceData($ws.Range('A4:E9'))
  $ch.PlotBy = 2   # xlColumns: one series per quarter, regions on the category axis
  $ch.HasTitle = $true; $ch.ChartTitle.Text = 'Quarterly sales'

  # callout shape under the table (msoShapeRoundedRectangularCallout = 106)
  $cTop = $ws.Range('A12').Top + 6
  $s = $ws.Shapes.AddShape(106, $ws.Range('D12').Left, $cTop, 190, 44)
  $s.Name = 'onlineCallout'
  $s.Fill.ForeColor.RGB = Rgb 255 242 204
  $s.Line.ForeColor.RGB = Rgb 191 144 0
  $s.Adjustments.Item(1) = 0.35; $s.Adjustments.Item(2) = -1.05
  $tr = $s.TextFrame2.TextRange
  $tr.Text = 'Online grew 48% this year'
  $tr.Font.Size = 11; $tr.Font.Bold = 1; $tr.Font.Fill.ForeColor.RGB = Rgb 127 96 0
  $tr.ParagraphFormat.Alignment = 2   # centre
  $s.TextFrame2.VerticalAnchor = 3    # middle

  # second sheet, so the tabs show more than one sheet
  $notes.Range('A1').Value2 = 'Notes'
  $notes.Range('A1').Font.Bold = $true; $notes.Range('A1').Font.Size = 14
  $notes.Range('A3').Value2 = 'Sample data made for the XSheet README. The numbers are invented.'

  $ws.Activate(); [void]$ws.Range('A1').Select()
  if (Test-Path $out) { Remove-Item $out }
  $wb.SaveAs($out, 51)   # xlOpenXMLWorkbook
  $wb.Close($false)
  Write-Output "wrote $out"
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
