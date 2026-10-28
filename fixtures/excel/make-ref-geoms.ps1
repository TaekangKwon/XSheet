param([string]$Dir)
# ISSUE-91: ref-geoms.xlsx (Excel 16 COM). Shapes: every AutoShapeType 1..183, adjusted presets, rotation/flip, text in oval/diamond,
# freeforms (BuildFreeform / AddCurve / AddPolyline), outline weight/dash. Writes ref-geoms.xlsx and com.json into $Dir.
# Run: powershell -File make-ref-geoms.ps1 -Dir <out dir>   (PowerShell 5.1, file saved as UTF-8 with BOM)
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$notes = New-Object System.Collections.ArrayList
$inputs = @{}
$ws = $null
function Inp($name, $obj) { $script:inputs[$name] = $obj }
function Try-Do($label, [scriptblock]$sb) { try { & $sb } catch { [void]$script:notes.Add("$label failed: $($_.Exception.Message)") } }
function Safe($sb) { try { & $sb } catch { $null } }
# style: no fill, 1pt outline
function Plain($s) { $s.Fill.Visible = 0; $s.Line.Weight = 1 }
function Rec($s) {
  $o = [ordered]@{}
  $o.Name = $s.Name; $o.Type = $s.Type
  $o.Left = $s.Left; $o.Top = $s.Top; $o.Width = $s.Width; $o.Height = $s.Height
  $o.Rotation = $s.Rotation; $o.HorizontalFlip = $s.HorizontalFlip; $o.VerticalFlip = $s.VerticalFlip
  $o.AutoShapeType = Safe { $s.AutoShapeType }
  $o.AdjCount = Safe { $s.Adjustments.Count }
  if ($o.AdjCount) { $adj = @(); for ($i = 1; $i -le $o.AdjCount; $i++) { $adj += $s.Adjustments.Item($i) }; $o.Adjustments = $adj }
  $o.NodeCount = Safe { $s.Nodes.Count }
  $l = $s.Line
  $o.Line = [ordered]@{ Visible = $l.Visible; Weight = $l.Weight; DashStyle = $l.DashStyle; BeginArrowheadStyle = $l.BeginArrowheadStyle; EndArrowheadStyle = $l.EndArrowheadStyle }
  $o.FillVisible = Safe { $s.Fill.Visible }
  $o.HasText = Safe { $s.TextFrame2.HasText }
  $o.Text = Safe { if ($s.TextFrame2.HasText) { $s.TextFrame2.TextRange.Text } else { $null } }
  $o
}

try {
  $wb = $xl.Workbooks.Add()
  $wb.Styles.Item('Normal').Font.Name = 'Calibri'
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'geoms'
  $ws.Cells.Item(1, 1).Value2 = 'ref-geoms'
  $W = 50; $H = 40

  # ========== 1 every AutoShapeType 1..183, 13 per row, pitch 70 x 60pt, from y=60 ==========
  for ($t = 1; $t -le 183; $t++) {
    $k = $t - 1
    $x = 20 + 70 * ($k % 13); $y = 60 + 60 * [math]::Floor($k / 13)
    Try-Do ("preset$t") {
      $s = $ws.Shapes.AddShape($t, $x, $y, $W, $H); $s.Name = 'p' + $t; Plain $s
      Inp $s.Name @{ api = 'AddShape'; type = $t; x = $x; y = $y; w = $W; h = $H }
    }
  }

  # ========== 2 adjusted presets (y = 60 + 60*15 + 40) ==========
  $y = 1000
  $s = $ws.Shapes.AddShape(33, 20, $y, 90, 40); $s.Name = 'adjRightArrow'; Plain $s
  Try-Do 'adjRightArrow' { $s.Adjustments.Item(1) = 0.3; $s.Adjustments.Item(2) = 0.8 }
  Inp 'adjRightArrow' @{ api = 'AddShape(33)+Adjustments'; adj = @(0.3, 0.8) }
  $s = $ws.Shapes.AddShape(5, 130, $y, 90, 40); $s.Name = 'adjRoundRectSmall'; Plain $s
  Try-Do 'adjRoundRectSmall' { $s.Adjustments.Item(1) = 0.1 }
  Inp 'adjRoundRectSmall' @{ api = 'AddShape(5)+Adjustments'; adj = @(0.1) }
  $s = $ws.Shapes.AddShape(5, 240, $y, 90, 40); $s.Name = 'adjRoundRectMax'; Plain $s
  Try-Do 'adjRoundRectMax' { $s.Adjustments.Item(1) = 0.5 }
  Inp 'adjRoundRectMax' @{ api = 'AddShape(5)+Adjustments'; adj = @(0.5) }
  $s = $ws.Shapes.AddShape(105, 350, $y, 90, 40); $s.Name = 'adjCallout'; Plain $s
  Try-Do 'adjCallout' { $s.Adjustments.Item(1) = -0.4; $s.Adjustments.Item(2) = 0.9 }
  Inp 'adjCallout' @{ api = 'AddShape(105)+Adjustments'; adj = @(-0.4, 0.9) }
  $s = $ws.Shapes.AddShape(13, 460, $y, 60, 60); $s.Name = 'adjCan'; Plain $s
  Try-Do 'adjCan' { $s.Adjustments.Item(1) = 0.5 }
  Inp 'adjCan' @{ api = 'AddShape(13)+Adjustments'; adj = @(0.5) }
  $s = $ws.Shapes.AddShape(33, 540, $y, 90, 40); $s.Name = 'adjRightArrowNeg'; Plain $s
  Try-Do 'adjRightArrowNeg' { $s.Adjustments.Item(1) = 0.9; $s.Adjustments.Item(2) = 0.1 }
  Inp 'adjRightArrowNeg' @{ api = 'AddShape(33)+Adjustments'; adj = @(0.9, 0.1) }

  # ========== 3 rotation / flip ==========
  $y = 1100
  $s = $ws.Shapes.AddShape(33, 20, $y, 90, 40); $s.Name = 'rot30'; Plain $s; $s.Rotation = 30
  $s = $ws.Shapes.AddShape(33, 130, $y, 90, 40); $s.Name = 'flipH'; Plain $s; $s.Flip(0)
  $s = $ws.Shapes.AddShape(105, 240, $y, 90, 40); $s.Name = 'flipV'; Plain $s; $s.Flip(1)
  $s = $ws.Shapes.AddShape(13, 350, $y, 60, 60); $s.Name = 'rot90'; Plain $s; $s.Rotation = 90
  $s = $ws.Shapes.AddShape(33, 450, $y, 90, 40); $s.Name = 'flipHRot45'; Plain $s; $s.Flip(0); $s.Rotation = 45
  $s = $ws.Shapes.AddShape(33, 580, $y, 90, 40); $s.Name = 'flipHVRot270'; Plain $s; $s.Flip(0); $s.Flip(1); $s.Rotation = 270

  # ========== 4 text in oval / diamond (default fill and outline) ==========
  $y = 1200
  $s = $ws.Shapes.AddShape(9, 20, $y, 90, 50); $s.Name = 'textOval'; $s.TextFrame2.TextRange.Text = 'oval'
  $s = $ws.Shapes.AddShape(4, 130, $y, 90, 50); $s.Name = 'textDiamond'; $s.TextFrame2.TextRange.Text = 'dia'
  $s = $ws.Shapes.AddShape(9, 240, $y, 90, 50); $s.Name = 'emptyOval'
  $s = $ws.Shapes.AddShape(4, 350, $y, 90, 50); $s.Name = 'emptyDiamond'; $s.Line.Visible = 0; $s.Fill.Visible = 0

  # ========== 5 freeforms ==========
  $y = 1300
  Try-Do 'ffTriangle' {
    $b = $ws.Shapes.BuildFreeform(0, 20, $y + 40)
    $b.AddNodes(0, 0, 70, $y); $b.AddNodes(0, 0, 120, $y + 40); $b.AddNodes(0, 0, 20, $y + 40)
    $s = $b.ConvertToShape(); $s.Name = 'ffTriangle'; Plain $s
  }
  Try-Do 'ffCurve' {
    $b = $ws.Shapes.BuildFreeform(1, 150, $y + 40)
    $b.AddNodes(1, 1, 170, $y - 10, 210, $y + 90, 250, $y + 40)
    $s = $b.ConvertToShape(); $s.Name = 'ffCurve'; Plain $s
  }
  Try-Do 'ffMixed' {
    $b = $ws.Shapes.BuildFreeform(0, 280, $y + 40)
    $b.AddNodes(0, 0, 280, $y); $b.AddNodes(1, 0, 340, $y); $b.AddNodes(0, 0, 340, $y + 40); $b.AddNodes(0, 0, 280, $y + 40)
    $s = $b.ConvertToShape(); $s.Name = 'ffMixed'; Plain $s
  }
  Try-Do 'curve4' {
    $pts = New-Object 'single[,]' 4, 2
    $pts[0, 0] = 380; $pts[0, 1] = $y + 40; $pts[1, 0] = 400; $pts[1, 1] = $y - 10
    $pts[2, 0] = 440; $pts[2, 1] = $y + 90; $pts[3, 0] = 480; $pts[3, 1] = $y + 40
    $s = $ws.Shapes.AddCurve($pts); $s.Name = 'curve4'; Plain $s
  }
  Try-Do 'curve7' {
    $pts = New-Object 'single[,]' 7, 2
    $xs = @(510, 520, 550, 570, 590, 620, 630); $ys = @(0, -20, -20, 0, 20, 20, 0)
    for ($i = 0; $i -lt 7; $i++) { $pts[$i, 0] = $xs[$i]; $pts[$i, 1] = $y + 20 + $ys[$i] }
    $s = $ws.Shapes.AddCurve($pts); $s.Name = 'curve7'; Plain $s
  }
  $y = 1400
  Try-Do 'polyOpen' {
    $pts = New-Object 'single[,]' 4, 2
    $pts[0, 0] = 20; $pts[0, 1] = $y; $pts[1, 0] = 60; $pts[1, 1] = $y + 40
    $pts[2, 0] = 100; $pts[2, 1] = $y; $pts[3, 0] = 140; $pts[3, 1] = $y + 40
    $s = $ws.Shapes.AddPolyline($pts); $s.Name = 'polyOpen'; Plain $s
  }
  Try-Do 'polyClosed' {
    $pts = New-Object 'single[,]' 5, 2
    $pts[0, 0] = 170; $pts[0, 1] = $y; $pts[1, 0] = 230; $pts[1, 1] = $y; $pts[2, 0] = 230; $pts[2, 1] = $y + 40
    $pts[3, 0] = 200; $pts[3, 1] = $y + 20; $pts[4, 0] = 170; $pts[4, 1] = $y
    $s = $ws.Shapes.AddPolyline($pts); $s.Name = 'polyClosed'; Plain $s
  }
  Try-Do 'polyArrow' {
    $pts = New-Object 'single[,]' 3, 2
    $pts[0, 0] = 260; $pts[0, 1] = $y + 40; $pts[1, 0] = 300; $pts[1, 1] = $y; $pts[2, 0] = 340; $pts[2, 1] = $y + 40
    $s = $ws.Shapes.AddPolyline($pts); $s.Name = 'polyArrow'; Plain $s
    $s.Line.EndArrowheadStyle = 2; $s.Line.BeginArrowheadStyle = 5
  }
  Try-Do 'polyFilled' {
    $pts = New-Object 'single[,]' 4, 2
    $pts[0, 0] = 370; $pts[0, 1] = $y + 40; $pts[1, 0] = 410; $pts[1, 1] = $y; $pts[2, 0] = 450; $pts[2, 1] = $y + 40; $pts[3, 0] = 370; $pts[3, 1] = $y + 40
    $s = $ws.Shapes.AddPolyline($pts); $s.Name = 'polyFilled'
  }
  Try-Do 'curveArrow' {
    $pts = New-Object 'single[,]' 4, 2
    $pts[0, 0] = 480; $pts[0, 1] = $y + 40; $pts[1, 0] = 500; $pts[1, 1] = $y - 10; $pts[2, 0] = 540; $pts[2, 1] = $y + 90; $pts[3, 0] = 580; $pts[3, 1] = $y + 40
    $s = $ws.Shapes.AddCurve($pts); $s.Name = 'curveArrow'; Plain $s; $s.Line.EndArrowheadStyle = 2
  }
  Try-Do 'ffRot' {
    $b = $ws.Shapes.BuildFreeform(0, 620, $y + 40)
    $b.AddNodes(0, 0, 660, $y); $b.AddNodes(0, 0, 700, $y + 40); $b.AddNodes(0, 0, 620, $y + 40)
    $s = $b.ConvertToShape(); $s.Name = 'ffRot'; Plain $s; $s.Rotation = 40; $s.Flip(0)
  }

  # ========== 6 outline weight / dash ==========
  $y = 1500
  $wts = @(0.25, 1, 3, 6, 12)
  for ($i = 0; $i -lt $wts.Count; $i++) {
    $s = $ws.Shapes.AddShape(1, 20 + 100 * $i, $y, 80, 40); $s.Name = 'w_' + $i; $s.Fill.Visible = 0; $s.Line.Weight = [single]$wts[$i]
    Inp $s.Name @{ api = 'AddShape(1)'; weight = $wts[$i] }
  }
  $y = 1600
  $dash = @(@(1, 'solid'), @(2, 'squareDot'), @(3, 'roundDot'), @(4, 'dash'), @(5, 'dashDot'), @(6, 'dashDotDot'), @(7, 'longDash'), @(8, 'longDashDot'))
  for ($i = 0; $i -lt $dash.Count; $i++) {
    $s = $ws.Shapes.AddShape(1, 20 + 100 * $i, $y, 80, 40); $s.Name = 'd_' + $dash[$i][1]; $s.Fill.Visible = 0; $s.Line.Weight = 2; $s.Line.DashStyle = $dash[$i][0]
    Inp $s.Name @{ api = 'AddShape(1)'; weight = 2; dashStyle = $dash[$i][0] }
  }
  $s = $ws.Shapes.AddShape(1, 20, 1700, 80, 40); $s.Name = 'lineNone'; $s.Fill.Visible = 0; $s.Line.Visible = 0
  $s = $ws.Shapes.AddShape(1, 120, 1700, 80, 40); $s.Name = 'defaultRect'
  $s = $ws.Shapes.AddTextbox(1, 220, 1700, 80, 40); $s.Name = 'tbBordered'; $s.Line.Visible = -1; $s.Line.Weight = 3

  # ========== record ==========
  $list = New-Object System.Collections.ArrayList
  foreach ($sh in $ws.Shapes) { [void]$list.Add((Rec $sh)) }
  $wb.SaveAs((Join-Path $Dir 'ref-geoms.xlsx'), 51); $wb.Close($false)
  $out = [ordered]@{ excelVersion = $xl.Version; notes = $notes; inputs = $inputs; shapes = $list }
  $json = $out | ConvertTo-Json -Depth 12
  [IO.File]::WriteAllText((Join-Path $Dir 'com.json'), $json, (New-Object Text.UTF8Encoding($false)))
  "ok: $($list.Count) shapes; notes=$($notes.Count)"
  $notes
} finally { try { $xl.Quit() } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
