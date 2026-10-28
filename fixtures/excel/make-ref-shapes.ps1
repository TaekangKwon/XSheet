param([string]$Dir)
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$notes = New-Object System.Collections.ArrayList
$inputs = @{}
$ws = $null
$OFF = 100
function Inp($name, $obj) { $script:inputs[$name] = $obj }
function Try-Do($label, [scriptblock]$sb) { try { & $sb } catch { [void]$script:notes.Add("$label failed: $($_.Exception.Message)") } }
function Line2($name, $type, $x1, $y1, $x2, $y2) {
  $s = $script:ws.Shapes.AddConnector($type, $x1, $y1, $x2, $y2); $s.Name = $name
  Inp $name @{ api = 'AddConnector'; type = $type; x1 = $x1; y1 = $y1; x2 = $x2; y2 = $y2 }
  $s
}
function Safe($sb) { try { & $sb } catch { $null } }
function Rec($s) {
  $o = [ordered]@{}
  $o.Name = $s.Name; $o.Type = $s.Type
  $o.Left = $s.Left; $o.Top = $s.Top; $o.Width = $s.Width; $o.Height = $s.Height
  $o.Rotation = $s.Rotation; $o.HorizontalFlip = $s.HorizontalFlip; $o.VerticalFlip = $s.VerticalFlip
  $o.Visible = $s.Visible
  $o.IsConnector = [bool](Safe { $s.Connector })
  $o.AutoShapeType = Safe { $s.AutoShapeType }
  $o.AdjCount = Safe { $s.Adjustments.Count }
  if ($o.AdjCount) { $adj = @(); for ($i = 1; $i -le $o.AdjCount; $i++) { $adj += $s.Adjustments.Item($i) }; $o.Adjustments = $adj }
  $o.TopLeftCell = Safe { $s.TopLeftCell.Address($false, $false) }
  $o.BottomRightCell = Safe { $s.BottomRightCell.Address($false, $false) }
  $o.Placement = Safe { $s.Placement }
  if ($o.IsConnector) {
    $cf = $s.ConnectorFormat
    $o.ConnectorType = Safe { $cf.Type }
    $o.BeginConnected = Safe { $cf.BeginConnected }
    $o.EndConnected = Safe { $cf.EndConnected }
    if ($o.BeginConnected) { $o.BeginShape = $cf.BeginConnectedShape.Name; $o.BeginSite = $cf.BeginConnectionSite }
    if ($o.EndConnected) { $o.EndShape = $cf.EndConnectedShape.Name; $o.EndSite = $cf.EndConnectionSite }
  }
  if ($s.Type -ne 6 -and $s.Type -ne 3) {
    $l = $s.Line
    $o.Line = [ordered]@{ Visible = $l.Visible; Weight = $l.Weight; DashStyle = $l.DashStyle
      BeginArrowheadStyle = $l.BeginArrowheadStyle; BeginArrowheadLength = $l.BeginArrowheadLength; BeginArrowheadWidth = $l.BeginArrowheadWidth
      EndArrowheadStyle = $l.EndArrowheadStyle; EndArrowheadLength = $l.EndArrowheadLength; EndArrowheadWidth = $l.EndArrowheadWidth }
    $o.FillVisible = Safe { $s.Fill.Visible }
  }
  if ($s.Type -ne 6 -and $s.Type -ne 3 -and -not $o.IsConnector) {
    $has = Safe { $s.TextFrame2.HasText }
    $tf = Safe { $s.TextFrame2 }
    if ($tf) {
      $t = [ordered]@{ HasText = $has }
      $t.VerticalAnchor = $tf.VerticalAnchor; $t.Orientation = $tf.Orientation; $t.WordWrap = $tf.WordWrap; $t.AutoSize = $tf.AutoSize
      $t.MarginLeft = $tf.MarginLeft; $t.MarginTop = $tf.MarginTop; $t.MarginRight = $tf.MarginRight; $t.MarginBottom = $tf.MarginBottom
      if ($has) {
        $tr = $tf.TextRange
        $t.Text = $tr.Text
        $paras = New-Object System.Collections.ArrayList
        $n = ($tr.Text -split "`r").Count
        for ($i = 1; $i -le $n; $i++) {
          $p = $tr.Paragraphs($i)
          $pp = [ordered]@{ Text = $p.Text; Size = $p.Font.Size; Bold = $p.Font.Bold; Italic = $p.Font.Italic; Underline = $p.Font.UnderlineStyle; Strike = $p.Font.Strike
            Alignment = $p.ParagraphFormat.Alignment; BulletVisible = $p.ParagraphFormat.Bullet.Visible; BulletType = Safe { $p.ParagraphFormat.Bullet.Type } }
          $runs = New-Object System.Collections.ArrayList
          Try-Do 'runs' { $rc = $p.Runs().Count; for ($j = 1; $j -le $rc; $j++) { $r = $p.Runs($j); [void]$runs.Add([ordered]@{ Text = $r.Text; Size = $r.Font.Size; Bold = $r.Font.Bold; Italic = $r.Font.Italic }) } }
          $pp.Runs = $runs
          [void]$paras.Add($pp)
        }
        $t.Paragraphs = $paras
      }
      $o.TextFrame2 = $t
    }
    $o.Formula = Safe { $s.DrawingObject.Formula }
  }
  if ($s.Type -eq 6) {
    $items = New-Object System.Collections.ArrayList
    foreach ($g in $s.GroupItems) { [void]$items.Add((Rec $g)) }
    $o.GroupItems = $items
  }
  $o
}

try {
  $wb = $xl.Workbooks.Add()
  $wb.Styles.Item('Normal').Font.Name = 'Calibri'
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'shapes'
  $ws.Cells.Item(1, 1).Value2 = 'linked text'
  $ws.Cells.Item(2, 1).Value2 = 'a'; $ws.Cells.Item(2, 2).Value2 = 10
  $ws.Cells.Item(3, 1).Value2 = 'b'; $ws.Cells.Item(3, 2).Value2 = -3
  $ws.Cells.Item(4, 1).Value2 = 'c'; $ws.Cells.Item(4, 2).Value2 = 5
  $ws.Cells.Item(5, 1).Value2 = 'd'; $ws.Cells.Item(5, 2).Value2 = 8
  # 図形は pt 座標で配置。先頭 100pt（A1:B5 の表の下）から 55pt 刻みの帯に置く

  # ========== 1 直線（反転・回転） ==========
  $y = $OFF + 20
  [void](Line2 'lnPlain'  1 20 $y 100 ($y + 30))
  [void](Line2 'lnFlipV'  1 120 ($y + 30) 200 $y)
  [void](Line2 'lnFlipH'  1 320 $y 270 ($y + 30))
  [void](Line2 'lnFlipHV' 1 420 ($y + 30) 340 $y)
  $y = $OFF + 20 + 55 + 20
  $s = $ws.Shapes.AddLine(20, $y, 100, $y + 40); $s.Name = 'lnRot90';  $s.Rotation = 90;  Inp 'lnRot90'  @{ api = 'AddLine'; x1 = 20; y1 = $y; x2 = 100; y2 = $y + 40; rotation = 90 }
  $s = $ws.Shapes.AddLine(140, $y, 220, $y + 40); $s.Name = 'lnRot180'; $s.Rotation = 180; Inp 'lnRot180' @{ api = 'AddLine'; x1 = 140; y1 = $y; x2 = 220; y2 = $y + 40; rotation = 180 }
  $s = $ws.Shapes.AddLine(260, $y, 340, $y + 40); $s.Name = 'lnRot270'; $s.Rotation = 270; Inp 'lnRot270' @{ api = 'AddLine'; x1 = 260; y1 = $y; x2 = 340; y2 = $y + 40; rotation = 270 }
  $s = Line2 'lnFlipVRot90' 1 380 $y 460 ($y + 40); $s.Flip(1); $s.Rotation = 90; Inp 'lnFlipVRot90' @{ api = 'AddConnector+Flip(1)+Rotation'; type = 1; x1 = 380; y1 = $y; x2 = 460; y2 = $y + 40; flip = 'V'; rotation = 90 }
  $s = Line2 'lnFlipHRot180' 1 500 $y 580 ($y + 40); $s.Flip(0); $s.Rotation = 180; Inp 'lnFlipHRot180' @{ api = 'AddConnector+Flip(0)+Rotation'; type = 1; x1 = 500; y1 = $y; x2 = 580; y2 = $y + 40; flip = 'H'; rotation = 180 }
  $s = $ws.Shapes.AddLine(620, $y, 700, $y); $s.Name = 'lnHoriz'; Inp 'lnHoriz' @{ api = 'AddLine'; x1 = 620; y1 = $y; x2 = 700; y2 = $y }
  $s = $ws.Shapes.AddLine(720, $y, 720, $y + 40); $s.Name = 'lnVert'; Inp 'lnVert' @{ api = 'AddLine'; x1 = 720; y1 = $y; x2 = 720; y2 = $y + 40 }
  $s = $ws.Shapes.AddLine(760, $y, 840, $y + 40); $s.Name = 'lnRot45'; $s.Rotation = 45; Inp 'lnRot45' @{ api = 'AddLine'; x1 = 760; y1 = $y; x2 = 840; y2 = $y + 40; rotation = 45 }

  # ========== 2 カギ線・曲線（自由端） ==========
  $y = $OFF + 20 + 55 * 2
  $s = Line2 'bent3'      2 20 $y 120 ($y + 40);  $s.Line.EndArrowheadStyle = 2
  $s = Line2 'bent3FlipV' 2 160 ($y + 40) 260 $y; $s.Line.EndArrowheadStyle = 2
  $s = Line2 'bent3FlipH' 2 400 $y 300 ($y + 40); $s.Line.EndArrowheadStyle = 2
  $s = Line2 'bent3Adj'   2 440 $y 540 ($y + 40); $s.Line.EndArrowheadStyle = 2
  Try-Do 'bent3Adj' { $s.Adjustments.Item(1) = 0.3 }
  $s = Line2 'bent3AdjNeg' 2 580 $y 680 ($y + 40); $s.Line.EndArrowheadStyle = 2
  Try-Do 'bent3AdjNeg' { $s.Adjustments.Item(1) = -0.4 }
  $s = Line2 'bent3AdjBig' 2 720 $y 800 ($y + 40); $s.Line.EndArrowheadStyle = 2
  Try-Do 'bent3AdjBig' { $s.Adjustments.Item(1) = 1.6 }
  $y = $OFF + 20 + 55 * 3
  $s = Line2 'curve'      3 20 $y 120 ($y + 40);  $s.Line.EndArrowheadStyle = 4
  $s = Line2 'curveFlipV' 3 160 ($y + 40) 260 $y; $s.Line.EndArrowheadStyle = 4
  $s = Line2 'curveFlipH' 3 400 $y 300 ($y + 40); $s.Line.EndArrowheadStyle = 4
  $s = Line2 'curveAdj'   3 440 $y 540 ($y + 40); $s.Line.EndArrowheadStyle = 4
  Try-Do 'curveAdj' { $s.Adjustments.Item(1) = 0.2 }
  $s = Line2 'curveRot90' 3 580 $y 660 ($y + 40); $s.Rotation = 90

  # ========== 3 矢じり 5 種 × 大きさ ==========
  $y = $OFF + 20 + 55 * 4
  $styles = @(@(2, 'tri'), @(3, 'open'), @(4, 'stealth'), @(5, 'diamond'), @(6, 'oval'))
  $k = 0
  foreach ($st in $styles) {
    $s = Line2 ('arrow_' + $st[1]) 1 (20 + 160 * $k) ($y + 10) (150 + 160 * $k) ($y + 10)
    $s.Line.BeginArrowheadStyle = $st[0]; $s.Line.EndArrowheadStyle = $st[0]
    $s.Line.BeginArrowheadLength = 1; $s.Line.BeginArrowheadWidth = 1
    $s.Line.EndArrowheadLength = 3; $s.Line.EndArrowheadWidth = 3
    $k++
  }
  $y = $OFF + 20 + 55 * 5
  $k = 0
  foreach ($st in $styles) {
    $s = Line2 ('arrowMed_' + $st[1]) 1 (20 + 160 * $k) ($y + 10) (150 + 160 * $k) ($y + 10)
    $s.Line.EndArrowheadStyle = $st[0]; $s.Line.EndArrowheadLength = 2; $s.Line.EndArrowheadWidth = 2
    $k++
  }
  $y = $OFF + 20 + 55 * 6
  $s = Line2 'arrowShortWide' 1 20 ($y + 10) 150 ($y + 10);  $s.Line.EndArrowheadStyle = 2; $s.Line.EndArrowheadLength = 1; $s.Line.EndArrowheadWidth = 3
  $s = Line2 'arrowLongNarrow' 1 180 ($y + 10) 310 ($y + 10); $s.Line.EndArrowheadStyle = 2; $s.Line.EndArrowheadLength = 3; $s.Line.EndArrowheadWidth = 1
  $s = Line2 'arrowThick' 1 340 ($y + 10) 470 ($y + 10);  $s.Line.Weight = 5; $s.Line.EndArrowheadStyle = 2

  # ========== 4 太線・破線・点線・不可視 ==========
  $y = $OFF + 20 + 55 * 7
  $dash = @(@(1, 'solid'), @(2, 'squareDot'), @(3, 'roundDot'), @(4, 'dash'), @(5, 'dashDot'), @(6, 'dashDotDot'), @(7, 'longDash'), @(8, 'longDashDot'))
  $k = 0
  foreach ($d in $dash) {
    $s = $ws.Shapes.AddLine(20 + 100 * $k, $y, 100 + 100 * $k, $y); $s.Name = 'dash_' + $d[1]; $s.Line.DashStyle = $d[0]; $s.Line.Weight = 2
    Inp $s.Name @{ api = 'AddLine'; x1 = 20 + 100 * $k; y1 = $y; x2 = 100 + 100 * $k; y2 = $y; weight = 2 }
    $k++
  }
  $y = $OFF + 20 + 55 * 8
  $s = $ws.Shapes.AddLine(20, $y, 100, $y); $s.Name = 'wThin'; $s.Line.Weight = 0.25; Inp 'wThin' @{ api = 'AddLine'; weight = 0.25 }
  $s = $ws.Shapes.AddLine(120, $y, 200, $y); $s.Name = 'wDefault'; Inp 'wDefault' @{ api = 'AddLine' }
  $s = $ws.Shapes.AddLine(220, $y, 300, $y); $s.Name = 'w3pt'; $s.Line.Weight = 3; Inp 'w3pt' @{ api = 'AddLine'; weight = 3 }
  $s = $ws.Shapes.AddLine(320, $y, 400, $y); $s.Name = 'w6pt'; $s.Line.Weight = 6; Inp 'w6pt' @{ api = 'AddLine'; weight = 6 }
  $s = $ws.Shapes.AddLine(420, $y, 500, $y); $s.Name = 'w12pt'; $s.Line.Weight = 12; Inp 'w12pt' @{ api = 'AddLine'; weight = 12 }
  $s = $ws.Shapes.AddLine(520, $y, 600, $y); $s.Name = 'lnInvisible'; $s.Line.Visible = 0; Inp 'lnInvisible' @{ api = 'AddLine'; lineVisible = 0 }
  $s = $ws.Shapes.AddShape(1, 640, $y - 10, 60, 20); $s.Name = 'shapeHidden'; $s.TextFrame2.TextRange.Text = 'hidden'; $s.Visible = 0; Inp 'shapeHidden' @{ api = 'AddShape(1)+Visible=0' }

  # ========== 5 文字 ==========
  $y = $OFF + 20 + 55 * 9
  $t = $ws.Shapes.AddTextbox(1, 20, $y, 130, 70); $t.Name = 'txtMulti'
  $t.TextFrame2.TextRange.Text = "Title`rline2" + [char]11 + "brk`r`rlast"
  $t.TextFrame2.TextRange.Paragraphs(1).Font.Bold = -1
  $t.TextFrame2.TextRange.Paragraphs(1).ParagraphFormat.Alignment = 2
  $t.TextFrame2.VerticalAnchor = 3
  $t.TextFrame2.TextRange.Paragraphs(1).Font.Size = 14
  $t = $ws.Shapes.AddTextbox(1, 170, $y, 130, 70); $t.Name = 'txtStyle'
  $t.TextFrame2.TextRange.Text = "italic`runder`rstrike`r" + [string]([char]0x65E5) + [string]([char]0x672C) + [string]([char]0x8A9E) + " 8pt`rbig24"
  $t.TextFrame2.TextRange.Paragraphs(1).Font.Italic = -1
  $t.TextFrame2.TextRange.Paragraphs(2).Font.UnderlineStyle = 2
  $t.TextFrame2.TextRange.Paragraphs(3).Font.Strike = 1
  $t.TextFrame2.TextRange.Paragraphs(4).Font.Size = 8
  $t.TextFrame2.TextRange.Paragraphs(5).Font.Size = 24
  $t.TextFrame2.TextRange.ParagraphFormat.Alignment = 3
  $t.TextFrame2.VerticalAnchor = 4
  $t = $ws.Shapes.AddTextbox(1, 320, $y, 130, 40); $t.Name = 'txtRuns'
  $t.TextFrame2.TextRange.Text = 'AAA bbb CCC'
  $t.TextFrame2.TextRange.Characters(1, 3).Font.Bold = -1
  $t.TextFrame2.TextRange.Characters(9, 3).Font.Size = 20
  $t = $ws.Shapes.AddTextbox(1, 470, $y, 100, 60); $t.Name = 'txtJust'
  $t.TextFrame2.TextRange.Text = 'justify justify justify justify justify'
  $t.TextFrame2.TextRange.ParagraphFormat.Alignment = 4
  $t.TextFrame2.MarginLeft = 0; $t.TextFrame2.MarginTop = 0; $t.TextFrame2.MarginRight = 0; $t.TextFrame2.MarginBottom = 0
  $t = $ws.Shapes.AddTextbox(1, 590, $y, 100, 40); $t.Name = 'txtDist'
  $t.TextFrame2.TextRange.Text = 'dist'
  $t.TextFrame2.TextRange.ParagraphFormat.Alignment = 5

  # 入れ子の四角形（日本語は char コードで）
  $y = $OFF + 20 + 55 * 10 + 20
  $sysName = [string]([char]0x30B7) + [string]([char]0x30B9) + [string]([char]0x30C6) + [string]([char]0x30E0)
  $srvName = [string]([char]0x30B5) + [string]([char]0x30FC) + [string]([char]0x30D0)
  $o = $ws.Shapes.AddShape(1, 20, $y, 220, 110); $o.Name = 'outer'; $o.TextFrame2.TextRange.Text = $sysName; $o.TextFrame2.VerticalAnchor = 1
  $i = $ws.Shapes.AddShape(1, 50, $y + 40, 90, 30); $i.Name = 'inner'; $i.TextFrame2.TextRange.Text = $srvName
  $i2 = $ws.Shapes.AddShape(1, 150, $y + 40, 70, 30); $i2.Name = 'inner2'; $i2.TextFrame2.TextRange.Text = 'DB'

  # ラベルほか
  $lb = $ws.Shapes.AddTextbox(1, 270, $y, 60, 20); $lb.Name = 'label'; $lb.TextFrame2.TextRange.Text = 'HTTPS'; $lb.Line.Visible = 0; $lb.Fill.Visible = 0
  $nl = $ws.Shapes.AddShape(1, 350, $y, 70, 24); $nl.Name = 'noLineRect'; $nl.Line.Visible = 0; $nl.TextFrame2.TextRange.Text = 'noLine'
  $nf = $ws.Shapes.AddShape(1, 440, $y, 70, 24); $nf.Name = 'noFillRect'; $nf.Fill.Visible = 0; $nf.TextFrame2.TextRange.Text = 'noFill'
  $dr = $ws.Shapes.AddShape(1, 530, $y, 70, 24); $dr.Name = 'defaultRect'; $dr.TextFrame2.TextRange.Text = 'default'
  $em = $ws.Shapes.AddShape(1, 620, $y, 70, 24); $em.Name = 'emptyRect'
  $tbd = $ws.Shapes.AddTextbox(1, 710, $y, 70, 24); $tbd.Name = 'emptyTextbox'
  $tbd2 = $ws.Shapes.AddTextbox(1, 800, $y, 70, 24); $tbd2.Name = 'textboxBordered'; $tbd2.Line.Visible = -1; $tbd2.TextFrame2.TextRange.Text = 'tb border'

  # 角丸・楕円・ひし形・ブロック矢印・折り返し無し・自動調整・余白
  $y = $OFF + 20 + 55 * 12 + 10
  $a1 = $ws.Shapes.AddShape(5, 20, $y, 80, 40); $a1.Name = 'roundRect'; $a1.TextFrame2.TextRange.Text = 'R'; $a1.TextFrame2.VerticalAnchor = 4; $a1.TextFrame2.TextRange.ParagraphFormat.Alignment = 3; $a1.TextFrame2.MarginLeft = 0; $a1.TextFrame2.MarginTop = 0
  $a2 = $ws.Shapes.AddShape(9, 120, $y, 80, 40); $a2.Name = 'ellipse'; $a2.TextFrame2.TextRange.Text = 'oval'
  $a3 = $ws.Shapes.AddShape(4, 220, $y, 80, 40); $a3.Name = 'diamond'; $a3.TextFrame2.TextRange.Text = 'dia'
  $a4 = $ws.Shapes.AddShape(33, 320, $y, 80, 40); $a4.Name = 'rightArrowBlock'; $a4.TextFrame2.TextRange.Text = 'go'
  $a5 = $ws.Shapes.AddShape(1, 420, $y, 80, 40); $a5.Name = 'wrapNone'; $a5.TextFrame2.TextRange.Text = 'a long text without wrap'; $a5.TextFrame2.WordWrap = 0
  $a6 = $ws.Shapes.AddTextbox(1, 520, $y, 40, 20); $a6.Name = 'noWrapTb'; $a6.TextFrame2.TextRange.Text = 'a long text without wrap'; $a6.TextFrame2.WordWrap = 0
  $a7 = $ws.Shapes.AddTextbox(1, 580, $y, 60, 40); $a7.Name = 'autofitShrink'; $a7.TextFrame2.TextRange.Text = 'shrink this long text please'; $a7.TextFrame2.AutoSize = 2
  $a8 = $ws.Shapes.AddTextbox(1, 660, $y, 60, 40); $a8.Name = 'autofitResize'; $a8.TextFrame2.TextRange.Text = 'resize shape'; $a8.TextFrame2.AutoSize = 1
  $a9 = $ws.Shapes.AddShape(1, 740, $y, 100, 40); $a9.Name = 'insets'; $a9.TextFrame2.TextRange.Text = 'insets'; $a9.TextFrame2.MarginLeft = 20; $a9.TextFrame2.MarginTop = 10; $a9.TextFrame2.MarginRight = 5; $a9.TextFrame2.MarginBottom = 2

  # セル参照・箇条書き・縦書き
  $y = $OFF + 20 + 55 * 13 + 20
  $lk = $ws.Shapes.AddShape(1, 20, $y, 100, 24); $lk.Name = 'linkedText'
  Try-Do 'linkedText' { $lk.DrawingObject.Formula = '=$A$1' }
  $bl = $ws.Shapes.AddTextbox(1, 140, $y, 100, 50); $bl.Name = 'bullets'; $bl.TextFrame2.TextRange.Text = "one`rtwo`rthree"
  $bl.TextFrame2.TextRange.ParagraphFormat.Bullet.Visible = -1
  Try-Do 'bullets' { $bl.TextFrame2.TextRange.ParagraphFormat.Bullet.Type = 1; $bl.TextFrame2.TextRange.ParagraphFormat.Bullet.Character = 8226 }
  $nb = $ws.Shapes.AddTextbox(1, 260, $y, 100, 50); $nb.Name = 'numbered'; $nb.TextFrame2.TextRange.Text = "first`rsecond`rthird`rplain"
  Try-Do 'numbered' {
    $r = $nb.TextFrame2.TextRange.Paragraphs(1, 3).ParagraphFormat.Bullet
    $r.Visible = -1; $r.Type = 2; $r.Style = 3; $r.StartValue = 3
  }
  $vt = @(@(2, 'vertUp'), @(3, 'vertDown'), @(4, 'vertFarEast'), @(5, 'vertStack'), @(6, 'vertHorzRotFE'))
  $k = 0
  foreach ($v in $vt) {
    $tb = $ws.Shapes.AddTextbox(1, 400 + 70 * $k, $y, 30, 90); $tb.Name = $v[1]
    $tb.TextFrame2.TextRange.Text = 'ABC' + [string]([char]0x7E26) + [string]([char]0x66F8) + [string]([char]0x304D)
    $ori = $v[0]; $nm = $v[1]
    Try-Do $nm { $tb.TextFrame2.Orientation = $ori }
    $k++
  }

  # 回転・反転した文字入りの図形
  $y = $OFF + 20 + 55 * 16
  $r1 = $ws.Shapes.AddShape(1, 20, $y, 100, 30); $r1.Name = 'rot90'; $r1.TextFrame2.TextRange.Text = 'r90'; $r1.Rotation = 90
  $r2 = $ws.Shapes.AddShape(1, 170, $y, 100, 30); $r2.Name = 'rot30'; $r2.TextFrame2.TextRange.Text = 'r30'; $r2.Rotation = 30
  $r3 = $ws.Shapes.AddShape(1, 320, $y, 100, 30); $r3.Name = 'rot270'; $r3.TextFrame2.TextRange.Text = 'r270'; $r3.Rotation = 270
  $r4 = $ws.Shapes.AddShape(1, 470, $y, 100, 30); $r4.Name = 'rot180'; $r4.TextFrame2.TextRange.Text = 'r180'; $r4.Rotation = 180
  $r5 = $ws.Shapes.AddTextbox(1, 620, $y, 100, 30); $r5.Name = 'rotTb45'; $r5.TextFrame2.TextRange.Text = 'tb45'; $r5.Rotation = 45
  $r6 = $ws.Shapes.AddShape(1, 770, $y, 100, 30); $r6.Name = 'flipHText'; $r6.TextFrame2.TextRange.Text = 'flipH'; $r6.Flip(0)
  $r7 = $ws.Shapes.AddShape(1, 920, $y, 100, 30); $r7.Name = 'flipVText'; $r7.TextFrame2.TextRange.Text = 'flipV'; $r7.Flip(1)

  # ========== 6 接続した線 ==========
  $y = $OFF + 20 + 55 * 18
  $src = $ws.Shapes.AddTextbox(1, 20, $y, 70, 24); $src.Name = 'srcTb'; $src.TextFrame2.TextRange.Text = 'client'
  $dst = $ws.Shapes.AddShape(1, 150, $y + 60, 70, 30); $dst.Name = 'dstRect'; $dst.TextFrame2.TextRange.Text = 'API'
  $k = Line2 'linkElbow' 2 0 0 10 10;    $k.ConnectorFormat.BeginConnect($src, 3); $k.ConnectorFormat.EndConnect($dst, 2); $k.Line.EndArrowheadStyle = 2
  $k = Line2 'linkStraight' 1 0 0 10 10; $k.ConnectorFormat.BeginConnect($src, 4); $k.ConnectorFormat.EndConnect($dst, 1); $k.Line.EndArrowheadStyle = 2
  $k = Line2 'linkCurve' 3 0 0 10 10;    $k.ConnectorFormat.BeginConnect($src, 3); $k.ConnectorFormat.EndConnect($dst, 1); $k.Line.EndArrowheadStyle = 2
  $ua = $ws.Shapes.AddShape(1, 300, $y, 60, 20); $ua.Name = 'uA'
  $ub = $ws.Shapes.AddShape(1, 300, $y + 60, 100, 20); $ub.Name = 'uB'
  $k = Line2 'uRightRight' 2 0 0 10 10; $k.ConnectorFormat.BeginConnect($ua, 4); $k.ConnectorFormat.EndConnect($ub, 4); $k.Line.EndArrowheadStyle = 2; $k.Line.DashStyle = 4
  $k = Line2 'uLeftLeft' 2 0 0 10 10;   $k.ConnectorFormat.BeginConnect($ua, 2); $k.ConnectorFormat.EndConnect($ub, 2); $k.Line.EndArrowheadStyle = 2
  $uc = $ws.Shapes.AddShape(1, 480, $y + 60, 60, 20); $uc.Name = 'uC'
  $ud = $ws.Shapes.AddShape(1, 600, $y + 60, 60, 20); $ud.Name = 'uD'
  $k = Line2 'uTopTop' 2 0 0 10 10;     $k.ConnectorFormat.BeginConnect($uc, 1); $k.ConnectorFormat.EndConnect($ud, 1); $k.Line.EndArrowheadStyle = 2
  $k = Line2 'uBottomBottom' 2 0 0 10 10; $k.ConnectorFormat.BeginConnect($uc, 3); $k.ConnectorFormat.EndConnect($ud, 3); $k.Line.EndArrowheadStyle = 2
  $v1 = $ws.Shapes.AddShape(1, 720, $y, 60, 24); $v1.Name = 'vTop'
  $v2 = $ws.Shapes.AddShape(1, 780, $y + 70, 60, 24); $v2.Name = 'vBottom'
  $k = Line2 'vElbow' 2 0 0 10 10;       $k.ConnectorFormat.BeginConnect($v1, 3); $k.ConnectorFormat.EndConnect($v2, 1); $k.Line.EndArrowheadStyle = 2
  $k = Line2 'bentRightTop' 2 0 0 10 10; $k.ConnectorFormat.BeginConnect($v1, 4); $k.ConnectorFormat.EndConnect($v2, 1); $k.Line.EndArrowheadStyle = 2
  $w1 = $ws.Shapes.AddShape(1, 900, $y + 60, 60, 24); $w1.Name = 'wA'
  $w2 = $ws.Shapes.AddShape(1, 860, $y, 60, 24); $w2.Name = 'wB'
  $k = Line2 'bentTopBottom' 2 0 0 10 10;    $k.ConnectorFormat.BeginConnect($w1, 1); $k.ConnectorFormat.EndConnect($w2, 3); $k.Line.EndArrowheadStyle = 2
  $k = Line2 'bentLeftRight' 2 0 0 10 10;    $k.ConnectorFormat.BeginConnect($w1, 2); $k.ConnectorFormat.EndConnect($w2, 4); $k.Line.EndArrowheadStyle = 2
  $k = Line2 'bentBack' 2 0 0 10 10;         $k.ConnectorFormat.BeginConnect($w2, 4); $k.ConnectorFormat.EndConnect($w1, 2); $k.Line.EndArrowheadStyle = 2

  # ---- 6b bent4・bent5・curved2/4/5（B は A の右下 +130,+80。接続点 1 上 2 左 3 下 4 右）----
  $y = $OFF + 20 + 55 * 20
  $combos = @(@('bent4', 2, 1, 2), @('curved4', 3, 1, 2), @('bent5', 2, 1, 3), @('curved5', 3, 1, 3), @('curved2', 3, 3, 2))
  $q = 0
  foreach ($cb in $combos) {
    $ax = 20 + 240 * $q
    $pa = $ws.Shapes.AddShape(1, $ax, $y, 60, 30); $pa.Name = 'pA_' + $cb[0]
    $pb = $ws.Shapes.AddShape(1, $ax + 130, $y + 80, 60, 30); $pb.Name = 'pB_' + $cb[0]
    $k = Line2 $cb[0] $cb[1] 0 0 10 10; $k.ConnectorFormat.BeginConnect($pa, $cb[2]); $k.ConnectorFormat.EndConnect($pb, $cb[3]); $k.Line.EndArrowheadStyle = 2
    $q++
  }

  # ========== 7 グラフ ==========
  $y = $OFF + 20 + 55 * 23
  $ws.Range('A2:B5').Select() | Out-Null
  Try-Do 'chartBar' { $c = $ws.Shapes.AddChart2(-1, 51, 20, $y, 280, 180); $c.Name = 'chartBar' }
  $ws.Range('A2:B5').Select() | Out-Null
  Try-Do 'chartexWaterfall' { $c = $ws.Shapes.AddChart2(-1, 119, 330, $y, 300, 180); $c.Name = 'chartexWaterfall' }

  # ========== 8 グループ ==========
  $y = $OFF + 20 + 55 * 27
  $g1 = $ws.Shapes.AddTextbox(1, 20, $y, 60, 20); $g1.Name = 'gTb'; $g1.TextFrame2.TextRange.Text = 'gtext'
  $g2 = $ws.Shapes.AddShape(1, 140, $y + 30, 60, 20); $g2.Name = 'gRect'; $g2.TextFrame2.TextRange.Text = 'grect'
  $g3 = Line2 'gConn' 1 0 0 10 10; $g3.ConnectorFormat.BeginConnect($g1, 4); $g3.ConnectorFormat.EndConnect($g2, 2); $g3.Line.EndArrowheadStyle = 2
  $grp = $ws.Shapes.Range([object[]]@('gTb', 'gRect', 'gConn')).Group(); $grp.Name = 'grpScaled'
  Try-Do 'grpScaled' { $grp.Width = $grp.Width * 1.5 }
  $h1 = $ws.Shapes.AddShape(1, 320, $y, 60, 30); $h1.Name = 'hRect'; $h1.TextFrame2.TextRange.Text = 'h1'
  $h2 = $ws.Shapes.AddShape(9, 400, $y + 20, 60, 30); $h2.Name = 'hOval'; $h2.TextFrame2.TextRange.Text = 'h2'
  $h3 = $ws.Shapes.AddLine(320, $y + 60, 460, $y + 60); $h3.Name = 'hLine'; $h3.Line.EndArrowheadStyle = 2
  $grp2 = $ws.Shapes.Range([object[]]@('hRect', 'hOval', 'hLine')).Group(); $grp2.Name = 'grpPlain'
  $i1 = $ws.Shapes.AddShape(1, 520, $y, 60, 30); $i1.Name = 'iRect'; $i1.TextFrame2.TextRange.Text = 'i1'
  $i2b = $ws.Shapes.AddShape(1, 600, $y + 20, 60, 30); $i2b.Name = 'iRect2'; $i2b.TextFrame2.TextRange.Text = 'i2'
  $grp3 = $ws.Shapes.Range([object[]]@('iRect', 'iRect2')).Group(); $grp3.Name = 'grpRot'; $grp3.Rotation = 30
  $n1 = $ws.Shapes.AddShape(1, 720, $y, 50, 30); $n1.Name = 'nRect1'; $n1.TextFrame2.TextRange.Text = 'n1'
  $n2 = $ws.Shapes.AddShape(1, 780, $y, 50, 30); $n2.Name = 'nRect2'; $n2.TextFrame2.TextRange.Text = 'n2'
  $gi = $ws.Shapes.Range([object[]]@('nRect1', 'nRect2')).Group(); $gi.Name = 'grpInner'
  $n3 = $ws.Shapes.AddShape(1, 720, $y + 50, 110, 30); $n3.Name = 'nRect3'; $n3.TextFrame2.TextRange.Text = 'n3'
  Try-Do 'grpNested' { $go = $ws.Shapes.Range([object[]]@('grpInner', 'nRect3')).Group(); $go.Name = 'grpOuter' }

  # ========== 記録 ==========
  $list = New-Object System.Collections.ArrayList
  foreach ($sh in $ws.Shapes) { [void]$list.Add((Rec $sh)) }
  $wb.SaveAs((Join-Path $Dir 'ref-shapes.xlsx'), 51); $wb.Close($false)
  $out = [ordered]@{ excelVersion = $xl.Version; notes = $notes; inputs = $inputs; shapes = $list }
  $json = $out | ConvertTo-Json -Depth 12
  [IO.File]::WriteAllText((Join-Path $Dir 'com.json'), $json, (New-Object Text.UTF8Encoding($false)))
  "ok: $($list.Count) shapes; notes=$($notes.Count)"
  $notes
} finally { try { $xl.Quit() } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
