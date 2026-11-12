# Issue #25：曜日の書式（aaa・aaaa）の正解データを Excel 16（ja-JP、COM）で作る。
# 作るもの（$Dir に書く。ほかのファイルは触らない）：
#   ref-weekday.xlsx            1900 年基準。書式のグループごとのシート（basic・count・upper・skip・ddd・mixed）と、式のシート weekday
#   ref-weekday-1904.xlsx       1904 年基準（Workbook.Date1904 = $true）。シートの構成は同じ
#   ref-weekday.expected.json   セルごとの { sheet, cell, serial, format, formatRead, formatCode, text } と、ブックが 1904 基準か
# シートの形：1 行目＝書式コード（文字列）、A 列＝シリアル値（標準の書式）、B 列から右＝各書式を設定したセル。
# 書式は NumberFormat（英語側の表記）に設定し、読み戻した NumberFormat を formatRead に、NumberFormatLocal を formats[].formatLocal に残す。
# 設定でエラーになったら同じ文字列を NumberFormatLocal で試す（通れば setVia = 'NumberFormatLocal'）。それも通らなければ formats[].error に残し、その列には値を入れない。
# formatCode／numFmtId は保存した xlsx の styles.xml から読んだもの（XSheet が実際に読む文字列）。NumberFormat の読み戻しと違うことがある。
# ja-JP の COM の NumberFormat では「\」が円記号の文字（styles.xml では "¥"）になり、1 文字の逃がしは「!」になる（styles.xml では \）。
# そのため \ の逃がしは '!aaaa'（styles.xml では \aaaa）で試す。
# 埋め（*）の .Text は列幅で変わるので、セルごとに columnWidth も残す。
# 列幅：AutoFit のあと「#」だけの表示が残るセルは、幅 255 でも「#」のままか（＝値が日付の範囲外）を確かめ、
#       それ以外の「#」が消える幅まで広げてから .Text を読む。幅 255 でも「#」のセルは hashAtMaxWidth = true。
# 実行（CI では実行しない。Excel 16 が要る）：
#   powershell -ExecutionPolicy Bypass -File make-ref-weekday.ps1 [-Dir <出力先>] [-Force]
#   -Dir を省くとこのスクリプトのフォルダ。出力先に同名のファイルがあるときは -Force が無ければ何もせずに止まる。
# このファイルは UTF-8（BOM 付き）で保存する（PowerShell 5.1 が日本語の書式コードを読めるように）。
param([string]$Dir = $PSScriptRoot, [switch]$Force)
$ErrorActionPreference = 'Stop'
if (-not [IO.Path]::IsPathRooted($Dir)) { $Dir = Join-Path (Get-Location).Path $Dir }
$Dir = [IO.Path]::GetFullPath($Dir)
New-Item -ItemType Directory -Force -Path $Dir | Out-Null
$path1900 = Join-Path $Dir 'ref-weekday.xlsx'
$path1904 = Join-Path $Dir 'ref-weekday-1904.xlsx'
$pathJson = Join-Path $Dir 'ref-weekday.expected.json'
# 上書きの前に止める（Excel を起動する前）
foreach ($p in @($path1900, $path1904, $pathJson)) {
  if ((Test-Path $p) -and -not $Force) { throw "既にある（上書きするなら -Force）: $p" }
}

# 書式のグループ（シート名 → 書式コード）
$groups = [ordered]@{
  basic = @('aaa', 'aaaa', 'm/d(aaa)', 'yyyy/m/d aaaa', 'yyyy"年"m"月"d"日"(aaa)', 'm/d(aaa) h:mm', 'aaa h:mm AM/PM', 'aaa A/P')
  count = @('a', 'aa', 'aaaaa', 'aaaaaa')
  upper = @('AAA', 'AAAA', 'A', 'AA', 'AAAAA')
  skip  = @('"aaa"aaa', '\aaaa', '!aaaa', '[$-411]aaa', '[$-ja-JP]aaaa', '[Red]aaa', '[Red]m/d', '[Color10]aaa', '[赤]aaa', 'aaa;@', 'm/d;aaa', 'aaa;;;', '[<=45000]aaa;aaaa')
  ddd   = @('ddd', 'dddd', 'm/d(ddd)')
  # 数値の桁との混在・空白（_）と埋め（*）の後の a・小文字の時刻記号・大文字小文字の混在
  mixed = @('0 aaa', '#,##0 aaa', 'General aaa', '0.00;aaa', 'aaa_a', 'aaa*a', '_aaaa', 'aaa am/pm', 'aaa a/p', 'aAa', 'Aaaa')
}
# シリアル値（extra = 依頼の範囲の外で足したもの：負の第 2 区画・上限の 1 つ先）
$serials1900 = @(1, 2, 7, 59, 60, 61, 62, 366, 45607, 45607.75, 45973, 2958465, 0, -1, -45607, 2958466)
$extra1900 = @(-45607, 2958466)
$serials1904 = @(0, 1, 2, 7, 1461, 44145, 44145.5, -1, -1461, 2957003, 2957004)
$extra1904 = @(-1, -1461, 2957003, 2957004)

$cells = New-Object System.Collections.ArrayList
$formats = New-Object System.Collections.ArrayList
$weekday = New-Object System.Collections.ArrayList

# 列番号 → 列の文字（1 → A）
function Col-Name([int]$c) {
  $s = ''
  while ($c -gt 0) { $m = ($c - 1) % 26; $s = [char](65 + $m) + $s; $c = [math]::Floor(($c - 1) / 26) }
  $s
}
function Is-Hash($t) { ($null -ne $t) -and ([string]$t -match '^#+$') }
# 列 $c の 2..$last 行で「#」だけの表示になっている行の一覧（カンマ区切りの文字列）
function Hash-Rows($ws, [int]$c, [int]$last) {
  $r = @()
  for ($i = 2; $i -le $last; $i++) { if (Is-Hash $ws.Cells.Item($i, $c).Text) { $r += $i } }
  $r -join ','
}
# 列幅を決める：幅 255 でも「#」の行（値が範囲外）だけが「#」として残る、いちばん狭い幅にする
function Fit-Column($ws, [int]$c, [int]$last) {
  $col = $ws.Columns.Item($c)
  $col.ColumnWidth = 255
  $h255 = Hash-Rows $ws $c $last
  [void]$col.AutoFit()
  if ((Hash-Rows $ws $c $last) -ne $h255) {
    foreach ($w in @(20, 30, 45, 60, 90, 130, 200, 255)) {
      if ($w -le [double]$col.ColumnWidth) { continue }
      $col.ColumnWidth = $w
      if ((Hash-Rows $ws $c $last) -eq $h255) { break }
    }
  }
  $h255
}

# 保存した xlsx の styles.xml から、セルごとの numFmtId と formatCode を読む（キーは「シート名!A1」）
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Read-FormatCodes([string]$file) {
  $map = @{}
  $zip = [IO.Compression.ZipFile]::OpenRead($file)
  try {
    $xmlOf = {
      param($name)
      $sr = New-Object IO.StreamReader($zip.GetEntry($name).Open(), [Text.Encoding]::UTF8)
      try { $x = New-Object Xml.XmlDocument; $x.LoadXml($sr.ReadToEnd()); $x } finally { $sr.Dispose() }
    }
    $wbx = & $xmlOf 'xl/workbook.xml'
    $rels = & $xmlOf 'xl/_rels/workbook.xml.rels'
    $sty = & $xmlOf 'xl/styles.xml'
    $codes = @{}
    foreach ($n in $sty.GetElementsByTagName('numFmt')) { $codes[[int]$n.GetAttribute('numFmtId')] = $n.GetAttribute('formatCode') }
    $xfs = @($sty.GetElementsByTagName('cellXfs').Item(0).GetElementsByTagName('xf'))
    $target = @{}
    foreach ($r in $rels.GetElementsByTagName('Relationship')) { $target[$r.GetAttribute('Id')] = $r.GetAttribute('Target') }
    foreach ($s in $wbx.GetElementsByTagName('sheet')) {
      $tg = $target[$s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')]
      if ($tg.StartsWith('/')) { $part = $tg.Substring(1) } else { $part = 'xl/' + $tg }
      $sx = & $xmlOf $part
      foreach ($c in $sx.GetElementsByTagName('c')) {
        $si = 0; if ($c.HasAttribute('s')) { $si = [int]$c.GetAttribute('s') }
        $id = [int]$xfs[$si].GetAttribute('numFmtId')
        $map[$s.GetAttribute('name') + '!' + $c.GetAttribute('r')] = @{ numFmtId = $id; formatCode = $codes[$id] }
      }
    }
  } finally { $zip.Dispose() }
  $map
}

# 1 冊を作る：書式のシート群と式のシート
function Build-Book($xl, [bool]$is1904, $serials, $extras, [string]$file) {
  $wb = $xl.Workbooks.Add()
  $wb.Date1904 = $is1904
  $first = $wb.Worksheets.Item(1)
  $prev = $null
  $book = [IO.Path]::GetFileName($file)
  $last = $serials.Count + 1
  foreach ($g in $script:groups.Keys) {
    if ($null -eq $prev) { $ws = $first } else { $ws = $wb.Worksheets.Add([Type]::Missing, $prev) }
    $ws.Name = $g; $prev = $ws
    $ws.Rows.Item(1).NumberFormat = '@'
    $ws.Cells.Item(1, 1).Value2 = 'serial'
    for ($i = 0; $i -lt $serials.Count; $i++) { $ws.Cells.Item($i + 2, 1).Value2 = [double]$serials[$i] }
    $c = 1
    foreach ($fmt in $script:groups[$g]) {
      $c++
      $ws.Cells.Item(1, $c).Value2 = $fmt
      $rng = $ws.Range($ws.Cells.Item(2, $c), $ws.Cells.Item($last, $c))
      $f = [ordered]@{ book = $book; date1904 = $is1904; sheet = $g; column = (Col-Name $c); group = $g; format = $fmt }
      try {
        $rng.NumberFormat = $fmt
        $f.setVia = 'NumberFormat'
      } catch {
        $f.error = $_.Exception.Message
        try { $rng.NumberFormatLocal = $fmt; $f.setVia = 'NumberFormatLocal' } catch { $f.errorLocal = $_.Exception.Message }
      }
      if (-not $f.Contains('setVia')) { [void]$script:formats.Add($f); continue }
      $f.formatRead = [string]$ws.Cells.Item(2, $c).NumberFormat
      $f.formatLocal = [string]$ws.Cells.Item(2, $c).NumberFormatLocal
      [void]$script:formats.Add($f)
      for ($i = 0; $i -lt $serials.Count; $i++) { $ws.Cells.Item($i + 2, $c).Value2 = [double]$serials[$i] }
    }
    [void]$ws.Columns.Item(1).AutoFit()
    for ($cc = 2; $cc -le $c; $cc++) {
      $fmt = $script:groups[$g][$cc - 2]
      $fe = $script:formats | Where-Object { $_.book -eq $book -and $_.sheet -eq $g -and $_.format -eq $fmt } | Select-Object -First 1
      if (-not $fe.Contains('setVia')) { continue }
      $h255 = Fit-Column $ws $cc $last
      $hs = @(); if ($h255) { $hs = $h255 -split ',' | ForEach-Object { [int]$_ } }
      for ($i = 0; $i -lt $serials.Count; $i++) {
        $r = $i + 2
        $o = [ordered]@{ book = $book; date1904 = $is1904; sheet = $g; cell = ((Col-Name $cc) + $r); serial = [double]$serials[$i]
          format = $fmt; formatRead = $fe.formatRead; text = [string]$ws.Cells.Item($r, $cc).Text; columnWidth = [double]$ws.Columns.Item($cc).ColumnWidth }
        if ($fe.setVia -ne 'NumberFormat') { $o.setVia = $fe.setVia }
        if ($hs -contains $r) { $o.hashAtMaxWidth = $true }
        if ($extras -contains $serials[$i]) { $o.extra = $true }
        [void]$script:cells.Add($o)
      }
    }
  }
  # 式のシート：A＝シリアル値、B＝WEEKDAY(値)、C＝TEXT(値,"aaa")、D＝TEXT(値,"aaaa")（式には値をそのまま書く）
  $ws = $wb.Worksheets.Add([Type]::Missing, $prev); $ws.Name = 'weekday'
  $ws.Rows.Item(1).NumberFormat = '@'
  $ws.Cells.Item(1, 1).Value2 = 'serial'; $ws.Cells.Item(1, 2).Value2 = 'WEEKDAY'; $ws.Cells.Item(1, 3).Value2 = 'TEXT aaa'; $ws.Cells.Item(1, 4).Value2 = 'TEXT aaaa'
  for ($i = 0; $i -lt $serials.Count; $i++) {
    $r = $i + 2; $v = [double]$serials[$i]; $lit = $v.ToString([Globalization.CultureInfo]::InvariantCulture)
    # ここで .Value2 = <double> にすると PowerShell 5.1 の COM の束縛が「Double を String に変換できない」で落ちた（原因は未特定）。数の文字列を Formula に入れる
    $ws.Cells.Item($r, 1).Formula = $lit
    $ws.Cells.Item($r, 2).Formula = "=WEEKDAY($lit)"
    $ws.Cells.Item($r, 3).Formula = "=TEXT($lit,`"aaa`")"
    $ws.Cells.Item($r, 4).Formula = "=TEXT($lit,`"aaaa`")"
  }
  $xl.Calculate()
  [void]$ws.Columns.Item('A:D').AutoFit()
  for ($i = 0; $i -lt $serials.Count; $i++) {
    $r = $i + 2
    foreach ($k in @(@(2, 'WEEKDAY'), @(3, 'TEXT_aaa'), @(4, 'TEXT_aaaa'))) {
      $cl = $ws.Cells.Item($r, $k[0])
      $v2 = $cl.Value2
      $o = [ordered]@{ book = $book; date1904 = $is1904; sheet = 'weekday'; cell = ((Col-Name $k[0]) + $r); serial = [double]$serials[$i]
        kind = $k[1]; formula = [string]$cl.Formula; value2 = $v2; text = [string]$cl.Text }
      # エラー値（#NUM! など）は Value2 が Int32 のエラー番号になる
      if ($v2 -is [int]) { $o.value2 = $null; $o.errorCode = $v2 }
      if ($extras -contains $serials[$i]) { $o.extra = $true }
      [void]$script:weekday.Add($o)
    }
  }
  $first.Activate()
  $wb.SaveAs($file, 51)
  $wb.Close($false)
  # styles.xml の formatCode（組み込みの書式は numFmtId だけで formatCode は null）
  $fc = Read-FormatCodes $file
  foreach ($o in $script:cells) {
    if ($o.book -ne $book) { continue }
    $k = $fc[$o.sheet + '!' + $o.cell]
    $o.Insert(7, 'formatCode', $k.formatCode); $o.Insert(8, 'numFmtId', $k.numFmtId)
  }
  foreach ($f in $script:formats) {
    if ($f.book -ne $book -or -not $f.Contains('setVia')) { continue }
    $k = $fc[$f.sheet + '!' + $f.column + '2']
    $f.formatCode = $k.formatCode; $f.numFmtId = $k.numFmtId
  }
}

# ConvertTo-Json（5.1）がバックスラッシュ u の 4 桁にする ' < > & を素の文字に戻す（どれも JSON の文字列にそのまま書ける）
function Unescape-Json([string]$j) {
  $bs = [string][char]92
  foreach ($k in @(@('0027', "'"), @('003c', '<'), @('003e', '>'), @('0026', '&'))) { $j = $j.Replace($bs + 'u' + $k[0], $k[1]) }
  $j
}
# 配列を 1 要素 1 行の JSON にする
function Json-Lines($arr) {
  $lines = @()
  foreach ($x in $arr) {
    $j = $x | ConvertTo-Json -Compress -Depth 5
    $j = Unescape-Json $j
    $lines += '    ' + $j
  }
  "[`n" + ($lines -join ",`n") + "`n  ]"
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $excel = [ordered]@{
    version = [string]$xl.Version; build = [string]$xl.Build
    # International(1)＝xlCountryCode（81＝日本）、International(2)＝xlCountrySetting。LanguageID(2)＝msoLanguageIDUI（1041＝日本語）
    countryCode = $xl.International(1); countrySetting = $xl.International(2); uiLanguageId = $xl.LanguageSettings.LanguageID(2)
  }
  Build-Book $xl $false $serials1900 $extra1900 $path1900
  Build-Book $xl $true $serials1904 $extra1904 $path1904
  $about = 'Issue #25 曜日の書式（aaa・aaaa）の正解。make-ref-weekday.ps1 が Excel の COM で作った。text は Range.Text（Excel の表示文字）、formatCode は xlsx の styles.xml の文字列'
  $books = @([ordered]@{ file = 'ref-weekday.xlsx'; date1904 = $false }, [ordered]@{ file = 'ref-weekday-1904.xlsx'; date1904 = $true })
  $cj = { param($v) Unescape-Json (ConvertTo-Json -InputObject $v -Compress -Depth 5) }
  $json = "{`n  `"about`": " + (& $cj $about) + ",`n  `"excel`": " + (& $cj $excel) + ",`n  `"books`": " + (& $cj $books) +
    ",`n  `"formats`": " + (Json-Lines $formats) + ",`n  `"cells`": " + (Json-Lines $cells) + ",`n  `"weekday`": " + (Json-Lines $weekday) + "`n}`n"
  [IO.File]::WriteAllText($pathJson, $json, (New-Object Text.UTF8Encoding($false)))
  "ok: formats=$($formats.Count) cells=$($cells.Count) weekday=$($weekday.Count)"
} finally {
  try { foreach ($w in @($xl.Workbooks)) { $w.Close($false) } } catch {}
  try { $xl.Quit() } catch {}
  [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  $xl = $null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
