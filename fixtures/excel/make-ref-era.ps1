# Issue #32：和暦（g・e）と組み込みの日付の書式 27〜36・50〜58 の正解データを Excel 16（ja-JP、COM）で作る。
# 作るもの（$Dir に書く。ほかのファイルは触らない）：
#   ref-era.xlsx            1900 年基準。このスクリプトが XML から組み立てる合成のファイル（Excel では保存しない）
#   ref-era-1904.xlsx       1904 年基準（workbookPr date1904="1"）。シートの構成は同じ
#   ref-era.expected.json   セルごとの { book, sheet, cell, serial, numFmtId, formatCode, formatRead, formatLocal, text }
# シート：custom＝独自の書式（numFmtId 164 から）、builtin＝組み込みの番号（styles.xml に numFmt を書かない）。
# 形：1 行目＝書式コードか「id N」（文字列）、A 列＝シリアル値（標準）、B 列から右＝各書式のセル。
# Excel は開くだけ（保存しない）ので、XSheet が読むファイルと Excel が表示したファイルは同じバイト列。
# text は Range.Text（幅 60 で読み、「#」だけなら幅 255 で読み直す。幅 255 でも「#」なら hashAtMaxWidth = true）。
# formatRead／formatLocal は Excel が読み取った書式（NumberFormat・NumberFormatLocal）。組み込みの番号の中身を知るために残す。
# 実行（CI では実行しない。Excel 16 が要る）：
#   powershell -ExecutionPolicy Bypass -File make-ref-era.ps1 [-Dir <出力先>] [-Force]
#   -Dir を省くとこのスクリプトのフォルダ。出力先に同名のファイルがあるときは -Force が無ければ何もせずに止まる。
# このファイルは UTF-8（BOM 付き）で保存する（PowerShell 5.1 が日本語の書式コードを読めるように）。
param([string]$Dir = $PSScriptRoot, [switch]$Force)
$ErrorActionPreference = 'Stop'
if (-not [IO.Path]::IsPathRooted($Dir)) { $Dir = Join-Path (Get-Location).Path $Dir }
$Dir = [IO.Path]::GetFullPath($Dir)
New-Item -ItemType Directory -Force -Path $Dir | Out-Null
$path1900 = Join-Path $Dir 'ref-era.xlsx'
$path1904 = Join-Path $Dir 'ref-era-1904.xlsx'
$pathJson = Join-Path $Dir 'ref-era.expected.json'
foreach ($p in @($path1900, $path1904, $pathJson)) {
  if ((Test-Path $p) -and -not $Force) { throw "既にある（上書きするなら -Force）: $p" }
}

# 独自の書式（和暦の字句 g・e・r と言語の指定の組み合わせ。依頼のファイルの書式 [$-ja-JP]ggge"年"m"月"d"日" を含む）
$custom = @(
  '[$-411]ggge"年"m"月"d"日"', '[$-ja-JP]ggge"年"m"月"d"日"', '[$-ja-JP-x-gannen]ggge"年"m"月"d"日"', '[$-411]ge.m.d', '[$-411]gge.m.d',
  '[$-411]g', '[$-411]gg', '[$-411]ggg', '[$-411]gggg', '[$-411]e', '[$-411]ee', '[$-411]eee',
  'ggge"年"m"月"d"日"', 'e', 'ee', 'ge', '[$-409]ggge', '[$-409]e', '[$-412]ggge', '[$-411]r', '[$-411]rr',
  '[$-411]ggge"年"m"月"d"日"(aaa)', '[$-411]ggge"年"m"月"d"日" h:mm', '"ggg"e', '[$-411]"g"e', '[$-411]yyyy"年"', '[$-411]G', '[$-411]E',
  '[$-411]ggge"年";@', '[Red][$-411]ggge',
  '[$-ja-JP-x-gannen]e', '[$-ja-JP-x-gannen]ee', '[$-ja-JP-x-gannen]ge.m.d', '[$-411]rrr', '[$-ja]ggge', '[$-0411]ggge', '[$-ja-JP]e',
  '[$-ja-JP-x-gannen]e"年"', '[$-ja-JP-x-gannen]ee"年"', '[$-ja-JP-x-gannen]e" 年"', '[$-ja-JP-x-gannen]e\年', '[$-ja-JP-x-gannen]e"年度"', '[$-ja-JP-x-gannen]e"/"m',
  '[$-ja-JP-x-gannen]ggge', '[$-ja-JP-x-gannen]ggge"年"', '[$-ja-JP-x-gannen]ge"年"', '[$-ja-JP-x-gannen]gge"年"', '[$-ja-JP-x-gannen]gggee"年"',
  '[$-ja-JP-x-gannen]ggg" "e"年"', '[$-ja-JP-x-gannen]e"年"ggg', '[$-ja-JP-x-gannen]ggge"年度"'
)
# 組み込みの番号（14・22 は西暦の確かめ）
$builtin = @(14, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58)

# 日付 → 1900 年基準のシリアル値（1900-03-01 以降。それより前は直接書く）
function Serial([int]$y, [int]$m, [int]$d) { [double]((New-Object DateTime($y, $m, $d)).ToOADate()) }
$dates = @(
  @(1900, 7, 29), @(1912, 7, 29), @(1912, 7, 30), @(1926, 12, 24), @(1926, 12, 25), @(1927, 1, 1), @(1989, 1, 7), @(1989, 1, 8),
  @(1990, 1, 1), @(2019, 4, 30), @(2019, 5, 1), @(2019, 12, 31), @(2020, 1, 1), @(2026, 11, 11), @(9999, 12, 31)
)
$serials1900 = @(0, 1, 59, 60, 61) + ($dates | ForEach-Object { Serial $_[0] $_[1] $_[2] }) + @(((Serial 2026 11 11) + 0.5625), -1, 2958466)
# 1904 年基準：同じ日付から 1462 を引く（1904-01-01 が 0）
$serials1904 = @(0, 1, -1, -1462) + ($dates | Select-Object -Skip 1 | ForEach-Object { (Serial $_[0] $_[1] $_[2]) - 1462 })

function Col-Name([int]$c) {
  $s = ''
  while ($c -gt 0) { $m = ($c - 1) % 26; $s = [char](65 + $m) + $s; $c = [math]::Floor(($c - 1) / 26) }
  $s
}
function Esc([string]$s) { [Security.SecurityElement]::Escape($s) }
function Num([double]$v) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) }

# シート 1 枚の XML。cols＝[{ head, xf }]
function Sheet-Xml($cols, $serials) {
  $sb = New-Object Text.StringBuilder
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
  [void]$sb.Append('<cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="' + ($cols.Count + 1) + '" width="40" customWidth="1"/></cols><sheetData>')
  [void]$sb.Append('<row r="1"><c r="A1" t="inlineStr"><is><t>serial</t></is></c>')
  for ($i = 0; $i -lt $cols.Count; $i++) {
    $a = (Col-Name ($i + 2)) + '1'
    [void]$sb.Append('<c r="' + $a + '" t="inlineStr"><is><t xml:space="preserve">' + (Esc $cols[$i].head) + '</t></is></c>')
  }
  [void]$sb.Append('</row>')
  for ($r = 0; $r -lt $serials.Count; $r++) {
    $row = $r + 2; $v = Num $serials[$r]
    [void]$sb.Append('<row r="' + $row + '"><c r="A' + $row + '"><v>' + $v + '</v></c>')
    for ($i = 0; $i -lt $cols.Count; $i++) {
      [void]$sb.Append('<c r="' + (Col-Name ($i + 2)) + $row + '" s="' + $cols[$i].xf + '"><v>' + $v + '</v></c>')
    }
    [void]$sb.Append('</row>')
  }
  [void]$sb.Append('</sheetData></worksheet>')
  $sb.ToString()
}

# 合成の xlsx を組み立てる（custom・builtin の 2 枚）
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Build-Xlsx([string]$file, [bool]$is1904, $serials) {
  $xfIds = New-Object System.Collections.ArrayList
  [void]$xfIds.Add(0)
  $customCols = @(); $numFmts = @()
  for ($i = 0; $i -lt $script:custom.Count; $i++) {
    $id = 164 + $i
    $numFmts += '<numFmt numFmtId="' + $id + '" formatCode="' + (Esc $script:custom[$i]) + '"/>'
    $customCols += @{ head = $script:custom[$i]; xf = $xfIds.Count }
    [void]$xfIds.Add($id)
  }
  $builtinCols = @()
  foreach ($id in $script:builtin) {
    $builtinCols += @{ head = "id $id"; xf = $xfIds.Count }
    [void]$xfIds.Add($id)
  }
  $xfs = ($xfIds | ForEach-Object {
      $apply = ''; if ($_ -ne 0) { $apply = ' applyNumberFormat="1"' }
      '<xf numFmtId="' + $_ + '" fontId="0" fillId="0" borderId="0" xfId="0"' + $apply + '/>'
    }) -join ''
  $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="' + $numFmts.Count + '">' + ($numFmts -join '') + '</numFmts>' +
    '<fonts count="1"><font><sz val="11"/><name val="游ゴシック"/><family val="3"/><charset val="128"/><scheme val="minor"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="' + $xfIds.Count + '">' + $xfs + '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
  $pr = '<workbookPr/>'; if ($is1904) { $pr = '<workbookPr date1904="1"/>' }
  $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  $rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  $ct = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
  $head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  $parts = [ordered]@{
    '[Content_Types].xml'        = $head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="' + $ct + '.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="' + $ct + '.worksheet+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="' + $ct + '.worksheet+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="' + $ct + '.styles+xml"/></Types>'
    '_rels/.rels'                = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="' + $rel + '/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    'xl/workbook.xml'            = $head + '<workbook xmlns="' + $ns + '" xmlns:r="' + $rel + '">' + $pr +
    '<sheets><sheet name="custom" sheetId="1" r:id="rId1"/><sheet name="builtin" sheetId="2" r:id="rId2"/></sheets></workbook>'
    'xl/_rels/workbook.xml.rels' = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="' + $rel + '/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="' + $rel + '/worksheet" Target="worksheets/sheet2.xml"/>' +
    '<Relationship Id="rId3" Type="' + $rel + '/styles" Target="styles.xml"/></Relationships>'
    'xl/worksheets/sheet1.xml'   = (Sheet-Xml $customCols $serials)
    'xl/worksheets/sheet2.xml'   = (Sheet-Xml $builtinCols $serials)
    'xl/styles.xml'              = $styles
  }
  if (Test-Path $file) { Remove-Item $file }
  $zip = [IO.Compression.ZipFile]::Open($file, 'Create')
  try {
    $utf8 = New-Object Text.UTF8Encoding($false)
    foreach ($k in $parts.Keys) {
      $e = $zip.CreateEntry($k, [IO.Compression.CompressionLevel]::Optimal)
      $w = New-Object IO.StreamWriter($e.Open(), $utf8)
      try { $w.Write($parts[$k]) } finally { $w.Dispose() }
    }
  } finally { $zip.Dispose() }
  @{ custom = $customCols; builtin = $builtinCols; xfIds = $xfIds }
}

function Is-Hash($t) { ($null -ne $t) -and ([string]$t -match '^#+$') }

$cells = New-Object System.Collections.ArrayList
# Excel で開いて各セルの表示を読む（読み取り専用で開き、保存しない）
function Read-Book($xl, [string]$file, [bool]$is1904, $serials, $layout) {
  $wb = $xl.Workbooks.Open($file, 0, $true)
  try {
    $book = [IO.Path]::GetFileName($file)
    foreach ($sheetName in @('custom', 'builtin')) {
      $ws = $wb.Worksheets.Item($sheetName)
      $cols = $layout[$sheetName]
      for ($i = 0; $i -lt $cols.Count; $i++) {
        $c = $i + 2
        $col = $ws.Columns.Item($c)
        $col.ColumnWidth = 255
        $hash255 = @{}
        for ($r = 0; $r -lt $serials.Count; $r++) { if (Is-Hash $ws.Cells.Item($r + 2, $c).Text) { $hash255[$r] = $true } }
        $col.ColumnWidth = 60
        $id = $layout.xfIds[$cols[$i].xf]
        for ($r = 0; $r -lt $serials.Count; $r++) {
          $cell = $ws.Cells.Item($r + 2, $c)
          $text = [string]$cell.Text
          if ((Is-Hash $text) -and -not $hash255.ContainsKey($r)) { $col.ColumnWidth = 255; $text = [string]$cell.Text; $col.ColumnWidth = 60 }
          $fc = $null; if ($sheetName -eq 'custom') { $fc = $script:custom[$i] }
          $o = [ordered]@{ book = $book; date1904 = $is1904; sheet = $sheetName; cell = ((Col-Name $c) + ($r + 2)); serial = [double]$serials[$r]
            numFmtId = $id; formatCode = $fc; formatRead = [string]$cell.NumberFormat; formatLocal = [string]$cell.NumberFormatLocal; text = $text
          }
          if ($hash255.ContainsKey($r)) { $o.hashAtMaxWidth = $true }
          [void]$script:cells.Add($o)
        }
      }
    }
  } finally { $wb.Close($false) }
}

# ConvertTo-Json（5.1）がバックスラッシュ u の 4 桁にする ' < > & を素の文字に戻す
function Unescape-Json([string]$j) {
  $bs = [string][char]92
  foreach ($k in @(@('0027', "'"), @('003c', '<'), @('003e', '>'), @('0026', '&'))) { $j = $j.Replace($bs + 'u' + $k[0], $k[1]) }
  $j
}
$nl = [string][char]10
# 配列を 1 要素 1 行の JSON にする
function Json-Lines($arr) {
  $lines = @()
  foreach ($x in $arr) { $lines += '    ' + (Unescape-Json ($x | ConvertTo-Json -Compress -Depth 5)) }
  '[' + $nl + ($lines -join (',' + $nl)) + $nl + '  ]'
}

$layout1900 = Build-Xlsx $path1900 $false $serials1900
$layout1904 = Build-Xlsx $path1904 $true $serials1904
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  # International(1)＝xlCountryCode（81＝日本）、International(2)＝xlCountrySetting。LanguageID(2)＝msoLanguageIDUI（1041＝日本語）
  $excel = [ordered]@{ version = [string]$xl.Version; build = [string]$xl.Build; countryCode = $xl.International(1); countrySetting = $xl.International(2); uiLanguageId = $xl.LanguageSettings.LanguageID(2) }
  Read-Book $xl $path1900 $false $serials1900 $layout1900
  Read-Book $xl $path1904 $true $serials1904 $layout1904
  $about = 'Issue #32 和暦（g・e）と組み込みの日付の書式の正解。make-ref-era.ps1 が組み立てた xlsx を Excel の COM で開き（保存しない）、Range.Text を読んだ。formatCode は styles.xml に書いた文字列（組み込みの番号は null）'
  $cj = { param($v) Unescape-Json (ConvertTo-Json -InputObject $v -Compress -Depth 5) }
  $json = '{' + $nl + '  "about": ' + (& $cj $about) + ',' + $nl + '  "excel": ' + (& $cj $excel) + ',' + $nl + '  "cells": ' + (Json-Lines $cells) + $nl + '}' + $nl
  [IO.File]::WriteAllText($pathJson, $json, (New-Object Text.UTF8Encoding($false)))
  "ok: cells=$($cells.Count)"
} finally {
  try { foreach ($w in @($xl.Workbooks)) { $w.Close($false) } } catch {}
  try { $xl.Quit() } catch {}
  [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  $xl = $null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
