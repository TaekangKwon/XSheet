# ISSUE-95: ref-colors.xlsx (Excel 16 COM) and the pixels Excel painted for every shape (U7: colour conversion, U8: face shading).
# Shapes: Excel defaults, theme colours (12 slots x brightness steps), srgb fills, transparency, gradients (2-3 stops), patterns,
# a picture fill, glyph text colours (Wingdings square), 6pt outlines and lines, and the shaded faces of can / cube / foldedCorner / bevel.
# For each shape: CopyPicture -> paste into a chart (background magenta) -> Chart.Export PNG -> read pixels (a C# helper, System.Drawing).
# Writes into $Dir (nothing else is touched): ref-colors.xlsx, com.json (COM values + pixels), png\*.png (not committed), pic-fill.png (picture-fill source).
# Run (do NOT run on CI; needs Excel 16, screen zoom 125% like make-ref-shapes.ps1):
#   powershell -ExecutionPolicy Bypass -File make-ref-colors.ps1 -Dir <out dir>   (PowerShell 5.1, file saved as UTF-8 with BOM)
# Notes from the runs: the CopyPicture bitmap has a WHITE background rectangle (transparent fills blend over white, not over the magenta bg),
# and Chart.Export writes indexed PNGs. So "run" (first stroke) is unusable; run make-ref-colors.extract.mjs <out dir> to read the PNGs and add "stroke".
# After the run: node make-ref-colors.extract.mjs <out dir>; copy ref-colors.xlsx and expected.json (as ref-colors.expected.json) into __fixtures__/excel.
param([string]$Dir)
$ErrorActionPreference = 'Stop'
# relative -Dir is resolved against the PowerShell location (not the process cwd), then made absolute
if (-not [IO.Path]::IsPathRooted($Dir)) { $Dir = Join-Path (Get-Location).Path $Dir }
$Dir = [IO.Path]::GetFullPath($Dir)
New-Item -ItemType Directory -Force -Path $Dir | Out-Null
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
public static class RefPixel {
  // ARGB pixels of a PNG (rows packed, no stride padding)
  static int[] Load(string path, out int w, out int h) {
    using (Bitmap bmp = new Bitmap(path)) {
      w = bmp.Width; h = bmp.Height;
      BitmapData d = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
      int[] buf = new int[w * h];
      for (int y = 0; y < h; y++) System.Runtime.InteropServices.Marshal.Copy(IntPtr.Add(d.Scan0, y * d.Stride), buf, y * w, w);
      bmp.UnlockBits(d);
      return buf;
    }
  }
  static bool IsBg(int p, int bg) {
    int dr = Math.Abs(((p >> 16) & 255) - ((bg >> 16) & 255));
    int dg = Math.Abs(((p >> 8) & 255) - ((bg >> 8) & 255));
    int db = Math.Abs((p & 255) - (bg & 255));
    return dr <= 6 && dg <= 6 && db <= 6;
  }
  // {width, height, x0, y0, x1, y1}: bounding box of the pixels that are not the background (x1,y1 inclusive); x0 = -1 if none
  public static int[] Info(string path, int bgR, int bgG, int bgB) {
    int w, h; int[] b = Load(path, out w, out h);
    int bg = (bgR << 16) | (bgG << 8) | bgB;
    int x0 = -1, y0 = -1, x1 = -1, y1 = -1;
    for (int y = 0; y < h; y++) for (int x = 0; x < w; x++) {
      if (IsBg(b[y * w + x], bg)) continue;
      if (x0 < 0 || x < x0) x0 = x; if (x1 < 0 || x > x1) x1 = x;
      if (y0 < 0 || y < y0) y0 = y; if (y1 < 0 || y > y1) y1 = y;
    }
    return new int[] { w, h, x0, y0, x1, y1 };
  }
  // mean {r, g, b} over the rectangle [x0..x1] x [y0..y1] (clipped to the image)
  public static int[] Mean(string path, int x0, int y0, int x1, int y1) {
    int w, h; int[] b = Load(path, out w, out h);
    long sr = 0, sg = 0, sb = 0, n = 0;
    for (int y = Math.Max(y0, 0); y <= Math.Min(y1, h - 1); y++) for (int x = Math.Max(x0, 0); x <= Math.Min(x1, w - 1); x++) {
      int p = b[y * w + x]; sr += (p >> 16) & 255; sg += (p >> 8) & 255; sb += p & 255; n++;
    }
    if (n == 0) return new int[] { -1, -1, -1 };
    return new int[] { (int)Math.Round((double)sr / n), (int)Math.Round((double)sg / n), (int)Math.Round((double)sb / n) };
  }
  // First run of non-background pixels along row y (horizontal=true) or column (horizontal=false) starting at `from`.
  // {start, end, r, g, b}: r,g,b = mean of the middle third of the run (stroke centre). start = -1 if none.
  public static int[] Run(string path, bool horizontal, int fixedPos, int from, int bgR, int bgG, int bgB) {
    int w, h; int[] b = Load(path, out w, out h);
    int bg = (bgR << 16) | (bgG << 8) | bgB;
    int len = horizontal ? w : h;
    int s = -1, e = -1;
    for (int i = Math.Max(from, 0); i < len; i++) {
      int p = horizontal ? b[fixedPos * w + i] : b[i * w + fixedPos];
      bool bgPix = IsBg(p, bg);
      if (s < 0) { if (!bgPix) s = i; } else if (bgPix) { e = i - 1; break; }
    }
    if (s < 0) return new int[] { -1, -1, -1, -1, -1 };
    if (e < 0) e = len - 1;
    int third = Math.Max((e - s + 1) / 3, 0);
    int a = s + third, z = e - third;
    long sr = 0, sg = 0, sb = 0, n = 0;
    for (int i = a; i <= z; i++) {
      int p = horizontal ? b[fixedPos * w + i] : b[i * w + fixedPos];
      sr += (p >> 16) & 255; sg += (p >> 8) & 255; sb += p & 255; n++;
    }
    return new int[] { s, e, (int)Math.Round((double)sr / n), (int)Math.Round((double)sg / n), (int)Math.Round((double)sb / n) };
  }
}
"@

$xl = New-Object -ComObject Excel.Application
# Visible: CopyPicture + Chart.Paste silently pasted nothing while Excel was hidden (first run: every PNG was pure background)
$xl.Visible = $true; $xl.DisplayAlerts = $false
$notes = New-Object System.Collections.ArrayList
$records = New-Object System.Collections.ArrayList
$ws = $null
$wsw = $null
$pngDir = Join-Path $Dir 'png'
New-Item -ItemType Directory -Force -Path $pngDir | Out-Null
# chart background (magenta): the pasted picture is found as the pixels that differ from it. Excel RGB int = R + G*256 + B*65536
$BG = @(255, 0, 255)
$BGINT = 255 + 0 * 256 + 255 * 65536
$script:slot = 0
$script:pasteNoted = $false

function Safe($sb) { try { & $sb } catch { $null } }
function Try-Do($label, [scriptblock]$sb) { try { & $sb } catch { [void]$script:notes.Add("$label failed: $($_.Exception.Message)") } }
# Excel colour integer (BGR) -> 'RRGGBB'
function Hex-Of($v) {
  if ($null -eq $v) { return $null }
  $v = [int]$v
  '{0:X2}{1:X2}{2:X2}' -f ($v -band 255), (($v -shr 8) -band 255), (($v -shr 16) -band 255)
}
# 'RRGGBB' -> Excel colour integer
function Rgb-Int([string]$hex) {
  [Convert]::ToInt32($hex.Substring(0, 2), 16) + [Convert]::ToInt32($hex.Substring(2, 2), 16) * 256 + [Convert]::ToInt32($hex.Substring(4, 2), 16) * 65536
}
# next free spot on the sheet: 10 columns, pitch 90 x 90 pt
function Next-Pos() {
  $k = $script:slot; $script:slot++
  , @((20 + 90 * ($k % 10)), (20 + 90 * [math]::Floor($k / 10)))
}
function New-Rect($name, $w, $h, $type) {
  $p = Next-Pos
  $s = $script:ws.Shapes.AddShape($type, $p[0], $p[1], $w, $h); $s.Name = $name
  $s
}
function New-HLine($name, $len) {
  $p = Next-Pos
  $s = $script:ws.Shapes.AddLine($p[0], $p[1] + 20, $p[0] + $len, $p[1] + 20); $s.Name = $name
  $s
}
# COM view of the fill and the line (what Excel says the colour is; compared to the pixels in the test)
function Com-Fill($s) {
  $f = $s.Fill
  $o = [ordered]@{ Visible = Safe { $f.Visible }; Type = Safe { $f.Type }; Transparency = Safe { $f.Transparency }
    ForeRGB = Hex-Of (Safe { $f.ForeColor.RGB }); ForeTheme = Safe { $f.ForeColor.ObjectThemeColor }; ForeBrightness = Safe { $f.ForeColor.Brightness }
    ForeTintAndShade = Safe { $f.ForeColor.TintAndShade }; BackRGB = Hex-Of (Safe { $f.BackColor.RGB }); Pattern = Safe { $f.Pattern } }
  $n = Safe { $f.GradientStops.Count }
  if ($n) {
    $st = @()
    for ($i = 1; $i -le $n; $i++) { $g = $f.GradientStops.Item($i); $st += [ordered]@{ Position = $g.Position; RGB = Hex-Of $g.Color.RGB; Transparency = $g.Transparency } }
    $o.GradientStops = $st
  }
  $o
}
function Com-Line($s) {
  $l = $s.Line
  [ordered]@{ Visible = Safe { $l.Visible }; Weight = Safe { $l.Weight }; Transparency = Safe { $l.Transparency }
    ForeRGB = Hex-Of (Safe { $l.ForeColor.RGB }); ForeTheme = Safe { $l.ForeColor.ObjectThemeColor }; ForeBrightness = Safe { $l.ForeColor.Brightness } }
}
# Export one shape to PNG: CopyPicture -> paste into a magenta chart (on the work sheet) -> Chart.Export
function Export-Png($s, [string]$name) {
  $w = [math]::Ceiling([double]$s.Width) + 60
  $h = [math]::Ceiling([double]$s.Height) + 60
  $co = $script:wsw.ChartObjects().Add(0, 0, $w, $h)
  try {
    $ch = $co.Chart
    $ch.ChartArea.Format.Fill.Visible = -1
    $ch.ChartArea.Format.Fill.Solid()
    $ch.ChartArea.Format.Fill.ForeColor.RGB = $script:BGINT
    $ch.ChartArea.Format.Line.Visible = 0
    # CopyPicture(xlScreen=1, xlBitmap=2) -> Paste into the chart. Chart.Paste can paste nothing without an error, so check Shapes.Count
    # and retry with the sheets activated / the chart area selected (the strategy that worked is noted once).
    $ok = $false
    foreach ($strategy in @('plain', 'activate', 'select')) {
      $script:ws.Activate()
      $s.CopyPicture(1, 2)
      Start-Sleep -Milliseconds 150
      if ($strategy -ne 'plain') { $script:wsw.Activate(); $co.Activate() }
      if ($strategy -eq 'select') { $ch.ChartArea.Select() | Out-Null }
      $ch.Paste()
      if ($ch.Shapes.Count -gt 0) {
        $ok = $true
        if (-not $script:pasteNoted) { $script:pasteNoted = $true; [void]$script:notes.Add("paste strategy: $strategy") }
        break
      }
    }
    if (-not $ok) { [void]$script:notes.Add("paste failed (all strategies): $name") }
    $path = Join-Path $script:pngDir ($name + '.png')
    [void]$ch.Export($path, 'PNG')
    $path
  } finally { $co.Delete() | Out-Null }
}
# Record one shape. $pts: array of @(fx, fy, radiusPx) = sample points as fractions of the picture box (patch mean).
# $runs: 'h' = first stroke along the middle row from the left edge, 'v' = first stroke down the middle column from the top edge.
function Probe($s, [string]$group, $pts, [string]$runs, $inp) {
  $name = $s.Name
  $o = [ordered]@{ name = $name; group = $group; input = $inp; bg = $script:BG }
  $o.fill = Com-Fill $s
  $o.line = Com-Line $s
  Try-Do "probe $name" {
    $png = Export-Png $s $name
    $inf = [RefPixel]::Info($png, $script:BG[0], $script:BG[1], $script:BG[2])
    $o.box = $inf
    if ($inf[2] -lt 0) { [void]$script:notes.Add("no pixels (picture not pasted): $name") }
    if ($inf[2] -ge 0) {
      $bw = $inf[4] - $inf[2] + 1; $bh = $inf[5] - $inf[3] + 1
      $samples = @()
      foreach ($p in $pts) {
        $cx = [int][math]::Round($inf[2] + $p[0] * ($bw - 1)); $cy = [int][math]::Round($inf[3] + $p[1] * ($bh - 1)); $r = [int]$p[2]
        $m = [RefPixel]::Mean($png, $cx - $r, $cy - $r, $cx + $r, $cy + $r)
        $samples += [ordered]@{ fx = $p[0]; fy = $p[1]; radiusPx = $r; rgb = ('{0:X2}{1:X2}{2:X2}' -f $m[0], $m[1], $m[2]) }
      }
      $o.samples = $samples
      if ($runs -eq 'h') {
        $my = [int][math]::Round(($inf[3] + $inf[5]) / 2)
        $rn = [RefPixel]::Run($png, $true, $my, $inf[2], $script:BG[0], $script:BG[1], $script:BG[2])
        $o.run = [ordered]@{ axis = 'h'; start = $rn[0]; end = $rn[1]; rgb = ('{0:X2}{1:X2}{2:X2}' -f $rn[2], $rn[3], $rn[4]) }
      } elseif ($runs -eq 'v') {
        $mx = [int][math]::Round(($inf[2] + $inf[4]) / 2)
        $rn = [RefPixel]::Run($png, $false, $mx, $inf[3], $script:BG[0], $script:BG[1], $script:BG[2])
        $o.run = [ordered]@{ axis = 'v'; start = $rn[0]; end = $rn[1]; rgb = ('{0:X2}{1:X2}{2:X2}' -f $rn[2], $rn[3], $rn[4]) }
      }
    }
  }
  [void]$script:records.Add($o)
}
# Wingdings 'n' = a solid square glyph: the centre pixel of the text is the text colour
function Set-Glyph($s, [string]$chars) {
  $tf = $s.TextFrame2
  $tf.WordWrap = 0; $tf.VerticalAnchor = 3
  $tf.TextRange.Text = $chars
  $tf.TextRange.Font.Name = 'Wingdings'; $tf.TextRange.Font.Size = 72
  $tf.TextRange.ParagraphFormat.Alignment = 2
}
function Solid-Fill($s, [string]$hex) { $s.Fill.Visible = -1; $s.Fill.Solid(); $s.Fill.ForeColor.RGB = (Rgb-Int $hex); $s.Line.Visible = 0 }

try {
  $wb = $xl.Workbooks.Add()
  $wb.Styles.Item('Normal').Font.Name = 'Calibri'
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'colors'
  $wsw = $wb.Worksheets.Add([Type]::Missing, $ws)
  $wsw.Name = 'work'
  $ws.Activate()
  $ws.Cells.Item(1, 1).Value2 = 'ref-colors'
  $CENTER = @(, @(0.5, 0.5, 4))

  # ========== 1 Excel defaults (style refs: fillRef/lnRef/fontRef) ==========
  $s = New-Rect 'defShape' 60 40 1;  Probe $s 'default' $CENTER 'h' @{ api = 'AddShape(1)' }
  $s = New-Rect 'defShapeThick' 60 40 1; $s.Line.Weight = 6; Probe $s 'default' $CENTER 'h' @{ api = 'AddShape(1)+Line.Weight=6' }
  $p = Next-Pos; $s = $ws.Shapes.AddTextbox(1, $p[0], $p[1], 60, 40); $s.Name = 'defTextbox'
  Probe $s 'default' $CENTER 'h' @{ api = 'AddTextbox(1)' }
  $s = New-HLine 'defLine' 60;  Probe $s 'default' @() 'v' @{ api = 'AddLine' }
  $s = New-HLine 'defLineThick' 60; $s.Line.Weight = 6; Probe $s 'default' @() 'v' @{ api = 'AddLine+Weight=6' }
  $p = Next-Pos; $s = $ws.Shapes.AddConnector(1, $p[0], $p[1] + 20, $p[0] + 60, $p[1] + 20); $s.Name = 'defConnector'; $s.Line.Weight = 6
  Probe $s 'default' @() 'v' @{ api = 'AddConnector(1)+Weight=6' }
  # default text colours: shape text = fontRef lt1 (white), textbox text = dk1 (black)
  $s = New-Rect 'textDefShape' 100 100 1; Set-Glyph $s 'n'; Probe $s 'text' @(@(0.5, 0.5, 4), @(0.08, 0.08, 3)) '' @{ api = 'AddShape(1)+glyph'; expect = 'glyph lt1 on accent1' }
  $p = Next-Pos; $s = $ws.Shapes.AddTextbox(1, $p[0], $p[1], 100, 100); $s.Name = 'textDefTextbox'; Set-Glyph $s 'n'; $s.Height = 100
  Probe $s 'text' $CENTER '' @{ api = 'AddTextbox(1)+glyph'; expect = 'glyph dk1' }

  # ========== 2 theme colours: 12 slots x brightness steps (Excel's colour picker: lighter 80/60/40, darker 25/50) ==========
  $slots = @('dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink')
  $steps = @(0, 0.8, 0.6, 0.4, -0.25, -0.5)
  for ($i = 0; $i -lt 12; $i++) {
    foreach ($b in $steps) {
      $nm = 'th_' + $slots[$i] + '_' + ('{0:+0.00;-0.00;+0.00}' -f $b)
      $s = New-Rect $nm 36 24 1
      Try-Do $nm {
        $s.Fill.Visible = -1; $s.Fill.Solid(); $s.Fill.ForeColor.ObjectThemeColor = ($i + 1)
        if ($b -ne 0) { $s.Fill.ForeColor.Brightness = $b }
        $s.Line.Visible = 0
      }
      Probe $s 'theme' $CENTER '' @{ slot = $slots[$i]; themeIndex = ($i + 1); brightness = $b }
    }
  }

  # ========== 3 srgb fills, transparency ==========
  foreach ($c in @('FF0000', '00B050', '7030A0', '123456', 'FFC000')) {
    $s = New-Rect ('srgb_' + $c) 60 40 1; Solid-Fill $s $c
    Probe $s 'srgb' $CENTER '' @{ rgb = $c }
  }
  foreach ($t in @(0.25, 0.5, 0.75)) {
    $s = New-Rect ('alpha_FF0000_' + $t) 60 40 1; Solid-Fill $s 'FF0000'; $s.Fill.Transparency = $t
    Probe $s 'alpha' $CENTER '' @{ rgb = 'FF0000'; transparency = $t; note = 'blended over the magenta background (bg)' }
  }
  $s = New-Rect 'alpha_accent1_0.5' 60 40 1; $s.Fill.Visible = -1; $s.Fill.Solid(); $s.Fill.ForeColor.ObjectThemeColor = 5; $s.Fill.Transparency = 0.5; $s.Line.Visible = 0
  Probe $s 'alpha' $CENTER '' @{ slot = 'accent1'; transparency = 0.5 }

  # ========== 4 gradients (2 and 3 stops): the centre pixel is the 50% position ==========
  foreach ($g in @(@('grad2_red_blue', 'FF0000', '0000FF'), @('grad2_black_white', '000000', 'FFFFFF'), @('grad2_accent', '4472C4', 'ED7D31'))) {
    $s = New-Rect $g[0] 60 40 1
    Try-Do $g[0] {
      $s.Fill.Visible = -1; $s.Fill.TwoColorGradient(1, 1)
      $s.Fill.GradientStops.Item(1).Color.RGB = (Rgb-Int $g[1]); $s.Fill.GradientStops.Item(2).Color.RGB = (Rgb-Int $g[2]); $s.Line.Visible = 0
    }
    Probe $s 'gradient' $CENTER '' @{ stops = @(@{ pos = 0; rgb = $g[1] }, @{ pos = 1; rgb = $g[2] }) }
  }
  $s = New-Rect 'grad3_black_red_white' 60 40 1
  Try-Do 'grad3' {
    $s.Fill.Visible = -1; $s.Fill.TwoColorGradient(1, 1)
    $s.Fill.GradientStops.Item(1).Color.RGB = (Rgb-Int '000000'); $s.Fill.GradientStops.Item(2).Color.RGB = (Rgb-Int 'FFFFFF')
    $s.Fill.GradientStops.Insert((Rgb-Int 'FF0000'), 0.3); $s.Line.Visible = 0
  }
  Probe $s 'gradient' $CENTER '' @{ stops = @(@{ pos = 0; rgb = '000000' }, @{ pos = 0.3; rgb = 'FF0000' }, @{ pos = 1; rgb = 'FFFFFF' }) }

  # ========== 5 patterns (fore black / back white; MsoPatternType numbers) — the patch is 24x24 px so the pattern averages out ==========
  $pat = @(@('pct5', 1), @('pct25', 4), @('pct50', 7), @('pct75', 10), @('pct90', 12), @('horz', 49), @('dkHorz', 13), @('smCheck', 17))
  foreach ($pp in $pat) {
    $s = New-Rect ('patt_' + $pp[0]) 60 40 1
    Try-Do ('patt_' + $pp[0]) {
      $s.Fill.Visible = -1; $s.Fill.Patterned($pp[1]); $s.Fill.ForeColor.RGB = (Rgb-Int '000000'); $s.Fill.BackColor.RGB = (Rgb-Int 'FFFFFF'); $s.Line.Visible = 0
    }
    Probe $s 'pattern' @(, @(0.5, 0.5, 12)) '' @{ prst = $pp[0]; msoPattern = $pp[1]; fg = '000000'; bg = 'FFFFFF' }
  }
  $s = New-Rect 'patt_pct25_red_blue' 60 40 1
  Try-Do 'patt color' { $s.Fill.Visible = -1; $s.Fill.Patterned(4); $s.Fill.ForeColor.RGB = (Rgb-Int 'FF0000'); $s.Fill.BackColor.RGB = (Rgb-Int '0000FF'); $s.Line.Visible = 0 }
  Probe $s 'pattern' @(, @(0.5, 0.5, 12)) '' @{ prst = 'pct25'; msoPattern = 4; fg = 'FF0000'; bg = '0000FF' }

  # ========== 6 picture fill (the source: a 16x16 green PNG written next to the output; the spec paints it #D9D9D9, the pixel is recorded for the record) ==========
  $src = Join-Path $Dir 'pic-fill.png'
  $bmp = New-Object System.Drawing.Bitmap(16, 16)
  $gr = [System.Drawing.Graphics]::FromImage($bmp); $gr.Clear([System.Drawing.Color]::FromArgb(0, 160, 80)); $gr.Dispose()
  $bmp.Save($src, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
  $s = New-Rect 'blip_green' 60 40 1
  Try-Do 'blip' { $s.Fill.Visible = -1; $s.Fill.UserPicture($src); $s.Line.Visible = 0 }
  Probe $s 'blip' $CENTER '' @{ source = 'pic-fill.png'; sourceRgb = '00A050' }

  # ========== 7 glyph text colours: explicit RGB, theme + brightness, per-run ==========
  $s = New-Rect 'text_rgb_FF0000' 100 100 1; $s.Fill.Visible = 0; $s.Line.Visible = 0; Set-Glyph $s 'n'; $s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = (Rgb-Int 'FF0000')
  Probe $s 'text' $CENTER '' @{ rgb = 'FF0000' }
  $s = New-Rect 'text_accent2_-0.25' 100 100 1; $s.Fill.Visible = 0; $s.Line.Visible = 0; Set-Glyph $s 'n'
  Try-Do 'text theme' { $f = $s.TextFrame2.TextRange.Font.Fill.ForeColor; $f.ObjectThemeColor = 6; $f.Brightness = -0.25 }
  Probe $s 'text' $CENTER '' @{ slot = 'accent2'; brightness = -0.25 }
  $s = New-Rect 'text_runs_green_blue' 100 100 1; $s.Fill.Visible = 0; $s.Line.Visible = 0; Set-Glyph $s 'nn'
  Try-Do 'text runs' {
    $tr = $s.TextFrame2.TextRange
    $tr.Characters(1, 1).Font.Fill.ForeColor.RGB = (Rgb-Int '00B050'); $tr.Characters(2, 1).Font.Fill.ForeColor.RGB = (Rgb-Int '0000FF')
  }
  Probe $s 'text' @(@(0.32, 0.5, 4), @(0.68, 0.5, 4)) '' @{ runs = @('00B050', '0000FF'); note = 'sample points are guesses for the two square glyphs' }

  # ========== 8 6pt outlines (fill none): the first stroke along the middle row / column is sampled ==========
  $s = New-Rect 'ln_rgb_FF0000' 60 40 1; $s.Fill.Visible = 0; $s.Line.Weight = 6; $s.Line.ForeColor.RGB = (Rgb-Int 'FF0000')
  Probe $s 'outline' @() 'h' @{ rgb = 'FF0000' }
  $s = New-Rect 'ln_accent1_-0.5' 60 40 1; $s.Fill.Visible = 0; $s.Line.Weight = 6
  Try-Do 'ln theme' { $s.Line.ForeColor.ObjectThemeColor = 5; $s.Line.ForeColor.Brightness = -0.5 }
  Probe $s 'outline' @() 'h' @{ slot = 'accent1'; brightness = -0.5 }
  $s = New-Rect 'ln_dk1' 60 40 1; $s.Fill.Visible = 0; $s.Line.Weight = 6
  Try-Do 'ln dk1' { $s.Line.ForeColor.ObjectThemeColor = 1 }
  Probe $s 'outline' @() 'h' @{ slot = 'dk1' }
  $s = New-Rect 'ln_alpha_0.5_FF0000' 60 40 1; $s.Fill.Visible = 0; $s.Line.Weight = 6; $s.Line.ForeColor.RGB = (Rgb-Int 'FF0000'); $s.Line.Transparency = 0.5
  Probe $s 'outline' @() 'h' @{ rgb = 'FF0000'; transparency = 0.5; note = 'blended over the magenta background (bg)' }
  $s = New-HLine 'lnh_rgb_00B050' 60; $s.Line.Weight = 6; $s.Line.ForeColor.RGB = (Rgb-Int '00B050')
  Probe $s 'outline' @() 'v' @{ rgb = '00B050' }
  $s = New-HLine 'lnh_accent6_+0.4' 60; $s.Line.Weight = 6
  Try-Do 'lnh theme' { $s.Line.ForeColor.ObjectThemeColor = 10; $s.Line.ForeColor.Brightness = 0.4 }
  Probe $s 'outline' @() 'v' @{ slot = 'accent6'; brightness = 0.4 }

  # ========== 9 shaded faces (U8): can / cube / foldedCorner / bevel, fill colours below, no outline. Sample fractions are of the picture box ==========
  $faces = @(
    @('can', 13, 60, 80, @(@('top', 0.5, 0.09), @('body', 0.5, 0.6))),
    @('cube', 14, 70, 70, @(@('front', 0.3, 0.65), @('top', 0.5, 0.12), @('side', 0.87, 0.4))),
    @('foldedCorner', 16, 70, 50, @(@('body', 0.4, 0.4), @('fold1', 0.92, 0.86), @('fold2', 0.95, 0.92))),
    @('bevel', 15, 60, 60, @(@('center', 0.5, 0.5), @('top', 0.5, 0.06), @('left', 0.06, 0.5), @('right', 0.94, 0.5), @('bottom', 0.5, 0.94)))
  )
  foreach ($f in $faces) {
    foreach ($c in @('4472C4', 'FF0000', '808080', 'FFC000')) {
      $s = New-Rect ('face_' + $f[0] + '_' + $c) $f[2] $f[3] $f[1]; Solid-Fill $s $c
      $pts = @(); $names = @()
      foreach ($q in $f[4]) { $pts += , @($q[1], $q[2], 3); $names += $q[0] }
      Probe $s 'face' $pts '' @{ shape = $f[0]; autoShapeType = $f[1]; rgb = $c; sampleNames = $names }
    }
  }

  # ========== record ==========
  $wb.SaveAs((Join-Path $Dir 'ref-colors.xlsx'), 51); $wb.Close($false)
  $out = [ordered]@{ excelVersion = $xl.Version; bg = $BG; notes = $notes; shapes = $records }
  $json = $out | ConvertTo-Json -Depth 10
  [IO.File]::WriteAllText((Join-Path $Dir 'com.json'), $json, (New-Object Text.UTF8Encoding($false)))
  "ok: $($records.Count) shapes; notes=$($notes.Count)"
  $notes
} finally { try { $xl.Quit() } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }
