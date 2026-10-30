/// <reference types="node" />
/**
 * PNG から目盛線・軸線の色を読む器（`__fixtures__/excel/make-ref-charts-png.mjs`・ISSUE-105）の単体テスト。
 * Excel の PNG が無くても動くよう、合成の画像（縦棒・横棒・円）で次を固定する：
 * 線の行の検出（棒の端・系列の色を線と取り違えない）・役割（主の目盛線・補助の目盛線・項目軸の線・値の軸の線）・ink（背景との差の合計）・
 * 2 行にまたがる線の ink・円（軸なし）・`run` の入出力（png-index → ref-charts-png.json）。
 */
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { encodePng } from '../../fixtures/excel/geom-compare.mjs';
import {
  analyzeChartPng,
  run,
  type PngIndexEntry,
} from '../../fixtures/excel/make-ref-charts-png.mjs';

interface Img {
  width: number;
  height: number;
  data: Uint8Array;
}

/** 白い背景の画像 */
function blank(w: number, h: number): Img {
  return { width: w, height: h, data: new Uint8Array(w * h * 4).fill(255) };
}
const rgb = (hex: string): [number, number, number] => [
  parseInt(hex.slice(0, 2), 16),
  parseInt(hex.slice(2, 4), 16),
  parseInt(hex.slice(4, 6), 16),
];
function put(img: Img, x: number, y: number, hex: string): void {
  const i = (y * img.width + x) * 4;
  const [r, g, b] = rgb(hex);
  img.data[i] = r;
  img.data[i + 1] = g;
  img.data[i + 2] = b;
  img.data[i + 3] = 255;
}
function rect(img: Img, x0: number, y0: number, x1: number, y1: number, hex: string): void {
  for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) put(img, x, y, hex);
}

/** 縮尺 2（300×180pt → 600×360px）。プロットの内側 = (34,24)pt から 240×120pt ＝ px の (68,48)-(548,288) */
const entry = (over: Partial<PngIndexEntry> = {}): PngIndexEntry => ({
  Name: 't',
  Source: 'test',
  File: 't.png',
  WidthPt: 300,
  HeightPt: 180,
  PlotInside: { Left: 30, Top: 20, Width: 240, Height: 120 },
  HasAxes: true,
  Horizontal: false,
  ValueAxis: { Min: 0, Max: 20, Major: 5, Minor: 1 },
  ExcludeRGB: ['4472C4'],
  ...over,
});

/** 縦棒の絵：主の目盛線 D9D9D9（1 行）・項目軸の線 2 行（808080・E0E0E0）・値の軸の線 000000（1 列）・棒 1 本（上の縁は薄まった色） */
function columnPicture(): Img {
  const img = blank(600, 360);
  for (const y of [48, 108, 168, 228]) rect(img, 68, y, 548, y, 'D9D9D9');
  rect(img, 68, 288, 548, 288, '808080');
  rect(img, 68, 289, 548, 289, 'E0E0E0');
  rect(img, 68, 48, 68, 288, '000000');
  rect(img, 100, 150, 160, 287, '4472C4');
  rect(img, 100, 149, 160, 149, '8FA8DC'); // 棒の上の縁（アンチエイリアス）。線ではない
  return img;
}

describe('analyzeChartPng（縦棒）', () => {
  const r = analyzeChartPng(columnPicture(), entry());

  it('縮尺とプロットの外接を PNG の幅と COM の値から出す', () => {
    expect(r.scale).toBe(2);
    expect(r.bg).toBe('FFFFFF');
    expect(r.plotPx).toEqual({ x0: 68, y0: 48, x1: 548, y1: 288 });
  });

  it('主の目盛線を 4 本（上端・3 本の内側）と、色・ink を読む', () => {
    const grids = r.lines.filter((l) => l.role === 'majorGrid');
    expect(grids.map((l) => l.pos)).toEqual([48, 108, 168, 228]);
    expect(grids[0]!.peak).toBe('D9D9D9');
    expect(grids[0]!.ink).toEqual([38, 38, 38]);
    expect(r.summary.majorGrid).toMatchObject({ count: 4, peak: 'D9D9D9', rows: 1 });
  });

  it('2 行にまたがる項目軸の線は 1 本にまとめ、ink は行の合計', () => {
    const axis = r.lines.filter((l) => l.role === 'catAxisLine');
    expect(axis).toHaveLength(1);
    expect(axis[0]!.rows.map((x) => x.p)).toEqual([288, 289]);
    expect(axis[0]!.peak).toBe('808080');
    expect(axis[0]!.ink).toEqual([127 + 31, 127 + 31, 127 + 31]);
    expect(axis[0]!.pos).toBeGreaterThan(288);
    expect(axis[0]!.pos).toBeLessThan(289);
  });

  it('値の軸の線（左端の縦線）を読み、棒の縁と系列の色は線にしない', () => {
    const v = r.lines.filter((l) => l.dir === 'v');
    expect(v.map((l) => [l.role, l.peak])).toEqual([['valAxisLine', '000000']]);
    expect(
      r.lines.some((l) => l.rows.some((x) => x.color === '8FA8DC' || x.color === '4472C4')),
    ).toBe(false);
  });

  it('補助の目盛線は主の目盛と重ならない位置で minorGrid になる', () => {
    const img = columnPicture();
    for (const y of [60, 72]) rect(img, 68, y, 548, y, 'F2F2F2'); // 値 19・18（1 刻みの補助）
    const m = analyzeChartPng(img, entry());
    expect(m.lines.filter((l) => l.role === 'minorGrid').map((l) => l.pos)).toEqual([60, 72]);
    expect(m.summary.minorGrid).toMatchObject({ count: 2, peak: 'F2F2F2' });
  });
});

describe('analyzeChartPng（横棒・円）', () => {
  it('横棒は値の方向が縦線：項目軸の線は 0 の位置の縦線・値の軸の線は下端の横線', () => {
    // 同じ座標系で転置：値 0〜20 が x の 68〜548。主の目盛線は x=68+120k（k=1..4）。項目軸の線は x=68、値の軸の線は y=288
    const img = blank(600, 360);
    for (const x of [188, 308, 428, 548]) rect(img, x, 48, x, 288, 'D9D9D9');
    rect(img, 68, 48, 68, 288, '595959');
    rect(img, 68, 288, 548, 288, 'BFBFBF');
    const r = analyzeChartPng(img, entry({ Horizontal: true }));
    expect(r.lines.filter((l) => l.role === 'majorGrid').map((l) => l.pos)).toEqual([
      188, 308, 428, 548,
    ]);
    expect(
      r.lines.filter((l) => l.role === 'catAxisLine').map((l) => [l.dir, l.pos, l.peak]),
    ).toEqual([['v', 68, '595959']]);
    expect(
      r.lines.filter((l) => l.role === 'valAxisLine').map((l) => [l.dir, l.pos, l.peak]),
    ).toEqual([['h', 288, 'BFBFBF']]);
  });

  it('円（軸なし）は線を探さない', () => {
    const r = analyzeChartPng(columnPicture(), entry({ HasAxes: false }));
    expect(r.lines).toEqual([]);
    expect(r.plotPx).toBeUndefined();
  });

  it('目盛の範囲が無いと線は取っても役割は other・プロットの値が無いと線は無い', () => {
    const noScale = analyzeChartPng(columnPicture(), entry({ ValueAxis: null }));
    expect(noScale.lines.length).toBeGreaterThan(0);
    expect(noScale.lines.filter((l) => l.dir === 'h').every((l) => l.role === 'other')).toBe(true);
    const noPlot = analyzeChartPng(columnPicture(), entry({ PlotInside: {} }));
    expect(noPlot.lines).toEqual([]);
  });
});

describe('run（png-index → ref-charts-png.json）', () => {
  it('png-index*.json の全グラフを解き、PNG が無いものは notes に書く', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ref-charts-png-'));
    try {
      const png = join(dir, 'png');
      mkdirSync(png);
      const img = columnPicture();
      writeFileSync(join(png, 't.png'), encodePng(img.width, img.height, img.data));
      const idx = {
        excelVersion: '16.0',
        source: 'x',
        notes: [],
        charts: [entry({ Name: 'a', File: 't.png' }), entry({ Name: 'b', File: 'missing.png' })],
      };
      writeFileSync(join(png, 'png-index.x.json'), JSON.stringify(idx));
      // PowerShell の ConvertTo-Json は 1 件の配列を値にすることがある：charts が 1 件のオブジェクトでも読む
      writeFileSync(
        join(png, 'png-index.y.json'),
        JSON.stringify({
          excelVersion: '16.0',
          notes: 'n1',
          charts: entry({ Name: 'c', File: 't.png', HasAxes: false }),
        }),
      );
      const out = run(dir);
      expect(Object.keys(out.charts).sort()).toEqual(['a', 'c']);
      expect(out.charts.a!.summary.majorGrid!.count).toBe(4);
      expect(out.charts.c!.lines).toEqual([]);
      expect(out.notes).toEqual(['b: PNG が無い（missing.png）', 'n1']);
      const saved = JSON.parse(
        readFileSync(join(dir, 'ref-charts-png.json'), 'utf8'),
      ) as typeof out;
      expect(saved.charts.a!.bg).toBe('FFFFFF');
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('png-index が無ければエラー', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ref-charts-png-'));
    try {
      mkdirSync(join(dir, 'png'));
      expect(() => run(dir)).toThrow(/png-index/);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
