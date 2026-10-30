/// <reference types="node" />
/**
 * PNG の読み取りの器（`__fixtures__/excel/make-ref-charts-png.mjs`）の、ISSUE-110（3 回目の照合）で足した部分の単体テスト。
 * Excel の PNG が無くても動くよう合成の画像で次を固定する：面の角（scanCorner）・項目軸の文字の塊と角度（scanLabels）・数値軸の目盛の印（scanTicks）・
 * index の印（Corner・LabelBand・Ticks）からの呼び出し・`run` の出力名と index の名前の指定（既定のままなら 2 回目と同じ出力）。
 */
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { encodePng } from '../../fixtures/excel/geom-compare.mjs';
import {
  analyzeChartPng,
  run,
  scanCorner,
  scanLabels,
  scanTicks,
  type PngIndexEntry,
} from '../../fixtures/excel/make-ref-charts-png.mjs';

interface Img {
  width: number;
  height: number;
  data: Uint8Array;
}

/** 白い背景（不透明）の画像 */
function blank(w: number, h: number): Img {
  return { width: w, height: h, data: new Uint8Array(w * h * 4).fill(255) };
}
function put(img: Img, x: number, y: number, rgb = 0x000000): void {
  const i = (y * img.width + x) * 4;
  img.data[i] = (rgb >> 16) & 255;
  img.data[i + 1] = (rgb >> 8) & 255;
  img.data[i + 2] = rgb & 255;
  img.data[i + 3] = 255;
}
function rect(img: Img, x0: number, y0: number, x1: number, y1: number, rgb = 0x000000): void {
  for (let y = y0; y <= y1; y++) for (let x = x0; x <= x1; x++) put(img, x, y, rgb);
}
/** 左上の角を半径 r の四分円で丸め、外側を透明にする */
function roundTopLeft(img: Img, r: number): void {
  for (let y = 0; y < r; y++)
    for (let x = 0; x < r; x++)
      if ((r - x) ** 2 + (r - y) ** 2 > r * r) img.data[(y * img.width + x) * 4 + 3] = 0;
}

describe('scanCorner（面の角）', () => {
  it('直角の面：(0,0) が枠線の色 → rounded=false・半径 0', () => {
    const img = blank(100, 60);
    rect(img, 0, 0, 99, 0, 0x0000ff);
    rect(img, 0, 0, 0, 59, 0x0000ff);
    expect(scanCorner(img)).toEqual({
      rounded: false,
      transparent: false,
      outside: null,
      radiusPx: 0,
      runs: [],
    });
  });

  it('角丸の面（外側が透明）：半径 10 → radiusPx 10・1 行目の外側の連なりが最も長い', () => {
    const img = blank(100, 60);
    rect(img, 10, 0, 99, 0, 0x0000ff);
    roundTopLeft(img, 10);
    const c = scanCorner(img);
    expect(c.rounded).toBe(true);
    expect(c.transparent).toBe(true);
    expect(c.radiusPx).toBe(10);
    expect(c.runs[0]).toBeGreaterThan(c.runs[5]!);
    expect(c.runs[9]).toBe(1);
  });

  it('角丸の面（外側が不透明な背景色）：外側の色を返す', () => {
    const img = blank(100, 60);
    rect(img, 0, 0, 99, 59, 0xffffff);
    rect(img, 8, 0, 99, 0, 0x0000ff); // 上の枠線（中ほどは白と違う色）
    const c = scanCorner(img);
    expect(c.rounded).toBe(true);
    expect(c.transparent).toBe(false);
    expect(c.outside).toBe('FFFFFF');
  });
});

describe('scanLabels（項目軸の文字の塊）', () => {
  const region = { x0: 0, y0: 90, x1: 299, y1: 199 };

  it('水平・縦・右上がりの 3 つの塊を数え、角度を主軸から出す。目盛の印のような細い線は除く', () => {
    const img = blank(300, 200);
    rect(img, 20, 100, 40, 105); // 水平（21×6）
    rect(img, 100, 100, 105, 120); // 縦（6×21）
    for (let i = 0; i <= 24; i++) put(img, 160 + i, 150 - i); // 右上がりの斜め線（45°）
    rect(img, 250, 92, 250, 97); // 目盛の印（幅 1・画素 6）
    const r = scanLabels(img, 0xffffff, region, 3);
    expect(r.count).toBe(3);
    const [h, v, d] = r.clusters;
    expect(h!.w).toBe(21);
    expect(Math.abs(h!.angle)).toBeLessThan(5);
    expect(Math.abs(v!.angle)).toBeGreaterThan(85);
    expect(d!.angle).toBeGreaterThan(40);
    expect(d!.angle).toBeLessThan(50);
    expect(r.medianW).toBe(21); // 幅 21・6・25 の中央値
  });

  it('文字の間（g 画素以内）は 1 つの塊、g より離れれば別の塊', () => {
    const img = blank(300, 200);
    rect(img, 20, 100, 24, 110);
    rect(img, 27, 100, 31, 110); // 隙間 2 画素（g=3 で 1 つ）
    rect(img, 60, 100, 64, 110); // 隙間 28 画素（別）
    expect(scanLabels(img, 0xffffff, region, 3).count).toBe(2);
    expect(scanLabels(img, 0xffffff, region, 1).count).toBe(3);
  });

  it('範囲が画像の外・空なら塊は 0', () => {
    const img = blank(300, 200);
    expect(scanLabels(img, 0xffffff, { x0: 400, y0: 0, x1: 500, y1: 10 }, 3).count).toBe(0);
  });
});

describe('scanTicks（数値軸の目盛の印）', () => {
  // 縦の軸：x0=100・y0=50・y1=250。縮尺 2（cap=14）。主の目盛は 50・150・250、補助は 100・200（と主の位置）
  const geo = {
    vert: true,
    x0: 100,
    y0: 50,
    x1: 400,
    y1: 250,
    scale: 2,
    majorPos: [50, 150, 250],
    minorPos: [50, 100, 150, 200, 250],
    tol: 3,
  };

  it('外側の主の印（長さ 8）3 本と補助の印（長さ 5）2 本を数え、内側は 0', () => {
    const img = blank(500, 300);
    for (const y of [50, 150, 250]) rect(img, 92, y, 99, y);
    for (const y of [100, 200]) rect(img, 95, y, 99, y);
    const t = scanTicks(img, 0xffffff, geo);
    expect(t.cap).toBe(14);
    expect(t.out).toMatchObject({ major: 3, minor: 2, other: 0, majorLen: 8, minorLen: 5 });
    expect(t.in).toMatchObject({ major: 0, minor: 0, other: 0 });
  });

  it('内側の補助の印。目盛線のような長い線（cap 以上）と、位置が合わない印の扱い', () => {
    const img = blank(500, 300);
    rect(img, 101, 100, 104, 100); // 内側の補助の印（長さ 4）
    rect(img, 101, 150, 400, 150); // 主の目盛線（長い）→ 印にしない
    rect(img, 92, 120, 99, 120); // 外側・どの目盛の位置でもない
    const t = scanTicks(img, 0xffffff, geo);
    expect(t.in).toMatchObject({ major: 0, minor: 1, other: 0, minorLen: 4 });
    expect(t.out).toMatchObject({ major: 0, minor: 0, other: 1 });
  });

  it('横の軸（横棒の値の軸）：下端から下が外側', () => {
    const img = blank(500, 300);
    rect(img, 150, 251, 150, 258); // 下へ 8（外側）
    rect(img, 200, 245, 200, 249); // 上へ 5（内側）
    const t = scanTicks(img, 0xffffff, {
      vert: false,
      x0: 100,
      y0: 50,
      x1: 400,
      y1: 250,
      scale: 2,
      majorPos: [150],
      minorPos: [200],
      tol: 3,
    });
    expect(t.out).toMatchObject({ major: 1, majorLen: 8 });
    expect(t.in).toMatchObject({ minor: 1, minorLen: 5 });
  });
});

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

describe('analyzeChartPng の ISSUE-110 の印（Corner・LabelBand・Ticks）', () => {
  function picture(): Img {
    const img = blank(600, 360);
    for (const y of [288, 228, 168, 108, 48]) rect(img, 62, y, 67, y); // 数値軸の外側の主の印（長さ 6・軸の線なし）
    rect(img, 100, 296, 140, 303); // 項目軸の文字の塊（プロットの下）
    roundTopLeft(img, 12);
    return img;
  }

  it('印が無いグラフには corner・labels・valTicks を付けない（2 回目の結果と同じ）', () => {
    const r = analyzeChartPng(picture(), entry());
    expect(r.corner).toBeUndefined();
    expect(r.labels).toBeUndefined();
    expect(r.valTicks).toBeUndefined();
  });

  it('印があれば読む：角の半径・項目軸の文字の塊（1 つ）・数値軸の目盛の印（主 5 本）', () => {
    const r = analyzeChartPng(
      picture(),
      entry({
        Corner: true,
        Ticks: true,
        LabelBand: { Left: 30, Top: 140, Width: 240, Height: 30 },
      }),
    );
    expect(r.corner).toMatchObject({ rounded: true, radiusPx: 12 });
    expect(r.labels!.count).toBe(1);
    expect(r.labels!.clusters[0]).toMatchObject({ x0: 100, x1: 140, y0: 296, y1: 303 });
    expect(r.valTicks!.out).toMatchObject({ major: 5, majorLen: 6 });
  });
});

describe('run の出力名・index の名前の指定', () => {
  it('既定は ref-charts-png.json・png-index*.json。opts.out / opts.index で別のファイルにできる', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ref-charts-png3-'));
    try {
      const png = join(dir, 'png');
      mkdirSync(png);
      const img = blank(600, 360);
      writeFileSync(join(png, 't.png'), encodePng(img.width, img.height, img.data));
      const idx = (name: string) =>
        JSON.stringify({
          excelVersion: '16.0',
          notes: [],
          charts: [entry({ Name: name, File: 't.png' })],
        });
      writeFileSync(join(png, 'png-index.ref-charts3.json'), idx('a'));
      writeFileSync(join(png, 'png-index.other.json'), idx('b'));
      const all = run(dir);
      expect(Object.keys(all.charts).sort()).toEqual(['a', 'b']);
      const only = run(dir, {
        out: 'ref-charts3-png.json',
        index: '^png-index\\.ref-charts3\\.json$',
      });
      expect(Object.keys(only.charts)).toEqual(['a']);
      const saved = JSON.parse(
        readFileSync(join(dir, 'ref-charts3-png.json'), 'utf8'),
      ) as typeof only;
      expect(Object.keys(saved.charts)).toEqual(['a']);
      // 既定の出力は、2 つ目の呼び出しで壊れていない
      const first = JSON.parse(
        readFileSync(join(dir, 'ref-charts-png.json'), 'utf8'),
      ) as typeof all;
      expect(Object.keys(first.charts).sort()).toEqual(['a', 'b']);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
