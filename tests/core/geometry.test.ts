import { describe, expect, it } from 'vitest';
import {
  colWidthToPx,
  createGeometry,
  defaultColWidthToPx,
  ptToPx,
  type Geometry,
} from '../../src/core/geometry';
import { emptyLayout, type SheetLayout } from '../../src/core/types';

/** 指定した項目だけ替えた layout（既定は emptyLayout） */
const layoutOf = (over: Partial<SheetLayout> = {}): SheetLayout => ({ ...emptyLayout(), ...over });

describe('TD-83-1 換算式（pt・文字数 → px）', () => {
  it('行高：round(pt × 96 / 72)（15pt＝20px、18.75pt＝25px）', () => {
    expect(ptToPx(15)).toBe(20);
    expect(ptToPx(18.75)).toBe(25);
    expect(ptToPx(0)).toBe(0);
    expect(ptToPx(409.5)).toBe(546);
  });

  it('列幅：trunc(((256·w + trunc(128/MDW)) / 256) · MDW)', () => {
    // 余白込みの保存値 9.140625（Calibri 11 の既定の 8.43 文字）は 64px
    expect(colWidthToPx(9.140625, 7)).toBe(64);
    expect(colWidthToPx(10, 7)).toBe(70);
    expect(colWidthToPx(9, 8)).toBe(72);
    expect(colWidthToPx(0, 7)).toBe(0);
    expect(colWidthToPx(255, 7)).toBe(1785);
  });

  it('既定の列幅（defaultColWidth 無し）：ceil((base·MDW + 5) / 8) · 8（MDW 7＝64px、MDW 8＝72px）', () => {
    expect(defaultColWidthToPx(8, 7)).toBe(64);
    expect(defaultColWidthToPx(8, 8)).toBe(72);
  });

  it('既定の列幅の余白は 2·ceil(MDW/4)+1（MDW 9 以上では 5 より大きい：MDW 9 は 80px）', () => {
    expect(defaultColWidthToPx(8, 9)).toBe(80);
    expect(defaultColWidthToPx(8, 5)).toBe(48);
  });
});

describe('TD-83-2 createGeometry：既定値のみ', () => {
  it('寸法の指定が無いとき、MDW 7 は 64×20px、MDW 8 は 72×20px で並ぶ', () => {
    const g7 = createGeometry(layoutOf(), 10, 5);
    expect([g7.colWidth(1), g7.rowHeight(1)]).toEqual([64, 20]);
    expect([g7.colLeft(1), g7.colLeft(2), g7.colLeft(5)]).toEqual([0, 64, 256]);
    expect([g7.rowTop(1), g7.rowTop(2), g7.rowTop(10)]).toEqual([0, 20, 180]);
    expect([g7.totalWidth, g7.totalHeight]).toEqual([320, 200]);

    const g8 = createGeometry(layoutOf({ mdw: 8 }), 3, 2);
    expect([g8.colWidth(1), g8.totalWidth]).toEqual([72, 144]);
  });

  it('defaultRowHeight・defaultColWidth を指定すると、既定の寸法が替わる', () => {
    const g = createGeometry(layoutOf({ defaultRowHeight: 18.75, defaultColWidth: 10 }), 4, 3);
    expect([g.rowHeight(2), g.totalHeight]).toEqual([25, 100]);
    expect([g.colWidth(2), g.totalWidth]).toEqual([70, 210]);
  });

  it('範囲外の番号：1 未満の位置は 0、rows／cols より後ろも同じ既定の寸法で延びる', () => {
    const g = createGeometry(layoutOf(), 2, 2);
    expect([g.rowTop(0), g.rowTop(-3), g.colLeft(0)]).toEqual([0, 0, 0]);
    expect([g.rowTop(3), g.rowTop(11), g.rowHeight(11)]).toEqual([40, 200, 20]);
    expect([g.colLeft(3), g.colWidth(9)]).toEqual([128, 64]);
    expect([g.rowHeight(0), g.colWidth(0)]).toEqual([0, 0]);
  });

  it('行・列が 0 のとき totalWidth／totalHeight は 0', () => {
    const g = createGeometry(layoutOf(), 0, 0);
    expect([g.totalWidth, g.totalHeight]).toEqual([0, 0]);
    expect([g.rowAt(5), g.colAt(5)]).toEqual([0, 0]);
  });
});

describe('TD-83-3 createGeometry：疎な列幅・行高・非表示', () => {
  it('列幅の指定（範囲指定）：範囲を展開せず、累積が既定との差分で出る', () => {
    const g = createGeometry(
      layoutOf({ cols: [{ min: 2, max: 3, width: 10, hidden: false }] }),
      1,
      5,
    );
    expect([1, 2, 3, 4, 5].map((c) => g.colWidth(c))).toEqual([64, 70, 70, 64, 64]);
    expect([1, 2, 3, 4, 5].map((c) => g.colLeft(c))).toEqual([0, 64, 134, 204, 268]);
    expect(g.totalWidth).toBe(332);
  });

  it('非表示の列は幅 0（後ろの列が左へ詰まる）。width が無い列は既定の幅', () => {
    const g = createGeometry(
      layoutOf({
        cols: [
          { min: 2, max: 2, width: null, hidden: true },
          { min: 4, max: 4, width: null, hidden: false },
        ],
      }),
      1,
      4,
    );
    expect([1, 2, 3, 4].map((c) => g.colWidth(c))).toEqual([64, 0, 64, 64]);
    expect([1, 2, 3, 4].map((c) => g.colLeft(c))).toEqual([0, 64, 64, 128]);
    expect(g.totalWidth).toBe(192);
  });

  it('行高の指定（pt）と非表示の行（高さ 0）', () => {
    const g = createGeometry(
      layoutOf({
        rows: [
          { min: 2, max: 2, ht: 30, hidden: false, custom: true },
          { min: 4, max: 4, ht: 99, hidden: true, custom: true },
        ],
      }),
      5,
      1,
    );
    expect([1, 2, 3, 4, 5].map((r) => g.rowHeight(r))).toEqual([20, 40, 20, 0, 20]);
    expect([1, 2, 3, 4, 5].map((r) => g.rowTop(r))).toEqual([0, 20, 60, 80, 80]);
    expect(g.totalHeight).toBe(100);
  });

  it('列の指定が重なる・順不同でも、min の小さい範囲を優先して崩れない', () => {
    const g = createGeometry(
      layoutOf({
        cols: [
          { min: 3, max: 4, width: 10, hidden: false },
          { min: 1, max: 3, width: 20, hidden: false },
        ],
      }),
      1,
      5,
    );
    expect([1, 2, 3, 4, 5].map((c) => g.colWidth(c))).toEqual([140, 140, 140, 70, 64]);
    expect(g.totalWidth).toBe(140 * 3 + 70 + 64);
  });

  it('結合範囲の幅・高さは、端の left／top の差で出る（合計）', () => {
    const g = createGeometry(
      layoutOf({
        cols: [{ min: 2, max: 3, width: 10, hidden: false }],
        rows: [
          { min: 2, max: 2, ht: 30, hidden: false, custom: true },
          { min: 3, max: 3, ht: null, hidden: true, custom: true },
        ],
      }),
      5,
      5,
    );
    // B2:D4（列 2〜4・行 2〜4）
    expect(g.colLeft(5) - g.colLeft(2)).toBe(70 + 70 + 64);
    expect(g.rowTop(5) - g.rowTop(2)).toBe(40 + 0 + 20);
  });
});

describe('TD-83-4 createGeometry：zeroHeight', () => {
  it('<row> の無い行は高さ 0、<row> のある行は表示する（ht があればその高さ）', () => {
    const g = createGeometry(
      layoutOf({
        zeroHeight: true,
        rows: [
          { min: 1, max: 1, ht: null, hidden: false, custom: true },
          { min: 3, max: 3, ht: null, hidden: false, custom: true },
          { min: 4, max: 4, ht: 30, hidden: false, custom: true },
        ],
      }),
      5,
      1,
    );
    expect([1, 2, 3, 4, 5].map((r) => g.rowHeight(r))).toEqual([20, 0, 20, 40, 0]);
    expect([1, 2, 3, 4, 5].map((r) => g.rowTop(r))).toEqual([0, 20, 20, 40, 80]);
    expect(g.totalHeight).toBe(80);
  });

  it('zeroHeight でも hidden のある行は 0。<row> がひとつも無ければ全体が 0', () => {
    const g = createGeometry(
      layoutOf({
        zeroHeight: true,
        rows: [{ min: 1, max: 1, ht: null, hidden: true, custom: true }],
      }),
      3,
      1,
    );
    expect(g.totalHeight).toBe(0);
    expect(createGeometry(layoutOf({ zeroHeight: true }), 3, 1).totalHeight).toBe(0);
  });
});

describe('TD-83-5 rowAt／colAt：二分探索と境界', () => {
  const g = createGeometry(
    layoutOf({
      cols: [{ min: 2, max: 2, width: null, hidden: true }],
      rows: [
        { min: 2, max: 2, ht: 30, hidden: false, custom: true },
        { min: 4, max: 4, ht: null, hidden: true, custom: true },
      ],
    }),
    5,
    4,
  );
  // 行の top：0, 20, 60, 80, 80（行 4 は非表示）／列の left：0, 64, 64, 128（列 2 は非表示）

  it('範囲の中：その位置を含む行・列を返す（境目は下の行・右の列）', () => {
    expect([0, 19, 20, 59, 60, 79].map((y) => g.rowAt(y))).toEqual([1, 1, 2, 2, 3, 3]);
    expect([0, 63, 64, 127, 128, 191].map((x) => g.colAt(x))).toEqual([1, 1, 3, 3, 4, 4]);
  });

  it('非表示の行・列は返らず、次の表示行・列へ進む', () => {
    expect(g.rowAt(80)).toBe(5);
    expect(g.colAt(64)).toBe(3);
  });

  it('先頭・末尾・範囲外：負・NaN は先頭、下端以降は最後の行・列へ丸める', () => {
    expect([g.rowAt(-1), g.rowAt(NaN), g.colAt(-100), g.colAt(NaN)]).toEqual([1, 1, 1, 1]);
    expect([g.rowAt(99), g.rowAt(100), g.rowAt(1e9)]).toEqual([5, 5, 5]);
    expect([g.colAt(191), g.colAt(192), g.colAt(1e9)]).toEqual([4, 4, 4]);
  });

  it('rowAt と rowTop は整合する（全行で rowAt(rowTop(r)) が r の高さ内に収まる）', () => {
    for (let r = 1; r <= 5; r++) {
      if (g.rowHeight(r) === 0) continue;
      expect(g.rowAt(g.rowTop(r))).toBe(r);
      expect(g.rowAt(g.rowTop(r) + g.rowHeight(r) - 1)).toBe(r);
    }
  });
});

describe('TD-83-6 巨大なシート：全行・全列の配列を作らず即時に返る', () => {
  const ROWS = 1_048_576;
  const COLS = 16_384;

  it('1,048,576 行・16,384 列（範囲指定の列幅・疎な行高）でも一瞬で作れ、値が正しい', () => {
    const t0 = performance.now();
    const g = createGeometry(
      layoutOf({
        cols: [{ min: 1, max: COLS, width: 10, hidden: false }],
        rows: [{ min: 1_000_000, max: 1_000_000, ht: 30, hidden: false, custom: true }],
      }),
      ROWS,
      COLS,
    );
    expect(g.totalWidth).toBe(COLS * 70);
    expect(g.totalHeight).toBe(ROWS * 20 + 20);
    expect(g.rowTop(1_000_001)).toBe(999_999 * 20 + 40);
    expect(g.rowAt(g.rowTop(1_000_000) + 39)).toBe(1_000_000);
    expect(g.rowAt(g.rowTop(1_000_001))).toBe(1_000_001);
    expect(g.rowAt(g.totalHeight - 1)).toBe(ROWS);
    expect(g.colAt(g.totalWidth - 1)).toBe(COLS);
    expect(g.colLeft(COLS)).toBe((COLS - 1) * 70);
    expect(performance.now() - t0).toBeLessThan(200);
  });

  it('zeroHeight でも、<row> のある行の数に比例するだけで返る', () => {
    const g = createGeometry(
      layoutOf({
        zeroHeight: true,
        rows: [{ min: 500_000, max: 500_000, ht: null, hidden: false, custom: true }],
      }),
      ROWS,
      1,
    );
    expect(g.totalHeight).toBe(20);
    expect(g.rowAt(0)).toBe(500_000);
  });
});

describe('TD-83-12 極端な値でも座標が有限（Excel の上限で丸める）', () => {
  it('列幅 1e307・defaultColWidth 1e307・baseColWidth 1e307 でも totalWidth が有限（255 文字で丸める）', () => {
    const g = createGeometry(
      layoutOf({
        cols: [{ min: 1, max: 2, width: 1e307, hidden: false }],
        defaultColWidth: 1e307,
      }),
      1,
      16_384,
    );
    expect(Number.isFinite(g.totalWidth)).toBe(true);
    expect(g.colWidth(1)).toBe(1785); // 255 文字（MDW 7）
    expect(g.colWidth(3)).toBe(1785);
    const b = createGeometry(layoutOf({ baseColWidth: 1e307 }), 1, 16_384);
    expect(Number.isFinite(b.totalWidth)).toBe(true);
  });

  it('行高 1e307・defaultRowHeight 1e307 でも totalHeight が有限（409.5pt＝546px で丸める）', () => {
    const g = createGeometry(
      layoutOf({
        defaultRowHeight: 1e307,
        rows: [{ min: 1, max: 1, ht: 1e307, hidden: false, custom: true }],
      }),
      1_048_576,
      1,
    );
    expect(Number.isFinite(g.totalHeight)).toBe(true);
    expect(g.rowHeight(1)).toBe(546);
    expect(g.rowHeight(2)).toBe(546);
    expect(g.totalHeight).toBe(1_048_576 * 546);
  });

  it('Infinity・NaN・負の値は既定に戻す', () => {
    const g = createGeometry(
      layoutOf({
        defaultRowHeight: Infinity,
        defaultColWidth: NaN,
        cols: [{ min: 1, max: 1, width: -5, hidden: false }],
        mdw: Infinity,
      }),
      2,
      2,
    );
    expect([g.rowHeight(1), g.colWidth(1), g.colWidth(2)]).toEqual([20, 0, 0]);
    expect(Number.isFinite(g.totalWidth + g.totalHeight)).toBe(true);
  });
});

describe('TD-83-13 既定の行高の規則（customHeight・フォントの既定の行高）', () => {
  const rowsPx = (g: Geometry, n: number) =>
    Array.from({ length: n }, (_, i) => g.rowHeight(i + 1));

  it('customHeight の無い defaultRowHeight は、フォントの既定の行高（fontRowPx）が優先される', () => {
    const g = createGeometry(layoutOf({ defaultRowHeight: 30, fontRowPx: 20 }), 3, 1);
    expect(rowsPx(g, 3)).toEqual([20, 20, 20]);
  });

  it('sheetFormatPr の customHeight="1" なら、defaultRowHeight がそのまま既定の行高', () => {
    const g = createGeometry(
      layoutOf({ defaultRowHeight: 30, customHeight: true, fontRowPx: 20 }),
      3,
      1,
    );
    expect(rowsPx(g, 3)).toEqual([40, 40, 40]);
  });

  it('fontRowPx が無ければ defaultRowHeight の換算値（無ければ 15pt＝20px）', () => {
    expect(rowsPx(createGeometry(layoutOf({ defaultRowHeight: 18 }), 1, 1), 1)).toEqual([24]);
    expect(rowsPx(createGeometry(layoutOf(), 1, 1), 1)).toEqual([20]);
  });

  it('行の ht：customHeight 付きは ptToPx(ht)。無しで ±0.01 以内の差ならシートの既定、違えば ptToPx(ht)', () => {
    const g = createGeometry(
      layoutOf({
        defaultRowHeight: 14.4,
        fontRowPx: 25,
        rows: [
          { min: 1, max: 1, ht: 14.4, hidden: false, custom: true }, // 19
          { min: 2, max: 2, ht: 14.4, hidden: false, custom: false }, // 既定（25）
          { min: 3, max: 3, ht: 14.405, hidden: false, custom: false }, // ±0.01 以内 → 既定
          { min: 4, max: 4, ht: 30, hidden: false, custom: false }, // 違う → 40
          { min: 5, max: 5, ht: 30, hidden: true, custom: true }, // 非表示 → 0
        ],
      }),
      6,
      1,
    );
    expect(rowsPx(g, 6)).toEqual([19, 25, 25, 40, 0, 25]);
  });

  it('zeroHeight：<row> のある行は、ht が無ければシートの既定（fontRowPx 優先）で表示', () => {
    const g = createGeometry(
      layoutOf({
        zeroHeight: true,
        fontRowPx: 25,
        rows: [{ min: 1, max: 2, ht: null, hidden: false, custom: false }],
      }),
      3,
      1,
    );
    expect(rowsPx(g, 3)).toEqual([25, 25, 0]);
  });
});

describe('TD-83-14 隣り合う同じ寸法の範囲はまとめる', () => {
  it('1,048,576 行を 1 行ずつ指定しても（同じ高さ）、軸は 1 件に畳まれ、一瞬で返る', () => {
    const rows = Array.from({ length: 1_048_576 }, (_, i) => ({
      min: i + 1,
      max: i + 1,
      ht: 30,
      hidden: false,
      custom: true,
    }));
    const t0 = performance.now();
    const g = createGeometry(layoutOf({ rows }), 1_048_576, 1);
    expect(g.totalHeight).toBe(1_048_576 * 40);
    expect(g.rowTop(1_000_000)).toBe(999_999 * 40);
    expect(g.rowAt(40 * 500_000)).toBe(500_001);
    expect(performance.now() - t0).toBeLessThan(2000);
  });
});
