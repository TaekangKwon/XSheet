import { describe, expect, it } from 'vitest';
import { openXlsx } from '../../src/core/index';
import {
  borderXml,
  buildThemeXml,
  buildXlsx,
  colLetters,
  type CellSpec,
  type XlsxSpec,
} from '../../src/testing/make-xlsx';
import type { SheetRow, XlsxSheet } from '../../src/core/types';

/** 罫線の表：0＝罫線なし、1＝下の細線、2＝4 辺の太線（xf 番号 s＝borderId と同じ番号にする） */
const BORDERS = [
  borderXml({ left: {}, right: {}, top: {}, bottom: {}, diagonal: {} }),
  borderXml({ bottom: { style: 'thin', color: 'auto="1"' } }),
  borderXml({
    left: { style: 'thick' },
    right: { style: 'thick' },
    top: { style: 'thick' },
    bottom: { style: 'thick' },
  }),
];

/** s＝0（罫線なし）・1（下の細線）・2（4 辺の太線）の書式を持つブックの 1 枚目のシートを開く */
async function open(
  rows: CellSpec[][],
  extra: Partial<XlsxSpec> = {},
  sheetExtra: { merges?: string[]; sheetDataXml?: string } = {},
): Promise<XlsxSheet> {
  const bytes = await buildXlsx({
    sheets: [{ name: 'S', rows, ...sheetExtra }],
    xfs: [0, 0, 0],
    borders: BORDERS,
    xfBorderIds: [0, 1, 2],
    ...extra,
  });
  return (await openXlsx(bytes)).openSheet(1);
}

/** 行の罫線の番号を、borders の下の辺の線種・全辺の太線の目印（'b'＝thin の下辺、'T'＝太線の 4 辺）の配列にする */
function marks(sheet: XlsxSheet, row: SheetRow): (string | undefined)[] {
  return (row.borders ?? []).map((i) => {
    if (i === undefined) return undefined;
    const b = sheet.borders[i]!;
    return b.left?.style === 'thick' ? 'T' : b.bottom?.style === 'thin' ? 'b' : '?';
  });
}

describe('getRows の borders：セルごとの罫線の番号', () => {
  it('値のあるセル・値の無い書式だけのセル・罫線なし（s=0 と s なし）を区別して番号を返す', async () => {
    const sheet = await open([
      [{ v: 'a', t: 'inlineStr', s: 1 }, { s: 2 }, { v: 1, s: 0 }, { v: 'x', t: 'inlineStr' }],
    ]);
    const [row] = sheet.getRows(1, 1);
    expect(row!.cells.map((c) => c?.text)).toEqual(['a', undefined, '1', 'x']);
    expect(row!.borders).toHaveLength(4);
    expect(marks(sheet, row!)).toEqual(['b', 'T', undefined, undefined]);
  });

  it('同じ罫線は同じ番号。XlsxSheet.borders の位置を指す（線の無い罫線は番号を使わない）', async () => {
    const sheet = await open([[{ s: 1 }, { s: 2 }, { s: 1 }]]);
    expect(sheet.borders).toHaveLength(2);
    const b = getRow(sheet, 1).borders!;
    expect(b[0]).toBe(b[2]);
    expect(b[0]).not.toBe(b[1]);
  });

  it('結合で隠れるセル（null）の罫線も返す。結合の左上は CellView のまま', async () => {
    const sheet = await open(
      [
        [{ v: 'm', t: 'inlineStr', s: 1 }, { s: 2 }],
        [{ s: 1 }, { s: 2 }],
      ],
      {},
      { merges: ['A1:B2'] },
    );
    const rows = sheet.getRows(1, 2);
    expect(rows[0]!.cells[0]).toMatchObject({ text: 'm', colSpan: 2, rowSpan: 2 });
    expect(rows[0]!.cells[1]).toBeNull();
    expect(rows[1]!.cells).toEqual([null, null]);
    expect(marks(sheet, rows[0]!)).toEqual(['b', 'T']);
    expect(marks(sheet, rows[1]!)).toEqual(['b', 'T']);
  });

  it('罫線のあるセルが無い行は borders を持たない。返す配列の長さは cells と同じ', async () => {
    const sheet = await open([
      [
        { v: 'a', t: 'inlineStr' },
        { v: 'b', t: 'inlineStr' },
        { v: 'c', t: 'inlineStr', s: 1 },
      ],
      [{ v: 'x', t: 'inlineStr' }, { s: 0 }],
      [{ s: 2 }],
    ]);
    const rows = sheet.getRows(1, 3);
    expect(rows[0]!.borders).toHaveLength(rows[0]!.cells.length);
    expect(marks(sheet, rows[0]!)).toEqual([undefined, undefined, 'b']);
    expect(rows[1]!.borders).toBeUndefined();
    expect(marks(sheet, rows[2]!)).toEqual(['T', undefined, undefined]);
  });

  it('範囲の行だけを返し、範囲外の行の罫線は含めない', async () => {
    const rows: CellSpec[][] = [];
    for (let r = 1; r <= 10; r++) rows.push([{ v: r, s: r % 2 }]);
    const sheet = await open(rows);
    const got = sheet.getRows(3, 5);
    expect(got.map((r) => r.r)).toEqual([3, 4, 5]);
    expect(got.map((r) => marks(sheet, r)[0])).toEqual(['b', undefined, 'b']);
  });

  it('r 属性の無いセルは直前の次の列に置く。r が範囲外のセルは読まない', async () => {
    const sheet = await open(
      [
        [
          { v: 'a', t: 'inlineStr' },
          { s: 1, r: null },
          { s: 2, r: null },
        ],
      ],
      {},
    );
    expect(marks(sheet, getRow(sheet, 1))).toEqual([undefined, 'b', 'T']);
    const sheet2 = await open([
      [
        { v: 'a', t: 'inlineStr' },
        { s: 1, r: 'XFE1' },
      ],
    ]);
    expect(sheet2.getRows(1, 1)[0]!.borders).toBeUndefined();
  });

  it('s が不正・範囲外（cellXfs に無い）のセルは罫線なし', async () => {
    const sheet = await open([[{ v: 'a', t: 'inlineStr', attrs: ' s="99"' }, { attrs: ' s="x"' }]]);
    expect(sheet.getRows(1, 1)[0]!.borders).toBeUndefined();
  });

  it('名前空間の接頭辞付きの XML でも読む', async () => {
    const sheet = await open([[{ v: 'a', t: 'inlineStr' }, { s: 2 }]], { prefix: 'x' });
    expect(marks(sheet, getRow(sheet, 1))).toEqual([undefined, 'T']);
  });
});

describe('表の範囲：罫線だけがある末尾の行・列まで広げる', () => {
  const text = (v: string): CellSpec => ({ v, t: 'inlineStr' });

  it('末尾に書式だけの罫線の行があれば、rowCount がその行まで広がる', async () => {
    const sheet = await open([[text('a')], [text('b')], [{ s: 1 }], [{ s: 0 }]]);
    expect(sheet.rowCount).toBe(3);
    expect(sheet.colCount).toBe(1);
    const rows = sheet.getRows(1, sheet.rowCount);
    expect(marks(sheet, rows[2]!)).toEqual(['b']);
  });

  it('右に書式だけの罫線の列があれば、colCount がその列まで広がる（どの行のものでも）', async () => {
    const sheet = await open([[text('a')], [null, null, null, { s: 2 }], [text('c')]]);
    expect(sheet.rowCount).toBe(3);
    expect(sheet.colCount).toBe(4);
    const rows = sheet.getRows(1, 3);
    expect(rows.every((r) => r.cells.length === 4)).toBe(true);
    expect(marks(sheet, rows[1]!)).toEqual([undefined, undefined, undefined, 'T']);
  });

  it('値の無い罫線だけのシートも rowCount・colCount を持つ（空のシートではない）', async () => {
    const sheet = await open([
      [{ s: 1 }, { s: 1 }],
      [{ s: 2 }, { s: 2 }],
    ]);
    expect([sheet.rowCount, sheet.colCount]).toEqual([2, 2]);
  });

  it('罫線の無い書式（s=0 や、罫線なしの罫線を指す s）だけの末尾は広げない', async () => {
    const sheet = await open([[text('a')], [{ s: 0 }, { s: 0 }, { s: 0 }]]);
    expect([sheet.rowCount, sheet.colCount]).toEqual([1, 1]);
    expect(sheet.getRows(1, 1)[0]!.borders).toBeUndefined();
  });

  it('罫線を持つセルが 1 つも無いブック（borders の指定なし）の行数・列数・cells は今と同じ', async () => {
    const rows: CellSpec[][] = [
      [text('a'), { s: 1 }],
      [{ s: 1 }, { s: 1 }, { s: 1 }],
    ];
    const plain = await open(rows, { borders: undefined, xfBorderIds: undefined });
    expect([plain.rowCount, plain.colCount]).toEqual([1, 1]);
    expect(plain.borders).toEqual([]);
    const got = plain.getRows(1, 1)[0]!;
    expect(got.borders).toBeUndefined();
    expect(got.cells).toEqual([{ text: 'a', kind: 's', colSpan: 1, rowSpan: 1 }]);
  });

  it('罫線のセルが範囲を決めても、CellView の作り方（値の無いセルは undefined）は変えない', async () => {
    const sheet = await open([[text('a'), { s: 1 }, { s: 1 }]]);
    expect(sheet.getRows(1, 1)[0]!.cells).toEqual([
      { text: 'a', kind: 's', colSpan: 1, rowSpan: 1 },
      undefined,
      undefined,
    ]);
  });

  it('結合の範囲と罫線のセルの大きい方に合わせる', async () => {
    const sheet = await open([[text('a')], [{ s: 1 }]], {}, { merges: ['A1:C1'] });
    expect([sheet.rowCount, sheet.colCount]).toEqual([2, 3]);
  });
});

describe('ブックでの共有', () => {
  it('styles.xml が無いブックは罫線の表が空・borders の無い行', async () => {
    const bytes = await buildXlsx({
      sheets: [{ name: 'S', rows: [[{ v: 'a', t: 'inlineStr', s: 1 }]] }],
      noStyles: true,
    });
    const sheet = await (await openXlsx(bytes)).openSheet(1);
    expect(sheet.borders).toEqual([]);
    expect(sheet.getRows(1, 1)[0]!.borders).toBeUndefined();
  });

  it('全シートが同じ罫線の表（同じ配列）を共有する', async () => {
    const bytes = await buildXlsx({
      sheets: [
        { name: 'A', rows: [[{ s: 1 }]] },
        { name: 'B', rows: [[{ s: 2 }]] },
      ],
      xfs: [0, 0, 0],
      borders: BORDERS,
      xfBorderIds: [0, 1, 2],
    });
    const book = await openXlsx(bytes);
    const a = await book.openSheet(1);
    const b = await book.openSheet(2);
    expect(a.borders).toBe(b.borders);
    expect(a.borders).toHaveLength(2);
  });

  it('グラフシート（セルを持たない）は空の表', async () => {
    const bytes = await buildXlsx({ sheets: [{ name: 'C', chartsheet: true }] });
    const sheet = await (await openXlsx(bytes)).openSheet(1);
    expect(sheet.borders).toEqual([]);
  });
});

describe('罫線の色とテーマ', () => {
  it('theme の色の罫線は、ブックのテーマで解く（テーマが無ければ既定の 12 色）', async () => {
    const themed = [borderXml({ bottom: { style: 'thin', color: 'theme="4"' } })];
    const spec = {
      sheets: [{ name: 'S', rows: [[{ s: 1 }]] }],
      xfs: [0, 0],
      borders: [borderXml({}), ...themed],
      xfBorderIds: [0, 1],
    };
    const withTheme = await buildXlsx({
      ...spec,
      theme: { xml: buildThemeXml({ accent1: '<a:srgbClr val="102030"/>' }) },
    });
    const a = await (await openXlsx(withTheme)).openSheet(1);
    expect(a.borders[0]!.bottom?.color?.rgb).toBe(0x102030);
    const b = await (await openXlsx(await buildXlsx(spec))).openSheet(1);
    expect(b.borders[0]!.bottom?.color?.rgb).toBe(0x4472c4);
  });
});

describe('線形時間', () => {
  /** rows 行 × cols 列の、すべてのセルが書式 s を持つ値の無いセルと、1 列目の値を持つ sheetData */
  function grid(rows: number, cols: number, s: number): string {
    const parts: string[] = [];
    for (let r = 1; r <= rows; r++) {
      const cells: string[] = [`<c r="A${r}" s="${s}" t="inlineStr"><is><t>v${r}</t></is></c>`];
      for (let c = 2; c <= cols; c++) cells.push(`<c r="${colLetters(c)}${r}" s="${s}"/>`);
      parts.push(`<row r="${r}">${cells.join('')}</row>`);
    }
    return parts.join('');
  }

  /** シートを開いて全行を展開するまでの時間（ms）。3 回測って最短 */
  async function timeOf(sheetDataXml: string, withBorders: boolean): Promise<number> {
    const bytes = await buildXlsx({
      sheets: [{ name: 'S', sheetDataXml }],
      xfs: [0, 0, 0],
      ...(withBorders ? { borders: BORDERS, xfBorderIds: [0, 1, 2] } : {}),
    });
    const book = await openXlsx(bytes);
    let best = Infinity;
    for (let i = 0; i < 3; i++) {
      const t0 = performance.now();
      const sheet = await book.openSheet(1);
      const rows = sheet.getRows(1, sheet.rowCount);
      best = Math.min(best, performance.now() - t0);
      expect(rows).toHaveLength(sheet.rowCount);
    }
    return best;
  }

  it('罫線を持つセル 10 万個のシートの読み取りは、罫線の無い（書式に罫線が無い）同じシートの数倍に収まる（線形）', async () => {
    const plain = await timeOf(grid(1000, 100, 2), false);
    const bordered = await timeOf(grid(1000, 100, 2), true);
    // 目安は 2 倍以内。CI の揺れを見込んで余裕を持たせる
    expect(bordered).toBeLessThan(Math.max(plain * 3, 300));
  }, 60_000);
});

/** n 行目だけを展開して返す */
function getRow(sheet: XlsxSheet, n: number): SheetRow {
  return sheet.getRows(n, n)[0]!;
}
