import { describe, expect, it, vi } from 'vitest';
import { openXlsx } from '../../src/core/index';
import {
  type CellSpec,
  type ColSpec,
  type SheetSpec,
  type XlsxSpec,
  buildThemeXml,
  buildXlsx,
  colLetters,
  fontXml,
  patternFillXml,
} from '../../src/testing/make-xlsx';
import type { CellFormat, SheetRow, XlsxSheet } from '../../src/core/types';

// loadTheme を呼んだ回数を数えるため、実体を包む（結果は変えない）
vi.mock('../../src/core/theme', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../src/core/theme')>();
  return { ...actual, loadTheme: vi.fn(actual.loadTheme) };
});
import { loadTheme } from '../../src/core/theme';

const RED = 0xff0000;
const BLUE = 0x0000ff;

/**
 * 書式のブック：xf 0＝書式なし、xf 1＝赤の塗り、xf 2＝太字、xf 3＝青の塗り＋太字、xf 4＝斜体だけ（塗りなし）。
 * fonts[0] は既定のフォント（sz 11）、fonts[1] は太字、fonts[2] は斜体
 */
const BASE: Partial<XlsxSpec> = {
  fonts: [fontXml('<sz val="11"/>'), fontXml('<b/>'), fontXml('<i/>')],
  fills: [
    patternFillXml('none'),
    patternFillXml('solid', 'rgb="FFFF0000"'),
    patternFillXml('solid', 'rgb="FF0000FF"'),
  ],
  xfs: [0, 0, 0, 0, 0],
  xfFontIds: [0, 0, 1, 1, 2],
  xfFillIds: [0, 1, 0, 2, 0],
};

const text = (v: string): CellSpec => ({ v, t: 'inlineStr' });

interface SheetOpts {
  merges?: string[];
  cols?: ColSpec[];
  rowAttrs?: SheetSpec['rowAttrs'];
  sheetDataXml?: string;
}

async function open(
  rows: CellSpec[][],
  extra: Partial<XlsxSpec> = {},
  sheetExtra: SheetOpts = {},
): Promise<XlsxSheet> {
  const bytes = await buildXlsx({
    sheets: [{ name: 'S', rows, ...sheetExtra }],
    ...BASE,
    ...extra,
  });
  return (await openXlsx(bytes)).openSheet(1);
}

/** 書式の目印：塗りが赤 'R'・青 'U'、太字は 'b'、斜体は 'i' を後ろに足す（書式なしは undefined） */
function mark(sheet: XlsxSheet, idx: number | undefined): string | undefined {
  if (idx === undefined) return undefined;
  const f: CellFormat = sheet.formats[idx]!;
  return (
    (f.fill?.rgb === RED ? 'R' : f.fill?.rgb === BLUE ? 'U' : '') +
    (f.bold ? 'b' : '') +
    (f.italic ? 'i' : '')
  );
}

const marks = (sheet: XlsxSheet, row: SheetRow): (string | undefined)[] =>
  (row.formats ?? []).map((i) => mark(sheet, i));

const getRow = (sheet: XlsxSheet, n: number): SheetRow => sheet.getRows(n, n)[0]!;

describe('getRows の formats：セルごとの書式の番号', () => {
  it('値のあるセル・値の無い書式だけのセル・書式なし（s=0 と s なし）を区別して番号を返す', async () => {
    const sheet = await open([
      [
        { v: 'a', t: 'inlineStr', s: 1 },
        { s: 3 },
        { v: 1, s: 0 },
        { v: 'x', t: 'inlineStr', s: 4 },
      ],
    ]);
    const [row] = sheet.getRows(1, 1);
    expect(row!.cells.map((c) => c?.text)).toEqual(['a', undefined, '1', 'x']);
    expect(row!.formats).toHaveLength(4);
    expect(marks(sheet, row!)).toEqual(['R', 'Ub', undefined, 'i']);
  });

  it('同じ書式は同じ番号。XlsxSheet.formats の位置を指す（書式の無い xf は番号を使わない）', async () => {
    const sheet = await open([[{ s: 1 }, { s: 3 }, { s: 1 }, { s: 0 }]]);
    expect(sheet.formats).toHaveLength(4); // ブックの表（見える書式の xf は 1〜4 の 4 種）
    const f = getRow(sheet, 1).formats!;
    expect(f[0]).toBe(f[2]);
    expect(f[0]).not.toBe(f[1]);
    expect(f[3]).toBeUndefined();
  });

  it('結合で隠れるセル（null）の書式も返す。結合の左上は CellView のまま', async () => {
    const sheet = await open(
      [
        [{ v: 'm', t: 'inlineStr', s: 1 }, { s: 3 }],
        [{ s: 1 }, { s: 3 }],
      ],
      {},
      { merges: ['A1:B2'] },
    );
    const rows = sheet.getRows(1, 2);
    expect(rows[0]!.cells[0]).toMatchObject({ text: 'm', colSpan: 2, rowSpan: 2 });
    expect(rows[0]!.cells[1]).toBeNull();
    expect(marks(sheet, rows[0]!)).toEqual(['R', 'Ub']);
    expect(marks(sheet, rows[1]!)).toEqual(['R', 'Ub']);
  });

  it('書式のあるセルが無い行は formats を持たない。返す配列の長さは cells と同じ', async () => {
    const sheet = await open([
      [text('a'), text('b'), { v: 'c', t: 'inlineStr', s: 1 }],
      [text('x'), { s: 0 }],
      [{ s: 3 }],
    ]);
    const rows = sheet.getRows(1, 3);
    expect(rows[0]!.formats).toHaveLength(rows[0]!.cells.length);
    expect(marks(sheet, rows[0]!)).toEqual([undefined, undefined, 'R']);
    expect(rows[1]!.formats).toBeUndefined();
    expect(marks(sheet, rows[2]!)).toEqual(['Ub', undefined, undefined]);
  });

  it('範囲の行だけを返し、r 属性の無いセルは直前の次の列。範囲外・不正な s は書式なし', async () => {
    const rows: CellSpec[][] = [];
    for (let r = 1; r <= 6; r++) rows.push([{ v: r, s: r % 2 }]);
    const sheet = await open(rows);
    expect(sheet.getRows(3, 4).map((r) => marks(sheet, r)[0])).toEqual(['R', undefined]);

    const s2 = await open([[text('a'), { s: 1, r: null }, { s: 3, r: null }]]);
    expect(marks(s2, getRow(s2, 1))).toEqual([undefined, 'R', 'Ub']);
    const s3 = await open([[{ v: 'a', t: 'inlineStr', attrs: ' s="99"' }, { attrs: ' s="x"' }]]);
    expect(getRow(s3, 1).formats).toBeUndefined();
  });

  it('名前空間の接頭辞付きの XML でも読む', async () => {
    const sheet = await open([[text('a'), { s: 3 }]], { prefix: 'x' });
    expect(marks(sheet, getRow(sheet, 1))).toEqual([undefined, 'Ub']);
  });
});

describe('行と列の書式（customFormat の行の s・<col style>）', () => {
  /** 4 列目に値のある行 2 を足して colCount を 4 にする */
  const WIDE: CellSpec[] = [null, null, null, text('d')];

  it('customFormat="1" の行の s は、セルの要素の無い位置だけに当たる（セルの要素は s が優先・s の無いセルの要素には当たらない）', async () => {
    const sheet = await open(
      [[text('a'), null, { s: 2 }], WIDE],
      {},
      { rowAttrs: { 1: { s: 1, customFormat: true } } },
    );
    expect(sheet.colCount).toBe(4);
    // A は s の無い値のセル（xf 0）、B・D は要素が無い、C は s=2（太字）
    expect(marks(sheet, getRow(sheet, 1))).toEqual([undefined, 'R', 'b', 'R']);
    expect(getRow(sheet, 2).formats).toBeUndefined();
  });

  it('customFormat が無い行の s は当てない', async () => {
    const sheet = await open([[text('a')], WIDE], {}, { rowAttrs: { 1: { s: 1 } } });
    expect(getRow(sheet, 1).formats).toBeUndefined();
  });

  it('列の style は、セルの要素の無い位置に当たる（min〜max を展開し、表の列数の中だけ）', async () => {
    const sheet = await open(
      [[text('a'), null, { s: 4 }], WIDE],
      {},
      {
        cols: [
          { min: 2, max: 3, style: 1 },
          { min: 4, max: 99, style: 3 },
        ],
      },
    );
    expect(sheet.colCount).toBe(4);
    // 1 行目：A は値（s 無し）、B は要素なし、C は s=4（斜体）、D は要素なし。2 行目：A〜C は要素なし、D は値（s 無し）
    expect(marks(sheet, getRow(sheet, 1))).toEqual([undefined, 'R', 'i', 'Ub']);
    expect(marks(sheet, getRow(sheet, 2))).toEqual([undefined, 'R', 'R', undefined]);
  });

  it('行の書式は列の書式より優先する（customFormat の行の s が書式なし xf 0 でも、列の書式は当てない）', async () => {
    const sheet = await open(
      [
        [null, null],
        [text('a'), null],
        [null, null],
        [null, text('z')],
      ],
      {},
      {
        cols: [{ min: 1, max: 2, style: 1 }],
        rowAttrs: { 1: { s: 2, customFormat: true }, 3: { s: 0, customFormat: true } },
      },
    );
    expect(marks(sheet, getRow(sheet, 1))).toEqual(['b', 'b']);
    expect(marks(sheet, getRow(sheet, 2))).toEqual([undefined, 'R']);
    expect(getRow(sheet, 3).formats).toBeUndefined();
    expect(marks(sheet, getRow(sheet, 4))).toEqual(['R', undefined]);
  });

  it('<row> の無い行（XML に出てこない行）には列の書式が当たる', async () => {
    const cell = (ref: string): string => `<c r="${ref}" t="inlineStr"><is><t>v</t></is></c>`;
    const sheet = await open(
      [],
      {},
      {
        cols: [{ min: 1, max: 1, style: 1 }],
        sheetDataXml: `<row r="1">${cell('A1')}</row><row r="3">${cell('A3')}</row>`,
      },
    );
    expect(sheet.rowCount).toBe(3);
    expect(marks(sheet, getRow(sheet, 2))).toEqual(['R']);
  });

  it('行・列の書式では表の範囲は広がらない', async () => {
    const sheet = await open(
      [[text('a')], [null]],
      {},
      { cols: [{ min: 1, max: 5, style: 1 }], rowAttrs: { 2: { s: 3, customFormat: true } } },
    );
    expect([sheet.rowCount, sheet.colCount]).toEqual([1, 1]);
    expect(sheet.getRows(1, 5)).toHaveLength(1);
  });

  it('列の style が範囲外・不正・書式の無い xf のときは当てない', async () => {
    const sheet = await open(
      [
        [text('a'), null, null],
        [null, null, text('z')],
      ],
      {},
      {
        cols: [
          { min: 2, max: 2, style: 99 },
          { min: 3, max: 3, style: 0 },
        ],
      },
    );
    expect(getRow(sheet, 1).formats).toBeUndefined();
    expect(getRow(sheet, 2).formats).toBeUndefined();
  });

  it('重なった列の指定は、文書で先のものを採る', async () => {
    const sheet = await open(
      [
        [text('a'), null],
        [null, text('b')],
      ],
      {},
      {
        cols: [
          { min: 2, max: 2, style: 1 },
          { min: 1, max: 2, style: 3 },
        ],
      },
    );
    expect(marks(sheet, getRow(sheet, 1))).toEqual([undefined, 'R']);
    expect(marks(sheet, getRow(sheet, 2))).toEqual(['Ub', undefined]);
  });
});

describe('余白の位置の行・列の書式（rowFormat・colFormat。表の範囲の外でも引ける）', () => {
  it('colFormat：<col style> の書式の番号を、表の列数（colCount）より右の列でも返す。指定の無い列・範囲外・書式なしは -1', async () => {
    const sheet = await open(
      [[text('a'), null, null]],
      {},
      {
        cols: [
          { min: 2, max: 3, style: 1 },
          { min: 4, max: 99, style: 3 },
        ],
      },
    );
    expect(sheet.colCount).toBe(1);
    expect(sheet.colFormat(1)).toBe(-1); // 指定なし
    expect(mark(sheet, sheet.colFormat(2))).toBe('R');
    expect(mark(sheet, sheet.colFormat(3))).toBe('R');
    expect(mark(sheet, sheet.colFormat(4))).toBe('Ub');
    expect(mark(sheet, sheet.colFormat(99))).toBe('Ub');
    expect(sheet.colFormat(100)).toBe(-1);
    expect(sheet.colFormat(0)).toBe(-1);
    expect(sheet.colFormat(Number.NaN)).toBe(-1);
  });

  it('colFormat：重なった指定は文書で先のものを採る（getRows と同じ）。書式なしの xf 0 も先に取れば -1 のまま', async () => {
    const sheet = await open(
      [[text('a')]],
      {},
      {
        cols: [
          { min: 2, max: 2, style: 1 },
          { min: 1, max: 2, style: 3 },
          { min: 5, max: 5, style: 0 },
          { min: 5, max: 5, style: 1 },
        ],
      },
    );
    expect(mark(sheet, sheet.colFormat(2))).toBe('R');
    expect(mark(sheet, sheet.colFormat(1))).toBe('Ub');
    expect(sheet.colFormat(5)).toBe(-1);
  });

  it('rowFormat：customFormat="1" の行の s の書式の番号。表の行数（rowCount）より下の <row> も引ける。書式なしの xf は -1', async () => {
    const cell = (ref: string): string => `<c r="${ref}" t="inlineStr"><is><t>v</t></is></c>`;
    const sheet = await open(
      [],
      {},
      {
        sheetDataXml:
          `<row r="1">${cell('A1')}</row>` +
          '<row r="3" s="1" customFormat="1"/>' +
          '<row r="4" s="1"/>' +
          '<row r="5" s="0" customFormat="1"/>' +
          '<row r="9" s="3" customFormat="1"/>',
      },
    );
    expect(sheet.rowCount).toBe(1);
    expect(sheet.rowFormat(1)).toBeUndefined(); // customFormat が無い
    expect(mark(sheet, sheet.rowFormat(3))).toBe('R');
    expect(sheet.rowFormat(4)).toBeUndefined(); // customFormat が無い行の s は当てない
    expect(sheet.rowFormat(5)).toBe(-1); // 書式なし（列の書式も打ち消す目印）
    expect(mark(sheet, sheet.rowFormat(9))).toBe('Ub');
    expect(sheet.rowFormat(2)).toBeUndefined(); // <row> の無い行
    expect(sheet.rowFormat(1000)).toBeUndefined();
    expect(sheet.rowFormat(0)).toBeUndefined();
  });

  it('見える書式の無いブックは、行・列の書式があっても colFormat は -1・rowFormat は undefined', async () => {
    const sheet = await open(
      [[text('a')]],
      {
        fonts: [fontXml('<sz val="11"/>')],
        fills: [patternFillXml('none')],
        xfs: [0, 0],
        xfFontIds: [0, 0],
        xfFillIds: [0, 0],
      },
      { cols: [{ min: 1, max: 5, style: 1 }], rowAttrs: { 1: { s: 1, customFormat: true } } },
    );
    expect(sheet.formats).toHaveLength(0);
    expect(sheet.colFormat(2)).toBe(-1);
    expect(sheet.rowFormat(1)).toBeUndefined();
  });

  it('release 後も例外にしない（rowFormat は行の XML を手放すので undefined、colFormat は layout だけで引ける）', async () => {
    const sheet = await open(
      [[text('a')]],
      {},
      { cols: [{ min: 1, max: 5, style: 1 }], rowAttrs: { 1: { s: 1, customFormat: true } } },
    );
    (sheet as unknown as { release(): void }).release();
    expect(sheet.rowFormat(1)).toBeUndefined();
    expect(() => sheet.colFormat(2)).not.toThrow();
  });

  it('大量に重なった列の style でも colFormat は 1 回ごとに高速（8000 指定 × 100 列を 1 秒未満）', async () => {
    const cols: ColSpec[] = [];
    for (let i = 0; i < 8000; i++) cols.push({ min: 1, max: 16384, style: 1 });
    const sheet = await open([[text('a')]], {}, { cols });
    const t0 = performance.now();
    for (let c = 1; c <= 100; c++) sheet.colFormat(c);
    expect(performance.now() - t0).toBeLessThan(1000);
  });
});

describe('表の範囲：塗りだけがある末尾の行・列まで広げる', () => {
  it('末尾に塗りだけの行があれば rowCount がその行まで広がり、装飾だけ（塗りなし）の行では広がらない', async () => {
    const sheet = await open([[text('a')], [text('b')], [{ s: 1 }], [{ s: 2 }, { s: 4 }]]);
    expect(sheet.rowCount).toBe(3);
    expect(sheet.colCount).toBe(1);
    expect(marks(sheet, getRow(sheet, 3))).toEqual(['R']);
  });

  it('右に塗りだけの列があれば colCount がその列まで広がる（どの行のものでも）', async () => {
    const sheet = await open([[text('a')], [null, null, null, { s: 3 }], [text('c')]]);
    expect([sheet.rowCount, sheet.colCount]).toEqual([3, 4]);
    expect(marks(sheet, getRow(sheet, 2))).toEqual([undefined, undefined, undefined, 'Ub']);
    const bold = await open([[text('a')], [null, null, null, { s: 2 }]]);
    expect([bold.rowCount, bold.colCount]).toEqual([1, 1]);
  });

  it('CellView の作り方（値の無いセルは undefined）は変えない。結合の範囲と大きい方に合わせる', async () => {
    const sheet = await open([[text('a'), { s: 1 }, { s: 1 }]]);
    expect(sheet.getRows(1, 1)[0]!.cells).toEqual([
      { text: 'a', kind: 's', colSpan: 1, rowSpan: 1 },
      undefined,
      undefined,
    ]);
    const m = await open([[text('a')], [{ s: 1 }]], {}, { merges: ['A1:C1'] });
    expect([m.rowCount, m.colCount]).toEqual([2, 3]);
  });

  it('塗りは行にも列にも広げる（別々のセルが行と列のそれぞれの末尾を決める）', async () => {
    const sheet = await open([[text('a')], [{ s: 1 }], [null, { s: 1 }], [{ s: 2 }]], {}, {});
    expect([sheet.rowCount, sheet.colCount]).toEqual([3, 2]);
  });

  /** xf 0＝書式なし、xf 1＝赤の塗り、xf 2＝不透明な白の塗り（紙）、xf 3＝白に近い塗り（#FEFEFE） */
  const PAPER: Partial<XlsxSpec> = {
    fills: [
      patternFillXml('none'),
      patternFillXml('solid', 'rgb="FFFF0000"'),
      patternFillXml('solid', 'rgb="FFFFFFFF"'),
      patternFillXml('solid', 'rgb="FFFEFEFE"'),
    ],
    xfs: [0, 0, 0, 0],
    xfFontIds: [0, 0, 0, 0],
    xfFillIds: [0, 1, 2, 3],
  };

  it('不透明な白の塗り（紙）だけの末尾の行・列では広がらない（グリッド線消しの白塗りの帯で範囲を膨らませない）', async () => {
    const sheet = await open(
      [[text('a')], [{ s: 2 }, { s: 2 }, { s: 2 }], [null, null, null, { s: 2 }]],
      PAPER,
    );
    expect([sheet.rowCount, sheet.colCount]).toEqual([1, 1]);
    // 白の塗りの書式は書式として残る（描画側が紙として描く）
    expect(sheet.formats.some((f) => f.fill?.rgb === 0xffffff)).toBe(true);
  });

  it('白以外の塗りは広げる。白の塗りが混ざっても、広がりは白以外のセルだけで決まる', async () => {
    const sheet = await open(
      [
        [text('a')],
        [{ s: 2 }, { s: 2 }, { s: 2 }, { s: 2 }],
        [null, { s: 1 }],
        [null, null, null, null, { s: 2 }],
      ],
      PAPER,
    );
    expect([sheet.rowCount, sheet.colCount]).toEqual([3, 2]);
    // #FEFEFE は紙ではない（広げる）
    const half = await open([[text('a')], [null, { s: 3 }]], PAPER);
    expect([half.rowCount, half.colCount]).toEqual([2, 2]);
  });
});

describe('書式の無いブック', () => {
  const rows: CellSpec[][] = [
    [text('a'), { s: 1 }],
    [{ s: 1 }, { s: 1 }, { s: 1 }],
  ];

  it('書式の指定（fonts・fills）が無いブックの行数・列数・cells・borders は今と同じ', async () => {
    const plain = await open(rows, {
      fonts: undefined,
      fills: undefined,
      xfFontIds: undefined,
      xfFillIds: undefined,
    });
    expect([plain.rowCount, plain.colCount]).toEqual([1, 1]);
    expect(plain.formats).toEqual([]);
    expect(plain.borders).toEqual([]);
    const got = plain.getRows(1, 1)[0]!;
    expect(got.formats).toBeUndefined();
    expect(got.borders).toBeUndefined();
    expect(got.cells).toEqual([{ text: 'a', kind: 's', colSpan: 1, rowSpan: 1 }]);
  });

  it('書式の無いブックでは行の customFormat・列の style があっても formats を持たない', async () => {
    const plain = await open(
      [[text('a'), null]],
      { fonts: undefined, fills: undefined, xfFontIds: undefined, xfFillIds: undefined },
      { cols: [{ min: 1, max: 2, style: 1 }], rowAttrs: { 1: { s: 1, customFormat: true } } },
    );
    expect(getRow(plain, 1).formats).toBeUndefined();
  });

  it('styles.xml が無いブック・グラフシートは空の表', async () => {
    const bytes = await buildXlsx({
      sheets: [
        { name: 'S', rows: [[{ v: 'a', t: 'inlineStr', s: 1 }]] },
        { name: 'C', chartsheet: true },
      ],
      noStyles: true,
    });
    const book = await openXlsx(bytes);
    const sheet = await book.openSheet(1);
    expect(sheet.formats).toEqual([]);
    expect(getRow(sheet, 1).formats).toBeUndefined();
    expect((await book.openSheet(2)).formats).toEqual([]);
  });

  it('全シートが同じ書式の表（同じ配列）を共有する', async () => {
    const bytes = await buildXlsx({
      sheets: [
        { name: 'A', rows: [[{ s: 1 }]] },
        { name: 'B', rows: [[{ s: 3 }]] },
      ],
      ...BASE,
    });
    const book = await openXlsx(bytes);
    const a = await book.openSheet(1);
    const b = await book.openSheet(2);
    expect(a.formats).toBe(b.formats);
    expect(a.formats).toHaveLength(4);
  });
});

describe('書式の色とテーマ（TD-94-8 を守る）', () => {
  const calls = (): number => vi.mocked(loadTheme).mock.calls.length;
  const spec = (over: Partial<XlsxSpec>): XlsxSpec => ({
    sheets: [{ name: 'S', rows: [[{ s: 1 }]] }],
    xfs: [0, 0],
    xfFontIds: [0, 1],
    xfFillIds: [0, 1],
    theme: { xml: buildThemeXml({ accent1: '<a:srgbClr val="102030"/>' }) },
    ...over,
  });

  it('theme の塗り・文字の色は、ブックのテーマで解く', async () => {
    const bytes = await buildXlsx(
      spec({
        fonts: [fontXml(), fontXml('<color theme="4"/>')],
        fills: [patternFillXml('none'), patternFillXml('solid', 'theme="4"')],
      }),
    );
    const sheet = await (await openXlsx(bytes)).openSheet(1);
    expect(sheet.formats[0]).toMatchObject({ fill: { rgb: 0x102030 }, color: { rgb: 0x102030 } });
  });

  it('<fonts>・<fills>・<borders> のどれにも theme= が無ければ、テーマを読まない', async () => {
    const bytes = await buildXlsx(
      spec({
        fonts: [fontXml(), fontXml('<b/><color rgb="FFFF0000"/>')],
        fills: [patternFillXml('none'), patternFillXml('solid', 'rgb="FF00FF00"')],
      }),
    );
    vi.mocked(loadTheme).mockClear();
    const sheet = await (await openXlsx(bytes)).openSheet(1);
    expect(sheet.formats).toHaveLength(1);
    expect(calls()).toBe(0);
  });

  it('theme= があるときは、どのシートを開いてもブックで 1 回だけ読む', async () => {
    const fillOnly = await buildXlsx(
      spec({
        sheets: [
          { name: 'A', rows: [[{ s: 1 }]] },
          { name: 'B', rows: [[{ s: 1 }]] },
        ],
        fonts: [fontXml(), fontXml('<b/>')],
        fills: [patternFillXml('none'), patternFillXml('solid', 'theme="4"')],
      }),
    );
    vi.mocked(loadTheme).mockClear();
    const book = await openXlsx(fillOnly);
    await book.openSheet(1);
    await book.openSheet(2);
    expect(calls()).toBe(1);
    book.close();

    const fontOnly = await buildXlsx(
      spec({
        fonts: [fontXml(), fontXml('<color theme="4"/>')],
        fills: [patternFillXml('none')],
        xfFillIds: [0, 0],
      }),
    );
    vi.mocked(loadTheme).mockClear();
    await (await openXlsx(fontOnly)).openSheet(1);
    expect(calls()).toBe(1);
  });
});

describe('線形時間', () => {
  it('重なった列の style を大量に並べても、列数にほぼ比例の時間で読む', async () => {
    const cols: ColSpec[] = [];
    for (let i = 0; i < 8000; i++) cols.push({ min: 1, max: 16384, style: 1 });
    const sheet = await open(
      [[null, text('b')], [{ v: 'z', t: 'inlineStr', r: 'XFD2' }]],
      {},
      { cols },
    );
    const t0 = performance.now();
    const row = sheet.getRows(1, 1)[0]!;
    expect(performance.now() - t0).toBeLessThan(2000);
    expect(sheet.colCount).toBe(16384);
    expect(row.formats!.filter((f) => f !== undefined)).toHaveLength(16383);
  }, 30_000);

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

  /** 罫線の表：xf 1＝下の細線 */
  const BORDERS = [
    '<border><left/><right/><top/><bottom/><diagonal/></border>',
    '<border><bottom style="thin"><color auto="1"/></bottom></border>',
  ];

  /** シートを開いて全行を展開するまでの時間（ms）。3 回測って最短 */
  async function timeOf(sheetDataXml: string, extra: Partial<XlsxSpec>): Promise<number> {
    const bytes = await buildXlsx({ sheets: [{ name: 'S', sheetDataXml }], ...extra });
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

  it('書式を持つセル 10 万個のシートの読み取りは、書式の無い同じシートの数倍に収まる（線形）', async () => {
    const plain = await timeOf(grid(1000, 100, 1), { xfs: [0, 0] });
    const formatted = await timeOf(grid(1000, 100, 1), {
      xfs: [0, 0],
      fills: BASE.fills,
      xfFillIds: [0, 1],
    });
    // 目安は 2 倍以内。CI の揺れを見込んで余裕を持たせる
    expect(formatted).toBeLessThan(Math.max(plain * 3, 300));
  }, 60_000);

  it('塗りが 1 つでもあるブックで、罫線も塗りも持たないセルばかりの行は、塗りの無いブックと同じ程度の時間で読む（範囲の走査を省く）', async () => {
    // 書式の無い値のセルが並ぶ 1000 行 × 100 列。塗りの xf はあるが、どの行も参照しない
    const parts: string[] = [];
    for (let r = 1; r <= 1000; r++) {
      const cells: string[] = [];
      for (let c = 1; c <= 100; c++)
        cells.push(`<c r="${colLetters(c)}${r}" t="inlineStr"><is><t>v</t></is></c>`);
      parts.push(`<row r="${r}">${cells.join('')}</row>`);
    }
    const data = parts.join('');
    const without = await timeOf(data, { xfs: [0, 0] });
    const withFill = await timeOf(data, { xfs: [0, 0], fills: BASE.fills, xfFillIds: [0, 1] });
    expect(withFill).toBeLessThan(Math.max(without * 2, 300));
  }, 60_000);

  it('罫線と書式の両方を持つシートは、罫線だけのシートの数倍に収まる（走査は 1 回）', async () => {
    const bordersOnly = await timeOf(grid(1000, 100, 1), {
      xfs: [0, 0],
      borders: BORDERS,
      xfBorderIds: [0, 1],
    });
    const both = await timeOf(grid(1000, 100, 1), {
      xfs: [0, 0],
      borders: BORDERS,
      xfBorderIds: [0, 1],
      fills: BASE.fills,
      xfFillIds: [0, 1],
    });
    // 目安は 1.3 倍以内。CI の揺れを見込んで余裕を持たせる
    expect(both).toBeLessThan(Math.max(bordersOnly * 2, 300));
  }, 60_000);
});
