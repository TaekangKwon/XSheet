import { describe, expect, it } from 'vitest';
import { colName, openXlsx, XlsxError } from '../../src/core/index';
import { buildXlsx, buildZip } from '../../src/testing/make-xlsx';

const open = async (spec: Parameters<typeof buildXlsx>[0]) => openXlsx(await buildXlsx(spec));

describe('シートの一覧（workbook.xml）', () => {
  it('workbook の順に 1 から番号を付ける（ファイル名の順ではない）', async () => {
    const book = await open({
      sheets: [
        { name: 'First', file: 'sheet3.xml', rows: [['one']] },
        { name: 'Second', file: 'sheet1.xml', rows: [['two']] },
        { name: 'Third', file: 'sheet2.xml', rows: [['three']] },
      ],
    });
    expect(book.sheets.map((s) => [s.index, s.name])).toEqual([
      [1, 'First'],
      [2, 'Second'],
      [3, 'Third'],
    ]);
    // 3 番目＝ sheet2.xml の中身（rels を引いている）
    expect((await book.openSheet(3)).getRows(1, 1)[0]!.cells[0]!.text).toBe('three');
    expect((await book.openSheet(1)).getRows(1, 1)[0]!.cells[0]!.text).toBe('one');
  });

  it('非表示（hidden・veryHidden）の印を付け、番号は非表示も数える', async () => {
    const book = await open({
      sheets: [
        { name: 'Shown', rows: [['a']] },
        { name: 'Hidden', state: 'hidden', rows: [['b']] },
        { name: 'Very', state: 'veryHidden', rows: [['c']] },
        { name: 'Last', rows: [['d']] },
      ],
    });
    expect(book.sheets.map((s) => s.hidden)).toEqual([false, true, true, false]);
    expect(book.sheets[3]!.index).toBe(4);
    expect((await book.openSheet(2)).getRows(1, 1)[0]!.cells[0]!.text).toBe('b');
  });

  it('シート名の実体参照を戻す', async () => {
    const book = await open({ sheets: [{ name: 'A&B <1>', rows: [] }] });
    expect(book.sheets[0]!.name).toBe('A&B <1>');
  });

  it('グラフシートも 1 枚に数え、開くと空のシート', async () => {
    const book = await open({
      sheets: [
        { name: 'Data', rows: [['x']] },
        { name: 'Chart', chartsheet: true, xml: '<chartsheet/>' },
      ],
    });
    expect(book.sheets).toHaveLength(2);
    const chart = await book.openSheet(2);
    expect([chart.rowCount, chart.colCount, chart.getRows(1, 5)]).toEqual([0, 0, []]);
  });

  it('範囲外・整数でないシート番号は RangeError', async () => {
    const book = await open({ sheets: [{ name: 'S', rows: [] }] });
    await expect(book.openSheet(0)).rejects.toBeInstanceOf(RangeError);
    await expect(book.openSheet(2)).rejects.toBeInstanceOf(RangeError);
    await expect(book.openSheet(1.5)).rejects.toBeInstanceOf(RangeError);
  });

  it('styles・sharedStrings が無くても開ける', async () => {
    const book = await open({
      sheets: [{ name: 'S', rows: [['t', 5]] }],
      noStyles: true,
      noSharedStrings: true,
    });
    expect((await book.openSheet(1)).getRows(1, 1)[0]!.cells.map((c) => c?.text)).toEqual([
      't',
      '5',
    ]);
  });

  it('close の後は openSheet が例外、開いたシートの getRows は空', async () => {
    const book = await open({ sheets: [{ name: 'S', rows: [['t']] }] });
    const sheet = await book.openSheet(1);
    book.close();
    await expect(book.openSheet(1)).rejects.toBeInstanceOf(XlsxError);
    expect(sheet.getRows(1, 1)).toEqual([]);
  });
});

describe('実験が想定していなかった形', () => {
  const rows = [['text', 42, { t: 'b' as const, v: 1 }, { t: 'e' as const, v: '#DIV/0!' }]];

  it('名前空間の接頭辞付き XML（<x:c>・<x:row>・<x:sheet>）を読む', async () => {
    const book = await open({
      prefix: 'x',
      sharedStrings: ['<x:t>shared</x:t>'],
      sheets: [
        { name: 'P1', rows: [...rows, [{ t: 's', v: 0 }]], merges: ['A2:B2'] },
        { name: 'P2', state: 'hidden', rows: [['z']] },
      ],
    });
    expect(book.sheets.map((s) => [s.name, s.hidden])).toEqual([
      ['P1', false],
      ['P2', true],
    ]);
    const sheet = await book.openSheet(1);
    const got = sheet.getRows(1, 2);
    expect(got[0]!.cells.map((c) => c?.text)).toEqual(['text', '42', 'TRUE', '#DIV/0!']);
    expect(got[1]!.cells[0]).toEqual({ text: 'shared', colSpan: 2, rowSpan: 1 });
  });

  it('接頭辞付きの styles（<x:numFmt>・<x:xf>）の表示形式も当たる', async () => {
    const book = await open({
      prefix: 'x',
      xfs: [0, 3],
      sheets: [{ name: 'S', rows: [[{ v: 1234567, s: 1 }]] }],
    });
    expect((await book.openSheet(1)).getRows(1, 1)[0]!.cells[0]!.text).toBe('1,234,567');
  });

  it('<c> の r 属性を省くと、列を数えて補う', async () => {
    const book = await open({
      omitRef: true,
      sheets: [{ name: 'S', rows: [['a', 'b', 'c'], ['d']] }],
    });
    const sheet = await book.openSheet(1);
    expect(sheet.colCount).toBe(3);
    expect(sheet.getRows(1, 2).map((r) => r.cells.map((c) => c?.text))).toEqual([
      ['a', 'b', 'c'],
      ['d', undefined, undefined],
    ]);
  });

  it('r を省いたセルは直前のセルの次の列（r のあるセルの後ろも数える）', async () => {
    const book = await open({
      sheets: [
        {
          name: 'S',
          rows: [
            [
              { v: 'first', t: 'inlineStr', r: 'C1' },
              { v: 'next', t: 'inlineStr', r: null },
              { v: 'far', t: 'inlineStr', r: 'F1' },
            ],
          ],
        },
      ],
    });
    const sheet = await book.openSheet(1);
    expect(sheet.colCount).toBe(6);
    expect(sheet.getRows(1, 1)[0]!.cells.map((c) => c?.text)).toEqual([
      undefined,
      undefined,
      'first',
      'next',
      undefined,
      'far',
    ]);
  });

  it('Strict 形式は unsupported', async () => {
    const bytes = await buildXlsx({ strict: true, sheets: [{ name: 'S', rows: [['a']] }] });
    await expect(openXlsx(bytes)).rejects.toMatchObject({ kind: 'unsupported' });
  });

  it('xlsb（workbook.bin）は unsupported、workbook が無い zip は broken', async () => {
    await expect(
      openXlsx(await buildZip({ 'xl/workbook.bin': 'x', '[Content_Types].xml': '<Types/>' })),
    ).rejects.toMatchObject({ kind: 'unsupported' });
    await expect(openXlsx(await buildZip({ 'readme.txt': 'hello' }))).rejects.toMatchObject({
      kind: 'broken',
    });
  });
});

describe('colName（列の記号）', () => {
  it('1→A、26→Z、27→AA、702→ZZ、703→AAA、16384→XFD', () => {
    expect([1, 26, 27, 52, 702, 703, 16384].map(colName)).toEqual([
      'A',
      'Z',
      'AA',
      'AZ',
      'ZZ',
      'AAA',
      'XFD',
    ]);
  });
  it('1 未満・整数でないものは空文字', () => {
    expect([0, -1, 1.5, NaN].map(colName)).toEqual(['', '', '', '']);
  });
});
