import { describe, expect, it } from 'vitest';
import { openXlsx, type XlsxSheet } from '../../src/core/index';
import { buildXlsx, colLetters } from '../../src/testing/make-xlsx';

/** 行番号 r の A 列に `row<r>`、C 列に r の数値を置いた rows を n 行ぶん作る（sheetData の XML） */
function bigSheetData(n: number): string {
  const parts: string[] = [];
  for (let r = 1; r <= n; r++) {
    parts.push(
      `<row r="${r}"><c r="A${r}" t="inlineStr"><is><t>row${r}</t></is></c><c r="C${r}"><v>${r}</v></c></row>`,
    );
  }
  return parts.join('');
}

async function sheetFromData(sheetDataXml: string): Promise<XlsxSheet> {
  const book = await openXlsx(await buildXlsx({ sheets: [{ name: 'S', sheetDataXml }] }));
  return book.openSheet(1);
}

describe('行の位置の索引と、範囲の行だけの展開', () => {
  it('rowCount・colCount と、範囲の行だけを展開する', async () => {
    const sheet = await sheetFromData(bigSheetData(1000));
    expect([sheet.rowCount, sheet.colCount]).toEqual([1000, 3]);
    const rows = sheet.getRows(500, 502);
    expect(rows.map((r) => r.r)).toEqual([500, 501, 502]);
    expect(rows.map((r) => r.cells.map((c) => c?.text))).toEqual([
      ['row500', undefined, '500'],
      ['row501', undefined, '501'],
      ['row502', undefined, '502'],
    ]);
  });

  it('呼ぶたびに新しい行を作る（展開済みの行をためこまない）', async () => {
    const sheet = await sheetFromData(bigSheetData(50));
    const a = sheet.getRows(10, 12);
    const b = sheet.getRows(10, 12);
    expect(a[0]).not.toBe(b[0]);
    expect(a[0]!.cells[0]).not.toBe(b[0]!.cells[0]);
    expect(a).toEqual(b);
  });

  it('範囲を行数・列数に収める（先頭より前・末尾より後・逆・数でない）', async () => {
    const sheet = await sheetFromData(bigSheetData(10));
    expect(sheet.getRows(-5, 2).map((r) => r.r)).toEqual([1, 2]);
    expect(sheet.getRows(9, 99).map((r) => r.r)).toEqual([9, 10]);
    expect(sheet.getRows(11, 20)).toEqual([]);
    expect(sheet.getRows(5, 4)).toEqual([]);
    expect(sheet.getRows(NaN, 3)).toEqual([]);
    expect(sheet.getRows(1.9, 2.9).map((r) => r.r)).toEqual([1, 2]);
  });

  it('行の無い番号も、長さ colCount の空の行として返す', async () => {
    const sheet = await sheetFromData(
      '<row r="2"><c r="B2" t="inlineStr"><is><t>b</t></is></c></row><row r="5"><c r="C5" t="inlineStr"><is><t>c</t></is></c></row>',
    );
    expect([sheet.rowCount, sheet.colCount]).toEqual([5, 3]);
    const rows = sheet.getRows(1, 5);
    expect(rows.map((r) => r.r)).toEqual([1, 2, 3, 4, 5]);
    expect(rows.every((r) => r.cells.length === 3)).toBe(true);
    expect(rows.map((r) => r.cells.map((c) => c?.text))).toEqual([
      [undefined, undefined, undefined],
      [undefined, 'b', undefined],
      [undefined, undefined, undefined],
      [undefined, undefined, undefined],
      [undefined, undefined, 'c'],
    ]);
  });

  it('行の r 属性が無ければ直前の行の次の番号', async () => {
    const sheet = await sheetFromData(
      '<row><c t="inlineStr"><is><t>one</t></is></c></row><row><c t="inlineStr"><is><t>two</t></is></c></row><row r="6"><c t="inlineStr"><is><t>six</t></is></c></row><row><c t="inlineStr"><is><t>seven</t></is></c></row>',
    );
    expect(sheet.rowCount).toBe(7);
    expect(sheet.getRows(1, 7).map((r) => r.cells[0]?.text)).toEqual([
      'one',
      'two',
      undefined,
      undefined,
      undefined,
      'six',
      'seven',
    ]);
  });

  it('行が昇順でない文書も、番号で引ける', async () => {
    const sheet = await sheetFromData(
      '<row r="3"><c r="A3" t="inlineStr"><is><t>three</t></is></c></row><row r="1"><c r="A1" t="inlineStr"><is><t>one</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>two</t></is></c></row>',
    );
    expect(sheet.rowCount).toBe(3);
    expect(sheet.getRows(1, 3).map((r) => r.cells[0]?.text)).toEqual(['one', 'two', 'three']);
  });

  it('末尾の「書式だけで値の無い行」は行数に数えない（行 1,048,576 まで色を塗った表で何百万行も出さない）', async () => {
    const sheet = await sheetFromData(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>data</t></is></c></row><row r="2"><c r="A2" s="1"/><c r="B2" s="1"/></row><row r="1048576" spans="1:2"><c r="A1048576" s="1"/></row>',
    );
    expect([sheet.rowCount, sheet.colCount]).toEqual([1, 1]);
  });

  it('範囲外の行番号・列は読まない', async () => {
    const sheet = await sheetFromData(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>ok</t></is></c><c r="XFE1" t="inlineStr"><is><t>too-wide</t></is></c></row><row r="2000000"><c r="A2000000" t="inlineStr"><is><t>too-tall</t></is></c></row>',
    );
    expect([sheet.rowCount, sheet.colCount]).toEqual([1, 1]);
    expect(sheet.getRows(1, 1)[0]!.cells.map((c) => c?.text)).toEqual(['ok']);
  });

  it('空のシートは 0 行・0 列', async () => {
    const empty = await sheetFromData('');
    expect([empty.rowCount, empty.colCount, empty.getRows(1, 10)]).toEqual([0, 0, []]);
    const book = await openXlsx(
      await buildXlsx({
        sheets: [
          {
            name: 'E',
            xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
          },
        ],
      }),
    );
    const selfClosing = await book.openSheet(1);
    expect([selfClosing.rowCount, selfClosing.colCount]).toEqual([0, 0]);
  });

  it('列は最も右の値のあるセルまで（行ごとに違っても）', async () => {
    const letters = (c: number) => colLetters(c);
    const sheet = await sheetFromData(
      `<row r="1"><c r="${letters(2)}1" t="inlineStr"><is><t>b</t></is></c></row><row r="2"><c r="${letters(28)}2" t="inlineStr"><is><t>ab</t></is></c></row><row r="3"><c r="${letters(5)}3" t="inlineStr"><is><t>e</t></is></c></row>`,
    );
    expect(sheet.colCount).toBe(28);
    expect(sheet.getRows(2, 2)[0]!.cells[27]!.text).toBe('ab');
  });

  it('閉じていない sheetData は broken', async () => {
    const xml =
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row>';
    const book = await openXlsx(await buildXlsx({ sheets: [{ name: 'S', xml }] }));
    await expect(book.openSheet(1)).rejects.toMatchObject({ kind: 'broken' });
  });
});

describe('速さ（小さめの合成シート。実測は XLSX_PERF=1 で表示）', () => {
  it('20 万セル（10,000 行 × 20 列・共有文字列あり）を、開く＋最初の画面が 3 秒以内', async () => {
    const rows = 10_000;
    const cols = 20;
    const parts: string[] = [];
    for (let r = 1; r <= rows; r++) {
      let cs = '';
      for (let c = 1; c <= cols; c++) {
        const ref = `${colLetters(c)}${r}`;
        if (c % 4 === 1) cs += `<c r="${ref}" t="s"><v>${(r * 7 + c) % 500}</v></c>`;
        else if (c % 4 === 2) cs += `<c r="${ref}" s="1"><v>${r * 13 + c}</v></c>`;
        else if (c % 4 === 3) cs += `<c r="${ref}" s="2"><v>${((r + c) % 1000) / 1000}</v></c>`;
        else cs += `<c r="${ref}" s="3"><v>${44000 + (r % 700)}</v></c>`;
      }
      parts.push(`<row r="${r}">${cs}</row>`);
    }
    const bytes = await buildXlsx({
      sharedStrings: Array.from({ length: 500 }, (_v, i) => `<t>text ${i}</t>`),
      xfs: [0, 3, 10, 14],
      sheets: [{ name: 'Big', sheetDataXml: parts.join(''), merges: ['A1:B2'] }],
    });
    const t0 = performance.now();
    const book = await openXlsx(bytes);
    const sheet = await book.openSheet(1);
    const t1 = performance.now();
    const first = sheet.getRows(1, 60);
    const t2 = performance.now();
    if (process.env.XLSX_PERF) {
      console.info(
        `xlsx perf: cells=${rows * cols} open+openSheet=${(t1 - t0).toFixed(0)}ms getRows(60)=${(t2 - t1).toFixed(1)}ms zip=${(bytes.length / 1e6).toFixed(1)}MB`,
      );
    }
    expect([sheet.rowCount, sheet.colCount]).toEqual([rows, cols]);
    expect(first).toHaveLength(60);
    expect(first[59]!.cells[1]!.text).toBe('782'); // 書式 #,##0 の列（c=2：60 × 13 + 2）
    expect(first[59]!.cells[2]!.text).toBe('6.30%'); // 書式 0.00% の列（c=3：63 / 1000 = 0.063）
    expect(first[59]!.cells[3]!.text).toBe('2020/8/17'); // 書式 14 の列（c=4：44000 + 60 = 44060）
    expect(t1 - t0).toBeLessThan(3000);
  }, 60_000);
});
