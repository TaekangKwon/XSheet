import { describe, expect, it } from 'vitest';
import { openXlsx } from '../../src/core/index';
import { parseAttrs } from '../../src/core/xml';
import { buildXlsx } from '../../src/testing/make-xlsx';

/** Object.prototype が汚れていないことを確かめる */
function expectCleanPrototype() {
  expect(Object.keys(Object.prototype)).toEqual([]);
  expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  expect(({} as Record<string, unknown>).text).toBeUndefined();
  expect(Object.getPrototypeOf({})).toBe(Object.prototype);
}

describe('__proto__ などの名前でプロトタイプが汚染されない', () => {
  it('属性の名前が __proto__・constructor のタグを読んでも汚染されない', () => {
    const attrs = parseAttrs('<c __proto__="polluted" constructor="x" polluted="y" r="A1">');
    expect(attrs.r).toBe('A1');
    expect(Object.getPrototypeOf(attrs)).toBeNull();
    expectCleanPrototype();
  });

  it('__proto__・constructor という名前のシートを、名前のまま一覧に出して開ける', async () => {
    const book = await openXlsx(
      await buildXlsx({
        sheets: [
          { name: '__proto__', rows: [['in proto sheet']] },
          { name: 'constructor', state: 'hidden', rows: [['in ctor sheet']] },
          { name: 'hasOwnProperty', rows: [['third']] },
        ],
      }),
    );
    expect(book.sheets.map((s) => [s.index, s.name, s.hidden])).toEqual([
      [1, '__proto__', false],
      [2, 'constructor', true],
      [3, 'hasOwnProperty', false],
    ]);
    expect((await book.openSheet(1)).getRows(1, 1)[0]!.cells[0]!.text).toBe('in proto sheet');
    expect((await book.openSheet(3)).getRows(1, 1)[0]!.cells[0]!.text).toBe('third');
    expectCleanPrototype();
  });

  it('セル・行・シートの要素に __proto__ 属性があっても汚染されず、値は読める', async () => {
    const xml =
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
      '<row r="1" __proto__="polluted" polluted="1"><c r="A1" t="inlineStr" __proto__="x" polluted="y"><is><t>safe</t></is></c></row>' +
      '</sheetData></worksheet>';
    const book = await openXlsx(await buildXlsx({ sheets: [{ name: 'S', xml }] }));
    const sheet = await book.openSheet(1);
    expect(sheet.getRows(1, 1)[0]!.cells[0]!.text).toBe('safe');
    expectCleanPrototype();
  });

  it('rels の Id・共有文字列・numFmtId に __proto__ があっても汚染されない', async () => {
    const bytes = await buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: [
            [
              { t: 's', v: 0 },
              { v: 1, s: 1 },
            ],
          ],
        },
      ],
      sharedStrings: ['<t>__proto__</t>'],
      xfs: [0, 164],
      numFmts: { 164: '"__proto__"0' },
      extraFiles: {
        'xl/_rels/workbook.xml.rels':
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="__proto__" Type="x/worksheet" Target="worksheets/nowhere.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId90" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId91" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
      },
    });
    const book = await openXlsx(bytes);
    const sheet = await book.openSheet(1);
    expect(sheet.getRows(1, 1)[0]!.cells.map((c) => c?.text)).toEqual(['__proto__', '__proto__1']);
    expectCleanPrototype();
  });

  it('セルの値は文字のまま返す（HTML として解釈しない・加工しない）', async () => {
    const payload = '<img src=x onerror=alert(1)>&lt;b&gt;';
    const sheet = await (
      await openXlsx(await buildXlsx({ sheets: [{ name: 'S', rows: [[payload]] }] }))
    ).openSheet(1);
    expect(sheet.getRows(1, 1)[0]!.cells[0]!.text).toBe(payload);
  });
});
