import { describe, expect, it } from 'vitest';
import { openXlsx } from '../../src/core/index';
import { MAX_COLS, MAX_ROWS } from '../../src/core/limits';
import { parseAreaRef, readPrintAreas, splitAreas } from '../../src/core/print-area';
import { buildXlsx } from '../../src/testing/make-xlsx';

const PA = '_xlnm.Print_Area';

describe('parseAreaRef（範囲 1 つ）', () => {
  it('セルの範囲・1 セル・逆向き・$ の有無', () => {
    expect(parseAreaRef('$A$1:$H$50')).toEqual({ r1: 1, c1: 1, r2: 50, c2: 8 });
    expect(parseAreaRef('B2')).toEqual({ r1: 2, c1: 2, r2: 2, c2: 2 });
    expect(parseAreaRef('$C$3:$A$1')).toEqual({ r1: 1, c1: 1, r2: 3, c2: 3 });
    expect(parseAreaRef('aa10:ab$12')).toEqual({ r1: 10, c1: 27, r2: 12, c2: 28 });
  });

  it('列だけ・行だけは、行・列の全体にする', () => {
    expect(parseAreaRef('$B:$D')).toEqual({ r1: 1, c1: 2, r2: MAX_ROWS, c2: 4 });
    expect(parseAreaRef('$3:$5')).toEqual({ r1: 3, c1: 1, r2: 5, c2: MAX_COLS });
  });

  it('読めない物・範囲外は null', () => {
    for (const s of ['#REF!', '', 'A0', 'XFE1', 'A1048577', 'A1:B', 'SUM(A1)', '0:3', 'A:XFE'])
      expect(parseAreaRef(s), s).toBeNull();
  });
});

describe('splitAreas（カンマで分ける）', () => {
  it("シート名の引用の中のカンマ・引用の中の '' では分けない", () => {
    expect(splitAreas("'a,b'!$A$1:$B$2,'a,b'!$D$1")).toEqual(["'a,b'!$A$1:$B$2", "'a,b'!$D$1"]);
    expect(splitAreas("'it''s'!A1,'it''s'!B2")).toEqual(["'it''s'!A1", "'it''s'!B2"]);
    expect(splitAreas('S!A1')).toEqual(['S!A1']);
  });
});

describe('readPrintAreas（workbook.xml）', () => {
  const wb = (names: string) =>
    `<workbook><sheets><sheet name="A"/><sheet name="B"/></sheets><definedNames>${names}</definedNames></workbook>`;

  it('localSheetId ごとに読み、ほかの名前・シートの無い名前・範囲外の id は読まない', () => {
    const m = readPrintAreas(
      wb(
        `<definedName name="_xlnm.Print_Area" localSheetId="1">B!$A$1:$C$3</definedName>` +
          `<definedName name="_xlnm.Print_Titles" localSheetId="0">A!$1:$1</definedName>` +
          `<definedName name="_xlnm.Print_Area">A!$A$1</definedName>` +
          `<definedName name="_xlnm.Print_Area" localSheetId="2">A!$A$1</definedName>` +
          `<definedName name="_xlnm.Print_Area" localSheetId="-1">A!$A$1</definedName>`,
      ),
      2,
      100,
    );
    expect([...m.entries()]).toEqual([[1, [{ r1: 1, c1: 1, r2: 3, c2: 3 }]]]);
  });

  it('実体参照を戻し、#REF! を捨て、上限を超えた範囲は読まない。同じシートの 2 つめは無視', () => {
    const m = readPrintAreas(
      wb(
        `<definedName name="_XLNM.PRINT_AREA" localSheetId="0">&apos;A &amp; B&apos;!$A$1:$B$2,#REF!,&apos;A &amp; B&apos;!$D$4,A!$E$5,A!$F$6</definedName>` +
          `<definedName name="_xlnm.Print_Area" localSheetId="0">A!$Z$1</definedName>`,
      ),
      2,
      3,
    );
    expect(m.get(0)).toEqual([
      { r1: 1, c1: 1, r2: 2, c2: 2 },
      { r1: 4, c1: 4, r2: 4, c2: 4 },
      { r1: 5, c1: 5, r2: 5, c2: 5 },
    ]);
  });

  it('読める範囲が 1 つも無ければ入れない', () => {
    const m = readPrintAreas(
      wb(`<definedName name="_xlnm.Print_Area" localSheetId="0">#REF!</definedName>`),
      2,
      100,
    );
    expect(m.size).toBe(0);
  });
});

describe('SheetInfo.printArea（openXlsx）', () => {
  it('非表示のシートも数えた位置で付き、無いシートには項目が無い', async () => {
    const book = await openXlsx(
      await buildXlsx({
        sheets: [
          { name: 'One', rows: [['a']] },
          { name: 'Hid', state: 'hidden', rows: [['b']] },
          { name: 'Sheet 3', rows: [['c']] },
        ],
        definedNames: [
          { name: PA, localSheetId: 2, ref: "'Sheet 3'!$B$2:$D$9,'Sheet 3'!$F:$G" },
          { name: PA, localSheetId: 1, ref: 'Hid!$A$1' },
        ],
      }),
    );
    expect('printArea' in book.sheets[0]!).toBe(false);
    expect(book.sheets[1]!.printArea).toEqual([{ r1: 1, c1: 1, r2: 1, c2: 1 }]);
    expect(book.sheets[2]!.printArea).toEqual([
      { r1: 2, c1: 2, r2: 9, c2: 4 },
      { r1: 1, c1: 6, r2: MAX_ROWS, c2: 7 },
    ]);
  });

  it('名前空間の接頭辞付きの workbook.xml も読む・maxPrintAreas を渡せる', async () => {
    const book = await openXlsx(
      await buildXlsx({
        prefix: 'x',
        sheets: [{ name: 'S', rows: [['a']] }],
        definedNames: [{ name: PA, localSheetId: 0, ref: 'S!$A$1,S!$B$2,S!$C$3' }],
      }),
      { limits: { maxPrintAreas: 2 } },
    );
    expect(book.sheets[0]!.printArea).toHaveLength(2);
  });
});
