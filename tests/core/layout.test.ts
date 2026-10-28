import { describe, expect, it } from 'vitest';
import { fontDefaultRowPx } from '../../src/core/font-metrics';
import { createGeometry } from '../../src/core/geometry';
import { openXlsx, type XlsxSheet } from '../../src/core/index';
import { buildXlsx, type SheetSpec, type XlsxSpec } from '../../src/testing/make-xlsx';
import { maxDigitWidth, parseStyles } from '../../src/core/styles';
import { emptyLayout } from '../../src/core/types';

/** シートを 1 枚だけ持つ xlsx を組んで、そのシートを開く */
async function open(sheet: Partial<SheetSpec>, book: Partial<XlsxSpec> = {}): Promise<XlsxSheet> {
  const bytes = await buildXlsx({ sheets: [{ name: 'S', ...sheet }], ...book });
  return (await openXlsx(bytes)).openSheet(1);
}

/** シート XML を直接指定して開く（不正な値の試験用） */
const openXml = (body: string): Promise<XlsxSheet> =>
  open({
    xml:
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      body +
      '</worksheet>',
  });

describe('TD-83-7 openSheet：<cols>・<sheetFormatPr>・行の ht／hidden を layout に読む', () => {
  it('欠落時は既定値（cols・rows は空、既定の列幅 null・baseColWidth 8・15pt・zeroHeight false・MDW 7）', async () => {
    const sheet = await open({ rows: [['a'], ['b']] });
    expect(sheet.layout).toEqual(emptyLayout());
    expect(sheet.layout.mdw).toBe(7);
  });

  it('<cols> の <col min max width hidden> を、範囲のまま疎に持つ', async () => {
    const sheet = await open({
      rows: [['a']],
      cols: [
        { min: 1, max: 1, width: 20.7109375 },
        { min: 3, max: 16384, width: 12 },
        { min: 2, max: 2, hidden: true },
      ],
    });
    expect(sheet.layout.cols).toEqual([
      { min: 1, max: 1, width: 20.7109375, hidden: false },
      { min: 3, max: 16384, width: 12, hidden: false },
      { min: 2, max: 2, width: null, hidden: true },
    ]);
  });

  it('不正な <col>（min・max が数でない・逆・0 以下）は捨て、max の上限超えは 16384 に収め、不正な幅は width なしにする', async () => {
    const sheet = await openXml(
      '<cols><col min="x" max="3" width="9"/><col min="5" max="2" width="9"/><col min="0" max="1" width="9"/>' +
        '<col min="2" max="99999" width="-4"/><col min="7" max="7" width="abc"/></cols>' +
        '<sheetData/>',
    );
    expect(sheet.layout.cols).toEqual([
      { min: 2, max: 16384, width: null, hidden: false },
      { min: 7, max: 7, width: null, hidden: false },
    ]);
  });

  it('<sheetFormatPr> の defaultColWidth・baseColWidth・defaultRowHeight・zeroHeight', async () => {
    const sheet = await open({
      rows: [['a']],
      formatPr: {
        defaultColWidth: 10.5,
        baseColWidth: 10,
        defaultRowHeight: 18.75,
        zeroHeight: true,
      },
    });
    expect(sheet.layout).toMatchObject({
      defaultColWidth: 10.5,
      baseColWidth: 10,
      defaultRowHeight: 18.75,
      zeroHeight: true,
    });
  });

  it('<sheetFormatPr> の一部だけの指定・不正値は、残りを既定にする', async () => {
    const some = await open({ rows: [['a']], formatPr: { defaultRowHeight: 13.5 } });
    expect(some.layout).toMatchObject({
      defaultColWidth: null,
      baseColWidth: 8,
      defaultRowHeight: 13.5,
      zeroHeight: false,
    });
    const bad = await openXml(
      '<sheetFormatPr defaultColWidth="zz" baseColWidth="-1" defaultRowHeight="NaN" zeroHeight="0"/><sheetData/>',
    );
    expect(bad.layout).toEqual(emptyLayout());
  });

  it('行の ht（pt）・hidden を読む。指定の無い行は持たない（疎）', async () => {
    const sheet = await open({
      rows: [['a'], ['b'], ['c'], ['d']],
      rowAttrs: { 2: { ht: 30 }, 4: { hidden: true }, 3: { ht: 18.75, hidden: true } },
    });
    expect(sheet.layout.rows).toEqual([
      { min: 2, max: 2, ht: 30, hidden: false, custom: true },
      { min: 3, max: 3, ht: 18.75, hidden: true, custom: true },
      { min: 4, max: 4, ht: null, hidden: true, custom: false },
    ]);
  });

  it('zeroHeight のときは <row> のあるすべての行を持つ（無い行を既定で非表示にするため）', async () => {
    const sheet = await open({
      rows: [['a'], [], ['c']],
      rowAttrs: { 3: { ht: 30 } },
      formatPr: { zeroHeight: true },
    });
    expect(sheet.layout.rows).toEqual([
      { min: 1, max: 2, ht: null, hidden: false, custom: false },
      { min: 3, max: 3, ht: 30, hidden: false, custom: true },
    ]);
  });

  it('r の省略・順不同・重複した行も、行番号の昇順・重複なしで持つ（重複は後の行を採る）', async () => {
    const sheet = await open({
      sheetDataXml:
        '<row r="5" ht="30"><c r="A5" t="inlineStr"><is><t>e</t></is></c></row>' +
        '<row ht="21"><c r="A6" t="inlineStr"><is><t>f</t></is></c></row>' +
        '<row r="2" hidden="1"><c r="A2" t="inlineStr"><is><t>b</t></is></c></row>' +
        '<row r="5" ht="45"><c r="A5" t="inlineStr"><is><t>e2</t></is></c></row>',
    });
    expect(sheet.layout.rows).toEqual([
      { min: 2, max: 2, ht: null, hidden: true, custom: false },
      { min: 5, max: 5, ht: 45, hidden: false, custom: false },
      { min: 6, max: 6, ht: 21, hidden: false, custom: false },
    ]);
  });

  it('名前空間の接頭辞付き XML（<x:cols>・<x:sheetFormatPr>・<x:row>）も読む', async () => {
    const sheet = await open(
      {
        rows: [['a'], ['b']],
        cols: [{ min: 2, max: 2, width: 15 }],
        formatPr: { defaultRowHeight: 18 },
        rowAttrs: { 2: { ht: 33 } },
      },
      { prefix: 'x' },
    );
    expect(sheet.layout.cols).toEqual([{ min: 2, max: 2, width: 15, hidden: false }]);
    expect(sheet.layout.defaultRowHeight).toBe(18);
    expect(sheet.layout.rows).toEqual([{ min: 2, max: 2, ht: 33, hidden: false, custom: true }]);
  });

  it('<sheetData/>（自己閉じ）・<sheetData> の無いシートでも、寸法は読める', async () => {
    const a = await openXml('<sheetFormatPr defaultRowHeight="18"/><sheetData/>');
    expect(a.layout.defaultRowHeight).toBe(18);
    const b = await openXml('<sheetFormatPr defaultRowHeight="21"/>');
    expect(b.layout.defaultRowHeight).toBe(21);
  });

  it('rowCount・colCount・getRows は寸法の指定に影響されない（既存の表示は不変）', async () => {
    const plain = await open({ rows: [['a', 'b'], ['c']] });
    const dim = await open({
      rows: [['a', 'b'], ['c']],
      cols: [{ min: 1, max: 5, width: 30 }],
      formatPr: { defaultRowHeight: 30 },
      rowAttrs: { 1: { ht: 40 } },
    });
    expect([dim.rowCount, dim.colCount]).toEqual([plain.rowCount, plain.colCount]);
    expect(dim.getRows(1, 2)).toEqual(plain.getRows(1, 2));
  });

  it('読んだ layout から createGeometry で px が出る（Calibri 11・MDW 7）', async () => {
    const sheet = await open(
      {
        rows: [['a'], ['b'], ['c']],
        cols: [{ min: 1, max: 1, width: 10 }],
        rowAttrs: { 2: { ht: 30 }, 3: { hidden: true } },
        formatPr: { defaultRowHeight: 15 },
      },
      { font: { name: 'Calibri', size: 11 } },
    );
    const g = createGeometry(sheet.layout, sheet.rowCount, sheet.colCount);
    expect([g.colWidth(1), g.rowHeight(1), g.rowHeight(2), g.rowHeight(3)]).toEqual([
      70, 20, 40, 0,
    ]);
    expect(g.totalHeight).toBe(60);
  });
});

describe('TD-83-8 styles：先頭の font → 最大桁幅 MDW', () => {
  it('Calibri 11＝7px、游ゴシック 11＝8px（Excel の既定列幅 64px／72px に対応）', () => {
    expect(maxDigitWidth('Calibri', 11)).toBe(7);
    expect(maxDigitWidth('游ゴシック', 11)).toBe(8);
    expect(maxDigitWidth('Yu Gothic', 11)).toBe(8);
    expect(maxDigitWidth('맑은 고딕', 11)).toBe(8);
    expect(maxDigitWidth('Malgun Gothic', 11)).toBe(8);
  });

  it('実測表のサイズはそのまま、間のサイズは線形補間（四捨五入）、範囲外は端に比例させる', () => {
    expect(maxDigitWidth('Calibri', 12)).toBe(8);
    expect(maxDigitWidth('Arial', 10)).toBe(8);
    expect(maxDigitWidth('Arial', 11)).toBe(8);
    expect(maxDigitWidth('Calibri', 13)).toBe(9); // 12:8 と 14:10 の中間
    expect(maxDigitWidth('Calibri', 22)).toBe(15); // 20:14 と 24:16 の中間
    expect(maxDigitWidth('Arial', 22)).toBe(17); // 20:15 と 24:18 の中間 16.5 → 17
    expect(maxDigitWidth('Calibri', 48)).toBe(32); // 範囲外：24:16 に比例
    expect(maxDigitWidth('Calibri', 3)).toBe(2); // 範囲外：6:4 に比例
    expect(maxDigitWidth('Meiryo', 11)).toBe(9);
    expect(maxDigitWidth('Dotum', 12)).toBe(8);
  });

  it('表に無いフォントは Calibri の行を使う', () => {
    expect(maxDigitWidth('Unknown Font', 11)).toBe(7);
    expect(maxDigitWidth('Unknown Font', 14)).toBe(10);
    expect(maxDigitWidth('', 12)).toBe(8);
  });

  it('現地語名・全角表記・空白の揺れを同じフォントとして扱う', () => {
    expect(maxDigitWidth('ＭＳ Ｐゴシック', 11)).toBe(8);
    expect(maxDigitWidth('MS PGothic', 11)).toBe(8);
    expect(maxDigitWidth('ＭＳ ゴシック', 11)).toBe(8);
    expect(maxDigitWidth('メイリオ', 11)).toBe(9);
    expect(maxDigitWidth('굴림', 20)).toBe(16);
    expect(maxDigitWidth('돋움', 10)).toBe(9);
    expect(maxDigitWidth('ＡＲＩＡＬ', 10)).toBe(8);
    expect(maxDigitWidth('Times New Roman', 14)).toBe(9);
    expect(maxDigitWidth('Segoe UI', 12)).toBe(9);
  });

  it('フォントの既定の行高（実測表）：Calibri 11＝20、游ゴシック 11＝25、맑은 고딕 11＝22、MS Pゴシック 11＝18、Arial 10＝17、他は null', () => {
    expect(fontDefaultRowPx('Calibri', 11)).toBe(20);
    expect(fontDefaultRowPx('游ゴシック', 11)).toBe(25);
    expect(fontDefaultRowPx('Yu Gothic', 11)).toBe(25);
    expect(fontDefaultRowPx('맑은 고딕', 11)).toBe(22);
    expect(fontDefaultRowPx('Malgun Gothic', 11)).toBe(22);
    expect(fontDefaultRowPx('ＭＳ Ｐゴシック', 11)).toBe(18);
    expect(fontDefaultRowPx('Arial', 10)).toBe(17);
    expect(fontDefaultRowPx('Calibri', 12)).toBeNull();
    expect(fontDefaultRowPx('Unknown', 11)).toBeNull();
  });

  it('サイズが不正なら 7px', () => {
    expect(maxDigitWidth('Calibri', NaN)).toBe(7);
    expect(maxDigitWidth('Calibri', 0)).toBe(7);
    expect(maxDigitWidth('Calibri', -3)).toBe(7);
  });

  it('parseStyles：styles.xml の先頭の font から mdw を出す。無い・読めないときは 7', () => {
    const mdw = (xml: string) => parseStyles(xml).mdw;
    expect(
      mdw(
        '<styleSheet><fonts count="2"><font><sz val="11"/><name val="游ゴシック"/></font><font><sz val="20"/><name val="Arial"/></font></fonts></styleSheet>',
      ),
    ).toBe(8);
    expect(
      mdw(
        '<styleSheet><fonts><font><name val="Calibri"/><sz val="11"/></font></fonts></styleSheet>',
      ),
    ).toBe(7);
    expect(
      mdw(
        '<x:styleSheet><x:fonts><x:font><x:sz val="12"/><x:name val="Calibri"/></x:font></x:fonts></x:styleSheet>',
      ),
    ).toBe(8);
    expect(mdw('<styleSheet><cellXfs count="1"><xf/></cellXfs></styleSheet>')).toBe(7);
    expect(mdw('<styleSheet><fonts><font/></fonts></styleSheet>')).toBe(7);
    expect(parseStyles(null).mdw).toBe(7);
  });

  it('合成 xlsx の font が layout.mdw になり、既定の列幅が 64px／72px になる', async () => {
    const calibri = await open({ rows: [['a']] }, { font: { name: 'Calibri', size: 11 } });
    const yu = await open({ rows: [['a']] }, { font: { name: '游ゴシック', size: 11 } });
    const none = await open({ rows: [['a']] });
    expect([calibri.layout.mdw, yu.layout.mdw, none.layout.mdw]).toEqual([7, 8, 7]);
    const px = (s: XlsxSheet) => createGeometry(s.layout, s.rowCount, s.colCount).colWidth(1);
    expect([px(calibri), px(yu), px(none)]).toEqual([64, 72, 64]);
  });
});

describe('TD-83-15 読み取りの丸め・ランレングス・重複・複数の <cols>', () => {
  it('極端な値は Excel の上限に丸める（列幅 255・行高 409.5・フォントサイズ 409）。座標は有限', async () => {
    const sheet = await openXml(
      '<sheetFormatPr defaultColWidth="1e307" baseColWidth="1e307" defaultRowHeight="1e307"/>' +
        '<cols><col min="1" max="2" width="1e307"/></cols>' +
        '<sheetData><row r="1" ht="1e307" customHeight="1"/></sheetData>',
    );
    expect(sheet.layout).toMatchObject({
      defaultColWidth: 255,
      baseColWidth: 255,
      defaultRowHeight: 409.5,
    });
    expect(sheet.layout.cols).toEqual([{ min: 1, max: 2, width: 255, hidden: false }]);
    expect(sheet.layout.rows).toEqual([{ min: 1, max: 1, ht: 409.5, hidden: false, custom: true }]);
    const g = createGeometry(sheet.layout, 1_048_576, 16_384);
    expect(Number.isFinite(g.totalWidth)).toBe(true);
    expect(Number.isFinite(g.totalHeight)).toBe(true);

    const mdw = parseStyles(
      '<styleSheet><fonts><font><sz val="1e307"/><name val="Calibri"/></font></fonts></styleSheet>',
    ).mdw;
    expect(Number.isFinite(mdw)).toBe(true);
    expect(mdw).toBeLessThan(300);
  });

  it('1,048,576 行すべてに同じ ht があっても、layout.rows は 1 件で、createGeometry は速い', async () => {
    const parts: string[] = [];
    for (let r = 1; r <= 1_048_576; r++) parts.push(`<row r="${r}" ht="30" customHeight="1"/>`);
    const sheet = await open({ sheetDataXml: parts.join('') });
    expect(sheet.layout.rows).toEqual([
      { min: 1, max: 1_048_576, ht: 30, hidden: false, custom: true },
    ]);
    const t0 = performance.now();
    const g = createGeometry(sheet.layout, 1_048_576, 1);
    expect(g.totalHeight).toBe(1_048_576 * 40);
    expect(performance.now() - t0).toBeLessThan(500);
  }, 60_000);

  it('連続する同じ指定の行は 1 件にまとめ、途切れたら分ける', async () => {
    const sheet = await open({
      rows: [['a'], ['b'], ['c'], ['d'], ['e'], ['f']],
      rowAttrs: {
        1: { ht: 30 },
        2: { ht: 30 },
        3: { ht: 30 },
        4: { hidden: true },
        5: { hidden: true },
        6: { ht: 30 },
      },
    });
    expect(sheet.layout.rows).toEqual([
      { min: 1, max: 3, ht: 30, hidden: false, custom: true },
      { min: 4, max: 5, ht: null, hidden: true, custom: false },
      { min: 6, max: 6, ht: 30, hidden: false, custom: true },
    ]);
  });

  it('既定と同じで非表示でもない行は持たない（ht がファイルの defaultRowHeight と同じで customHeight 無し）', async () => {
    const sheet = await open({
      rows: [['a'], ['b'], ['c']],
      formatPr: { defaultRowHeight: 14.4 },
      rowAttrs: { 1: { ht: 14.4, custom: false }, 2: { ht: 14.405, custom: false } },
    });
    expect(sheet.layout.rows).toEqual([]);
    // customHeight 付きは、既定と同じ値でも持つ（フォントの既定の行高と違いうるため）
    const custom = await open({
      rows: [['a']],
      formatPr: { defaultRowHeight: 14.4 },
      rowAttrs: { 1: { ht: 14.4 } },
    });
    expect(custom.layout.rows).toEqual([{ min: 1, max: 1, ht: 14.4, hidden: false, custom: true }]);
  });

  it('customHeight：<sheetFormatPr customHeight> と <row customHeight> を読む（ht が無い row の customHeight は無視）', async () => {
    const sheet = await openXml(
      '<sheetFormatPr defaultRowHeight="30" customHeight="1"/>' +
        '<sheetData><row r="1" ht="40" customHeight="1"/><row r="2" ht="40"/><row r="3" customHeight="1" hidden="1"/></sheetData>',
    );
    expect(sheet.layout.customHeight).toBe(true);
    expect(sheet.layout.rows).toEqual([
      { min: 1, max: 1, ht: 40, hidden: false, custom: true },
      { min: 2, max: 2, ht: 40, hidden: false, custom: false },
      { min: 3, max: 3, ht: null, hidden: true, custom: false },
    ]);
    const none = await open({ rows: [['a']], formatPr: { defaultRowHeight: 30 } });
    expect(none.layout.customHeight).toBe(false);
  });

  it('同じ行番号が複数ある行は後の行を採る（後の行に指定が無ければ、前の行の指定も取り除く）', async () => {
    const plainAfter = await open({
      sheetDataXml: '<row r="5" ht="30" customHeight="1"/><row r="5"/>',
    });
    expect(plainAfter.layout.rows).toEqual([]);
    // まとまった範囲の途中の行が取り除かれると、範囲が分かれる
    const middle = await open({
      sheetDataXml:
        '<row r="4" ht="30" customHeight="1"/><row r="5" ht="30" customHeight="1"/>' +
        '<row r="6" ht="30" customHeight="1"/><row r="5"/>',
    });
    expect(middle.layout.rows).toEqual([
      { min: 4, max: 4, ht: 30, hidden: false, custom: true },
      { min: 6, max: 6, ht: 30, hidden: false, custom: true },
    ]);
  });

  it('<cols> が複数あっても順にすべて読む。自己閉じの <cols/> は無視する', async () => {
    const sheet = await openXml(
      '<cols><col min="1" max="1" width="10"/></cols><cols/>' +
        '<cols><col min="3" max="3" width="20"/></cols><sheetData/>',
    );
    expect(sheet.layout.cols).toEqual([
      { min: 1, max: 1, width: 10, hidden: false },
      { min: 3, max: 3, width: 20, hidden: false },
    ]);
  });

  it('閉じの無い <cols> を大量に並べても、線形時間で読み捨てる', async () => {
    const t0 = performance.now();
    const sheet = await openXml(
      '<cols><col min="1" max="1" width="5"/>'.repeat(30_000) + '<sheetData/>',
    );
    expect(sheet.layout.cols).toEqual([]);
    expect(performance.now() - t0).toBeLessThan(2000);
  });

  it('styles.xml：先頭の font の既定の行高を fontRowPx に読む（実測表に無ければ null）', async () => {
    const calibri = await open({ rows: [['a']] }, { font: { name: 'Calibri', size: 11 } });
    expect(calibri.layout.fontRowPx).toBe(20);
    const arial12 = await open({ rows: [['a']] }, { font: { name: 'Arial', size: 12 } });
    expect(arial12.layout.fontRowPx).toBeNull();
    const none = await open({ rows: [['a']] });
    expect(none.layout.fontRowPx).toBeNull();
    expect(parseStyles(null).fontRowPx).toBeNull();
    expect(
      parseStyles(
        '<styleSheet><fonts><font><sz val="11"/><name val="游ゴシック"/></font></fonts></styleSheet>',
      ).fontRowPx,
    ).toBe(25);
  });

  it('styles.xml の <fonts> の閉じが無くても線形時間で読み、既定値にする', () => {
    const t0 = performance.now();
    const st = parseStyles('<styleSheet>' + '<fonts><font>'.repeat(30_000));
    expect([st.mdw, st.fontRowPx]).toEqual([7, null]);
    expect(performance.now() - t0).toBeLessThan(2000);
  });
});
