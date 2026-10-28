import { describe, expect, it } from 'vitest';
import { openXlsx, type XlsxSheet } from '../../src/core/index';
import { buildXlsx, type CellSpec, type XlsxSpec } from '../../src/testing/make-xlsx';

/** 1 シートの xlsx を組んで開き、シートを返す */
async function sheetOf(
  rows: CellSpec[][],
  extra: Partial<XlsxSpec> = {},
  merges?: string[],
): Promise<XlsxSheet> {
  const book = await openXlsx(await buildXlsx({ sheets: [{ name: 'S', rows, merges }], ...extra }));
  return book.openSheet(1);
}

/** 1 行目の文字だけを並べる */
const firstRow = (sheet: XlsxSheet) => sheet.getRows(1, 1)[0]!.cells.map((c) => c?.text);

describe('共有文字列', () => {
  const sharedStrings = [
    '<t>plain</t>',
    '<r><rPr><b/></rPr><t>bold</t></r><r><t xml:space="preserve"> and normal</t></r>',
    '<t xml:space="preserve">line1\nline2</t>',
    '<r><t>kanji</t></r><rPh sb="0" eb="1"><t>furigana</t></rPh>',
    '<t>a &amp; b &lt;c&gt; &#x41;&#66; _x0041_ _x005F_x0041_ end_x000D_</t>',
    '<t/>',
  ];
  const idx = (n: number): CellSpec => ({ t: 's', v: n });

  it('平文・リッチテキスト（連結）・改行・ふりがな（除く）・実体参照と _xHHHH_ を戻す', async () => {
    const sheet = await sheetOf([[idx(0), idx(1), idx(2), idx(3), idx(4), idx(5)]], {
      sharedStrings,
    });
    expect(firstRow(sheet)).toEqual([
      'plain',
      'bold and normal',
      'line1\nline2',
      'kanji',
      // &amp; → &、&lt;&gt; → <>、文字参照 → AB、_x0041_ → A、_x005F_ → _（その後ろは文字のまま）、_x000D_ → CR
      'a & b <c> AB A _x0041_ end\r',
      // 空の共有文字列（<t/>）は、見た目は空のまま「値のある文字列セル」（ISSUE-84。右隣へのはみ出しを止める）
      '',
    ]);
  });

  it('範囲外の番号・数でない番号は空のセル', async () => {
    const sheet = await sheetOf([[idx(99), { t: 's', v: 'abc' }, idx(0)]], { sharedStrings });
    expect(firstRow(sheet)).toEqual([undefined, undefined, 'plain']);
  });
});

describe('値の種類', () => {
  it('inline 文字列・数式の文字（t="str"）・真偽・エラー', async () => {
    const sheet = await sheetOf([
      [
        'inline text',
        { t: 'str', v: 'cached & <ok>', f: 'A1&B1' },
        { t: 'b', v: 1 },
        { t: 'b', v: 0 },
        true,
        { t: 'e', v: '#DIV/0!' },
        { t: 'e', v: '#N/A' },
      ],
    ]);
    expect(firstRow(sheet)).toEqual([
      'inline text',
      'cached & <ok>',
      'TRUE',
      'FALSE',
      'TRUE',
      '#DIV/0!',
      '#N/A',
    ]);
  });

  it('inline 文字列の改行・実体参照・先頭と末尾の空白を保つ', async () => {
    const sheet = await sheetOf([[' a\nb & c ']]);
    expect(firstRow(sheet)).toEqual([' a\nb & c ']);
  });

  it('数式の数値は保存済みの値を、書式を当てて出す', async () => {
    const sheet = await sheetOf(
      [
        [
          { v: 0.5, f: 'SUM(A2:A3)' },
          { v: 1234.5, f: 'B2*2', s: 1 },
        ],
      ],
      { xfs: [0, 4] },
    );
    expect(firstRow(sheet)).toEqual(['0.5', '1,234.50']);
  });

  it('値の無いセル（書式だけ）・空の <v> は空のセル', async () => {
    const sheet = await sheetOf([[{ s: 1 }, { t: 'str', v: '' }, { v: '' }, 'x']], { xfs: [0, 3] });
    expect(firstRow(sheet)).toEqual([undefined, undefined, undefined, 'x']);
    expect(sheet.colCount).toBe(4);
  });

  it('数でない <v> は書かれた文字のまま出す', async () => {
    const sheet = await sheetOf([[{ v: 'not-a-number' }]]);
    expect(firstRow(sheet)).toEqual(['not-a-number']);
  });

  it('1 つのセルの文字数は Excel の上限 32,767 で切る', async () => {
    const sheet = await sheetOf([['x'.repeat(40_000)]]);
    expect(firstRow(sheet)[0]!.length).toBe(32_767);
  });
});

describe('表示形式', () => {
  const numFmts = {
    164: '#,##0;\\-#,##0;"－";@', // 4 区画（正・負・0・文字）
    165: '0.0_);[Red]\\(0.0\\)',
    166: 'yyyy"年"m"月"d"日"',
    167: '0.00E+00',
  };
  // xf の位置 → numFmtId
  const xfs = [0, 3, 4, 2, 9, 10, 14, 22, 164, 165, 166, 167, 49, 21];
  const x = (n: number, v: number): CellSpec => ({ v, s: xfs.indexOf(n) });

  it('桁区切り・小数・％', async () => {
    const sheet = await sheetOf(
      [[x(3, 1234567), x(4, 1234.5), x(2, 0.5), x(9, 0.153), x(10, 0.153)]],
      { xfs, numFmts },
    );
    expect(firstRow(sheet)).toEqual(['1,234,567', '1,234.50', '0.50', '15%', '15.30%']);
  });

  it('4 区画の形式：正・負・0', async () => {
    const sheet = await sheetOf(
      [
        [x(164, 1234.5), x(164, -1234.5), x(164, 0)],
        [x(165, 3.14), x(165, -3.14)],
      ],
      { xfs, numFmts },
    );
    const rows = sheet.getRows(1, 2).map((r) => r.cells.map((c) => c?.text));
    expect(rows).toEqual([
      ['1,235', '-1,235', '－'],
      ['3.1 ', '(3.1)', undefined],
    ]);
  });

  it('日付・日時・時刻・日本語混じりの日付（日本語環境の組み込み 14・22）', async () => {
    const sheet = await sheetOf(
      [[x(14, 45000), x(22, 45000.75), x(21, 0.5), x(166, 45000), x(167, 1500000)]],
      { xfs, numFmts },
    );
    expect(firstRow(sheet)).toEqual([
      '2023/3/15',
      '2023/3/15 18:00',
      '12:00:00',
      '2023年3月15日',
      '1.50E+06',
    ]);
  });

  it('書式 General（既定）：浮動小数の誤差を丸め、整数は桁を落とさない', async () => {
    const sheet = await sheetOf([
      [0.30000000000000004, 1234.5, 12345678901, -7, 1.2345678901234567e19],
    ]);
    expect(firstRow(sheet)).toEqual(['0.3', '1234.5', '12345678901', '-7', '1.23457E+19']);
  });

  it('書式「@」の数値は数字のまま', async () => {
    const sheet = await sheetOf([[x(49, 12345)]], { xfs, numFmts });
    expect(firstRow(sheet)).toEqual(['12345']);
  });

  it('1904 年起点の日付', async () => {
    const sheet = await sheetOf([[x(14, 0), x(14, 1462)]], { xfs, numFmts, date1904: true });
    expect(firstRow(sheet)).toEqual(['1904/1/1', '1908/1/2']);
  });

  it('t="d"（ISO 日付）：日付の書式があればそれで、無ければ既定の日付書式で出す', async () => {
    const sheet = await sheetOf(
      [
        [
          { t: 'd', v: '2023-03-15' },
          { t: 'd', v: '2023-03-15T18:00:00', s: 0 },
          { t: 'd', v: '2023-03-15', s: xfs.indexOf(166) },
          { t: 'd', v: 'bad-date' },
        ],
      ],
      { xfs, numFmts },
    );
    expect(firstRow(sheet)).toEqual([
      '2023/3/15',
      '2023/3/15 18:00:00',
      '2023年3月15日',
      'bad-date',
    ]);
  });

  it('t="d" は 1904 年起点でも同じ日付に見える', async () => {
    const sheet = await sheetOf([[{ t: 'd', v: '2023-03-15' }]], { date1904: true });
    expect(firstRow(sheet)).toEqual(['2023/3/15']);
  });

  it('解釈できない書式・範囲外の s は、壊れずに数字を出す', async () => {
    const sheet = await sheetOf(
      [
        [
          { v: 5, s: 1 },
          { v: 6, s: 99 },
        ],
      ],
      { xfs: [0, 200], numFmts: { 200: '0;0;0;0;0;0;0' } },
    );
    expect(firstRow(sheet)).toEqual(['5', '6']);
  });

  // 書式コードの長さ上限（ISSUE-82・ADR-0025 決定 5）。超えたら書式を当てず値だけを出す
  it('TD-82-1: 長すぎる書式コード（255 文字超）は書式を当てず、値をそのまま出す', async () => {
    const longCode = '0' + '#'.repeat(300); // 301 文字
    const sheet = await sheetOf(
      [
        [
          { v: 1234.5, s: 1 },
          { v: 1234.5, s: 2 },
        ],
      ],
      {
        xfs: [0, 200, 201],
        numFmts: { 200: longCode, 201: '#,##0.00' },
      },
    );
    expect(firstRow(sheet)).toEqual(['1234.5', '1,234.50']);
  });
});

describe('結合セル', () => {
  const rows: CellSpec[][] = [
    ['A1 merged', 'stray B1', 'stray C1', 'D1'],
    ['stray A2', 'B2', 'C2', 'D2'],
    ['A3', 'B3', 'C3', 'D3'],
    ['A4', 'B4', 'C4', 'D4'],
  ];
  const text = (sheet: XlsxSheet, first: number, last: number) =>
    sheet
      .getRows(first, last)
      .map((r) =>
        r.cells.map((c) =>
          c === null ? null : c === undefined ? undefined : `${c.text}[${c.colSpan}x${c.rowSpan}]`,
        ),
      );

  it('左上だけ描き、colSpan・rowSpan で表す。左上以外に値が残っていても出さない', async () => {
    const sheet = await sheetOf(rows, {}, ['A1:C2']);
    expect(text(sheet, 1, 4)).toEqual([
      ['A1 merged[3x2]', null, null, 'D1[1x1]'],
      [null, null, null, 'D2[1x1]'],
      ['A3[1x1]', 'B3[1x1]', 'C3[1x1]', 'D3[1x1]'],
      ['A4[1x1]', 'B4[1x1]', 'C4[1x1]', 'D4[1x1]'],
    ]);
  });

  it('左上が範囲より上にあるとき、範囲の最初の行に左上の文字を置き、rowSpan を範囲内に縮める', async () => {
    const sheet = await sheetOf(rows, {}, ['A1:C4']);
    expect(text(sheet, 3, 3)).toEqual([['A1 merged[3x1]', null, null, 'D3[1x1]']]);
    expect(text(sheet, 2, 3)).toEqual([
      ['A1 merged[3x2]', null, null, 'D2[1x1]'],
      [null, null, null, 'D3[1x1]'],
    ]);
  });

  it('範囲の下で切れるときも rowSpan を範囲内に縮める', async () => {
    const sheet = await sheetOf(rows, {}, ['A2:B4']);
    expect(text(sheet, 1, 3)[1]![0]).toBe('stray A2[2x2]');
  });

  it('値の無い左上も結合として描く（文字は空）', async () => {
    const sheet = await sheetOf([[null, 'x'], ['y']], {}, ['A1:A2']);
    expect(
      sheet
        .getRows(1, 2)
        .map((r) => r.cells.map((c) => (c ? `${c.text}[${c.colSpan}x${c.rowSpan}]` : c))),
    ).toEqual([
      ['[1x2]', 'x[1x1]'],
      [null, undefined],
    ]);
  });

  it('結合が最後の行・列を決める（値が無くても）', async () => {
    const sheet = await sheetOf([['a']], {}, ['B2:D5']);
    expect([sheet.rowCount, sheet.colCount]).toEqual([5, 4]);
    // 最後の行だけ見ると、左上が範囲より上にあるので、その行の頭（B 列）に結合を置き直す
    expect(sheet.getRows(5, 5)[0]!.cells).toEqual([
      undefined,
      { text: '', colSpan: 3, rowSpan: 1 },
      null,
      null,
    ]);
  });

  it('1 セルだけの結合・不正な番地は結合として扱わない', async () => {
    const sheet = await sheetOf(rows, {}, ['A1:A1', 'B2', 'ZZZZ99:AAAA100', 'oops']);
    expect(text(sheet, 1, 1)[0]![0]).toBe('A1 merged[1x1]');
    expect([sheet.rowCount, sheet.colCount]).toEqual([4, 4]);
  });

  it('逆向きの範囲（C3:A1）も同じ結合', async () => {
    const sheet = await sheetOf(rows, {}, ['C2:A1']);
    expect(text(sheet, 1, 2)[0]!.slice(0, 3)).toEqual(['A1 merged[3x2]', null, null]);
  });
});

describe('セルの型（kind・ISSUE-84）', () => {
  /** 1 行目のセルの型を並べる（セルが無ければ undefined） */
  const kinds = (sheet: XlsxSheet) => sheet.getRows(1, 1)[0]!.cells.map((c) => c?.kind);

  it('TD-84-18: t 無し・n は数値、s・str・inlineStr は文字列、b は真偽、e はエラー、d は日付', async () => {
    const sheet = await sheetOf(
      [
        [
          { v: 1 },
          { t: 'n', v: 2 },
          'inline',
          { t: 'str', v: 'formula', f: 'A1' },
          { t: 's', v: 0 },
          { t: 'b', v: 1 },
          { t: 'e', v: '#N/A' },
          { t: 'd', v: '2024-01-15T00:00:00' },
        ],
      ],
      { sharedStrings: ['<t>shared</t>'] },
    );
    expect(kinds(sheet)).toEqual(['n', 'n', 's', 's', 's', 'b', 'e', 'd']);
  });

  it('TD-84-18: 書式で日付表示になる数値・通貨表示の数値は n のまま（数値なのではみ出さない）', async () => {
    const sheet = await sheetOf(
      [
        [
          { v: 45306, s: 1 },
          { v: 1234.5, s: 2 },
        ],
      ],
      { xfs: [0, 14, 4] },
    );
    expect(sheet.getRows(1, 1)[0]!.cells.map((c) => [c?.kind, c?.text])).toEqual([
      ['n', '2024/1/15'],
      ['n', '1,234.50'],
    ]);
  });

  it('TD-84-18: 数字だけの文字列（t="s"・inlineStr）・▲ 付きの文字列は s（文字列なので右へはみ出せる）', async () => {
    const sheet = await sheetOf([['12345678901234', '▲1,234', { t: 's', v: 0 }]], {
      sharedStrings: ['<t>00123</t>'],
    });
    expect(sheet.getRows(1, 1)[0]!.cells.map((c) => [c?.kind, c?.text])).toEqual([
      ['s', '12345678901234'],
      ['s', '▲1,234'],
      ['s', '00123'],
    ]);
  });

  it('TD-84-19: 数式の結果が空文字（t="str"・<f> と空の <v>）は、text が空の文字列セルとして残る', async () => {
    const sheet = await sheetOf([['a', { t: 'str', f: 'IF(A1="","x","")', v: '' }, 'c']]);
    const cells = sheet.getRows(1, 1)[0]!.cells;
    expect(cells[1]).toEqual({ text: '', kind: 's', colSpan: 1, rowSpan: 1 });
    expect(sheet.colCount).toBe(3);
  });

  it('TD-84-19: 空の共有文字列は文字列セル。範囲外・数でない番号・書式だけ・数式の無い空の str・空の inlineStr・数値の空の v は空のセル', async () => {
    const sheet = await sheetOf(
      [
        [
          { t: 's', v: 0 }, // 空の共有文字列 → 文字列セル
          { t: 's', v: 9 }, // 範囲外
          { t: 's', v: 'abc' }, // 数でない番号
          { s: 1 }, // 書式だけ
          { t: 'str', v: '' }, // 数式の無い空の str
          { t: 'inlineStr', v: '' }, // 空の inlineStr
          { v: '' }, // 空の <v>
          { t: 'str', f: 'A1' }, // 値の要素の無い数式（未計算）
          'x',
        ],
      ],
      { sharedStrings: ['<t/>'], xfs: [0, 3] },
    );
    expect(kinds(sheet)).toEqual([
      's',
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      's',
    ]);
    expect(sheet.getRows(1, 1)[0]!.cells[0]!.text).toBe('');
    expect(sheet.colCount).toBe(9);
  });
});
