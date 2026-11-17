import { describe, expect, it } from 'vitest';
import { openXlsx } from '../../src/core/index';
import { MAX_CELL_TEXT } from '../../src/core/limits';
import { readRuns } from '../../src/core/rich-text';
import { buildXlsx, fontXml } from '../../src/testing/make-xlsx';

/** リッチテキストの区切りの読み取り（M7-16・ADR-0009 §1） */

const bold = (body: string | null) => (body?.includes('<b') ? { bold: true as const } : {});

describe('readRuns', () => {
  it('<rPr> のある区切りに書式を付け、無い区切りは書式なし。空の区切りは捨てる', () => {
    const inner =
      '<r><t>検索</t></r><r><rPr><b val="1"/></rPr><t>テスト</t></r><r><rPr/><t/></r><r><t xml:space="preserve">（境目）</t></r>';
    expect(readRuns(inner, '検索テスト（境目）', bold)).toEqual([
      { text: '検索' },
      { text: 'テスト', font: { bold: true } },
      { text: '（境目）' },
    ]);
  });

  it('<rPr> が 1 つも無い・連結が平文と違う（<si> 直下の <t> が混ざる）なら持たない', () => {
    expect(readRuns('<r><t>a</t></r><r><t>b</t></r>', 'ab', bold)).toBeUndefined();
    expect(readRuns('<t>x</t><r><rPr><b/></rPr><t>a</t></r>', 'xa', bold)).toBeUndefined();
    expect(readRuns('<t>abc</t>', 'abc', bold)).toBeUndefined();
  });

  it('ふりがな（<rPh>）は区切りに入れない。実体参照を戻す', () => {
    const inner =
      '<r><rPr><b/></rPr><t>漢字&amp;</t></r><rPh sb="0" eb="2"><t>カンジ</t></rPh><phoneticPr fontId="1"/>';
    expect(readRuns(inner, '漢字&', bold)).toEqual([{ text: '漢字&', font: { bold: true } }]);
  });

  it('平文を長さの上限で切ったときは、区切りも同じ位置で切る', () => {
    expect(
      readRuns('<r><t>ab</t></r><r><rPr><b/></rPr><t>cde</t></r><r><t>f</t></r>', 'abc', bold),
    ).toEqual([{ text: 'ab' }, { text: 'c', font: { bold: true } }]);
  });
});

describe('CellView.runs（openXlsx）', () => {
  const styles = {
    fonts: [
      fontXml('<sz val="11"/><name val="游ゴシック"/>'),
      fontXml('<b/><color rgb="FF0000FF"/>'),
    ],
    xfs: [0, 0],
    xfFontIds: [0, 1],
  };

  it('inlineStr と共有文字列の区切りを読む。<rPr> は <font> と同じ読み方（色・大きさの比・下線・黒は自動）', async () => {
    const book = await openXlsx(
      await buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: [
              [
                {
                  t: 'inlineStr',
                  is: '<r><t>検索</t></r><r><rPr><b val="1"/><color rgb="FFFF0000"/></rPr><t>テスト</t></r><r><t>（書式の境目）</t></r>',
                },
                { t: 's', v: 0, s: 1 },
                { t: 's', v: 1 },
                'plain',
              ],
            ],
          },
        ],
        sharedStrings: [
          '<r><rPr><sz val="22"/><u val="double"/><strike/><i/><color rgb="FF000000"/></rPr><t>大</t></r><r><t>小</t></r>',
          '<t>平文</t>',
        ],
        ...styles,
      }),
    );
    const row = (await book.openSheet(1)).getRows(1, 1)[0]!;
    expect(row.cells[0]).toMatchObject({
      text: '検索テスト（書式の境目）',
      kind: 's',
      runs: [
        { text: '検索' },
        { text: 'テスト', font: { bold: true, color: { rgb: 0xff0000 } } },
        { text: '（書式の境目）' },
      ],
    });
    expect(row.cells[1]!.runs).toEqual([
      { text: '大', font: { sizeRatio: 2, underline: 'double', strike: true, italic: true } },
      { text: '小' },
    ]);
    expect(row.cells[2]!.runs).toBeUndefined();
    expect(row.cells[3]!.runs).toBeUndefined();
    book.close?.();
  });

  it('styles.xml の無いブックは区切りの書式を読まない（平文）', async () => {
    const book = await openXlsx(
      await buildXlsx({
        sheets: [{ name: 'S', rows: [[{ t: 'inlineStr', is: '<r><rPr><b/></rPr><t>a</t></r>' }]] }],
        noStyles: true,
      }),
    );
    const cell = (await book.openSheet(1)).getRows(1, 1)[0]!.cells[0]!;
    expect(cell.text).toBe('a');
    expect(cell.runs).toBeUndefined();
    book.close?.();
  });

  it('長い文字は MAX_CELL_TEXT で切り、区切りも合わせる', async () => {
    const long = 'x'.repeat(MAX_CELL_TEXT);
    const book = await openXlsx(
      await buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: [
              [{ t: 'inlineStr', is: `<r><t>${long}</t></r><r><rPr><b/></rPr><t>yy</t></r>` }],
            ],
          },
        ],
        ...styles,
      }),
    );
    const cell = (await book.openSheet(1)).getRows(1, 1)[0]!.cells[0]!;
    expect(cell.text).toHaveLength(MAX_CELL_TEXT);
    expect(cell.runs).toEqual([{ text: long }]);
    book.close?.();
  });
});
