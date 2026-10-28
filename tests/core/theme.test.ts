/// <reference types="node" />
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';
import { resolveColor } from '../../src/core/color';
import { openXlsx } from '../../src/core/index';
import { DEFAULT_LIMITS } from '../../src/core/limits';
import { buildThemeXml, buildXlsx, wsDr } from '../../src/testing/make-xlsx';
import { openPackage } from '../../src/core/open';
import {
  DEFAULT_THEME,
  THEME_SLOTS,
  loadTheme,
  parseTheme,
  type Theme,
} from '../../src/core/theme';

// loadTheme を呼んだ回数を数えるため、実体を包む（結果は変えない）
vi.mock('../../src/core/theme', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../../src/core/theme')>();
  return { ...actual, loadTheme: vi.fn(actual.loadTheme) };
});
import { loadWorkbook } from '../../src/core/workbook';

/** テーマを `#rrggbb` の配列（THEME_SLOTS の順）にする */
const hexes = (t: Theme): string[] =>
  THEME_SLOTS.map((s) => (t.get(s) ?? -1).toString(16).toUpperCase().padStart(6, '0'));

const DEFAULT_HEX = [
  '000000',
  'FFFFFF',
  '44546A',
  'E7E6E6',
  '4472C4',
  'ED7D31',
  'A5A5A5',
  'FFC000',
  '5B9BD5',
  '70AD47',
  '0563C1',
  '954F72',
];

/** zip を開いて、workbook からテーマの場所を引き、読む */
async function themeOf(
  spec: Parameters<typeof buildXlsx>[0],
  limits = DEFAULT_LIMITS,
): Promise<{ theme: Theme; path: string | null }> {
  const pkg = await openPackage(await buildXlsx(spec), limits);
  try {
    const wb = await loadWorkbook(pkg, limits);
    return { theme: await loadTheme(pkg, wb.themePath, limits), path: wb.themePath };
  } finally {
    pkg.dispose();
  }
}

describe('TD-94-5 既定の 12 色', () => {
  it('Office 2013〜2022 の値で、THEME_SLOTS の順に 12 色ある', () => {
    expect(THEME_SLOTS).toEqual([
      'dk1',
      'lt1',
      'dk2',
      'lt2',
      'accent1',
      'accent2',
      'accent3',
      'accent4',
      'accent5',
      'accent6',
      'hlink',
      'folHlink',
    ]);
    expect(hexes(DEFAULT_THEME)).toEqual(DEFAULT_HEX);
  });

  it('Excel 16 が書いた ref-shapes.xlsx のテーマは既定の 12 色と一致する', async () => {
    const zip = await JSZip.loadAsync(
      readFileSync(new URL('../../fixtures/excel/ref-shapes.xlsx', import.meta.url)),
    );
    const xml = await zip.file('xl/theme/theme1.xml')!.async('string');
    expect(hexes(parseTheme(xml))).toEqual(DEFAULT_HEX);
    // ブックを開く経路でも同じ（rels の theme の関係から引く）
    const pkg = await openPackage(
      readFileSync(new URL('../../fixtures/excel/ref-shapes.xlsx', import.meta.url)),
      DEFAULT_LIMITS,
    );
    const wb = await loadWorkbook(pkg, DEFAULT_LIMITS);
    expect(wb.themePath).toBe('xl/theme/theme1.xml');
    expect(hexes(await loadTheme(pkg, wb.themePath, DEFAULT_LIMITS))).toEqual(DEFAULT_HEX);
    pkg.dispose();
  });
});

describe('TD-94-6 テーマの 12 色の読み取り（parseTheme）', () => {
  it('srgbClr と sysClr（lastClr）を読む', () => {
    const xml = buildThemeXml({
      dk1: '<a:sysClr val="windowText" lastClr="111111"/>',
      lt1: '<a:srgbClr val="eeeeee"/>',
      dk2: '<a:srgbClr val="222222"/>',
      accent1: '<a:srgbClr val="AA0000"/>',
      folHlink: '<a:srgbClr val="00AA00"/>',
    });
    const t = parseTheme(xml);
    expect(t.get('dk1')).toBe(0x111111);
    expect(t.get('lt1')).toBe(0xeeeeee);
    expect(t.get('dk2')).toBe(0x222222);
    expect(t.get('accent1')).toBe(0xaa0000);
    expect(t.get('folHlink')).toBe(0x00aa00);
    // 書いていない色は既定
    expect(t.get('accent2')).toBe(0xed7d31);
    expect(t.size).toBe(12);
  });

  it('sysClr の lastClr が無ければ windowText＝黒・window＝白', () => {
    const t = parseTheme(
      buildThemeXml({
        dk1: '<a:sysClr val="windowText"/>',
        lt1: '<a:sysClr val="window"/>',
        dk2: '<a:sysClr val="menu"/>',
      }),
    );
    expect(t.get('dk1')).toBe(0x000000);
    expect(t.get('lt1')).toBe(0xffffff);
    // 解けない sysClr はその色だけ既定
    expect(t.get('dk2')).toBe(0x44546a);
  });

  it('欠けた色・不正な色・色の要素が無い色は、その色だけ既定で補う', () => {
    const t = parseTheme(
      buildThemeXml({
        accent1: '<a:srgbClr val="12345"/>',
        accent2: '<a:srgbClr val="GGGGGG"/>',
        accent3: '',
        accent4: '<a:schemeClr val="accent1"/>',
        accent5: '<a:srgbClr val="010203"/>',
      }),
    );
    expect(t.get('accent1')).toBe(0x4472c4);
    expect(t.get('accent2')).toBe(0xed7d31);
    expect(t.get('accent3')).toBe(0xa5a5a5);
    expect(t.get('accent4')).toBe(0xffc000);
    expect(t.get('accent5')).toBe(0x010203);
  });

  it('名前空間の接頭辞が違っても読める。別の場所（fmtScheme）にある同じ名前の要素は拾わない', () => {
    const xml =
      '<theme><themeElements><fmtScheme><accent1><srgbClr val="FF00FF"/></accent1></fmtScheme>' +
      '<t:clrScheme name="x"><t:accent1><t:srgbClr val="010203"/></t:accent1></t:clrScheme></themeElements></theme>';
    expect(parseTheme(xml).get('accent1')).toBe(0x010203);
  });

  it('null・clrScheme が無い・閉じが無い・空は既定の 12 色', () => {
    expect(hexes(parseTheme(null))).toEqual(DEFAULT_HEX);
    expect(hexes(parseTheme(''))).toEqual(DEFAULT_HEX);
    expect(hexes(parseTheme('<a:theme/>'))).toEqual(DEFAULT_HEX);
    expect(hexes(parseTheme('<a:clrScheme><a:accent1><a:srgbClr val="010203"/>'))).toEqual(
      DEFAULT_HEX,
    );
    expect(hexes(parseTheme('not xml at all'))).toEqual(DEFAULT_HEX);
  });

  it('__proto__・constructor などの名前の要素が混ざっても壊れない', () => {
    const xml =
      '<a:clrScheme><a:__proto__><a:srgbClr val="010203"/></a:__proto__>' +
      '<a:constructor><a:srgbClr val="040506"/></a:constructor>' +
      '<a:accent1><a:srgbClr val="070809"/></a:accent1></a:clrScheme>';
    const t = parseTheme(xml);
    expect(t.size).toBe(12);
    expect(t.get('accent1')).toBe(0x070809);
    expect(Object.keys(t)).toEqual([]);
    // schemeClr の名前としても解けない
    expect(resolveColor('<a:schemeClr val="__proto__"/>', t)).toBeNull();
    expect(resolveColor('<a:schemeClr val="constructor"/>', t)).toBeNull();
  });

  it('DEFAULT_THEME は変更できない（parseTheme は新しい Map を返す）', () => {
    const t = parseTheme(null);
    expect(t).not.toBe(DEFAULT_THEME);
    (t as Map<string, number>).set('accent1', 0);
    expect(DEFAULT_THEME.get('accent1')).toBe(0x4472c4);
  });

  it('巨大な clrScheme・大量の閉じの無い要素でも線形時間で既定で補う', () => {
    const N = 40_000;
    const LIMIT_MS = 1500;
    const xmls = [
      `<a:clrScheme>${'<a:accent1 '.repeat(N)}</a:clrScheme>`,
      `<a:clrScheme>${'<a:accent1>'.repeat(N)}</a:clrScheme>`,
      `<a:clrScheme><a:accent1>${'<a:srgbClr '.repeat(N)}</a:accent1></a:clrScheme>`,
      `<a:clrScheme>${'<a:accent1><a:srgbClr val="zzzzzz"/></a:accent1>'.repeat(N)}</a:clrScheme>`,
      `<a:clrScheme>${'<a:dk1>'.repeat(N)}${'<a:hlink>'.repeat(N)}`,
      '<a:clrScheme '.repeat(N),
      '<a:clrScheme>'.repeat(N),
    ];
    for (const xml of xmls) {
      const t0 = performance.now();
      const t = parseTheme(xml);
      expect(performance.now() - t0).toBeLessThan(LIMIT_MS);
      expect(t.size).toBe(12);
    }
  });
});

describe('TD-94-7 テーマの場所（workbook.xml.rels）と読み込み（loadTheme）', () => {
  it('theme の関係の Target を引いて読む', async () => {
    const { theme, path } = await themeOf({
      sheets: [{ name: 'S' }],
      theme: { xml: buildThemeXml({ accent1: '<a:srgbClr val="0A0B0C"/>' }) },
    });
    expect(path).toBe('xl/theme/theme1.xml');
    expect(theme.get('accent1')).toBe(0x0a0b0c);
    expect(theme.get('accent2')).toBe(0xed7d31);
  });

  it('別の名前・場所の theme でも rels で引ける', async () => {
    const { theme, path } = await themeOf({
      sheets: [{ name: 'S' }],
      theme: { xml: buildThemeXml({ accent1: '<a:srgbClr val="0A0B0C"/>' }), file: 'other.xml' },
    });
    expect(path).toBe('xl/theme/other.xml');
    expect(theme.get('accent1')).toBe(0x0a0b0c);
  });

  it('rels に theme の関係が無い（部品があっても）・部品が zip に無い・rels 自体が無いときは既定の 12 色', async () => {
    const noRel = await themeOf({
      sheets: [{ name: 'S' }],
      theme: { xml: buildThemeXml({ accent1: '<a:srgbClr val="0A0B0C"/>' }), omitRel: true },
    });
    expect(noRel.path).toBeNull();
    expect(hexes(noRel.theme)).toEqual(DEFAULT_HEX);

    const noPart = await themeOf({
      sheets: [{ name: 'S' }],
      theme: { xml: '', omitPart: true },
    });
    expect(noPart.path).toBe('xl/theme/theme1.xml');
    expect(hexes(noPart.theme)).toEqual(DEFAULT_HEX);

    const none = await themeOf({ sheets: [{ name: 'S' }] });
    expect(none.path).toBeNull();
    expect(hexes(none.theme)).toEqual(DEFAULT_HEX);
  });

  it('小さな部品の上限を超える theme は既定の 12 色（例外にしない）', async () => {
    const limits = { ...DEFAULT_LIMITS, maxSmallPartBytes: 2000 };
    const xml = buildThemeXml({ accent1: '<a:srgbClr val="0A0B0C"/>' }) + ' '.repeat(5000);
    // workbook.xml・rels が小さな上限に収まるよう、テーマだけが大きい合成にする
    const { theme } = await themeOf({ sheets: [{ name: 'S' }], theme: { xml } }, limits);
    expect(hexes(theme)).toEqual(DEFAULT_HEX);
  });

  it('壊れた XML の theme でも既定の 12 色（例外にしない）', async () => {
    const { theme } = await themeOf({
      sheets: [{ name: 'S' }],
      theme: { xml: '<<<>>> \u0000 <a:clrScheme><a:accent1>' },
    });
    expect(hexes(theme)).toEqual(DEFAULT_HEX);
  });

  it('theme の関係が外部（TargetMode=External）なら辿らない', async () => {
    const pkg = await openPackage(
      await buildXlsx({
        sheets: [{ name: 'S' }],
        extraFiles: {
          'xl/_rels/workbook.xml.rels':
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="https://example.com/theme.xml" TargetMode="External"/>' +
            '</Relationships>',
        },
      }),
      DEFAULT_LIMITS,
    );
    const wb = await loadWorkbook(pkg, DEFAULT_LIMITS);
    expect(wb.themePath).toBeNull();
    pkg.dispose();
  });
});

describe('TD-94-8 openXlsx への配線（テーマはブックで 1 回だけ・表示は変わらない）', () => {
  /** 図（空の drawing）を持つシートを 2 枚と、テーマを持つブック */
  const drawingBook = (theme?: Parameters<typeof buildXlsx>[0]['theme']) =>
    buildXlsx({
      sheets: [
        { name: 'A', rows: [['a']], drawing: { xml: wsDr([]), rels: [] } },
        { name: 'B', rows: [['b']], drawing: { xml: wsDr([]), rels: [] } },
        { name: 'C', rows: [['c']] },
      ],
      theme,
    });
  const themed = { xml: buildThemeXml({ accent1: '<a:srgbClr val="010203"/>' }) };
  const calls = () => vi.mocked(loadTheme).mock.calls.length;

  it('図のあるシートの openDrawings でテーマを 1 回だけ読む（別のシート・2 回目では読み直さない）', async () => {
    const book = await openXlsx(await drawingBook(themed));
    vi.mocked(loadTheme).mockClear();
    await book.openDrawings(1);
    expect(calls()).toBe(1);
    await book.openDrawings(2);
    await book.openDrawings(1);
    expect(calls()).toBe(1);
    // 呼び出しの引数：zip の中の theme の場所
    expect(vi.mocked(loadTheme).mock.calls[0]![1]).toBe('xl/theme/theme1.xml');
    book.close();
  });

  it('図の無いシートでは読まない（図の無いブックでテーマの読み取りを増やさない）', async () => {
    const book = await openXlsx(await drawingBook(themed));
    vi.mocked(loadTheme).mockClear();
    await book.openDrawings(3);
    await (await book.openSheet(3)).getRows(1, 1);
    expect(calls()).toBe(0);
    book.close();
  });

  it('close() の後は closed で失敗し、テーマを読み直さない', async () => {
    const book = await openXlsx(await drawingBook(themed));
    await book.openDrawings(1);
    book.close();
    vi.mocked(loadTheme).mockClear();
    await expect(book.openDrawings(1)).rejects.toMatchObject({ kind: 'broken' });
    await expect(book.openDrawings(2)).rejects.toThrow(/closed/);
    expect(calls()).toBe(0);
  });

  it('テーマがあってもなくても壊れても、図の読み取りの結果は変わらない', async () => {
    const results = [];
    for (const theme of [undefined, themed, { xml: '<<<' }]) {
      const book = await openXlsx(await drawingBook(theme));
      results.push(await book.openDrawings(1));
      book.close();
    }
    expect(results[1]).toEqual(results[0]);
    expect(results[2]).toEqual(results[0]);
  });
});
