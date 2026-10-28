import { describe, expect, it } from 'vitest';
import {
  EMPTY_FORMAT_TABLE,
  isPaperFill,
  parseCellFormats,
  stylesUseTheme,
  type FormatTable,
} from '../../src/core/cell-format';
import { DEFAULT_LIMITS, type XlsxLimits } from '../../src/core/limits';
import { fontXml, gradientFillXml, patternFillXml } from '../../src/testing/make-xlsx';
import { DEFAULT_THEME } from '../../src/core/theme';
import type { CellFormat } from '../../src/core/types';

/**
 * styles.xml（fonts・fills・cellXfs）を組む。xfs は cellXfs の `[fontId, fillId]` の並び（位置が xf 番号）。
 * cellStyleXfs にも `<xf>`（fontId・fillId が 1）を置く（cellXfs の数え方が混ざらないことの確認）
 */
function styles(opts: {
  fonts?: string[];
  fills?: string[];
  xfs: [number, number][];
  extra?: string;
}): string {
  const fonts = opts.fonts
    ? `<fonts count="${opts.fonts.length}">${opts.fonts.join('')}</fonts>`
    : '';
  const fills = opts.fills
    ? `<fills count="${opts.fills.length}">${opts.fills.join('')}</fills>`
    : '';
  return (
    `<styleSheet>${fonts}${fills}` +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="1" fillId="1"/></cellStyleXfs>' +
    `<cellXfs count="${opts.xfs.length}">${opts.xfs.map(([fo, fi]) => `<xf numFmtId="0" fontId="${fo}" fillId="${fi}" applyFill="0" applyFont="0"/>`).join('')}</cellXfs>` +
    `${opts.extra ?? ''}</styleSheet>`
  );
}

const parse = (xml: string | null, limits: XlsxLimits = DEFAULT_LIMITS): FormatTable =>
  parseCellFormats(xml, DEFAULT_THEME, limits);

/** xf 番号の書式（書式なしは undefined） */
function fmt(t: FormatTable, xf: number): CellFormat | undefined {
  const i = t.xfFormat[xf];
  return i === undefined || i < 0 ? undefined : t.formats[i];
}

/** fills の 1 件だけを xf 1 が指すスタイルを読み、xf 1 の塗りを返す（fills[0] は none） */
function fillOf(fill: string, extra = ''): CellFormat['fill'] {
  const t = parse(
    styles({
      fills: [patternFillXml('none'), fill],
      xfs: [
        [0, 0],
        [0, 1],
      ],
      extra,
    }),
  );
  return fmt(t, 1)?.fill;
}

describe('parseCellFormats：塗り（solid）', () => {
  it('solid は fgColor（rgb。先頭のアルファは無視）。bgColor は使わない', () => {
    expect(fillOf(patternFillXml('solid', 'rgb="FFFF0000"', 'rgb="FF0000FF"'))).toEqual({
      rgb: 0xff0000,
      alpha: 1,
    });
    expect(fillOf(patternFillXml('solid', 'rgb="00FFFF00"'))).toEqual({ rgb: 0xffff00, alpha: 1 });
  });

  it('fgColor は theme＋tint・indexed・<indexedColors> の上書きも解く', () => {
    expect(fillOf(patternFillXml('solid', 'theme="4"'))?.rgb).toBe(0x4472c4);
    // accent1(4472C4) に tint=0.8（Excel が「アクセント 1、白 + 基本色 80%」に書く値）。cell-color.ts と同じ式
    expect(fillOf(patternFillXml('solid', 'theme="4" tint="0.7999816888943144"'))?.rgb).toBe(
      0xdae3f3,
    );
    expect(fillOf(patternFillXml('solid', 'indexed="2"'))?.rgb).toBe(0xff0000);
    const palette = '<colors><indexedColors><rgbColor rgb="FF010203"/></indexedColors></colors>';
    expect(fillOf(patternFillXml('solid', 'indexed="0"'), palette)?.rgb).toBe(0x010203);
  });

  it('solid で fgColor が auto・解けない・無いときは黒（#000000）', () => {
    expect(fillOf(patternFillXml('solid', 'auto="1"'))).toEqual({ rgb: 0, alpha: 1 });
    expect(fillOf(patternFillXml('solid', 'rgb="zz"'))).toEqual({ rgb: 0, alpha: 1 });
    expect(fillOf(patternFillXml('solid', 'indexed="64"'))).toEqual({ rgb: 0, alpha: 1 });
    expect(fillOf(patternFillXml('solid'))).toEqual({ rgb: 0, alpha: 1 });
  });
});

describe('parseCellFormats：塗り（模様）', () => {
  it('patternType が無い・none・未知は塗りなし。塗りも装飾も無い xf は -1（formats は空）', () => {
    const t = parse(
      styles({
        fills: [
          patternFillXml('none', 'rgb="FFFF0000"'),
          patternFillXml(undefined, 'rgb="FFFF0000"'),
          patternFillXml('wavy', 'rgb="FFFF0000"'),
          '<fill/>',
        ],
        xfs: [
          [0, 0],
          [0, 1],
          [0, 2],
          [0, 3],
        ],
      }),
    );
    expect(Array.from(t.xfFormat)).toEqual([-1, -1, -1, -1]);
    expect(t.formats).toHaveLength(0);
  });

  it('gray125 は前景（黒）と背景（白）を 12.5% で混ぜる（色の既定は前景 黒・背景 白）', () => {
    expect(fillOf(patternFillXml('gray125'))?.rgb).toBe(0xdfdfdf);
  });

  it('mediumGray・darkHorizontal は 50%、lightGrid は 44%（7/16）。前景・背景を明示するとそれで混ぜる', () => {
    expect(fillOf(patternFillXml('mediumGray', 'rgb="FFFF0000"', 'rgb="FF0000FF"'))?.rgb).toBe(
      0x800080,
    );
    expect(fillOf(patternFillXml('darkHorizontal'))?.rgb).toBe(0x808080);
    expect(fillOf(patternFillXml('lightGrid'))?.rgb).toBe(0x8f8f8f);
  });

  it('密度の表：darkGray 75%・lightGray 25%・gray0625 6.25%・darkGrid 75%・lightTrellis 38%', () => {
    const g = (p: string): number | undefined => fillOf(patternFillXml(p))?.rgb;
    expect(g('darkGray')).toBe(0x404040); // 255 × 0.25 = 63.75
    expect(g('lightGray')).toBe(0xbfbfbf); // 255 × 0.75 = 191.25
    expect(g('gray0625')).toBe(0xefefef); // 255 × 0.9375 = 239.06
    expect(g('darkGrid')).toBe(0x404040);
    expect(g('lightTrellis')).toBe(0x9f9f9f); // 255 × 0.625 = 159.375
  });

  it('模様の色の auto・解けない色は既定（前景 黒・背景 白）に倒す', () => {
    expect(fillOf(patternFillXml('mediumGray', 'auto="1"', 'indexed="64"'))?.rgb).toBe(0x808080);
  });
});

describe('parseCellFormats：塗り（グラデーション）', () => {
  it('2 stop は線形 RGB の中央（赤→青 #BC00BC）', () => {
    const f = fillOf(
      gradientFillXml([
        [0, 'rgb="FFFF0000"'],
        [1, 'rgb="FF0000FF"'],
      ]),
    );
    expect(f?.rgb).toBe(0xbc00bc);
  });

  it('3 stop 以上は sRGB の混色で 50% の位置の色（stop は position の順に並べ替える）', () => {
    const f = fillOf(
      gradientFillXml([
        [1, 'rgb="FFFF0000"'],
        [0, 'rgb="FF000000"'],
        [0.25, 'rgb="FFFFFFFF"'],
      ]),
    );
    expect(f?.rgb).toBe(0xffaaaa);
  });

  it('stop が 1 つならその色。50% より後ろだけの stop はその色。stop が無ければ塗りなし', () => {
    expect(fillOf(gradientFillXml([[0.3, 'rgb="FF00FF00"']]))?.rgb).toBe(0x00ff00);
    expect(
      fillOf(
        gradientFillXml([
          [0.8, 'rgb="FF00FF00"'],
          [0.9, 'rgb="FF0000FF"'],
        ]),
      )?.rgb,
    ).toBe(0x00ff00);
    expect(fillOf('<fill><gradientFill/></fill>')).toBeUndefined();
  });

  it('stop の色の theme も解く', () => {
    const f = fillOf(gradientFillXml([[0, 'theme="4"']]));
    expect(f?.rgb).toBe(0x4472c4);
  });
});

/** 先頭の font（既定。sz 11）と inner の font を持つスタイルを読み、xf 1（font 1）の書式を返す */
function fontFmt(
  inner: string,
  base = '<sz val="11"/><name val="Calibri"/>',
): CellFormat | undefined {
  const t = parse(
    styles({
      fonts: [fontXml(base), fontXml(inner)],
      xfs: [
        [0, 0],
        [1, 0],
      ],
    }),
  );
  return fmt(t, 1);
}

describe('parseCellFormats：文字の装飾', () => {
  it('<b>・<i>・<strike> は val が無ければ真。val="0"・"false" は偽（書式なし）', () => {
    expect(fontFmt('<b/>')).toEqual({ bold: true });
    expect(fontFmt('<i/>')).toEqual({ italic: true });
    expect(fontFmt('<strike/>')).toEqual({ strike: true });
    expect(fontFmt('<b val="1"/>')).toEqual({ bold: true });
    expect(fontFmt('<b val="true"/>')).toEqual({ bold: true });
    expect(fontFmt('<b val="0"/><i val="false"/><strike val="0"/>')).toBeUndefined();
  });

  it('<u>：val 無し・single・singleAccounting は single、double・doubleAccounting は double、none は無し', () => {
    expect(fontFmt('<u/>')).toEqual({ underline: 'single' });
    expect(fontFmt('<u val="single"/>')).toEqual({ underline: 'single' });
    expect(fontFmt('<u val="singleAccounting"/>')).toEqual({ underline: 'single' });
    expect(fontFmt('<u val="double"/>')).toEqual({ underline: 'double' });
    expect(fontFmt('<u val="doubleAccounting"/>')).toEqual({ underline: 'double' });
    expect(fontFmt('<u val="none"/>')).toBeUndefined();
    expect(fontFmt('<u val="bogus"/>')).toBeUndefined();
  });

  it('5 種を同時に持てる（太字・斜体・取り消し線・下線・色）', () => {
    expect(fontFmt('<b/><i/><strike/><u val="double"/><color rgb="FFFF0000"/>')).toEqual({
      bold: true,
      italic: true,
      strike: true,
      underline: 'double',
      color: { rgb: 0xff0000, alpha: 1 },
    });
  });

  it('文字サイズは先頭の font との比（0.01 単位）。同じなら持たない。0.5〜4 に丸める', () => {
    expect(fontFmt('<sz val="11"/>')).toBeUndefined();
    expect(fontFmt('<sz val="22"/>')).toEqual({ sizeRatio: 2 });
    expect(fontFmt('<sz val="16.5"/>')).toEqual({ sizeRatio: 1.5 });
    expect(fontFmt('<sz val="12"/>')).toEqual({ sizeRatio: 1.09 });
    expect(fontFmt('<sz val="1"/>')).toEqual({ sizeRatio: 0.5 });
    expect(fontFmt('<sz val="100"/>')).toEqual({ sizeRatio: 4 });
  });

  it('sz が無い・数でない・0 以下の font、先頭の font に sz が無いときは比を持たない', () => {
    expect(fontFmt('<b/>')).toEqual({ bold: true });
    expect(fontFmt('<b/><sz val="x"/>')).toEqual({ bold: true });
    expect(fontFmt('<b/><sz val="0"/>')).toEqual({ bold: true });
    expect(fontFmt('<b/><sz val="20"/>', '<name val="Calibri"/>')).toEqual({ bold: true });
  });

  it('文字の色：明示色は持つ。auto・indexed 64・不透明な #000000・theme="1"（dk1 が黒）は自動（color なし）', () => {
    expect(fontFmt('<color rgb="FF0000FF"/>')).toEqual({ color: { rgb: 0x0000ff, alpha: 1 } });
    expect(fontFmt('<color theme="4"/>')).toEqual({ color: { rgb: 0x4472c4, alpha: 1 } });
    expect(fontFmt('<color rgb="FF000001"/>')).toEqual({ color: { rgb: 1, alpha: 1 } });
    expect(fontFmt('<color auto="1"/>')).toBeUndefined();
    expect(fontFmt('<color indexed="64"/>')).toBeUndefined();
    expect(fontFmt('<color rgb="FF000000"/>')).toBeUndefined();
    expect(fontFmt('<color theme="1"/>')).toBeUndefined();
    expect(fontFmt('<color rgb="zz"/>')).toBeUndefined();
  });

  it('フォント名だけ違う font は書式なし（name は読まない）', () => {
    expect(fontFmt('<sz val="11"/><name val="Arial"/><family val="2"/>')).toBeUndefined();
  });

  it('範囲外・不正な fontId は書式なし。<font/>（自己閉じ）も位置を数える', () => {
    const t = parse(
      styles({
        fonts: ['<font/>', fontXml('<b/>')],
        xfs: [
          [1, 0],
          [9, 0],
          [-1, 0],
        ],
      }),
    );
    expect(fmt(t, 0)).toEqual({ bold: true });
    expect(fmt(t, 1)).toBeUndefined();
    expect(fmt(t, 2)).toBeUndefined();
  });
});

describe('parseCellFormats：書式の表（xfFormat・重複・上限）', () => {
  const RED = patternFillXml('solid', 'rgb="FFFF0000"');

  it('xf 0（装飾も塗りも無い）は -1。cellStyleXfs の <xf> は数えない', () => {
    const t = parse(
      styles({
        fonts: [fontXml('<sz val="11"/>'), fontXml('<b/>')],
        fills: [patternFillXml('none'), RED],
        xfs: [
          [0, 0],
          [1, 1],
        ],
      }),
    );
    // cellStyleXfs の xf は fontId=1・fillId=1（書式が付く）だが、cellXfs の 2 件だけを数える
    expect(t.xfFormat).toHaveLength(2);
    expect(t.xfFormat[0]).toBe(-1);
    expect(fmt(t, 1)).toEqual({ fill: { rgb: 0xff0000, alpha: 1 }, bold: true });
  });

  it('塗りだけ・装飾だけの xf も書式になる。見た目が同じ書式は 1 つにまとまり、番号を共有する', () => {
    const t = parse(
      styles({
        fonts: [
          fontXml('<sz val="11"/>'),
          fontXml('<b/>'),
          fontXml('<b/><sz val="11"/><name val="Arial"/>'),
        ],
        fills: [patternFillXml('none'), RED, patternFillXml('solid', 'rgb="00FF0000"')],
        xfs: [
          [0, 1],
          [1, 0],
          [2, 0],
          [0, 2],
          [1, 1],
        ],
      }),
    );
    // xf 1 と 2（太字だけ）、xf 0 と 3（赤の塗りだけ。アルファの違いは無視）は同じ書式
    expect(t.formats).toHaveLength(3);
    expect(t.xfFormat[1]).toBe(t.xfFormat[2]);
    expect(t.xfFormat[0]).toBe(t.xfFormat[3]);
    expect(new Set(Array.from(t.xfFormat)).size).toBe(3);
  });

  it('fills・fonts が無い styles.xml や cellXfs が無い styles.xml は空の表', () => {
    expect(parse('<styleSheet/>')).toBe(EMPTY_FORMAT_TABLE);
    const noFills = parse(
      styles({
        xfs: [
          [0, 0],
          [0, 1],
        ],
      }),
    );
    expect(Array.from(noFills.xfFormat)).toEqual([-1, -1]);
    expect(noFills.formats).toHaveLength(0);
  });

  it('maxFills を超えた位置の塗りを参照する xf は塗りなし（maxFonts も同様に装飾なし）', () => {
    const fills = [patternFillXml('none'), RED, RED, patternFillXml('solid', 'rgb="FF00FF00"')];
    const t = parse(
      styles({
        fills,
        xfs: [
          [0, 1],
          [0, 2],
          [0, 3],
        ],
      }),
      { ...DEFAULT_LIMITS, maxFills: 3 },
    );
    expect(fmt(t, 0)?.fill?.rgb).toBe(0xff0000);
    expect(fmt(t, 2)).toBeUndefined();
    const fonts = [fontXml('<sz val="11"/>'), fontXml('<b/>'), fontXml('<i/>')];
    const f = parse(
      styles({
        fonts,
        xfs: [
          [1, 0],
          [2, 0],
        ],
      }),
      { ...DEFAULT_LIMITS, maxFonts: 2 },
    );
    expect(fmt(f, 0)).toEqual({ bold: true });
    expect(fmt(f, 1)).toBeUndefined();
  });

  it('maxFills は空の fill も数える', () => {
    const fills = [patternFillXml('none'), patternFillXml('none'), RED];
    const t = parse(styles({ fills, xfs: [[0, 2]] }), { ...DEFAULT_LIMITS, maxFills: 2 });
    expect(fmt(t, 0)).toBeUndefined();
  });

  it('まとめた後の書式の数が maxCellFormats を超えた書式は書式なし（超える前の書式・重複は残る）', () => {
    const fills = [patternFillXml('none')];
    for (let i = 1; i <= 4; i++)
      fills.push(patternFillXml('solid', `rgb="FF0000${i.toString(16).padStart(2, '0')}"`));
    // xf 0〜3 は別の色、xf 4 は xf 0 と同じ色（重複）
    const t = parse(
      styles({
        fills,
        xfs: [
          [0, 1],
          [0, 2],
          [0, 3],
          [0, 4],
          [0, 1],
        ],
      }),
      {
        ...DEFAULT_LIMITS,
        maxCellFormats: 3,
      },
    );
    expect(t.formats).toHaveLength(3);
    expect(Array.from(t.xfFormat)).toEqual([0, 1, 2, -1, 0]);
  });
});

describe('stylesUseTheme と空の表', () => {
  it('EMPTY_FORMAT_TABLE は空', () => {
    expect(EMPTY_FORMAT_TABLE.formats).toHaveLength(0);
    expect(EMPTY_FORMAT_TABLE.xfFormat).toHaveLength(0);
    expect(parse(null)).toBe(EMPTY_FORMAT_TABLE);
  });

  it('fonts・fills・borders のどれかに theme= があるときだけ true（cellXfs など外は見ない）', () => {
    const plain = styles({
      fonts: [fontXml('<color rgb="FF000000"/>')],
      fills: [patternFillXml('none')],
      xfs: [[0, 0]],
    });
    expect(stylesUseTheme(plain)).toBe(false);
    const inFont = styles({ fonts: [fontXml('<color theme="1"/>')], xfs: [[0, 0]] });
    const inFill = styles({ fills: [patternFillXml('solid', 'theme="4"')], xfs: [[0, 0]] });
    const inBorder =
      '<borders><border><left style="thin"><color theme="1"/></left></border></borders>' + plain;
    expect([stylesUseTheme(inFont), stylesUseTheme(inFill), stylesUseTheme(inBorder)]).toEqual([
      true,
      true,
      true,
    ]);
    const borderAuto =
      '<borders><border><left style="thin"><color auto="1"/></left></border></borders>' + plain;
    expect(stylesUseTheme(borderAuto)).toBe(false);
    const outside = plain.replace(
      '<cellStyleXfs',
      '<dxfs><dxf><font><color theme="1"/></font></dxf></dxfs><cellStyleXfs',
    );
    expect(stylesUseTheme(outside)).toBe(false);
    expect(stylesUseTheme(null)).toBe(false);
    expect(stylesUseTheme('<styleSheet/>')).toBe(false);
  });
});

describe('isPaperFill', () => {
  it('不透明な #FFFFFF だけ true。塗りなし・半透明の白・白に近い色は false', () => {
    expect(isPaperFill({ rgb: 0xffffff, alpha: 1 })).toBe(true);
    expect(isPaperFill({ rgb: 0xffffff, alpha: 0.5 })).toBe(false);
    expect(isPaperFill({ rgb: 0xfefefe, alpha: 1 })).toBe(false);
    expect(isPaperFill(undefined)).toBe(false);
  });
});
