import { describe, expect, it } from 'vitest';
import { EMPTY_BORDER_TABLE, parseBorders } from '../../src/core/borders';
import { DEFAULT_LIMITS } from '../../src/core/limits';
import { borderXml } from '../../src/testing/make-xlsx';
import { DEFAULT_THEME } from '../../src/core/theme';
import type { BorderStyle } from '../../src/core/types';

/** styles.xml（罫線と cellXfs だけ）を組む。cellStyleXfs にも `<xf>` を置く */
function styles(borders: string[], xfBorderIds: (number | string)[], extra = ''): string {
  return (
    `<styleSheet><borders count="${borders.length}">${borders.join('')}</borders>` +
    '<cellStyleXfs count="1"><xf numFmtId="0" borderId="3"/></cellStyleXfs>' +
    `<cellXfs count="${xfBorderIds.length}">${xfBorderIds.map((b) => `<xf numFmtId="0" borderId="${b}" applyBorder="0"/>`).join('')}</cellXfs>` +
    `${extra}</styleSheet>`
  );
}

const parse = (xml: string | null, limits = DEFAULT_LIMITS) =>
  parseBorders(xml, DEFAULT_THEME, limits);

const NONE = borderXml({ left: {}, right: {}, top: {}, bottom: {}, diagonal: {} });

describe('parseBorders：線種', () => {
  const ALL: BorderStyle[] = [
    'thin',
    'hair',
    'dotted',
    'dashed',
    'dashDot',
    'dashDotDot',
    'medium',
    'mediumDashed',
    'mediumDashDot',
    'mediumDashDotDot',
    'slantDashDot',
    'thick',
    'double',
  ];

  it('13 線種がそれぞれの BorderStyle になる', () => {
    const borders = ALL.map((s) => borderXml({ bottom: { style: s, color: 'auto="1"' } }));
    const t = parse(
      styles(
        borders,
        ALL.map((_, i) => i),
      ),
    );
    expect(t.borders.map((b) => b.bottom?.style)).toEqual(ALL);
  });

  it('未知・none・style 無しの辺は線なし。全辺が線なしの border は罫線なしでセルに番号を付けない', () => {
    const t = parse(
      styles(
        [
          borderXml({ left: { style: 'none' }, top: { style: 'wavy' }, right: {} }),
          NONE,
          borderXml({ left: { style: 'thin' }, top: { style: 'bogus' } }),
        ],
        [0, 1, 2],
      ),
    );
    expect(t.xfBorder[0]).toBe(-1);
    expect(t.xfBorder[1]).toBe(-1);
    expect(t.borders).toHaveLength(1);
    expect(t.xfBorder[2]).toBe(0);
    expect(t.borders[0]!.left?.style).toBe('thin');
    expect(t.borders[0]!.top).toBeUndefined();
  });

  it('自己閉じの <border/> も罫線なしで、位置は数える（次の border の borderId は 1）', () => {
    const t = parse(styles(['<border/>', borderXml({ top: { style: 'thin' } })], [0, 1]));
    expect(t.xfBorder[0]).toBe(-1);
    expect(t.borders[t.xfBorder[1]!]!.top?.style).toBe('thin');
  });
});

describe('parseBorders：辺・斜線', () => {
  it('left・right・top・bottom・diagonal を読み、start・end は左右になる', () => {
    const t = parse(
      styles(
        [
          borderXml({
            left: { style: 'thin' },
            right: { style: 'medium' },
            top: { style: 'thick' },
            bottom: { style: 'hair' },
          }),
          borderXml({ start: { style: 'dashed' }, end: { style: 'dotted' } }),
        ],
        [0, 1],
      ),
    );
    const a = t.borders[t.xfBorder[0]!]!;
    expect([a.left?.style, a.right?.style, a.top?.style, a.bottom?.style]).toEqual([
      'thin',
      'medium',
      'thick',
      'hair',
    ]);
    const b = t.borders[t.xfBorder[1]!]!;
    expect([b.left?.style, b.right?.style]).toEqual(['dashed', 'dotted']);
  });

  it('left があれば start より優先する', () => {
    const t = parse(
      styles([borderXml({ left: { style: 'thin' }, start: { style: 'thick' } })], [0]),
    );
    expect(t.borders[0]!.left?.style).toBe('thin');
  });

  it('斜線：diagonal の辺と diagonalUp・diagonalDown の向き（1・true が真）', () => {
    const edge = { diagonal: { style: 'thin', color: 'auto="1"' } };
    const t = parse(
      styles(
        [
          borderXml(edge, ' diagonalUp="1"'),
          borderXml(edge, ' diagonalDown="true"'),
          borderXml(edge, ' diagonalUp="1" diagonalDown="1"'),
          borderXml(edge),
        ],
        [0, 1, 2, 3],
      ),
    );
    const d = (i: number) => t.borders[t.xfBorder[i]!]!;
    expect([d(0).diagonalUp, d(0).diagonalDown]).toEqual([true, false]);
    expect([d(1).diagonalUp, d(1).diagonalDown]).toEqual([false, true]);
    expect([d(2).diagonalUp, d(2).diagonalDown]).toEqual([true, true]);
    expect([d(3).diagonalUp, d(3).diagonalDown]).toEqual([false, false]);
    expect(d(0).diagonal?.style).toBe('thin');
  });

  it('向きだけあって diagonal に線が無い border は罫線なし', () => {
    const t = parse(styles([borderXml({ diagonal: {} }, ' diagonalUp="1"')], [0]));
    expect(t.xfBorder[0]).toBe(-1);
    expect(t.borders).toHaveLength(0);
  });

  it('名前空間の接頭辞付きの styles.xml も読む', () => {
    const xml =
      '<x:styleSheet><x:borders count="1"><x:border><x:left style="thin"><x:color auto="1"/></x:left></x:border></x:borders>' +
      '<x:cellXfs count="1"><x:xf borderId="0"/></x:cellXfs></x:styleSheet>';
    const t = parse(xml);
    expect(t.borders[0]!.left).toEqual({ style: 'thin', color: null });
  });
});

describe('parseBorders：色', () => {
  it('auto・色なし・indexed="64" は null、rgb・theme＋tint・indexed を解く', () => {
    const e = (color?: string) => borderXml({ left: { style: 'thin', color } });
    const t = parse(
      styles(
        [
          e('auto="1"'),
          e(undefined),
          e('indexed="64"'),
          e('rgb="FFFF0000"'),
          e('theme="1"'),
          e('theme="0" tint="-0.249977111117893"'),
          e('indexed="2"'),
        ],
        [0, 1, 2, 3, 4, 5, 6],
      ),
    );
    const c = (i: number) => t.borders[t.xfBorder[i]!]!.left!.color;
    expect(c(0)).toBeNull();
    expect(c(1)).toBeNull();
    expect(c(2)).toBeNull();
    expect(c(3)).toEqual({ rgb: 0xff0000, alpha: 1 });
    expect(c(4)).toEqual({ rgb: 0x000000, alpha: 1 });
    expect(c(5)?.rgb).toBe(0xbfbfbf);
    expect(c(6)?.rgb).toBe(0xff0000);
  });

  it('<indexedColors> の上書きと、ブックのテーマを使う', () => {
    const theme = new Map(DEFAULT_THEME);
    theme.set('accent1', 0x102030);
    const xml = styles(
      [
        borderXml({ left: { style: 'thin', color: 'indexed="0"' } }),
        borderXml({ left: { style: 'thin', color: 'theme="4"' } }),
      ],
      [0, 1],
      '<colors><indexedColors><rgbColor rgb="FF0A0B0C"/></indexedColors></colors>',
    );
    const t = parseBorders(xml, theme, DEFAULT_LIMITS);
    expect(t.borders[0]!.left?.color?.rgb).toBe(0x0a0b0c);
    expect(t.borders[1]!.left?.color?.rgb).toBe(0x102030);
  });
});

describe('parseBorders：cellXfs と applyBorder', () => {
  it('cellStyleXfs の <xf> を数えない（cellXfs の位置が xf 番号）', () => {
    const t = parse(styles([NONE, NONE, NONE, borderXml({ top: { style: 'thin' } })], [3, 0]));
    // cellStyleXfs の xf は borderId="3" だが、xfBorder は cellXfs の 2 件だけ
    expect(t.xfBorder).toHaveLength(2);
    expect(t.xfBorder[0]).toBe(0);
    expect(t.xfBorder[1]).toBe(-1);
  });

  it('applyBorder="0" でも borderId で引く（E7 の暫定）', () => {
    const t = parse(styles([borderXml({ top: { style: 'thin' } })], [0]));
    expect(t.xfBorder[0]).toBe(0);
  });

  it('borderId が不正・範囲外・負は罫線なし', () => {
    const t = parse(styles([borderXml({ top: { style: 'thin' } })], [5, -1, 'x', 0.5, '']));
    expect(Array.from(t.xfBorder)).toEqual([-1, -1, -1, -1, -1]);
  });

  it('dxfs の中の <border> は <borders> に数えない', () => {
    const dxf = `<dxfs count="1"><dxf>${borderXml({ left: { style: 'thick' } })}</dxf></dxfs>`;
    const xml = `<styleSheet><borders count="1">${NONE}</borders><cellXfs><xf borderId="0"/></cellXfs>${dxf}</styleSheet>`;
    const t = parse(xml);
    expect(t.borders).toHaveLength(0);
  });

  it('同じ罫線を参照する複数の xf は同じ番号', () => {
    const t = parse(styles([borderXml({ top: { style: 'thin' } })], [0, 0, 0]));
    expect(Array.from(t.xfBorder)).toEqual([0, 0, 0]);
    expect(t.borders).toHaveLength(1);
  });
});

describe('parseBorders：無い・壊れた・上限', () => {
  it('xml が null・<borders> が無い・<cellXfs> が無いなら空の表', () => {
    expect(parse(null)).toEqual(EMPTY_BORDER_TABLE);
    expect(parse('<styleSheet/>').borders).toEqual([]);
    expect(parse(`<styleSheet><borders>${NONE}</borders></styleSheet>`).xfBorder).toHaveLength(0);
    expect(parse('<styleSheet><borders><border><left style="thin">').borders).toEqual([]);
  });

  it('maxBorders を超えた位置の border は読まず、それを参照するセルは罫線なし', () => {
    const b = borderXml({ top: { style: 'thin' } });
    const t = parse(styles([b, b, b], [0, 1, 2]), { ...DEFAULT_LIMITS, maxBorders: 2 });
    expect(t.borders).toHaveLength(2);
    expect(Array.from(t.xfBorder)).toEqual([0, 1, -1]);
  });

  it('maxBorders は空の border も数える', () => {
    const b = borderXml({ top: { style: 'thin' } });
    const t = parse(styles([NONE, NONE, b], [0, 1, 2]), { ...DEFAULT_LIMITS, maxBorders: 2 });
    expect(t.borders).toHaveLength(0);
    expect(Array.from(t.xfBorder)).toEqual([-1, -1, -1]);
  });
});
