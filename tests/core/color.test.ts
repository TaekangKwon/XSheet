import { describe, expect, it } from 'vitest';
import {
  PRESET_COLOR_COUNT,
  applyFillMode,
  findColorElement,
  fromLinear,
  lookupPresetColor,
  mixColors,
  resolveColor,
  toLinear,
  type Rgba,
} from '../../src/core/color';
import { MAX_COLOR_TRANSFORMS } from '../../src/core/limits';
import { DEFAULT_THEME, type Theme } from '../../src/core/theme';

/** 色の要素 1 つ（子の変換つき）を解いて `#rrggbb`（解けなければ null）にする */
const hex = (xml: string, theme: Theme = DEFAULT_THEME): string | null => {
  const c = resolveColor(xml, theme);
  return c === null ? null : c.rgb.toString(16).toUpperCase().padStart(6, '0');
};
const alphaOf = (xml: string): number | null => resolveColor(xml, DEFAULT_THEME)?.alpha ?? null;
const rgba = (rgb: number, alpha = 1): Rgba => ({ rgb, alpha });
/** 16 進の文字列 → 整数 */
const h = (s: string): number => parseInt(s, 16);
/** 整数 → 大文字 6 桁の 16 進 */
const x6 = (n: number): string => n.toString(16).toUpperCase().padStart(6, '0');

describe('TD-94-1 色の要素 6 種', () => {
  describe('srgbClr', () => {
    it('16 進 6 桁（大小文字どちらも）を整数にし、不透明度は 1', () => {
      expect(resolveColor('<a:srgbClr val="FF8000"/>', DEFAULT_THEME)).toEqual({
        rgb: 0xff8000,
        alpha: 1,
      });
      expect(hex('<a:srgbClr val="abcdef"/>')).toBe('ABCDEF');
      expect(hex('<a:srgbClr val="AbCdEf"/>')).toBe('ABCDEF');
      expect(hex('<a:srgbClr val="000000"/>')).toBe('000000');
    });

    it('5 桁・7 桁・16 進以外・# 付き・空・属性なしは解けない（null）', () => {
      for (const v of ['FF800', 'FF80000', 'GG0000', '#FF8000', '', ' FF8000', '0xFF8000'])
        expect(hex(`<a:srgbClr val="${v}"/>`), v).toBeNull();
      expect(hex('<a:srgbClr/>')).toBeNull();
    });
  });

  describe('schemeClr', () => {
    it('テーマの 12 色を引き、bg1→lt1・tx1→dk1・bg2→lt2・tx2→dk2 と読み替える', () => {
      const want: Record<string, string> = {
        dk1: '000000',
        lt1: 'FFFFFF',
        dk2: '44546A',
        lt2: 'E7E6E6',
        accent1: '4472C4',
        accent2: 'ED7D31',
        accent3: 'A5A5A5',
        accent4: 'FFC000',
        accent5: '5B9BD5',
        accent6: '70AD47',
        hlink: '0563C1',
        folHlink: '954F72',
        bg1: 'FFFFFF',
        tx1: '000000',
        bg2: 'E7E6E6',
        tx2: '44546A',
      };
      for (const [name, v] of Object.entries(want))
        expect(hex(`<a:schemeClr val="${name}"/>`), name).toBe(v);
    });

    it('渡したテーマの値を使う', () => {
      const theme: Theme = new Map([...DEFAULT_THEME, ['accent1', 0x123456]]);
      expect(hex('<a:schemeClr val="accent1"/>', theme)).toBe('123456');
      expect(hex('<a:schemeClr val="tx1"/>', theme)).toBe('000000');
    });

    it('phClr・未知の名前・__proto__・constructor・大文字違い・属性なしは解けない', () => {
      for (const v of ['phClr', 'foo', '__proto__', 'constructor', 'toString', 'ACCENT1', ''])
        expect(hex(`<a:schemeClr val="${v}"/>`), v).toBeNull();
      expect(hex('<a:schemeClr/>')).toBeNull();
    });
  });

  describe('sysClr', () => {
    it('lastClr（16 進 6 桁）を使う', () => {
      expect(hex('<a:sysClr val="windowText" lastClr="123456"/>')).toBe('123456');
      expect(hex('<a:sysClr val="menu" lastClr="abcdef"/>')).toBe('ABCDEF');
    });

    it('lastClr が無ければ windowText は黒・window は白、その他は解けない', () => {
      expect(hex('<a:sysClr val="windowText"/>')).toBe('000000');
      expect(hex('<a:sysClr val="window"/>')).toBe('FFFFFF');
      expect(hex('<a:sysClr val="menu"/>')).toBeNull();
      expect(hex('<a:sysClr/>')).toBeNull();
    });

    it('lastClr が不正なら無いものとして val で決める', () => {
      expect(hex('<a:sysClr val="window" lastClr="zzzzzz"/>')).toBe('FFFFFF');
      expect(hex('<a:sysClr val="menu" lastClr="12345"/>')).toBeNull();
    });
  });

  describe('scrgbClr', () => {
    it('線形 RGB（1/1000 %）を sRGB に戻す', () => {
      expect(hex('<a:scrgbClr r="0" g="0" b="0"/>')).toBe('000000');
      expect(hex('<a:scrgbClr r="100000" g="100000" b="100000"/>')).toBe('FFFFFF');
      // 線形 50% は sRGB で #BC（188）
      expect(hex('<a:scrgbClr r="50000" g="50000" b="50000"/>')).toBe('BCBCBC');
      expect(hex('<a:scrgbClr r="100000" g="0" b="0"/>')).toBe('FF0000');
    });

    it('範囲外は 0〜100% に収め、欠けた・整数でない成分は解けない', () => {
      expect(hex('<a:scrgbClr r="200000" g="-5000" b="0"/>')).toBe('FF0000');
      expect(hex('<a:scrgbClr r="0" g="0"/>')).toBeNull();
      expect(hex('<a:scrgbClr r="a" g="0" b="0"/>')).toBeNull();
      expect(hex('<a:scrgbClr r="1.5" g="0" b="0"/>')).toBeNull();
    });
  });

  describe('hslClr', () => {
    it('色相（1/60000 度）・彩度・明るさ（1/1000 %）から RGB にする', () => {
      expect(hex('<a:hslClr hue="0" sat="100000" lum="50000"/>')).toBe('FF0000');
      expect(hex('<a:hslClr hue="7200000" sat="100000" lum="50000"/>')).toBe('00FF00');
      expect(hex('<a:hslClr hue="14400000" sat="100000" lum="50000"/>')).toBe('0000FF');
      expect(hex('<a:hslClr hue="0" sat="0" lum="50000"/>')).toBe('808080');
      expect(hex('<a:hslClr hue="0" sat="100000" lum="0"/>')).toBe('000000');
      expect(hex('<a:hslClr hue="0" sat="100000" lum="100000"/>')).toBe('FFFFFF');
    });

    it('色相の 360°（21600000）は 0° と同じ・範囲外の彩度と明るさは収める', () => {
      expect(hex('<a:hslClr hue="21600000" sat="100000" lum="50000"/>')).toBe('FF0000');
      expect(hex('<a:hslClr hue="0" sat="900000" lum="50000"/>')).toBe('FF0000');
      expect(hex('<a:hslClr hue="0" sat="100000" lum="-1"/>')).toBe('000000');
    });

    it('欠けた・整数でない成分は解けない', () => {
      expect(hex('<a:hslClr hue="0" sat="100000"/>')).toBeNull();
      expect(hex('<a:hslClr hue="x" sat="1" lum="1"/>')).toBeNull();
    });
  });

  describe('prstClr', () => {
    it('仕様の表で引く（約 140 色）', () => {
      expect(hex('<a:prstClr val="red"/>')).toBe('FF0000');
      expect(hex('<a:prstClr val="white"/>')).toBe('FFFFFF');
      expect(hex('<a:prstClr val="black"/>')).toBe('000000');
      expect(hex('<a:prstClr val="green"/>')).toBe('008000');
      expect(hex('<a:prstClr val="lime"/>')).toBe('00FF00');
      expect(hex('<a:prstClr val="gray"/>')).toBe('808080');
      expect(hex('<a:prstClr val="dkBlue"/>')).toBe('00008B');
      expect(hex('<a:prstClr val="ltGray"/>')).toBe('D3D3D3');
      expect(hex('<a:prstClr val="medSeaGreen"/>')).toBe('3CB371');
      expect(hex('<a:prstClr val="yellowGreen"/>')).toBe('9ACD32');
      expect(PRESET_COLOR_COUNT).toBeGreaterThanOrEqual(140);
    });

    it('dk/lt/med は dark/light/medium の別名でも、Gray は Grey の綴りでも引ける', () => {
      expect(hex('<a:prstClr val="darkBlue"/>')).toBe('00008B');
      expect(hex('<a:prstClr val="lightGray"/>')).toBe('D3D3D3');
      expect(hex('<a:prstClr val="lightGrey"/>')).toBe('D3D3D3');
      expect(hex('<a:prstClr val="mediumSeaGreen"/>')).toBe('3CB371');
      expect(hex('<a:prstClr val="dimGrey"/>')).toBe('696969');
    });

    it('未知の名前・__proto__・constructor・大文字違い・属性なしは解けない', () => {
      for (const v of ['nothing', '__proto__', 'constructor', 'RED', ''])
        expect(hex(`<a:prstClr val="${v}"/>`), v).toBeNull();
      expect(hex('<a:prstClr/>')).toBeNull();
      expect(lookupPresetColor('hasOwnProperty')).toBeNull();
    });
  });

  describe('要素の見つけ方', () => {
    it('塗りの要素の中の色を取る（接頭辞あり・なし・別の接頭辞）', () => {
      expect(hex('<a:solidFill><a:srgbClr val="112233"/></a:solidFill>')).toBe('112233');
      expect(hex('<solidFill><srgbClr val="112233"/></solidFill>')).toBe('112233');
      expect(hex('<x:solidFill><x:schemeClr val="accent2"/></x:solidFill>')).toBe('ED7D31');
    });

    it('色の要素が無い・空・色でない要素だけなら null', () => {
      expect(resolveColor('', DEFAULT_THEME)).toBeNull();
      expect(resolveColor('<a:solidFill/>', DEFAULT_THEME)).toBeNull();
      expect(resolveColor('<a:srgbClrX val="112233"/>', DEFAULT_THEME)).toBeNull();
    });

    it('最初に見つけた色の要素だけを解く（解けなくても次の要素へは進まない）', () => {
      expect(hex('<a:schemeClr val="phClr"/><a:srgbClr val="112233"/>')).toBeNull();
      expect(hex('<a:srgbClr val="112233"/><a:srgbClr val="445566"/>')).toBe('112233');
    });

    it('閉じタグの無い要素でも、子の変換は読まずに色だけを返す（例外にしない）', () => {
      expect(hex('<a:srgbClr val="112233"><a:alpha val="50000"/>')).toBe('112233');
      expect(findColorElement('<a:srgbClr val="1">')?.body).toBeNull();
    });
  });
});

describe('TD-94-2 変換', () => {
  const ACC1 = (inner: string) => `<a:schemeClr val="accent1">${inner}</a:schemeClr>`;

  it('lumMod は HSL の明るさに掛ける（accent1 の「濃く 25%・50%」）', () => {
    expect(hex(ACC1('<a:lumMod val="75000"/>'))).toBe('2F5597');
    // PowerPoint の表示は #1F3864。HSL の明るさを丸めずに計算するため、ここでは 1 ずれて #203864（U7 で Excel の画素と照合する）
    expect(hex(ACC1('<a:lumMod val="50000"/>'))).toBe('203864');
  });

  it('lumMod＋lumOff は出現順に掛ける（accent1 の「薄く 40%・60%・80%」）', () => {
    expect(hex(ACC1('<a:lumMod val="60000"/><a:lumOff val="40000"/>'))).toBe('8FAADC');
    expect(hex(ACC1('<a:lumMod val="40000"/><a:lumOff val="60000"/>'))).toBe('B4C7E7');
    expect(hex(ACC1('<a:lumMod val="20000"/><a:lumOff val="80000"/>'))).toBe('DAE3F3');
  });

  it('順序が入れ替わると結果が変わる（lumOff してから lumMod）', () => {
    const a = hex(ACC1('<a:lumMod val="60000"/><a:lumOff val="40000"/>'));
    const b = hex(ACC1('<a:lumOff val="40000"/><a:lumMod val="60000"/>'));
    expect(a).not.toBe(b);
  });

  it('lum は明るさを置き換え、lumOff は負の値で下げ、収める（0〜1）', () => {
    expect(hex('<a:srgbClr val="FF0000"><a:lum val="100000"/></a:srgbClr>')).toBe('FFFFFF');
    expect(hex('<a:srgbClr val="FF0000"><a:lum val="0"/></a:srgbClr>')).toBe('000000');
    expect(hex('<a:srgbClr val="FF0000"><a:lumOff val="-100000"/></a:srgbClr>')).toBe('000000');
    expect(hex('<a:srgbClr val="808080"><a:lumOff val="900000"/></a:srgbClr>')).toBe('FFFFFF');
  });

  it('satMod・satOff・sat は HSL の彩度', () => {
    // 4472C4 の彩度 0.5 → satMod 0 で灰色（明るさは変わらない）
    expect(hex(ACC1('<a:satMod val="0"/>'))).toBe('848484');
    expect(hex('<a:srgbClr val="BF4040"><a:sat val="100000"/></a:srgbClr>')).toBe('FF0000');
    expect(hex('<a:srgbClr val="FF0000"><a:satOff val="-100000"/></a:srgbClr>')).toBe('808080');
    // 彩度の上限（1）を超えても収める
    expect(hex('<a:srgbClr val="FF0000"><a:satMod val="500000"/></a:srgbClr>')).toBe('FF0000');
  });

  it('hue・hueMod・hueOff は 360° で回る', () => {
    const red = (inner: string) => `<a:srgbClr val="FF0000">${inner}</a:srgbClr>`;
    // 赤（0°）＋120° ＝ 緑、＋240° ＝ 青、＋360° ＝ 赤
    expect(hex(red('<a:hueOff val="7200000"/>'))).toBe('00FF00');
    expect(hex(red('<a:hueOff val="14400000"/>'))).toBe('0000FF');
    expect(hex(red('<a:hueOff val="21600000"/>'))).toBe('FF0000');
    // 負の値は 360° の手前へ回る（0° − 120° ＝ 240° ＝ 青）
    expect(hex(red('<a:hueOff val="-7200000"/>'))).toBe('0000FF');
    // 350° ＋ 20° ＝ 10°
    expect(
      hex('<a:hslClr hue="21000000" sat="100000" lum="50000"><a:hueOff val="1200000"/></a:hslClr>'),
    ).toBe(hex('<a:hslClr hue="600000" sat="100000" lum="50000"/>'));
    expect(hex(red('<a:hue val="14400000"/>'))).toBe('0000FF');
    // 60° × 2 ＝ 120°
    expect(
      hex('<a:hslClr hue="3600000" sat="100000" lum="50000"><a:hueMod val="200000"/></a:hslClr>'),
    ).toBe('00FF00');
  });

  it('shade は線形 RGB で黒へ寄せる：accent1＋shade 15000 が #172C51、lt1＋shade 50000 が #BCBCBC', () => {
    // #172C51・#BCBCBC は ADR-0028 §決定 6.3 の計算値で、Excel の画素では未確認（U7・ISSUE-95 で照合）
    expect(hex('<a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr>')).toBe('172C51');
    expect(hex('<a:schemeClr val="lt1"><a:shade val="50000"/></a:schemeClr>')).toBe('BCBCBC');
    expect(hex('<a:srgbClr val="FFFFFF"><a:shade val="0"/></a:srgbClr>')).toBe('000000');
    expect(hex('<a:srgbClr val="4472C4"><a:shade val="100000"/></a:srgbClr>')).toBe('4472C4');
  });

  it('Excel の実測と一致する値（既定の図形の枠線・「薄く／濃く」の色）', () => {
    const t = (name: string, inner: string) =>
      hex(`<a:schemeClr val="${name}">${inner}</a:schemeClr>`);
    // 既定の図形の枠線（accent1＋shade 50000）
    expect(t('accent1', '<a:shade val="50000"/>')).toBe('2F528F');
    expect(t('accent1', '<a:tint val="40000"/>')).toBe('CFD5EA');
    expect(t('accent1', '<a:lumMod val="60000"/><a:lumOff val="40000"/>')).toBe('8FAADC');
    expect(t('accent2', '<a:lumMod val="75000"/>')).toBe('C55A11');
    expect(t('accent2', '<a:lumMod val="40000"/><a:lumOff val="60000"/>')).toBe('F8CBAD');
  });

  it('tint は線形 RGB で白へ寄せる', () => {
    expect(hex('<a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr>')).toBe('BCBCBC');
    expect(hex('<a:srgbClr val="FF0000"><a:tint val="50000"/></a:srgbClr>')).toBe('FFBCBC');
    expect(hex('<a:srgbClr val="4472C4"><a:tint val="0"/></a:srgbClr>')).toBe('FFFFFF');
    expect(hex('<a:srgbClr val="4472C4"><a:tint val="100000"/></a:srgbClr>')).toBe('4472C4');
  });

  it('alpha・alphaMod・alphaOff は不透明度（0〜1）', () => {
    expect(alphaOf('<a:srgbClr val="000000"/>')).toBe(1);
    expect(alphaOf('<a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr>')).toBeCloseTo(0.4);
    expect(
      alphaOf(
        '<a:srgbClr val="000000"><a:alpha val="80000"/><a:alphaMod val="50000"/></a:srgbClr>',
      ),
    ).toBeCloseTo(0.4);
    expect(alphaOf('<a:srgbClr val="000000"><a:alphaOff val="-20000"/></a:srgbClr>')).toBeCloseTo(
      0.8,
    );
    expect(alphaOf('<a:srgbClr val="000000"><a:alphaOff val="50000"/></a:srgbClr>')).toBe(1);
    expect(alphaOf('<a:srgbClr val="000000"><a:alpha val="-5"/></a:srgbClr>')).toBe(0);
    // 不透明度は色の値を変えない
    expect(hex('<a:srgbClr val="123456"><a:alpha val="10000"/></a:srgbClr>')).toBe('123456');
  });

  it('無視する変換（comp・inv・gray・gamma・invGamma・red・green・blue の系統）は色をそのままにし、数にも数える', () => {
    for (const t of [
      '<a:comp/>',
      '<a:inv/>',
      '<a:gray/>',
      '<a:gamma/>',
      '<a:invGamma/>',
      '<a:red val="50000"/>',
      '<a:redMod val="50000"/>',
      '<a:greenOff val="50000"/>',
      '<a:blue val="0"/>',
      '<a:unknownTransform val="1"/>',
    ])
      expect(hex(`<a:srgbClr val="4472C4">${t}</a:srgbClr>`), t).toBe('4472C4');
    // 無視する変換が間にあっても、他の変換は掛かる
    expect(hex(ACC1('<a:comp/><a:lumMod val="75000"/>'))).toBe('2F5597');
  });

  it('不正な値の変換（整数でない・空・属性なし・巨大）は飛ばし、後ろの変換は掛かる', () => {
    for (const v of ['abc', '', '1.5', '1e3', ' 5', '0x10', '9'.repeat(400)])
      expect(hex(ACC1(`<a:lumMod val="${v}"/><a:lumMod val="75000"/>`)), v).toBe('2F5597');
    expect(hex(ACC1('<a:lumMod/><a:lumMod val="75000"/>'))).toBe('2F5597');
  });

  it('巨大な値でも例外にならず、結果は収まる', () => {
    const big = '9'.repeat(300);
    for (const t of [
      'lumMod',
      'lumOff',
      'satMod',
      'satOff',
      'hueMod',
      'hueOff',
      'tint',
      'shade',
      'alphaMod',
    ]) {
      const c = resolveColor(
        `<a:srgbClr val="4472C4"><a:${t} val="${big}"/></a:srgbClr>`,
        DEFAULT_THEME,
      );
      expect(c, t).not.toBeNull();
      expect(c!.rgb).toBeGreaterThanOrEqual(0);
      expect(c!.rgb).toBeLessThanOrEqual(0xffffff);
      expect(c!.alpha).toBeGreaterThanOrEqual(0);
      expect(c!.alpha).toBeLessThanOrEqual(1);
    }
  });

  it(`1 色の変換は ${MAX_COLOR_TRANSFORMS} 個まで（17 個目以降は読まない）`, () => {
    const half = '<a:alphaMod val="50000"/>';
    const xml = (n: number) => `<a:srgbClr val="000000">${half.repeat(n)}</a:srgbClr>`;
    expect(alphaOf(xml(16))).toBeCloseTo(0.5 ** 16, 12);
    expect(alphaOf(xml(17))).toBeCloseTo(0.5 ** 16, 12);
    expect(alphaOf(xml(100_000))).toBeCloseTo(0.5 ** 16, 12);
    // 無視する変換・不正な値も 1 個と数える
    expect(alphaOf(`<a:srgbClr val="000000">${'<a:comp/>'.repeat(16)}${half}</a:srgbClr>`)).toBe(1);
  });

  it('線形化の式を切り替えられる（γ 2.2 の指定）', () => {
    expect(toLinear(0.5)).toBeCloseTo(0.214, 3);
    expect(toLinear(0.5, 2.2)).toBeCloseTo(0.5 ** 2.2, 10);
    expect(fromLinear(toLinear(0.3), null)).toBeCloseTo(0.3, 10);
    expect(fromLinear(toLinear(0.3, 2.3), 2.3)).toBeCloseTo(0.3, 10);
    // 端の値・範囲外
    expect(toLinear(0)).toBe(0);
    expect(toLinear(1)).toBeCloseTo(1, 12);
    expect(fromLinear(2)).toBeCloseTo(1, 12);
    expect(fromLinear(-1)).toBe(0);
    expect(toLinear(Number.NaN)).toBe(0);
  });
});

describe('TD-94-3 面の濃淡（darken・lighten）と混色', () => {
  const white = rgba(0xffffff);
  const black = rgba(0);
  const mid = rgba(h('4472C4'));

  it('norm は塗りの色のまま。未知の名前もそのまま', () => {
    expect(applyFillMode(mid, 'norm')).toEqual(mid);
    expect(applyFillMode(mid, 'foo')).toEqual(mid);
    expect(applyFillMode(mid, '__proto__')).toEqual(mid);
  });

  it('darken・darkenLess は sRGB の値に 0.6・0.805 を掛け、lighten・lightenLess は白へ 0.6・0.805 寄せる（Excel の面の画素から決めた U8）', () => {
    expect(x6(applyFillMode(white, 'darken').rgb)).toBe('999999');
    expect(x6(applyFillMode(white, 'darkenLess').rgb)).toBe('CDCDCD');
    expect(x6(applyFillMode(black, 'lighten').rgb)).toBe('666666');
    expect(x6(applyFillMode(black, 'lightenLess').rgb)).toBe('323232');
    // Excel 16 が描いた円柱・立方体・額縁の面（ref-colors.xlsx・accent1 #4472C4）。濃淡は線形 RGB の shade・tint とは別の式（sRGB の値のまま）
    expect(x6(applyFillMode(mid, 'darken').rgb)).toBe('294476');
    expect(x6(applyFillMode(mid, 'lighten').rgb)).toBe('8FAADC');
    expect(x6(applyFillMode(mid, 'darkenLess').rgb)).toBe('375C9E');
    expect(x6(applyFillMode(mid, 'lightenLess').rgb)).toBe('688DD0'); // Excel の画素は #698ED0（差 1）
  });

  it('不透明度は変えない', () => {
    expect(applyFillMode(rgba(0xffffff, 0.3), 'darken').alpha).toBe(0.3);
  });

  it('mixColors：割合 0 は a、1 は b、0.5 は中間（不透明度も混ぜる）', () => {
    const a = rgba(0x000000, 1);
    const b = rgba(0xffffff, 0);
    expect(mixColors(a, b, 0)).toEqual(a);
    expect(mixColors(a, b, 1)).toEqual(b);
    expect(mixColors(a, b, 0.5)).toEqual({ rgb: 0x808080, alpha: 0.5 });
    expect(x6(mixColors(rgba(0xff0000), rgba(0x0000ff), 0.5).rgb)).toBe('800080');
  });

  it('mixColors：範囲外は 0〜1 に収め、NaN は a のまま', () => {
    const a = rgba(0x102030);
    const b = rgba(0xa0b0c0);
    expect(mixColors(a, b, -3)).toEqual(a);
    expect(mixColors(a, b, 7)).toEqual(b);
    expect(mixColors(a, b, Number.NaN)).toEqual(a);
    expect(mixColors(a, b, Number.POSITIVE_INFINITY)).toEqual(b);
  });
});

/** かかった時間（ms） */
const elapsed = (fn: () => unknown): number => {
  const t0 = performance.now();
  fn();
  return performance.now() - t0;
};
const LIMIT_MS = 1500;
const N = 40_000;

describe('TD-94-4 細工した入力でも時間が線形', () => {
  it('> の無い開きタグ・閉じの無い要素・大量の子の列でも固まらず例外にならない', () => {
    const xmls = [
      '<a:srgbClr '.repeat(N),
      '<a:srgbClr>'.repeat(N),
      `<a:srgbClr val="000000">${'<a:lumMod '.repeat(N)}`,
      `<a:srgbClr val="000000">${'<a:lumMod>'.repeat(N)}</a:srgbClr>`,
      `<a:srgbClr val="000000">${'<a:lumMod val="50000"/>'.repeat(N)}</a:srgbClr>`,
      `<a:schemeClr val="accent1">${'<a:'.repeat(N)}</a:schemeClr>`,
      `<a:srgbClr val="${'f'.repeat(N)}"/>`,
      `<a:srgbClr ${'a'.repeat(N)}`,
      `<${'a'.repeat(N)}`,
      '<a:solidFill>'.repeat(N) + '<a:prstClr val="red"/>',
    ];
    for (const xml of xmls) {
      let r: Rgba | null = null;
      expect(
        elapsed(() => {
          r = resolveColor(xml, DEFAULT_THEME);
        }),
      ).toBeLessThan(LIMIT_MS);
      void r;
    }
  });

  it('同じ色の要素を数万個並べた入力（先頭の 1 つだけ解く）', () => {
    const xml = '<a:srgbClr val="112233"><a:lumMod val="50000"/></a:srgbClr>'.repeat(N);
    expect(elapsed(() => resolveColor(xml, DEFAULT_THEME))).toBeLessThan(LIMIT_MS);
    expect(hex(xml)).toBe(hex('<a:srgbClr val="112233"><a:lumMod val="50000"/></a:srgbClr>'));
  });
});
