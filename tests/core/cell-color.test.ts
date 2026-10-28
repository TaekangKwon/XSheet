import { describe, expect, it } from 'vitest';
import { INDEXED_COLORS, parseIndexedColors, resolveCellColor } from '../../src/core/cell-color';
import { DEFAULT_THEME, THEME_SLOTS } from '../../src/core/theme';
import { parseAttrs } from '../../src/core/xml';

/** `<color …/>` の開きタグを解く（既定のテーマ。palette 省略は既定の 64 色）。自動・解けないは null */
const solve = (tag: string, palette: readonly number[] | null = null): number | null => {
  const r = resolveCellColor(parseAttrs(tag), DEFAULT_THEME, palette);
  return r === null ? null : r.rgb;
};

describe('resolveCellColor：自動と属性なし', () => {
  it('auto="1"・属性なし・属性が undefined は null（自動）', () => {
    expect(solve('<color auto="1"/>')).toBeNull();
    expect(solve('<color/>')).toBeNull();
    expect(resolveCellColor(undefined, DEFAULT_THEME, null)).toBeNull();
  });

  it('indexed="64"（システムの前景色）は null、65（背景色）は白', () => {
    expect(solve('<color indexed="64"/>')).toBeNull();
    expect(solve('<color indexed="65"/>')).toBe(0xffffff);
  });
});

describe('resolveCellColor：rgb', () => {
  it('ARGB 8 桁は先頭のアルファを無視し、6 桁はそのまま読む。不透明度は 1', () => {
    expect(solve('<color rgb="FF112233"/>')).toBe(0x112233);
    expect(solve('<color rgb="00AABBCC"/>')).toBe(0xaabbcc);
    expect(solve('<color rgb="445566"/>')).toBe(0x445566);
    const c = resolveCellColor(parseAttrs('<color rgb="FF112233"/>'), DEFAULT_THEME, null);
    expect(c?.alpha).toBe(1);
  });

  it('桁数・文字が不正な rgb は null', () => {
    expect(solve('<color rgb="FF112"/>')).toBeNull();
    expect(solve('<color rgb="FF11223"/>')).toBeNull();
    expect(solve('<color rgb="GG112233"/>')).toBeNull();
    expect(solve('<color rgb="FFGG1122"/>')).toBeNull();
    expect(solve('<color rgb=""/>')).toBeNull();
  });
});

describe('resolveCellColor：theme の番号', () => {
  it('SpreadsheetML の順（0＝lt1・1＝dk1・2＝lt2・3＝dk2・4〜9＝accent・10＝hlink・11＝folHlink）で写す', () => {
    const slotOf = [
      'lt1',
      'dk1',
      'lt2',
      'dk2',
      'accent1',
      'accent2',
      'accent3',
      'accent4',
      'accent5',
      'accent6',
      'hlink',
      'folHlink',
    ] as const;
    for (let i = 0; i < 12; i++) {
      expect(solve(`<color theme="${i}"/>`), `theme=${i}`).toBe(DEFAULT_THEME.get(slotOf[i]!));
    }
    // 0 と 1・2 と 3 は THEME_SLOTS（dk1・lt1・dk2・lt2）の順とは入れ替わっている
    expect(THEME_SLOTS[0]).toBe('dk1');
    expect(solve('<color theme="0"/>')).toBe(0xffffff);
    expect(solve('<color theme="1"/>')).toBe(0x000000);
  });

  it('ブックのテーマの色を使う', () => {
    const theme = new Map(DEFAULT_THEME);
    theme.set('accent2', 0x123456);
    expect(resolveCellColor(parseAttrs('<color theme="5"/>'), theme, null)?.rgb).toBe(0x123456);
  });

  it('範囲外・整数でない・負の theme は null', () => {
    expect(solve('<color theme="12"/>')).toBeNull();
    expect(solve('<color theme="-1"/>')).toBeNull();
    expect(solve('<color theme="1.5"/>')).toBeNull();
    expect(solve('<color theme="x"/>')).toBeNull();
  });
});

describe('resolveCellColor：tint（HLS の明度）', () => {
  it('負は L × (1 + tint)：白に Excel が書く -0.25（-0.2499771…）は BFBFBF、-0.15 は D9D9D9', () => {
    expect(solve('<color theme="0" tint="-0.249977111117893"/>')).toBe(0xbfbfbf);
    expect(solve('<color theme="0" tint="-0.149998474121094"/>')).toBe(0xd9d9d9);
  });

  it('50% の境目（127.5）は 1 段階までの差を許す（Excel の 7F7F7F と四捨五入の 808080。E6 で確かめる）', () => {
    const v = solve('<color theme="0" tint="-0.499984740745262"/>')! & 255;
    expect(Math.abs(v - 0x7f)).toBeLessThanOrEqual(1);
  });

  it('正は L × (1 − tint) + tint：黒に 0.5 は 808080、赤 FF0000 に 0.5 は FF8080', () => {
    expect(solve('<color theme="1" tint="0.5"/>')).toBe(0x808080);
    expect(solve('<color rgb="FFFF0000" tint="0.5"/>')).toBe(0xff8080);
  });

  it('色相・彩度は保つ（accent1 4472C4 を暗くしても青系のまま）', () => {
    const rgb = solve('<color theme="4" tint="-0.5"/>')!;
    const r = (rgb >> 16) & 255;
    const g = (rgb >> 8) & 255;
    const b = rgb & 255;
    expect(b).toBeGreaterThan(g);
    expect(g).toBeGreaterThan(r);
    expect(b).toBeLessThan(0xc4);
  });

  it('tint が 0・数でないは色を変えない。-1 は黒、1 は白に丸める（範囲外も丸める）', () => {
    expect(solve('<color theme="4" tint="0"/>')).toBe(0x4472c4);
    expect(solve('<color theme="4" tint="x"/>')).toBe(0x4472c4);
    expect(solve('<color theme="4" tint="-1"/>')).toBe(0x000000);
    expect(solve('<color theme="4" tint="1"/>')).toBe(0xffffff);
    expect(solve('<color theme="4" tint="-5"/>')).toBe(0x000000);
    expect(solve('<color theme="4" tint="5"/>')).toBe(0xffffff);
  });

  it('自動（null）に tint を付けても null のまま', () => {
    expect(solve('<color auto="1" tint="0.5"/>')).toBeNull();
  });
});

describe('resolveCellColor：indexed（旧 64 色）', () => {
  it('既定の表：0＝黒・1＝白・2＝赤・22＝C0C0C0・63＝333333', () => {
    expect(INDEXED_COLORS).toHaveLength(64);
    expect(solve('<color indexed="0"/>')).toBe(0x000000);
    expect(solve('<color indexed="1"/>')).toBe(0xffffff);
    expect(solve('<color indexed="2"/>')).toBe(0xff0000);
    expect(solve('<color indexed="22"/>')).toBe(0xc0c0c0);
    expect(solve('<color indexed="63"/>')).toBe(0x333333);
  });

  it('範囲外・負・整数でない indexed は null', () => {
    expect(solve('<color indexed="66"/>')).toBeNull();
    expect(solve('<color indexed="-1"/>')).toBeNull();
    expect(solve('<color indexed="2.5"/>')).toBeNull();
    expect(solve('<color indexed="x"/>')).toBeNull();
  });

  it('palette（<indexedColors> の上書き）があればその値。64・65 は上書きされない', () => {
    const palette = parseIndexedColors(
      '<colors><indexedColors><rgbColor rgb="FF010203"/><rgbColor rgb="FF0A0B0C"/></indexedColors></colors>',
    );
    expect(palette).not.toBeNull();
    expect(solve('<color indexed="0"/>', palette)).toBe(0x010203);
    expect(solve('<color indexed="1"/>', palette)).toBe(0x0a0b0c);
    // 上書きの無い位置は既定の表のまま
    expect(solve('<color indexed="2"/>', palette)).toBe(0xff0000);
    expect(solve('<color indexed="64"/>', palette)).toBeNull();
    expect(solve('<color indexed="65"/>', palette)).toBe(0xffffff);
  });
});

describe('parseIndexedColors', () => {
  it('<indexedColors> が無い・空なら null（既定の表を使う）', () => {
    expect(parseIndexedColors('<styleSheet/>')).toBeNull();
    expect(parseIndexedColors('<colors><indexedColors/></colors>')).toBeNull();
    expect(parseIndexedColors('<colors><indexedColors></indexedColors></colors>')).toBeNull();
  });

  it('64 件まで読み、超えた分は読まない。不正な rgb はその位置だけ既定のまま', () => {
    const many = Array.from({ length: 70 }, () => '<rgbColor rgb="FF111111"/>');
    many[3] = '<rgbColor rgb="zz"/>';
    const xml = `<colors><indexedColors>${many.join('')}</indexedColors></colors>`;
    const p = parseIndexedColors(xml)!;
    expect(p).toHaveLength(64);
    expect(p[0]).toBe(0x111111);
    expect(p[3]).toBe(INDEXED_COLORS[3]);
    expect(p[63]).toBe(0x111111);
  });

  it('名前空間の接頭辞付きも読む', () => {
    const p = parseIndexedColors(
      '<x:colors><x:indexedColors><x:rgbColor rgb="FF0000FF"/></x:indexedColors></x:colors>',
    )!;
    expect(p[0]).toBe(0x0000ff);
  });
});
