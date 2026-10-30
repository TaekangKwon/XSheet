import { describe, expect, it } from 'vitest';
import { DEFAULTS, declaredValues, rule } from './style-css.helpers';

/**
 * 罫線の層の CSS（InterHub ISSUE-98・ADR-0029 §決定 1・6・8）：変数と重ね順の規則。
 * 正: InterHub xlsx-borders-css.test.ts（TD-98-12）。ダークの墨色（明るい灰色）は使う側が上書きする（examples に例を置く）。
 */

describe('TD-98-12 罫線の層の CSS：変数・重ね順・色', () => {
  it('墨色の変数 --xs-border-ink は明るい既定値が黒で、ほかはダーク（dark・auto）の明るい灰色だけ', () => {
    expect(declaredValues('--xs-border-ink')).toEqual(['#000000', '#d4d4d4', '#d4d4d4']);
    expect(rule(DEFAULTS)).toContain('--xs-border-ink: #000000');
  });

  it('罫線の svg は絶対配置・z-index: 0・クリックを受けない・はみ出しても切らない', () => {
    const r = rule('.xs-borders');
    expect(r).toContain('position: absolute');
    expect(r).toContain('z-index: 0');
    expect(r).toContain('pointer-events: none');
    expect(r).toContain('overflow: visible');
  });

  it('墨色の線は変数、二重線の隙間はセルの面の色の変数だけ（色のリテラルを書かない）', () => {
    expect(rule('.xs-border-path.is-ink')).toContain('stroke: var(--xs-border-ink)');
    expect(rule('.xs-border-path.is-gap')).toContain('stroke: var(--xs-cell-bg)');
    for (const sel of ['.xs-border-path', '.xs-border-path.is-ink', '.xs-border-path.is-gap']) {
      expect(rule(sel)).not.toMatch(/#[0-9a-f]{3,8}\b|rgb/i);
    }
    expect(rule('.xs-border-path')).toContain('fill: none');
  });

  it('罫線の層は図の層より先（下）。ともに z-index: 0 で DOM の順に重なる', () => {
    expect(rule('.xs-layer')).toContain('z-index: 0');
  });
});
