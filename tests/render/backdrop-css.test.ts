import { describe, expect, it } from 'vitest';
import { DEFAULTS, declaredValues, rule } from './style-css.helpers';

/**
 * ダーク表示の下地の CSS（InterHub ISSUE-96・ADR-0028 §決定 6.7）：変数と重ね順の規則。
 * 正: InterHub xlsx-backdrop-css.test.ts（TD-96-6）。InterHub はテーマの 4 つのブロック（ライト・ダーク）に値を置いたが、
 * `style.css` は明るい既定値だけを持つ（ダークは使う側が上書きする。docs/02_architecture.md §6）。
 */

describe('TD-96-6 下地の CSS：変数・z-index・isolation', () => {
  it('変数 --xs-backdrop は既定値のブロックの 1 か所だけにあり、明るい表示は透明', () => {
    expect(declaredValues('--xs-backdrop')).toEqual(['transparent']);
    expect(rule(DEFAULTS)).toContain('--xs-backdrop: transparent');
  });

  it('下地の表示 --xs-backdrop-display も既定値は none（ダークは使う側が block にする。DOM は変えない）', () => {
    expect(declaredValues('--xs-backdrop-display')).toEqual(['none']);
    expect(rule(DEFAULTS)).toContain('--xs-backdrop-display: none');
    expect(rule('.xs-backdrop')).toContain('display: var(--xs-backdrop-display, block)');
  });

  it('層は isolation: isolate で重ね順の文脈を作る', () => {
    expect(rule('.xs-layer')).toContain('isolation: isolate');
  });

  it('下地の svg は絶対配置・z-index: -1・はみ出しても切らない', () => {
    const r = rule('.xs-backdrop');
    expect(r).toContain('position: absolute');
    expect(r).toContain('z-index: -1');
    expect(r).toContain('overflow: visible');
  });

  it('下地の色は変数だけ（塗り・縁取り。色のリテラルを書かない）。図の要素に z-index・transform を付けない', () => {
    expect(rule('.xs-backdrop-part.is-fill')).toContain('fill: var(--xs-backdrop)');
    expect(rule('.xs-backdrop-part.is-stroke')).toContain('stroke: var(--xs-backdrop)');
    expect(rule('.xs-backdrop-part')).not.toMatch(/#[0-9a-f]{3,8}\b|rgb/i);
    expect(rule('.xs-figure')).not.toMatch(/z-index|transform|opacity/);
  });
});
