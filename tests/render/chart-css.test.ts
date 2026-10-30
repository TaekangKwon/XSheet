import { describe, expect, it } from 'vitest';
import { css, declaredValues, rule } from './style-css.helpers';

/**
 * グラフの CSS（InterHub ISSUE-103・ADR-0031 §決定 10）：jsdom は CSS を適用しないため、`style.css` の文面で確かめる（実際の見た目は目視＝TV-45）。
 * 正: InterHub xlsx-chart-css.test.ts（TD-103-15）。InterHub はテーマの 4 つのブロックに値を置いたが、`style.css` は明るい既定値を `:where()` に、
 * ダークの値をテーマの属性 `data-xs-theme`（M6-2）に持つ（backdrop-css.test.ts と同じ）。
 */

describe('TD-103-15 グラフの CSS（style.css の文面）', () => {
  it('svg は枠いっぱい。下地は図形と同じ svg.xs-backdrop（最背面・既存のテーマ変数だけ・色のリテラル／新しい変数なし）で、グラフ専用の下地の規則は無い', () => {
    const svg = rule('.xs-chart');
    expect(svg).toContain('width: 100%');
    expect(svg).toContain('height: 100%');
    // 下地は図形と共通の仕組み（z-index: -1・表示と色はテーマ変数）。グラフの svg の中の下地用クラスは無い
    expect(css).not.toContain('.xs-chart-backdrop');
    const back = rule('.xs-backdrop');
    expect(back).toContain('z-index: -1');
    expect(back).toContain('display: var(--xs-backdrop-display, block)');
    expect(rule('.xs-backdrop-part.is-fill')).toContain('fill: var(--xs-backdrop)');
    expect(back).not.toMatch(/#[0-9a-f]{3,8}\b|rgb/i);
    // テーマ変数の定義は増えていない（明るい既定値・ダーク・auto の 3 つ）
    expect(declaredValues('--xs-backdrop')).toHaveLength(3);
  });
});
