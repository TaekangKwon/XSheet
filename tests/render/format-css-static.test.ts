import { describe, expect, it } from 'vitest';
import { DEFAULTS, css, declaredValues, rule } from './style-css.helpers';

/**
 * 書式の描画に関わる `style.css` の規則（InterHub ISSUE-100・ADR-0030 §決定 9）。
 * 正: InterHub xlsx-format-css-static.test.ts（TD-100-19）。
 * 書式ごとの規則（`td.xs-xf-<番号>`）は実行時に `<style>` へ作るため、ここには無い（変数も増やさない）。
 */

describe('TD-100-19 is-over-fill の規則', () => {
  it('塗りが絡むはみ出しの文字の span は背景を持たない（透明。面の色でグリッド線・塗りを隠さない）', () => {
    const r = rule('.xs-cell.is-spill.is-over-fill .xs-text');
    expect(r).toContain('background: transparent');
  });

  it('is-over-fill の規則は通常のはみ出しの規則（.is-spill .xs-text）より後ろにあり、詳細度も強い', () => {
    const normal = css.indexOf('\n.xs-cell.is-spill .xs-text {');
    const over = css.indexOf('\n.xs-cell.is-spill.is-over-fill .xs-text {');
    expect(normal).toBeGreaterThanOrEqual(0);
    expect(over).toBeGreaterThan(normal);
  });

  it('書式のために変数を増やさない（--xs-cell-bg の定義は既定値のブロックと、ダークのテーマ（dark・auto）だけ）', () => {
    expect(declaredValues('--xs-cell-bg')).toEqual(['#f5f6f8', '#0f1115', '#0f1115']);
    expect(rule(DEFAULTS)).toContain('--xs-cell-bg:');
    expect(css).not.toMatch(/--xs-(fill|fmt|format)/);
  });
});
