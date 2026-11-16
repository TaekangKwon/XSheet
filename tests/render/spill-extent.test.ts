// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createGeometry, findText } from '../../src/core/index';
import { emptyLayout } from '../../src/core/internal';
import type { GridRow } from '../../src/render/grid';
import { rowNumWidth } from '../../src/render/grid';
import { estimateCellTextWidth, spillReachCols } from '../../src/render/spill-extent';
import type { CellSpec } from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  mount,
  openBook,
  px,
  scrollBody,
  stubClientSize,
  wait,
} from './render-sheet.helpers';

/**
 * 最後の列からのはみ出しの分だけ表を広げる（M7-11・ADR-0008 §4・Moobius の依頼 R8）。
 * 幅は文字数からの見積もりなので jsdom で確かめ、実機の見え方は依頼のサンプルで目視する。
 */

describe('文字の幅の見積もり', () => {
  it('全角 1em・ほかは 0.7em（13px）に余白 8px。大きさの比・太字で広がる', () => {
    expect(estimateCellTextWidth('検索')).toBeCloseTo(2 * 13 + 8);
    expect(estimateCellTextWidth('ab')).toBeCloseTo(1.4 * 13 + 8);
    expect(estimateCellTextWidth('ｱｲ')).toBeCloseTo(1.4 * 13 + 8); // 半角カナは全角に数えない
    expect(estimateCellTextWidth('検索', { sizeRatio: 2 })).toBeCloseTo(2 * 26 + 8);
    expect(estimateCellTextWidth('検索', { bold: true })).toBeCloseTo(2 * 13 * 1.1 + 8);
  });

  it('多めに出す：英字の大文字も実寸より広い', () => {
    // 13px の M は実寸で 11px 前後
    expect(estimateCellTextWidth('MMMMMMMMMM') - 8).toBeGreaterThan(10 * 11 * 0.8);
  });
});

describe('はみ出しが届く列', () => {
  const geo = createGeometry(emptyLayout(), 1_048_576, 16_384); // 列 64px
  const t = (text: string, kind?: 's' | 'n') => ({ text, kind, colSpan: 1, rowSpan: 1 });

  it('値のある最後のセルの文字だけを数える', () => {
    // 列 3 から 10 全角（130 + 8px）→ 128 + 138 = 266px は列 5
    const rows: Pick<GridRow, 'cells'>[] = [
      { cells: [t('a', 's'), undefined, t('あ'.repeat(10), 's')] },
    ];
    expect(spillReachCols(geo, rows, [])).toBe(5);
  });

  it('数値・空文字・結合・結合で隠れるセルははみ出さない', () => {
    const long = 'x'.repeat(100);
    expect(spillReachCols(geo, [{ cells: [t(long, 's'), t('1', 'n')] }], [])).toBe(0);
    expect(spillReachCols(geo, [{ cells: [t(long, 's'), t('', 's')] }], [])).toBe(0);
    expect(
      spillReachCols(
        geo,
        [{ cells: [{ text: long, kind: 's', colSpan: 2, rowSpan: 1 }, null] }],
        [],
      ),
    ).toBe(0);
    expect(spillReachCols(geo, [{ cells: [t(long, 's'), null] }], [])).toBe(0);
  });

  it('行の中の最大。書式の文字の大きさを使う', () => {
    const rows: Pick<GridRow, 'cells' | 'formats'>[] = [
      { cells: [t('あ'.repeat(4), 's')] }, // 52 + 8 = 60px → 列 1
      { cells: [t('あ'.repeat(4), 's')], formats: [0] }, // 104 + 8 = 112px → 列 2
    ];
    expect(spillReachCols(geo, rows, [{ sizeRatio: 2 }])).toBe(2);
    expect(spillReachCols(geo, [], [])).toBe(0);
  });

  it('右寄せ・折り返しは右へ届かない。中央はセルの中心から幅の半分', () => {
    const rows: Pick<GridRow, 'cells' | 'formats'>[] = [
      { cells: [t('あ'.repeat(20), 's')], formats: [0] },
    ];
    expect(spillReachCols(geo, rows, [{ hAlign: 'right' }])).toBe(0);
    expect(spillReachCols(geo, rows, [{ wrap: true }])).toBe(0);
    // 中心 32px から (260 + 8) / 2 = 134px → 166px は列 3
    expect(spillReachCols(geo, rows, [{ hAlign: 'center' }])).toBe(3);
    expect(spillReachCols(geo, rows, [{}])).toBe(5); // 268px は列 5
  });

  it('非表示の列の文字は数えない', () => {
    const g = createGeometry(
      { ...emptyLayout(), cols: [{ min: 1, max: 1, width: 10, hidden: true }] },
      10,
      16_384,
    );
    expect(spillReachCols(g, [{ cells: [t('あ'.repeat(20), 's')] }], [])).toBe(0);
  });
});

let restoreSize: () => void;
beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
});
afterEach(() => {
  cleanupViews();
  restoreSize();
});

/** 60 列（BH 列まで）。BH3・BH1500 に文字（依頼のサンプルの 06_検索 と同じ並び） */
async function lastColumnSheet(withRow3 = true) {
  const rows: CellSpec[][] = Array.from({ length: 1500 }, () => []);
  rows[0] = ['検索'];
  if (withRow3) {
    rows[2] = Array.from({ length: 60 }, (_, i) => (i === 59 ? '右端BH列の検索テスト' : null));
  }
  rows[1] = Array.from({ length: 60 }, (_, i) => (i === 0 ? 'a' : null)); // 列数を 60 にする
  rows[1499] = Array.from({ length: 60 }, (_, i) => (i === 59 ? '右下の検索テスト' : null));
  const book = await openBook({ sheets: [{ name: 'S', rows }] });
  const mounted = await mount(book);
  await wait();
  const { matches } = await findText(await book.openSheet(1), '検索テスト');
  return { ...mounted, matches };
}

const tableWidth = (root: HTMLElement) =>
  px(root.querySelector<HTMLElement>('table.xs-grid')!.style.width);
const NUM = rowNumWidth(1500);

describe('最後の列からのはみ出し（renderSheet）', () => {
  it('描いた行の文字が届く列まで表を広げる', async () => {
    const { root } = await lastColumnSheet();
    // BH3：8 全角＋2 半角 = 9.4em → 122.2 + 8px。BH の左端 3776px から 3906px → 列 62
    expect(tableWidth(root)).toBe(NUM + 62 * 64);
  });

  it('描いていない行へ移動するときに広げ、はみ出した文字の上の判定ははみ出し元', async () => {
    const { root, view, matches } = await lastColumnSheet(false);
    expect(tableWidth(root)).toBe(NUM + 60 * 64);
    view.setHighlights(matches);
    const index = matches.findIndex((m) => m.r === 1500);
    view.focusHighlight(index);
    // BH1500：8 全角 = 104 + 8px。3776px から 3888px → 列 61
    expect(tableWidth(root)).toBe(NUM + 61 * 64);
    const body = root.querySelector<HTMLElement>('.xs-body')!;
    // 表の右端まで動かせる：BI 列（61）の上の点
    scrollBody(root, { left: NUM + 61 * 64 - 800 });
    await wait();
    const x = NUM + 60 * 64 + 10 - body.scrollLeft; // BI 列の左端から 10px
    const y = 24 + (1500 - 1) * 20 + 10 - body.scrollTop;
    expect(view.cellAt(x, y)).toEqual({ r: 1500, c: 60 });
  });

  it('シートを替えると広げた分を戻す', async () => {
    const rows: CellSpec[][] = [
      Array.from({ length: 30 }, (_, i) => (i === 29 ? 'あ'.repeat(30) : null)),
    ];
    const book = await openBook({
      sheets: [
        { name: 'A', rows },
        { name: 'B', rows: [Array.from({ length: 30 }, (_, i) => (i === 29 ? 'z' : null))] },
      ],
    });
    const { root, view } = await mount(book);
    await wait();
    expect(tableWidth(root)).toBeGreaterThan(rowNumWidth(1) + 30 * 64);
    await view.showSheet(2);
    await wait();
    expect(tableWidth(root)).toBe(rowNumWidth(1) + 30 * 64);
  });
});
