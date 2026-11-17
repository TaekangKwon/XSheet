// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { findText } from '../../src/core/index';
import { alignClasses, resolveHAlign, spillSides } from '../../src/render/cell-align';
import type { CellSpec } from '../../src/testing/make-xlsx';
import { cleanupViews, mount, openBook, stubClientSize, wait } from './render-sheet.helpers';

/**
 * セルの配置と折り返しの描き方（M7-13・ADR-0008 §2・§3・Moobius の依頼 R9）。
 * クラスと入れ物の幅・位置を jsdom で確かめ、実機の見え方は依頼のサンプルで目視する。
 */

const t = (text: string, kind?: 's' | 'n') => ({ text, kind, colSpan: 1, rowSpan: 1 });

describe('配置のクラス', () => {
  it('標準は型で決め、指定があればそれ', () => {
    expect(resolveHAlign('s')).toBe('left');
    expect(resolveHAlign('n')).toBe('right');
    expect(resolveHAlign('b')).toBe('center');
    expect(resolveHAlign('n', { hAlign: 'left' })).toBe('left');
  });

  it('文字列の中央・右は safe の無いクラス、数値は今のクラス。縦と折り返し', () => {
    expect(alignClasses('s', { hAlign: 'center' })).toEqual(['is-h-center']);
    expect(alignClasses('s', { hAlign: 'right' })).toEqual(['is-h-right']);
    expect(alignClasses('n', { hAlign: 'center' })).toEqual(['is-center']);
    expect(alignClasses('s', { hAlign: 'left' })).toEqual([]);
    expect(alignClasses('s', { vAlign: 'top', wrap: true })).toEqual(['is-v-top', 'is-wrap']);
    expect(alignClasses('n', { vAlign: 'center', wrap: true })).toEqual(['is-v-center']); // 数値は折り返さない
    expect(alignClasses('s')).toEqual([]);
  });
});

describe('はみ出す向き', () => {
  const cells = [
    t('v', 's'),
    undefined,
    undefined,
    t('long', 's'),
    undefined,
    undefined,
    t('w', 's'),
  ];

  it('左寄せ・標準は右だけ、右寄せは左だけ、中央は両側。値のあるセルで止まる', () => {
    expect(spillSides(cells, 3, undefined)).toEqual({ left: 0, right: 2 });
    expect(spillSides(cells, 3, { hAlign: 'right' })).toEqual({ left: 2, right: 0 });
    expect(spillSides(cells, 3, { hAlign: 'center' })).toEqual({ left: 2, right: 2 });
  });

  it('描く範囲で止める。折り返し・数値・結合ははみ出さない', () => {
    expect(spillSides(cells, 3, { hAlign: 'center' }, 2, 5)).toEqual({ left: 1, right: 1 });
    expect(spillSides(cells, 3, { hAlign: 'center', wrap: true })).toEqual({ left: 0, right: 0 });
    expect(spillSides([undefined, t('1', 'n')], 1, { hAlign: 'right' })).toEqual({
      left: 0,
      right: 0,
    });
    expect(
      spillSides([undefined, { text: 'm', kind: 's', colSpan: 2, rowSpan: 1 }, null], 1, {
        hAlign: 'center',
      }),
    ).toEqual({ left: 0, right: 0 });
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

/** xf：0 なし・1 中央・2 右・3 上・4 縦中央・5 折り返し＋上・6 中央＋縦中央＋折り返し */
const ALIGNS = [
  undefined,
  'horizontal="center"',
  'horizontal="right"',
  'vertical="top"',
  'vertical="center"',
  'vertical="top" wrapText="1"',
  'horizontal="center" vertical="center" wrapText="1"',
];
const s = (v: string, xf: number): CellSpec => ({ v, t: 'inlineStr', s: xf });
const n = (v: number, xf: number): CellSpec => ({ v, s: xf });
const LONG = 'とても長い文字列がはみ出す';

async function alignSheet(merges: string[] = []) {
  const rows: CellSpec[][] = [
    [s('x', 0), s('c', 1), s('r', 2), n(5, 1), n(5, 0)],
    [],
    [null, null, null, s(LONG, 2)],
    [s('v', 0), null, s(LONG, 1), null, null, null, s('w', 0)],
    [null, s(`${LONG}\n2 行目`, 5)],
    [s('t', 3), s('m', 4)],
    [null, s(LONG, 6)],
  ];
  const book = await openBook({
    sheets: [{ name: 'S', rows, merges }],
    xfs: ALIGNS.map(() => 0),
    xfAlignments: ALIGNS,
  });
  const m = await mount(book);
  await wait();
  return m;
}

/** 行 r・列 c の td（結合の無い行だけ） */
function td(root: HTMLElement, r: number, c: number): HTMLTableCellElement {
  const tr = root.querySelector<HTMLTableRowElement>(`tbody tr[aria-rowindex="${r + 1}"]`)!;
  return [...tr.querySelectorAll<HTMLTableCellElement>('td:not(.xs-rownum)')][c - 1]!;
}
const box = (root: HTMLElement, r: number, c: number) =>
  td(root, r, c).querySelector<HTMLElement>(':scope > .xs-cell')!;
const classes = (el: HTMLElement) => [...el.classList].filter((c) => c !== 'xs-cell');

describe('配置の描き方（renderSheet）', () => {
  it('横・縦のクラス。数値の指定は safe のクラスで、型の既定のクラスは付けない', async () => {
    const { root } = await alignSheet();
    expect(classes(box(root, 1, 1))).toEqual([]);
    expect(classes(box(root, 1, 2))).toEqual(['is-h-center']);
    expect(classes(box(root, 1, 3))).toEqual(['is-h-right']);
    expect(classes(box(root, 1, 4))).toEqual(['is-center']); // 数値の中央
    expect(classes(box(root, 1, 5))).toEqual(['is-right']); // 数値の標準
    expect(classes(box(root, 6, 1))).toEqual(['is-v-top']);
    expect(classes(box(root, 6, 2))).toEqual(['is-v-center', 'is-spill']); // 右の空の列へはみ出す
  });

  it('右寄せは左へはみ出す：入れ物を左の空の列の分だけ広げる', async () => {
    const { root } = await alignSheet();
    const b = box(root, 3, 4);
    expect(classes(b)).toEqual(['is-h-right', 'is-spill']);
    expect(b.style.marginLeft).toBe(`${-64 * 3}px`);
    expect(b.style.width).toBe(`${64 * 4 - 1}px`);
    expect(b.querySelector(':scope > .xs-text')!.textContent).toBe(LONG);
  });

  it('中央は両側へはみ出し、左右の余白の差で文字の中心を元のセルの中心に合わせる', async () => {
    const { root } = await alignSheet();
    const b = box(root, 4, 3);
    expect(classes(b)).toEqual(['is-h-center', 'is-spill']);
    // 左は B の 1 列（A に値）、右は D〜F の 3 列（G に値）
    expect(b.style.marginLeft).toBe('-64px');
    expect(b.style.width).toBe(`${64 * 5 - 1}px`);
    expect(b.style.paddingRight).toBe(`${3 + 64 * 2}px`);
    expect(b.style.paddingLeft).toBe('');
  });

  it('折り返すセルははみ出さず、文字の改行を残す', async () => {
    const { root } = await alignSheet();
    const b = box(root, 5, 2);
    expect(classes(b)).toEqual(['is-v-top', 'is-wrap']);
    expect(b.style.width).toBe('');
    expect(b.textContent).toBe(`${LONG}\n2 行目`);
  });

  it('結合セルは結合の範囲いっぱいの入れ物に配置を当てる（B7:D8）', async () => {
    const { root } = await alignSheet(['B7:D8']);
    const cell = root.querySelector<HTMLTableCellElement>(
      'tbody tr[aria-rowindex="8"] td[colspan="3"]',
    )!;
    expect(cell.rowSpan).toBe(2);
    const b = cell.querySelector<HTMLElement>(':scope > .xs-cell')!;
    expect(classes(b)).toEqual(['is-h-center', 'is-v-center', 'is-wrap']);
    expect(b.style.height).toBe(`${20 * 2 - 1}px`);
  });

  it('左・両側へのはみ出しにもハイライトを付ける', async () => {
    const { root, view, book } = await alignSheet();
    const { matches } = await findText(await book.openSheet(1), 'はみ出す');
    view.setHighlights(matches);
    const hits = (r: number, c: number) =>
      [...box(root, r, c).querySelectorAll('.xs-text > .xs-hits > .xs-hit')].map(
        (e) => e.textContent,
      );
    expect(hits(3, 4)).toEqual(['はみ出す']);
    expect(hits(4, 3)).toEqual(['はみ出す']);
    // 折り返したセル（span.xs-text が無い）は入れ物の直下
    expect(
      [...box(root, 5, 2).querySelectorAll(':scope > .xs-hits > .xs-hit')].map(
        (e) => e.textContent,
      ),
    ).toEqual(['はみ出す']);
  });
});
