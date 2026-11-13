// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { findText } from '../../src/core/index';
import type { CellSpec } from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  headTexts,
  mount,
  mountWorkbook,
  openBook,
  px,
  stubClientSize,
  wait,
} from './render-sheet.helpers';

/** 見出しを隠す `headers: false` と `SheetView.showSheet`（M7-9・ADR-0007 §7） */

let restoreSize: () => void;
beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
});
afterEach(() => {
  cleanupViews();
  restoreSize();
});

const rows = (prefix: string): CellSpec[][] =>
  Array.from({ length: 50 }, (_, i) => [`${prefix}${i + 1} key`, i]);

describe('headers: false', () => {
  it('列見出しの行を作らず、行番号の列を幅 0 で畳む', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: rows('a') }] });
    const { root } = await mount(book, 1, { headers: false });
    await wait();
    expect(root.querySelector('thead')).toBeNull();
    expect(headTexts(root)).toEqual([]);
    const numCol = root.querySelector<HTMLElement>('col')!;
    expect(numCol.style.width).toBe('0px');
    expect(numCol.classList.contains('xs-collapsed')).toBe(true);
    const table = root.querySelector<HTMLElement>('table.xs-grid')!;
    const total = [...root.querySelectorAll<HTMLElement>('col')]
      .slice(1)
      .reduce((a, c) => a + px(c.style.width), 0);
    expect(px(table.style.width)).toBe(total);
  });

  it('既定（true）は今までどおり', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: rows('a') }] });
    const { root } = await mount(book);
    await wait();
    expect(headTexts(root).slice(0, 3)).toEqual(['', 'A', 'B']);
    expect(px(root.querySelector<HTMLElement>('col')!.style.width)).toBeGreaterThan(0);
  });

  it('位置の判定・移動は見出しの寸法を 0 として計算する', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: rows('a') }] });
    const { root, view } = await mount(book, 1, { headers: false });
    await wait();
    expect(view.cellAt(5, 5)).toEqual({ r: 1, c: 1 });
    expect(view.cellAt(64 + 5, 20 + 5)).toEqual({ r: 2, c: 2 });
    view.scrollToCell(20, 1);
    // 行 20 の上端 380px を、表示枠（400px）の上から 1/3 に
    expect(root.querySelector<HTMLElement>('.xs-body')!.scrollTop).toBe(Math.round(380 - 400 / 3));
  });

  it('幅に合わせる倍率は行番号の列を引かない', async () => {
    const book = await openBook({
      sheets: [{ name: 'S', rows: [Array.from({ length: 20 }, (_, i) => i)] }],
    });
    const { view } = await mount(book, 1, { headers: false, zoom: { fit: 'width' } });
    await wait();
    expect(view.zoom).toBeCloseTo(800 / 1280, 5);
  });
});

describe('SheetView.showSheet（renderSheet）', () => {
  const twoSheets = () =>
    openBook({
      sheets: [
        { name: 'A', rows: rows('a') },
        { name: 'B', rows: rows('b') },
      ],
    });
  const firstCell = (root: HTMLElement) =>
    root.querySelector('tbody tr[aria-rowindex="2"] td:not(.xs-rownum)')?.textContent;

  it('表示を作り直さずにシートを替える。倍率は保ち、ハイライトは消す', async () => {
    const book = await twoSheets();
    const { root, container, view } = await mount(book, 1, { zoom: 0.75 });
    await wait();
    const sheetRoot = container.querySelector('.xs-sheet');
    view.setHighlights((await findText(await book.openSheet(1), 'key')).matches);
    expect(root.querySelector('.xs-hit')).not.toBeNull();
    await view.showSheet(2);
    await wait();
    expect(container.querySelector('.xs-sheet')).toBe(sheetRoot);
    expect(firstCell(root)).toBe('b1 key');
    expect(view.zoom).toBe(0.75);
    expect(root.querySelector('.xs-hit')).toBeNull();
    expect(() => view.focusHighlight(0)).toThrow(RangeError);
  });

  it('表示中のシートは何もしない（ハイライトも残る）。ブックに無い番号は RangeError', async () => {
    const book = await twoSheets();
    const { root, view } = await mount(book);
    await wait();
    view.setHighlights([{ r: 1, c: 1, start: 0, end: 1 }]);
    await view.showSheet(1);
    expect(root.querySelector('.xs-hit')).not.toBeNull();
    await expect(view.showSheet(9)).rejects.toBeInstanceOf(RangeError);
  });

  it('renderWorkbook も headers を中の表示に渡す', async () => {
    const { root } = await mountWorkbook(
      { sheets: [{ name: 'A', rows: rows('a') }] },
      { headers: false },
    );
    await wait();
    expect(root.querySelector('thead')).toBeNull();
  });
});
