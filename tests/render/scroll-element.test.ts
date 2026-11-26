// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import type { CellSpec } from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  mount,
  mountWorkbook,
  openBook,
  stubClientSize,
  wait,
} from './render-sheet.helpers';

/**
 * スクロール枠を出す口 `scrollElement`（M7-19・ADR-0010 §1）。jsdom はレイアウトを持たないので、
 * 要素が同じであること・null になる場合・外から `scrollTop` を替えて scroll を送ると描き足すことを確かめる。
 * 2 つの表が一緒に動くことは examples（`scroll-sync.html`）で目視する。
 */

let restoreSize: () => void;
beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
});
afterEach(() => {
  cleanupViews();
  restoreSize();
});

/** 行 r（1 始まり）の A 列の td */
const rowCell = (root: HTMLElement, r: number): HTMLTableCellElement | null =>
  root.querySelector(`tbody tr[aria-rowindex="${r + 1}"] td:not(.xs-rownum)`);

const tallRows: CellSpec[][] = Array.from({ length: 300 }, (_, i) => [`r${i + 1}`]);

describe('renderWorkbook の scrollElement', () => {
  it('div.xs-body を返し、シートを替えても同じ要素。替えると左上に戻る', async () => {
    const { root, view } = await mountWorkbook({
      sheets: [
        { name: 'A', rows: tallRows },
        { name: 'B', rows: [['b']] },
      ],
    });
    const el = view.scrollElement;
    expect(el).not.toBeNull();
    expect(el).toBe(root.querySelector('.xs-sheet > .xs-body'));
    el!.scrollTop = 100;
    el!.scrollLeft = 30;
    await view.showSheet(2);
    expect(view.scrollElement).toBe(el);
    expect(el!.scrollTop).toBe(0);
    expect(el!.scrollLeft).toBe(0);
  });

  it('外から scrollTop を替えて scroll を送ると、利用者のスクロールと同じく描き足す', async () => {
    const { root, view } = await mountWorkbook({ sheets: [{ name: 'A', rows: tallRows }] });
    await wait();
    expect(rowCell(root, 280)).toBeNull(); // まだ描いていない
    const el = view.scrollElement!;
    el.scrollTop = 280 * 20;
    el.dispatchEvent(new Event('scroll'));
    await wait(60);
    expect(rowCell(root, 280)?.textContent).toBe('r280');
  });

  it('destroy の後は null', async () => {
    const { view } = await mountWorkbook({ sheets: [{ name: 'A', rows: [['a']] }] });
    expect(view.scrollElement).not.toBeNull();
    view.destroy();
    expect(view.scrollElement).toBeNull();
  });

  it('表の無い表示（シートが 0 枚・開けないファイル）は null', async () => {
    const empty = await mountWorkbook({ sheets: [] });
    expect(empty.view.scrollElement).toBeNull();
    const broken = await mountWorkbook(new TextEncoder().encode('plain text'));
    expect(broken.view.scrollElement).toBeNull();
  });
});

describe('renderSheet の scrollElement', () => {
  it('div.xs-body を返し、showSheet の後も同じ要素。destroy の後は null', async () => {
    const book = await openBook({
      sheets: [
        { name: 'A', rows: [['a']] },
        { name: 'B', rows: [['b']] },
      ],
    });
    const { root, view } = await mount(book);
    const el = view.scrollElement;
    expect(el).toBe(root.querySelector(':scope > .xs-body'));
    await view.showSheet(2);
    expect(view.scrollElement).toBe(el);
    view.destroy();
    expect(view.scrollElement).toBeNull();
  });
});
