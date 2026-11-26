// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { mergeRangeAt, type Selection } from '../../src/render/selection';
import type { RenderOptions } from '../../src/render/render-sheet';
import type { CellSpec } from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  mount,
  mountWorkbook,
  openBook,
  scrollBody,
  stubClientSize,
  wait,
} from './render-sheet.helpers';

/**
 * セル・行・列の選択（M7-24・ADR-0011 §1・§2・Issue #55）。jsdom はレイアウトを持たないので、
 * 位置は寸法（行 20px・列 64px・行番号の列 40px・列見出し 24px）から出し、帯の位置・見出しの印・知らせを DOM で確かめる。
 * 実機の見え方（重なりの順・スクロール・拡大縮小）は examples で目視する。
 */

let restore: () => void;
beforeEach(() => {
  restore = stubClientSize(800, 400);
});
afterEach(() => {
  cleanupViews();
  restore();
  document.body.replaceChildren();
});

const MAX_C = 16_384;
const MAX_R = 1_048_576;

/** r 行 × c 列に `r,c` を置いた行の並び */
function cells(rows: number, cols: number): CellSpec[][] {
  return Array.from({ length: rows }, (_, i) =>
    Array.from({ length: cols }, (_, j) => `${i + 1},${j + 1}`),
  );
}

async function setup(options: RenderOptions = {}, merges?: string[], rows = cells(30, 8)) {
  const book = await openBook({ sheets: [{ name: 'S', rows, merges }] });
  const changes: [Selection | null, string][] = [];
  const m = await mount(book, 1, {
    selection: 'row',
    onSelectionChange: (s, o) => changes.push([s, o]),
    ...options,
  });
  return { ...m, changes };
}

/** スクロール枠の中の位置 (x, y) を押して離す */
function click(root: HTMLElement, x: number, y: number, moveTo?: { x: number; y: number }): void {
  const body = root.querySelector<HTMLElement>('.xs-body')!;
  body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: x, clientY: y }));
  const end = moveTo ?? { x, y };
  body.dispatchEvent(
    new MouseEvent('click', { bubbles: true, clientX: end.x, clientY: end.y, button: 0 }),
  );
}

/** 行番号の列の幅（表の最初の <col>。行数の桁で変わる） */
const ox = (root: HTMLElement): number =>
  Number.parseFloat(root.querySelector<HTMLElement>('.xs-grid col')!.style.width);

/** セル (r, c) の中の位置 */
const at = (root: HTMLElement, r: number, c: number) => ({
  x: ox(root) + 64 * (c - 1) + 5,
  y: 24 + 20 * (r - 1) + 5,
});

/** 選択の帯（まだ置いていなければ、隠れた帯の代わりに hidden の要素） */
const band = (root: HTMLElement): HTMLElement =>
  root.querySelector<HTMLElement>('.xs-sel') ??
  Object.assign(document.createElement('div'), { hidden: true });
const selectedHeads = (root: HTMLElement) =>
  [...root.querySelectorAll('.is-selected')].map((el) => el.textContent);

describe('mergeRangeAt', () => {
  it('結合の中のセルは結合の範囲、そうでなければそのセル', async () => {
    const book = await openBook({
      sheets: [{ name: 'S', rows: cells(10, 6), merges: ['B2:D4', 'F1:F3'] }],
    });
    const sh = await book.openSheet(1);
    expect(mergeRangeAt(sh, 3, 3)).toEqual({ r1: 2, c1: 2, r2: 4, c2: 4 });
    expect(mergeRangeAt(sh, 2, 2)).toEqual({ r1: 2, c1: 2, r2: 4, c2: 4 });
    expect(mergeRangeAt(sh, 4, 4)).toEqual({ r1: 2, c1: 2, r2: 4, c2: 4 });
    expect(mergeRangeAt(sh, 3, 6)).toEqual({ r1: 1, c1: 6, r2: 3, c2: 6 });
    expect(mergeRangeAt(sh, 5, 3)).toEqual({ r1: 5, c1: 3, r2: 5, c2: 3 });
    expect(mergeRangeAt(sh, 3, 1)).toEqual({ r1: 3, c1: 1, r2: 3, c2: 1 });
    expect(mergeRangeAt(sh, 50, 1)).toEqual({ r1: 50, c1: 1, r2: 50, c2: 1 });
  });
});

describe('クリックで選ぶ', () => {
  it("selection: 'row' はセルのクリックで行の全体。帯は行の全幅、行番号に印", async () => {
    const { root, view, changes } = await setup();
    click(root, at(root, 3, 2).x, at(root, 3, 2).y);
    const sel = view.getSelection()!;
    expect(sel).toEqual({
      type: 'row',
      range: { r1: 3, c1: 1, r2: 3, c2: MAX_C },
      active: { r: 3, c: 2 },
    });
    expect(changes).toEqual([[sel, 'user']]);
    const b = band(root);
    expect(b.hidden).toBe(false);
    expect(b.className).toBe('xs-sel is-row');
    expect(b.style.top).toBe(`${24 + 40}px`);
    expect(b.style.left).toBe(`${ox(root)}px`);
    expect(b.style.height).toBe('20px');
    expect(Number.parseFloat(b.style.width)).toBeGreaterThanOrEqual(64 * 8);
    expect(selectedHeads(root)).toEqual(['3']);
    // 帯は図の層の前（罫線の層・表の後）
    expect(b.nextElementSibling?.classList.contains('xs-layer')).toBe(true);
  });

  it("selection: 'cell' はセル。行番号と列記号の両方に印、枠の帯", async () => {
    const { root, view } = await setup({ selection: 'cell' });
    click(root, at(root, 2, 3).x, at(root, 2, 3).y);
    expect(view.getSelection()).toEqual({
      type: 'cell',
      range: { r1: 2, c1: 3, r2: 2, c2: 3 },
      active: { r: 2, c: 3 },
    });
    const b = band(root);
    expect(b.className).toBe('xs-sel is-cell');
    expect([b.style.left, b.style.top, b.style.width, b.style.height]).toEqual([
      `${ox(root) + 128}px`,
      `${24 + 20}px`,
      '64px',
      '20px',
    ]);
    expect(selectedHeads(root).sort()).toEqual(['2', 'C']);
  });

  it('結合したセルは結合の範囲（行の選択なら結合の行の範囲）', async () => {
    const row = await setup({}, ['B2:C4']);
    click(row.root, at(row.root, 3, 3).x, at(row.root, 3, 3).y);
    expect(row.view.getSelection()!.range).toEqual({ r1: 2, c1: 1, r2: 4, c2: MAX_C });
    expect(row.view.getSelection()!.active).toEqual({ r: 2, c: 2 });
    expect(band(row.root).style.height).toBe('60px');
    expect(selectedHeads(row.root)).toEqual(['2', '3', '4']);

    const cell = await setup({ selection: 'cell' }, ['B2:C4']);
    click(cell.root, at(cell.root, 4, 2).x, at(cell.root, 4, 2).y);
    expect(cell.view.getSelection()!.range).toEqual({ r1: 2, c1: 2, r2: 4, c2: 3 });
  });

  it('行番号のクリックは行の全体、列記号のクリックは列の全体（今のセルの行・列を覚える）', async () => {
    const { root, view } = await setup({ selection: 'cell' });
    click(root, at(root, 5, 4).x, at(root, 5, 4).y);
    click(root, 10, at(root, 7, 1).y);
    expect(view.getSelection()).toEqual({
      type: 'row',
      range: { r1: 7, c1: 1, r2: 7, c2: MAX_C },
      active: { r: 7, c: 4 },
    });
    click(root, at(root, 1, 2).x, 10);
    const sel = view.getSelection()!;
    expect(sel).toEqual({
      type: 'column',
      range: { r1: 1, c1: 2, r2: MAX_R, c2: 2 },
      active: { r: 7, c: 2 },
    });
    const b = band(root);
    expect(b.className).toBe('xs-sel is-column');
    expect([b.style.left, b.style.top, b.style.width]).toEqual([
      `${ox(root) + 64}px`,
      '24px',
      '64px',
    ]);
    expect(selectedHeads(root)).toEqual(['B']);
  });

  it('左上の角・動いたクリック・右ボタン・文字を選んでいるときは選ばない', async () => {
    const { root, view, changes } = await setup();
    click(root, 10, 10);
    click(root, at(root, 2, 2).x, at(root, 2, 2).y, {
      x: at(root, 2, 2).x + 30,
      y: at(root, 2, 2).y,
    });
    const body = root.querySelector<HTMLElement>('.xs-body')!;
    body.dispatchEvent(new MouseEvent('mousedown', { button: 2, clientX: 50, clientY: 50 }));
    body.dispatchEvent(new MouseEvent('click', { button: 0, clientX: 50, clientY: 50 }));
    const text = root.querySelector('td:not(.xs-rownum) .xs-cell')!;
    const range = document.createRange();
    range.selectNodeContents(text);
    document.getSelection()!.removeAllRanges(); // 前のテストのフォーカスで残った位置を消す（addRange は範囲があると何もしない）
    document.getSelection()!.addRange(range);
    click(root, at(root, 2, 2).x, at(root, 2, 2).y);
    document.getSelection()!.removeAllRanges();
    expect(view.getSelection()).toBeNull();
    expect(changes).toEqual([]);
  });

  it('selection を指定しなければ選択は無い（クリックでも setSelection でも選ばない。今の利用者の動きを変えない）', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: cells(5, 5) }] });
    const { root, view } = await mount(book);
    click(root, at(root, 2, 2).x, at(root, 2, 2).y);
    click(root, 10, at(root, 2, 1).y);
    view.setSelection({ type: 'row', r: 1 });
    expect(view.getSelection()).toBeNull();
    expect(root.querySelector('.xs-sel')).toBeNull();
  });

  it('同じ物をもう一度選んでも知らせない', async () => {
    const { root, changes } = await setup();
    click(root, at(root, 2, 2).x, at(root, 2, 2).y);
    click(root, at(root, 2, 2).x, at(root, 2, 2).y);
    expect(changes).toHaveLength(1);
  });
});

describe('setSelection', () => {
  it('行・列・セルを選び、api で知らせる。null で外す', async () => {
    const { root, view, changes } = await setup({ selection: 'cell' });
    view.setSelection({ type: 'row', r: 5, r2: 3 });
    expect(view.getSelection()!.range).toEqual({ r1: 3, c1: 1, r2: 5, c2: MAX_C });
    expect(band(root).style.height).toBe('60px');
    view.setSelection({ type: 'column', c: 3 });
    expect(view.getSelection()!.range).toEqual({ r1: 1, c1: 3, r2: MAX_R, c2: 3 });
    view.setSelection({ type: 'cell', r: 2, c: 2 });
    expect(view.getSelection()!.type).toBe('cell');
    view.setSelection(null);
    expect(view.getSelection()).toBeNull();
    expect(band(root).hidden).toBe(true);
    expect(selectedHeads(root)).toEqual([]);
    expect(changes.map(([s, o]) => [s?.type ?? null, o])).toEqual([
      ['row', 'api'],
      ['column', 'api'],
      ['cell', 'api'],
      [null, 'api'],
    ]);
  });

  it.each([
    { type: 'row', r: 0 },
    { type: 'row', r: 1.5 },
    { type: 'row', r: 1, r2: MAX_R + 1 },
    { type: 'column', c: MAX_C + 1 },
    { type: 'cell', r: 1, c: Number.NaN },
    { type: 'range', r: 1, c: 1 },
  ])('正しくない選択は RangeError：%o', async (input) => {
    const { view } = await setup();
    expect(() => view.setSelection(input as never)).toThrow(RangeError);
    expect(view.getSelection()).toBeNull();
  });

  it('表の外の行は帯を出さず、表の範囲で帯を切る', async () => {
    const { root, view } = await setup();
    view.setSelection({ type: 'row', r: 5000 });
    expect(band(root).hidden).toBe(true);
    view.setSelection({ type: 'column', c: 2 });
    const h = Number.parseFloat(band(root).style.height);
    expect(h).toBeGreaterThan(0);
    expect(h).toBeLessThan(20 * 1000);
  });
});

describe('描き直し・拡大縮小・シートの切り替え', () => {
  it('スクロールで描き直しても印が残り、倍率を替えると帯の位置が合う', async () => {
    const { root, view } = await setup({}, undefined, cells(300, 5));
    view.setSelection({ type: 'row', r: 3 });
    scrollBody(root, { top: 3000 });
    await wait();
    scrollBody(root, { top: 0 });
    await wait();
    expect(selectedHeads(root)).toEqual(['3']);
    view.setZoom(2);
    const b = band(root);
    expect(b.style.top).toBe(`${24 + 80}px`);
    expect(b.style.height).toBe('40px');
  });

  it('シートを替えると外れ、api で null を知らせる', async () => {
    const changes: (Selection | null)[] = [];
    const { view } = await mountWorkbook(
      {
        sheets: [
          { name: 'A', rows: cells(5, 5) },
          { name: 'B', rows: cells(5, 5) },
        ],
      },
      { selection: 'cell', onSelectionChange: (s) => changes.push(s) },
    );
    view.setSelection({ type: 'cell', r: 2, c: 2 });
    await view.showSheet(2);
    expect(view.getSelection()).toBeNull();
    expect(changes.map((s) => s?.type ?? null)).toEqual(['cell', null]);
  });

  it('destroy の後は何もしない', async () => {
    const { view, root, changes } = await setup();
    const body = root.querySelector<HTMLElement>('.xs-body')!;
    const p = at(root, 2, 2);
    view.destroy();
    view.setSelection({ type: 'row', r: 1 });
    body.dispatchEvent(new MouseEvent('mousedown', { clientX: p.x, clientY: p.y }));
    body.dispatchEvent(new MouseEvent('click', { clientX: p.x, clientY: p.y }));
    expect(view.getSelection()).toBeNull();
    expect(changes).toEqual([]);
  });
});
