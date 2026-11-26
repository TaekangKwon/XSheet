// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createGeometry } from '../../src/core/index';
import { emptyLayout } from '../../src/core/internal';
import { keyTarget, nextVisible, type KeyContext } from '../../src/render/selection-keys';
import type { Selection } from '../../src/render/selection';
import type { CellSpec } from '../../src/testing/make-xlsx';
import { cleanupViews, mount, openBook, stubClientSize } from './render-sheet.helpers';

/**
 * キーボードで選択を移す（M7-25・ADR-0011 §4・Issue #56）。行き先の計算は寸法だけで確かめ、
 * 表示ではキーを送って選択・スクロールの量・既定の動きを止めたかを確かめる（寸法：行 20px・列 64px・列見出し 24px）。
 */

const MAX_C = 16_384;
const MAX_R = 1_048_576;

describe('nextVisible', () => {
  const size = (i: number) => ([3, 4].includes(i) ? 0 : 1);
  it('大きさ 0 を飛ばし、端より先は null', () => {
    expect(nextVisible(2, 1, 10, size)).toBe(5);
    expect(nextVisible(5, -1, 10, size)).toBe(2);
    expect(nextVisible(10, 1, 10, size)).toBeNull();
    expect(nextVisible(1, -1, 10, size)).toBeNull();
    expect(nextVisible(50, -1, 10, size)).toBeNull();
  });
});

describe('keyTarget', () => {
  const layout = {
    ...emptyLayout(),
    rows: [{ min: 5, max: 6, ht: null, custom: false, hidden: true }],
  };
  const geo = createGeometry(layout, 100, 20);
  const ctx: KeyContext = {
    geo,
    extent: { rows: 100, cols: 20 },
    used: { rows: 30, cols: 8 },
    pageHeight: 200,
  };
  const cell = (r1: number, c1: number, r2 = r1, c2 = c1): Selection => ({
    type: 'cell',
    range: { r1, c1, r2, c2 },
    active: { r: r1, c: c1 },
  });
  const row = (r: number, c: number, r2 = r): Selection => ({
    type: 'row',
    range: { r1: r, c1: 1, r2, c2: MAX_C },
    active: { r, c },
  });
  const col = (c: number, r: number): Selection => ({
    type: 'column',
    range: { r1: 1, c1: c, r2: MAX_R, c2: c },
    active: { r, c },
  });
  const k = (key: string, sel: Selection, ctrl = false) => keyTarget(key, ctrl, sel, ctx);

  it('矢印：セルは 4 方向、結合は範囲の端から、非表示の行は飛ばす、端は null', () => {
    expect(k('ArrowDown', cell(2, 2))).toEqual({ r: 3, c: 2 });
    expect(k('ArrowRight', cell(2, 2, 3, 4))).toEqual({ r: 2, c: 5 });
    expect(k('ArrowDown', cell(2, 2, 3, 4))).toEqual({ r: 4, c: 2 });
    expect(k('ArrowDown', cell(4, 1))).toEqual({ r: 7, c: 1 });
    expect(k('ArrowUp', cell(7, 1))).toEqual({ r: 4, c: 1 });
    expect(k('ArrowUp', cell(1, 1))).toBeNull();
    expect(k('ArrowLeft', cell(1, 1))).toBeNull();
    expect(k('ArrowRight', cell(1, 20))).toBeNull();
  });

  it('行の全体は ←・→ で今のセルの列を、列の全体は ↑・↓ で今のセルの行を動かす', () => {
    expect(k('ArrowRight', row(3, 2))).toEqual({ r: 3, c: 3 });
    expect(k('ArrowDown', row(3, 2, 4))).toEqual({ r: 7, c: 2 });
    expect(k('ArrowDown', col(3, 2))).toEqual({ r: 3, c: 3 });
    expect(k('ArrowRight', col(3, 2))).toEqual({ r: 2, c: 4 });
  });

  it('PageDown・PageUp は 1 ページの高さだけ、Home・Ctrl+Home・Ctrl+End', () => {
    // 行 1 の上端 0 から 200px → 行 11 の位置。行 5・6 は非表示なので 13 行目
    expect(k('PageDown', cell(1, 3))).toEqual({ r: 13, c: 3 });
    expect(k('PageUp', cell(13, 3))).toEqual({ r: 1, c: 3 });
    expect(k('PageUp', cell(1, 3))).toBeNull();
    expect(k('Home', cell(9, 7))).toEqual({ r: 9, c: 1 });
    expect(k('Home', cell(9, 7), true)).toEqual({ r: 1, c: 1 });
    expect(k('End', cell(9, 7), true)).toEqual({ r: 30, c: 8 });
    expect(k('End', cell(9, 7))).toBeUndefined();
    expect(k('Tab', cell(9, 7))).toBeUndefined();
  });
});

describe('表示でのキー', () => {
  let restore: () => void;
  beforeEach(() => {
    restore = stubClientSize(800, 400);
  });
  afterEach(() => {
    cleanupViews();
    restore();
    document.body.replaceChildren();
  });

  function rows(n: number, cols: number): CellSpec[][] {
    return Array.from({ length: n }, (_, i) =>
      Array.from({ length: cols }, (_, j) => `${i + 1},${j + 1}`),
    );
  }

  async function setup(selection: 'row' | 'cell') {
    const book = await openBook({ sheets: [{ name: 'S', rows: rows(200, 30) }] });
    const changes: (string | null)[] = [];
    const m = await mount(book, 1, {
      selection,
      onSelectionChange: (s, o) =>
        changes.push(s ? `${s.type}:${s.active.r},${s.active.c}:${o}` : null),
    });
    const body = m.root.querySelector<HTMLElement>('.xs-body')!;
    const press = (key: string, init: KeyboardEventInit = {}) => {
      const e = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init });
      body.dispatchEvent(e);
      return e.defaultPrevented;
    };
    return { ...m, body, press, changes };
  }

  it('スクロール枠はフォーカスを受け、最初のキーで見えている左上を選ぶ', async () => {
    const { body, view, press, changes } = await setup('cell');
    expect(body.tabIndex).toBe(0);
    expect(press('ArrowDown')).toBe(true);
    expect(view.getSelection()!.active).toEqual({ r: 1, c: 1 });
    expect(press('ArrowDown')).toBe(true);
    expect(press('ArrowRight')).toBe(true);
    expect(view.getSelection()!.range).toEqual({ r1: 2, c1: 2, r2: 2, c2: 2 });
    expect(changes).toEqual(['cell:1,1:user', 'cell:2,1:user', 'cell:2,2:user']);
  });

  it('見えない所へ動いたら最小の量だけスクロールする（下へは下端に、上へは上端に合わせる）', async () => {
    const { body, view, press } = await setup('cell');
    view.setSelection({ type: 'cell', r: 18, c: 1 });
    press('ArrowDown'); // 行 19：下端 380 ≤ 0 + 376 ではない → 4px
    expect(body.scrollTop).toBe(19 * 20 - (400 - 24));
    press('PageDown');
    const r = view.getSelection()!.active.r;
    expect(body.scrollTop).toBe(r * 20 - (400 - 24));
    press('Home', { ctrlKey: true });
    expect(view.getSelection()!.active).toEqual({ r: 1, c: 1 });
    expect(body.scrollTop).toBe(0);
    press('End', { ctrlKey: true });
    expect(view.getSelection()!.active).toEqual({ r: 200, c: 30 });
    expect(body.scrollLeft).toBeGreaterThan(0);
  });

  it("selection: 'row' の行の全体は ↓ で次の行、→ は列を覚えて横にスクロールする", async () => {
    const { body, view, press } = await setup('row');
    view.setSelection({ type: 'row', r: 3 });
    press('ArrowDown');
    expect(view.getSelection()!.range.r1).toBe(4);
    expect(view.getSelection()!.type).toBe('row');
    for (let i = 0; i < 15; i++) press('ArrowRight');
    expect(view.getSelection()!.active.c).toBe(16);
    expect(body.scrollLeft).toBeGreaterThan(0);
    expect(view.getSelection()!.type).toBe('row');
  });

  it('扱わないキー・Shift・Alt・端では既定の動きを止めない（端は止める）', async () => {
    const { view, press } = await setup('cell');
    view.setSelection({ type: 'cell', r: 1, c: 1 });
    expect(press('Tab')).toBe(false);
    expect(press('f', { ctrlKey: true })).toBe(false);
    expect(press('ArrowDown', { shiftKey: true })).toBe(false);
    expect(press('ArrowDown', { altKey: true })).toBe(false);
    expect(press('End')).toBe(false);
    expect(press('ArrowUp')).toBe(true); // 端：ページをスクロールさせない
    expect(view.getSelection()!.active).toEqual({ r: 1, c: 1 });
  });

  it('選択の番地を aria-live の要素に入れる', async () => {
    const { root, view } = await setup('cell');
    const live = root.querySelector('.xs-live')!;
    expect(live.getAttribute('aria-live')).toBe('polite');
    view.setSelection({ type: 'cell', r: 5, c: 3 });
    expect(live.textContent).toBe('C5');
    view.setSelection({ type: 'row', r: 2, r2: 4 });
    expect(live.textContent).toBe('2:4');
    view.setSelection({ type: 'column', c: 28 });
    expect(live.textContent).toBe('AB:AB');
  });

  it('destroy でフォーカスの属性とリスナを外す', async () => {
    const { body, view, press } = await setup('cell');
    view.destroy();
    expect(body.hasAttribute('tabindex')).toBe(false);
    expect(press('ArrowDown')).toBe(false);
  });
});
