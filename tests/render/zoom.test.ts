// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createGeometry } from '../../src/core/index';
import { emptyLayout } from '../../src/core/internal';
import {
  anchorScroll,
  contentWidth,
  fitZoom,
  resolveZoom,
  scrollAnchor,
  unzoomRect,
  zoomGeometry,
} from '../../src/render/zoom';
import type { CellSpec } from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  mount,
  mountWorkbook,
  openBook,
  px,
  scrollBody,
  stubClientSize,
  wait,
} from './render-sheet.helpers';

/**
 * 拡大縮小（M7-7・ADR-0007 §5）。寸法の包み・指定の確かめ・fit・左上のセルを残すことを jsdom で確かめ、
 * 罫線の層と td の境目が合うこと・文字の大きさは実機で見る。
 */

/** 行の高さ 20px・列の幅 64px（既定）に、行 3 を 25px・列 2 を 100px にした寸法 */
const layout = {
  ...emptyLayout(),
  rows: [{ min: 3, max: 3, ht: 18.75, custom: true, hidden: false }],
  cols: [{ min: 2, max: 2, width: 13.57, hidden: false }],
};
const base = createGeometry(layout, 100, 20);

describe('zoomGeometry', () => {
  it('倍率 1 は元の寸法そのもの', () => {
    expect(zoomGeometry(base, 1)).toBe(base);
  });

  it.each([0.5, 0.75, 1.5, 2])('境目は整数に丸め、高さ・幅は境目の差（倍率 %s）', (z) => {
    const g = zoomGeometry(base, z);
    for (let r = 1; r <= 100; r++) {
      expect(g.rowTop(r)).toBe(Math.round(base.rowTop(r) * z));
      expect(g.rowHeight(r)).toBe(g.rowTop(r + 1) - g.rowTop(r));
      expect(Number.isInteger(g.rowTop(r))).toBe(true);
    }
    for (let c = 1; c <= 20; c++) {
      expect(g.colLeft(c)).toBe(Math.round(base.colLeft(c) * z));
      expect(g.colWidth(c)).toBe(g.colLeft(c + 1) - g.colLeft(c));
    }
    expect(g.totalHeight).toBe(g.rowTop(101));
    expect(g.totalWidth).toBe(g.colLeft(21));
  });

  it.each([0.5, 0.75, 1.5])('rowAt・colAt は丸めた境目で引く（倍率 %s）', (z) => {
    const g = zoomGeometry(base, z);
    for (let r = 1; r <= 100; r++) {
      expect(g.rowAt(g.rowTop(r))).toBe(r);
      expect(g.rowAt(g.rowTop(r + 1) - 0.5)).toBe(r);
    }
    for (let c = 1; c <= 20; c++) {
      expect(g.colAt(g.colLeft(c))).toBe(c);
      expect(g.colAt(g.colLeft(c + 1) - 0.5)).toBe(c);
    }
    expect(g.rowAt(-5)).toBe(1);
    expect(g.rowAt(1e9)).toBe(100);
    expect(g.colAt(Number.NaN)).toBe(1);
  });
});

describe('倍率の指定', () => {
  it('数は 0.5〜2.0 に丸める。数でない値・正しくない指定は RangeError', () => {
    expect(resolveZoom(0.1)).toEqual({ fixed: 0.5 });
    expect(resolveZoom(3)).toEqual({ fixed: 2 });
    expect(resolveZoom(1.25)).toEqual({ fixed: 1.25 });
    expect(() => resolveZoom(Number.NaN)).toThrow(RangeError);
    expect(() => resolveZoom(Infinity)).toThrow(RangeError);
    expect(() => resolveZoom({ fit: 'height' } as never)).toThrow(RangeError);
    expect(() => resolveZoom(null as never)).toThrow(RangeError);
  });

  it('fit の min・max は既定 0.5・2.0。範囲に丸め、逆転は入れ替える', () => {
    expect(resolveZoom({ fit: 'width' })).toEqual({ fit: { min: 0.5, max: 2 } });
    expect(resolveZoom({ fit: 'width', min: 0.8, max: 0.6 })).toEqual({
      fit: { min: 0.6, max: 0.8 },
    });
    expect(resolveZoom({ fit: 'width', min: 0.1, max: Number.NaN })).toEqual({
      fit: { min: 0.5, max: 2 },
    });
  });

  it('fit の倍率：範囲に丸め、1% 未満の差は合わせ直さない', () => {
    expect(fitZoom(800, 1000, { min: 0.5, max: 2 }, 1)).toBe(0.8);
    expect(fitZoom(800, 100, { min: 0.5, max: 2 }, 1)).toBe(2);
    expect(fitZoom(805, 1000, { min: 0.5, max: 2 }, 0.8)).toBeNull();
    expect(fitZoom(0, 1000, { min: 0.5, max: 2 }, 1)).toBeNull();
    expect(fitZoom(800, 0, { min: 0.5, max: 2 }, 1)).toBeNull();
  });

  it('中身の幅はデータの列の右端と図の右端の大きいほう。判定の範囲は 100% に戻す', () => {
    expect(contentWidth(500, [])).toBe(500);
    expect(contentWidth(500, [{ index: 0, rect: { x: 400, y: 0, w: 300, h: 10 } }])).toBe(700);
    expect(unzoomRect({ x: 10, y: 20, w: 30, h: 40 }, 0.5)).toEqual({ x: 20, y: 40, w: 60, h: 80 });
  });

  it('左上のセルとセルの中の割合を、新しい寸法の位置にする', () => {
    const a = scrollAnchor(base, base.rowTop(50) + 10, base.colLeft(5) + 16);
    expect(a).toEqual({ r: 50, fy: 0.5, c: 5, fx: 0.25 });
    const g = zoomGeometry(base, 0.5);
    expect(anchorScroll(g, a)).toEqual({
      top: Math.round(g.rowTop(50) + 0.5 * g.rowHeight(50)),
      left: Math.round(g.colLeft(5) + 0.25 * g.colWidth(5)),
    });
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

const rows300 = (): CellSpec[][] =>
  Array.from({ length: 300 }, (_, i) => [`r${i + 1}`, 1234567, null, null]);
const bodyOf = (root: HTMLElement) => root.querySelector<HTMLElement>('.xs-body')!;
const colWidths = (root: HTMLElement) =>
  [...root.querySelectorAll<HTMLElement>('col')].map((c) => px(c.style.width));

describe('setZoom（renderSheet）', () => {
  it('セルの寸法に倍率を掛けて描き直す。見出しの高さ・行番号の列の幅は変えない', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: rows300() }] });
    const { root, view } = await mount(book);
    await wait();
    expect(view.zoom).toBe(1);
    const numWidth = colWidths(root)[0];
    view.setZoom(0.5);
    expect(view.zoom).toBe(0.5);
    expect(root.style.getPropertyValue('--xs-zoom')).toBe('0.5');
    const widths = colWidths(root);
    expect(widths[0]).toBe(numWidth); // 行番号の列
    expect(widths[1]).toBe(32); // 64px × 0.5
    const tr = root.querySelector<HTMLElement>('tbody tr[aria-rowindex="2"]')!;
    expect(tr.style.height).toBe('10px');
    view.setZoom(2);
    expect(colWidths(root)[1]).toBe(128);
    expect(root.querySelector<HTMLElement>('tbody tr[aria-rowindex="2"]')!.style.height).toBe(
      '40px',
    );
  });

  it('縮小すると枠を埋める行が増え、収まる数値の判定も倍率で（### になる）', async () => {
    const book = await openBook({
      sheets: [{ name: 'S', rows: [[123456789]], cols: [{ min: 1, max: 1, width: 10 }] }],
    });
    const { root, view } = await mount(book);
    await wait();
    const cell = () =>
      root.querySelector<HTMLElement>('tbody tr[aria-rowindex="2"] td:not(.xs-rownum) .xs-cell')!;
    expect(cell().textContent).toBe('123456789');
    const rowsAt1 = root.querySelectorAll('tbody tr:not(.xs-spacer)').length;
    view.setZoom(0.5);
    expect(cell().textContent).toMatch(/^#+$/);
    expect(root.querySelectorAll('tbody tr:not(.xs-spacer)').length).toBeGreaterThan(rowsAt1);
  });

  it('左上に見ていたセルを左上に残す', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: rows300() }] });
    const { root, view } = await mount(book);
    await wait();
    scrollBody(root, { top: 20 * 150 + 5, left: 64 + 16 });
    await wait(60);
    view.setZoom(0.5);
    expect(bodyOf(root).scrollTop).toBe(10 * 150 + 3); // 行 151 の上から 1/4
    expect(bodyOf(root).scrollLeft).toBe(32 + 8);
    expect(root.querySelector('tbody tr[aria-rowindex="152"]')).not.toBeNull();
  });

  it('位置の判定・移動も倍率を掛けた寸法で', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: rows300() }] });
    const { view } = await mount(book, 1, { zoom: 0.5 });
    await wait();
    expect(view.zoom).toBe(0.5);
    // 行番号の列 40px・見出し 24px は倍率なし。セルは 32 × 10
    expect(view.cellAt(40 + 33, 24 + 11)).toEqual({ r: 2, c: 2 });
    expect(view.cellAt(39, 24 + 11)).toBeNull();
  });

  it('幅に合わせる：データのある列の右端を枠の幅に合わせ、枠が変わると合わせ直す', async () => {
    const book = await openBook({
      sheets: [{ name: 'S', rows: [Array.from({ length: 20 }, (_, i) => i)] }],
    });
    const { view } = await mount(book, 1, { zoom: { fit: 'width' } });
    await wait();
    // 20 列 × 64px = 1280px を 800 − 32（行番号の列）に
    expect(view.zoom).toBeCloseTo(768 / 1280, 5);
    restoreSize();
    restoreSize = stubClientSize(1056, 400);
    view.setZoom({ fit: 'width' });
    expect(view.zoom).toBeCloseTo(1024 / 1280, 5);
    view.setZoom({ fit: 'width', max: 0.7 });
    expect(view.zoom).toBe(0.7);
    view.setZoom(1);
    expect(view.zoom).toBe(1);
  });

  it('正しくない倍率は RangeError（renderSheet は DOM を足さない）', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: [['a']] }] });
    const { view } = await mount(book);
    expect(() => view.setZoom(Number.NaN)).toThrow(RangeError);
    expect(view.zoom).toBe(1);
    await expect(mount(book, 1, { zoom: Number.NaN })).rejects.toBeInstanceOf(RangeError);
    expect(document.querySelectorAll('.xs-sheet')).toHaveLength(1);
  });
});

describe('setZoom（renderWorkbook）', () => {
  it('シートを替えても倍率を保つ', async () => {
    const { root, view } = await mountWorkbook(
      {
        sheets: [
          { name: 'A', rows: [['a']] },
          { name: 'B', rows: [['b']] },
        ],
      },
      { zoom: 1.5 },
    );
    await wait();
    expect(view.zoom).toBe(1.5);
    await view.showSheet(2);
    await wait();
    expect(view.zoom).toBe(1.5);
    expect(colWidths(root)[1]).toBe(96);
  });

  it('正しくない倍率はブックを開く前に RangeError。開けなかったブックは 1 のまま', async () => {
    await expect(
      mountWorkbook({ sheets: [{ name: 'A', rows: [] }] }, { zoom: Number.NaN }),
    ).rejects.toBeInstanceOf(RangeError);
    const { view } = await mountWorkbook(new Uint8Array([1, 2, 3]));
    view.setZoom(2);
    expect(view.zoom).toBe(1);
    expect(() => view.setZoom(Number.NaN)).toThrow(RangeError);
  });
});
