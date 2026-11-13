// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createGeometry, findText, type TextMatch } from '../../src/core/index';
import { emptyLayout } from '../../src/core/internal';
import type { GridRow } from '../../src/render/grid';
import {
  horizontalTarget,
  pointToCell,
  resolveDrawnCell,
  verticalTarget,
  type ViewBox,
} from '../../src/render/navigate';
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
 * 今のヒットへの移動と位置の判定（M7-6・ADR-0007 §3・§4）。スクロール先は寸法から出すので jsdom で確かめ、
 * 実機の見え方（smooth・はみ出す文字の幅を測る）は examples で目視する。
 */

/** 既定の寸法（行 20px・列 64px）の 1000 行 × 100 列の寸法 */
const geo = createGeometry(emptyLayout(), 1000, 100);
const view = (over: Partial<ViewBox> = {}): ViewBox => ({
  scrollTop: 0,
  scrollLeft: 0,
  clientWidth: 800,
  clientHeight: 400,
  originX: 40,
  originY: 24,
  ...over,
});

describe('スクロール先の計算', () => {
  it('縦：セルの下端から 1 行の高さを引いた位置を、セルの部分の上から 1/3 に置く', () => {
    // 行 100 の上端 1980px・セルの部分の高さ 376px → 1980 − 125.3
    expect(verticalTarget(geo, view(), 100)).toBe(1855);
    // 縦の結合（3 行）は最後の行の上端を基準にする
    expect(verticalTarget(geo, view(), 100, 3)).toBe(1895);
    // 上端・下端の先へは動かさない
    expect(verticalTarget(geo, view(), 2)).toBe(0);
    expect(verticalTarget(geo, view(), 1000)).toBe(24 + 20000 - 400);
  });

  it('縦：結合の最後の行が非表示なら、見える最後の行を基準にする', () => {
    const g = createGeometry(
      { ...emptyLayout(), rows: [{ min: 102, max: 102, ht: null, custom: false, hidden: true }] },
      1000,
      100,
    );
    // 行 100〜102（102 は非表示）：下端は 2020px、最後に見える行 101 の高さ 20px → 2000 − 125.3
    expect(verticalTarget(g, view(), 100, 3)).toBe(1875);
  });

  it('横：見えていれば動かさない。右なら右端に、左なら左端に寄せる', () => {
    // セルの部分の幅は 760px（列 1〜11 と少し）
    expect(horizontalTarget(geo, view(), { left: 64, right: 128 })).toBe(0);
    expect(horizontalTarget(geo, view(), { left: 64 * 19, right: 64 * 20 })).toBe(64 * 20 - 760);
    expect(horizontalTarget(geo, view({ scrollLeft: 2000 }), { left: 64, right: 128 })).toBe(64);
    // 枠より広い範囲は左端を合わせる
    expect(horizontalTarget(geo, view(), { left: 1000, right: 3000 })).toBe(1000);
  });

  it('画面の位置 → 行・列（見出し・枠の外・表の外は null）', () => {
    expect(pointToCell(geo, view(), 40 + 70, 24 + 25)).toEqual({ r: 2, c: 2 });
    expect(pointToCell(geo, view({ scrollTop: 200, scrollLeft: 64 }), 40, 24)).toEqual({
      r: 11,
      c: 2,
    });
    expect(pointToCell(geo, view(), 39, 100)).toBeNull(); // 行番号
    expect(pointToCell(geo, view(), 100, 23)).toBeNull(); // 列見出し
    expect(pointToCell(geo, view(), 800, 100)).toBeNull(); // スクロールバー
    expect(pointToCell(geo, view(), Number.NaN, 100)).toBeNull();
    const small = createGeometry(emptyLayout(), 3, 3);
    expect(pointToCell(small, view(), 40 + 64 * 3 + 1, 30)).toBeNull(); // 表の右
    expect(pointToCell(small, view(), 50, 24 + 60)).toBeNull(); // 表の下
  });
});

describe('描いた行から結合・はみ出しを引く', () => {
  const t = (text: string, kind?: 's') => ({ text, kind, colSpan: 1, rowSpan: 1 });
  const rows: GridRow[] = [
    // 行 1：B1:C2 の結合・D1 からはみ出す文字（E・F は空）・G1 に値
    {
      r: 1,
      cells: [
        t('a', 's'),
        { text: 'm', kind: 's', colSpan: 2, rowSpan: 2 },
        null,
        t('long text', 's'),
        undefined,
        t(''),
        t('g', 's'),
      ],
    },
    { r: 2, cells: [t('b', 's'), null, null, undefined] },
  ];

  it('結合で隠れるセルは左上（同じ行・上の行）', () => {
    expect(resolveDrawnCell(rows, 1, 3, 10)).toEqual({ r: 1, c: 2 });
    expect(resolveDrawnCell(rows, 2, 2, 10)).toEqual({ r: 1, c: 2 });
    expect(resolveDrawnCell(rows, 2, 3, 10)).toEqual({ r: 1, c: 2 });
    expect(resolveDrawnCell(rows, 2, 1, 10)).toEqual({ r: 2, c: 1 });
  });

  it('はみ出した文字の範囲ははみ出し元。値のあるセルで止まる', () => {
    expect(resolveDrawnCell(rows, 1, 5, 10)).toEqual({ r: 1, c: 4 });
    expect(resolveDrawnCell(rows, 1, 6, 10)).toEqual({ r: 1, c: 4 });
    expect(resolveDrawnCell(rows, 1, 7, 10)).toEqual({ r: 1, c: 7 });
    // 描いた最初の列より左の元は数えない（描くときもはみ出させない）
    expect(resolveDrawnCell(rows, 1, 5, 10, 5)).toEqual({ r: 1, c: 5 });
    // 行 2 の D2 は空、左は結合で隠れる（はみ出しの元にならない）
    expect(resolveDrawnCell(rows, 2, 4, 10)).toEqual({ r: 2, c: 4 });
    // 描いていない行はそのまま
    expect(resolveDrawnCell(rows, 9, 3, 10)).toEqual({ r: 9, c: 3 });
  });
});

let restoreSize: () => void;
beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
});
afterEach(() => {
  cleanupViews();
  restoreSize();
  vi.restoreAllMocks();
});

/** 300 行 × 30 列。各行の A 列と T 列（20 列目）に「key」 */
async function bigSheet() {
  const rows: CellSpec[][] = Array.from({ length: 300 }, (_, i) => {
    const row: CellSpec[] = Array.from({ length: 30 }, () => null);
    row[0] = `r${i + 1} key`;
    row[19] = `key ${i + 1}`;
    row[29] = 0;
    return row;
  });
  const book = await openBook({ sheets: [{ name: 'S', rows }] });
  const mounted = await mount(book);
  await wait();
  const { matches } = await findText(await book.openSheet(1), 'key');
  return { ...mounted, matches };
}

const bodyOf = (root: HTMLElement) => root.querySelector<HTMLElement>('.xs-body')!;
const current = (root: HTMLElement) =>
  [...root.querySelectorAll<HTMLElement>('.xs-hit.is-current')].map((e) => e.dataset.xsHit);

describe('focusHighlight（renderSheet）', () => {
  it('描いている範囲の中なら表を描き直さず、印だけを替える', async () => {
    const { root, view, matches } = await bigSheet();
    view.setHighlights(matches);
    const table = root.querySelector('table.xs-grid');
    view.focusHighlight(0, { scroll: false });
    expect(current(root)).toEqual(['0']);
    view.focusHighlight(2, { scroll: false });
    expect(current(root)).toEqual(['2']);
    expect(root.querySelector('table.xs-grid')).toBe(table);
    view.focusHighlight(-1);
    expect(current(root)).toEqual([]);
  });

  it('遠い行のヒットへスクロールし、描いた表で今のヒットになっている', async () => {
    const { root, view, matches } = await bigSheet();
    view.setHighlights(matches);
    // 行 280 の A 列（番号は行ごとに A・T の 2 つ）
    const index = matches.findIndex((m: TextMatch) => m.r === 280 && m.c === 1);
    view.focusHighlight(index);
    expect(bodyOf(root).scrollTop).toBe(verticalTarget(geo, view0(), 280));
    expect(bodyOf(root).scrollLeft).toBe(0); // A 列は見えている
    expect(current(root)).toEqual([String(index)]);
    // スクロールで描き直しても今のヒットは残る
    scrollBody(root, { top: bodyOf(root).scrollTop + 5 });
    await wait(60);
    expect(current(root)).toEqual([String(index)]);
  });

  it('見えていない列のヒットは横にも動かす（右端に寄せる）', async () => {
    const { root, view, matches } = await bigSheet();
    view.setHighlights(matches);
    const index = matches.findIndex((m: TextMatch) => m.r === 5 && m.c === 20);
    view.focusHighlight(index);
    // 文字の幅を測れない（jsdom）ときはセルの範囲：T 列の右端 20 × 64 を、セルの部分の右端（760px）に
    expect(bodyOf(root).scrollLeft).toBe(64 * 20 - (800 - 40));
    expect(bodyOf(root).scrollTop).toBe(0);
  });

  it('smooth は scrollTo を 1 回だけ呼ぶ', async () => {
    const { root, view, matches } = await bigSheet();
    view.setHighlights(matches);
    const body = bodyOf(root);
    const spy = vi.fn();
    body.scrollTo = spy as unknown as HTMLElement['scrollTo'];
    view.focusHighlight(matches.length - 2, { smooth: true });
    expect(spy).toHaveBeenCalledTimes(1);
    expect(spy.mock.calls[0]![0]).toMatchObject({ behavior: 'smooth', left: 0 });
  });

  it('番号の範囲外は RangeError。ハイライトが無ければ -1 だけ', async () => {
    const { view, matches } = await bigSheet();
    expect(() => view.focusHighlight(0)).toThrow(RangeError);
    expect(() => view.focusHighlight(-1)).not.toThrow();
    view.setHighlights(matches);
    expect(() => view.focusHighlight(matches.length)).toThrow(RangeError);
    expect(() => view.focusHighlight(-2)).toThrow(RangeError);
    expect(() => view.focusHighlight(0.5)).toThrow(RangeError);
  });

  it('setHighlights で一覧を替えると今のヒットは外れる', async () => {
    const { root, view, matches } = await bigSheet();
    view.setHighlights(matches);
    view.focusHighlight(0, { scroll: false });
    view.setHighlights(matches);
    expect(current(root)).toEqual([]);
  });
});

/** bigSheet の表示の寸法（表の行数 300 → 行番号の列 40px） */
const view0 = () => view();

describe('scrollToCell', () => {
  it('縦は上から 1/3、横は見えていないときだけ', async () => {
    const { root, view } = await bigSheet();
    view.scrollToCell(150, 3);
    expect(bodyOf(root).scrollTop).toBe(verticalTarget(geo, view0(), 150));
    expect(bodyOf(root).scrollLeft).toBe(0);
    view.scrollToCell(150, 25);
    expect(bodyOf(root).scrollLeft).toBe(64 * 25 - 760);
    expect(root.querySelector('tbody tr[aria-rowindex="151"]')).not.toBeNull();
    expect(() => view.scrollToCell(0, 1)).toThrow(RangeError);
    expect(() => view.scrollToCell(1, 1.5)).toThrow(RangeError);
  });
});

describe('cellAt・highlightAt', () => {
  it('寸法から行・列を出す（スクロールも見る）。見出しの上は null', async () => {
    const { view } = await bigSheet();
    expect(view.cellAt(40 + 5, 24 + 25)).toEqual({ r: 2, c: 1 });
    expect(view.cellAt(40 + 70, 24 + 25)).toEqual({ r: 2, c: 1 }); // B2 は A2 の文字のはみ出し
    expect(view.cellAt(10, 100)).toBeNull();
    expect(view.cellAt(100, 10)).toBeNull();
    view.scrollToCell(200, 1);
    const top = verticalTarget(geo, view0(), 200);
    expect(view.cellAt(41, 24 + 1)).toEqual({ r: Math.floor(top / 20) + 1, c: 1 });
  });

  it('結合は左上、はみ出した文字の上ははみ出し元', async () => {
    const book = await openBook({
      sheets: [
        {
          name: 'S',
          rows: [
            ['はみ出す長い文字', null, null, 'x'],
            [null, '結合', null],
            [null, null, null],
          ],
          merges: ['B2:C3'],
        },
      ],
    });
    const { view } = await mount(book);
    await wait();
    const at = (c: number, r: number) => view.cellAt(40 + 64 * (c - 1) + 5, 24 + 20 * (r - 1) + 5);
    expect(at(2, 1)).toEqual({ r: 1, c: 1 });
    expect(at(3, 1)).toEqual({ r: 1, c: 1 });
    expect(at(4, 1)).toEqual({ r: 1, c: 4 });
    expect(at(3, 3)).toEqual({ r: 2, c: 2 });
    expect(at(1, 3)).toEqual({ r: 3, c: 1 });
  });

  it('highlightAt は座標のいちばん上の片の番号（無ければ -1）', async () => {
    const { root, view, matches } = await bigSheet();
    view.setHighlights(matches);
    const piece = root.querySelector<HTMLElement>('.xs-hit[data-xs-hit="2"]')!;
    const spy = vi.fn(() => [piece, piece.parentElement!]);
    document.elementsFromPoint = spy as unknown as Document['elementsFromPoint'];
    expect(view.highlightAt(60, 60)).toBe(2);
    spy.mockReturnValue([piece.closest('td')!]);
    expect(view.highlightAt(60, 60)).toBe(-1);
    // 見出しの上は調べない
    spy.mockReturnValue([piece]);
    expect(view.highlightAt(10, 60)).toBe(-1);
    view.setHighlights(null);
    expect(view.highlightAt(60, 60)).toBe(-1);
    Reflect.deleteProperty(document, 'elementsFromPoint');
  });
});

describe('renderWorkbook', () => {
  it('中のシートの表示に渡す', async () => {
    const { root, view } = await mountWorkbook({
      sheets: [{ name: 'A', rows: Array.from({ length: 300 }, (_, i) => [`v${i + 1}`]) }],
    });
    await wait();
    view.setHighlights([{ r: 250, c: 1, start: 0, end: 1 }]);
    view.focusHighlight(0);
    const body = root.querySelector<HTMLElement>('.xs-body')!;
    expect(body.scrollTop).toBeGreaterThan(0);
    expect(root.querySelector('.xs-hit.is-current')).not.toBeNull();
    expect(view.cellAt(40 + 5, 24 + 5)).not.toBeNull();
  });

  it('開けなかったブック：判定は null・-1、番号の確かめだけは同じ', async () => {
    const { view } = await mountWorkbook(new Uint8Array([1, 2, 3]));
    expect(view.cellAt(100, 100)).toBeNull();
    expect(view.highlightAt(100, 100)).toBe(-1);
    expect(() => view.focusHighlight(-1)).not.toThrow();
    expect(() => view.focusHighlight(0)).toThrow(RangeError);
    expect(() => view.scrollToCell(1, 1)).not.toThrow();
    expect(() => view.scrollToCell(0, 1)).toThrow(RangeError);
  });
});

describe('横のスクロール先：左へ動かすとき', () => {
  it('範囲が枠に収まればセルの左端から見せる。収まらなければ範囲の左端', () => {
    const v = view({ scrollLeft: 3000 });
    expect(horizontalTarget(geo, v, { left: 100, right: 130, cellLeft: 64 })).toBe(64);
    expect(horizontalTarget(geo, v, { left: 1000, right: 1030, cellLeft: 64 })).toBe(1000);
  });
});
