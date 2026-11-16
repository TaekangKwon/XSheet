// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { findText } from '../../src/core/index';
import type { GridRow } from '../../src/render/grid';
import { resolveDrawnCell } from '../../src/render/navigate';
import type { CellSpec } from '../../src/testing/make-xlsx';
import { cleanupViews, mount, openBook, stubClientSize, wait } from './render-sheet.helpers';

/**
 * 配置に合わせた位置の判定と移動の横の位置（M7-14・ADR-0008 §3・§5・Moobius の依頼 R9）。
 */

describe('左・両側へのはみ出しの判定', () => {
  const t = (text: string, kind?: 's' | 'n') => ({ text, kind, colSpan: 1, rowSpan: 1 });
  // A に数値（はみ出さない）・B〜C 空・D に右寄せ（書式 0）・E 空・F に中央（書式 1）・G〜H 空・I に値
  const rows: GridRow[] = [
    {
      r: 1,
      cells: [
        t('1', 'n'),
        undefined,
        undefined,
        t('right', 's'),
        undefined,
        t('center', 's'),
        undefined,
        undefined,
        t('i', 's'),
      ],
      formats: [undefined, undefined, undefined, 0, undefined, 1],
    },
  ];
  const formats = [{ hAlign: 'right' as const }, { hAlign: 'center' as const }];

  it('右寄せの元は右から、中央の元は両側から引く', () => {
    expect(resolveDrawnCell(rows, 1, 2, 10, 1, formats)).toEqual({ r: 1, c: 4 });
    expect(resolveDrawnCell(rows, 1, 3, 10, 1, formats)).toEqual({ r: 1, c: 4 });
    // E は D（右寄せ）の右なので D からは届かず、F（中央）の左へのはみ出し
    expect(resolveDrawnCell(rows, 1, 5, 10, 1, formats)).toEqual({ r: 1, c: 6 });
    expect(resolveDrawnCell(rows, 1, 7, 10, 1, formats)).toEqual({ r: 1, c: 6 });
    expect(resolveDrawnCell(rows, 1, 8, 10, 1, formats)).toEqual({ r: 1, c: 6 });
  });

  it('書式が無ければ今までどおり（左寄せ）。描いた範囲の外の元は数えない', () => {
    expect(resolveDrawnCell(rows, 1, 2, 10)).toEqual({ r: 1, c: 2 });
    expect(resolveDrawnCell(rows, 1, 2, 10, 1, formats, 3)).toEqual({ r: 1, c: 2 });
  });
});

let restoreSize: () => void;
beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
  // 測る入れ物の幅を「文字数 × 10px」にする（jsdom はレイアウトを持たない）
  vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (
    this: HTMLElement,
  ) {
    const width = this.closest('.xs-measure') ? (this.textContent ?? '').length * 10 : 0;
    return { x: 0, y: 0, left: 0, top: 0, right: width, bottom: 0, width, height: 0 } as DOMRect;
  });
});
afterEach(() => {
  cleanupViews();
  restoreSize();
  vi.restoreAllMocks();
});

/** T5（20 列目）に「abcdefghij key」（14 文字＝140px）。xf：0 なし・1 右・2 中央・3 折り返し */
async function sheetWith(xf: number) {
  const rows: CellSpec[][] = Array.from({ length: 300 }, () => []);
  rows[0] = Array.from({ length: 30 }, (_, i) => (i === 0 ? 'x' : null));
  rows[4] = Array.from({ length: 20 }, (_, i) =>
    i === 19 ? { v: 'abcdefghij key', t: 'inlineStr', s: xf } : null,
  );
  const book = await openBook({
    sheets: [{ name: 'S', rows }],
    xfs: [0, 0, 0, 0],
    xfAlignments: [undefined, 'horizontal="right"', 'horizontal="center"', 'wrapText="1"'],
  });
  const m = await mount(book);
  await wait();
  const { matches } = await findText(await book.openSheet(1), 'key');
  m.view.setHighlights(matches);
  m.view.focusHighlight(0);
  return m.root.querySelector<HTMLElement>('.xs-body')!.scrollLeft;
}

describe('移動の横の位置（配置）', () => {
  // T 列は 1216〜1280px（入れ物の右端 1279px）。表は枠を埋める 19 行なので行番号の列は 32px、セルの部分の幅は 768px。
  // 一致は文字の 110〜140px
  it('左寄せ：セルの左端＋3px から', async () => {
    expect(await sheetWith(0)).toBe(1216 + 3 + 140 - 768);
  });

  it('右寄せ：入れ物の右端−3px−文字の幅から', async () => {
    expect(await sheetWith(1)).toBe(1279 - 3 - 140 + 140 - 768);
  });

  it('中央：セルの中心−文字の幅の半分から', async () => {
    expect(await sheetWith(2)).toBe(Math.ceil((1216 + 1279) / 2 - 70 + 140) - 768);
  });

  it('折り返し：セルの範囲', async () => {
    expect(await sheetWith(3)).toBe(1280 - 768);
  });
});
