// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { findText, type TextMatch } from '../../src/core/index';
import { createGridHits, decorateCell } from '../../src/render/highlight';
import { fontXml, patternFillXml, type CellSpec, type XlsxSpec } from '../../src/testing/make-xlsx';
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
 * 検索の一致のハイライト（M7-5・ADR-0007 §2）。jsdom は CSS とレイアウトを持たないので、DOM の形
 * （`span.xs-hits` 1 つ・片・番号・描き直しで残る・寸法が変わらない）を確かめ、色と見た目は examples で目視する。
 */

let restoreSize: () => void;
beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
});
afterEach(() => {
  cleanupViews();
  restoreSize();
});

const m = (r: number, c: number, start: number, end: number): TextMatch => ({ r, c, start, end });

/** 行 r（1 始まり）・列 c（1 始まり）の td（行番号の td・余白の行を除く） */
function cellAt(root: HTMLElement, r: number, c: number): HTMLTableCellElement | undefined {
  const tr = root.querySelector<HTMLTableRowElement>(`tbody tr[aria-rowindex="${r + 1}"]`);
  return tr
    ? [...tr.querySelectorAll<HTMLTableCellElement>('td:not(.xs-rownum)')][c - 1]
    : undefined;
}

/** セルの片を「番号:文字」の並びにする */
const pieces = (el: Element | undefined) =>
  [...(el?.querySelectorAll<HTMLElement>('.xs-hit') ?? [])].map(
    (p) => `${p.dataset.xsHit}:${p.textContent}`,
  );

describe('decorateCell（DOM の形）', () => {
  const box = (text: string) => {
    const el = document.createElement('div');
    el.className = 'xs-cell';
    el.textContent = text;
    return el;
  };

  it('文字を一致の範囲で分け、全体を span.xs-hits 1 つで包む（空白も文字のまま残る）', () => {
    const el = box(' ab  ab ');
    const hits = createGridHits([m(1, 1, 1, 3), m(1, 1, 5, 7)])!;
    decorateCell(el, ' ab  ab ', ' ab  ab ', hits.at(1, 1), -1);
    expect(el.children).toHaveLength(1);
    const wrap = el.firstElementChild!;
    expect(wrap.className).toBe('xs-hits');
    expect(el.textContent).toBe(' ab  ab ');
    expect([...wrap.childNodes].map((n) => n.textContent)).toEqual([' ', 'ab', '  ', 'ab', ' ']);
    expect(pieces(el)).toEqual(['0:ab', '1:ab']);
  });

  it('今のヒットに is-current。重なる一致は 1 つの片（番号は先頭）', () => {
    const el = box('abcdef');
    decorateCell(
      el,
      'abcdef',
      'abcdef',
      createGridHits([m(1, 1, 0, 3), m(1, 1, 2, 5)])!.at(1, 1),
      1,
    );
    expect(pieces(el)).toEqual(['0:abcde']);
    expect(el.querySelector('.xs-hit')!.classList.contains('is-current')).toBe(true);
  });

  it('### のセルはセル全体を光らせる（文字は変えない）', () => {
    const el = box('####');
    decorateCell(el, '123456', '####', createGridHits([m(1, 1, 2, 4)])!.at(1, 1), 0);
    expect(el.className).toBe('xs-cell xs-hit is-current');
    expect(el.dataset.xsHit).toBe('0');
    expect(el.textContent).toBe('####');
  });

  it('文字の範囲の外の一致・正しくない番号は付けない', () => {
    const el = box('ab');
    decorateCell(el, 'ab', 'ab', createGridHits([m(1, 1, 5, 7)])!.at(1, 1), -1);
    expect(el.querySelector('.xs-hits')).toBeNull();
    expect(
      createGridHits([m(0, 1, 0, 1), m(1, 1.5, 0, 1), m(1, 1, 2, 2)])!.at(1, 1),
    ).toBeUndefined();
    expect(createGridHits([])).toBeNull();
    expect(createGridHits(null)).toBeNull();
  });
});

describe('setHighlights（renderSheet）', () => {
  it('検索の一致を光らせ、null で消す。行の高さ・列の幅は変えない', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: [['見積 A', 'x'], ['再見積']] }] });
    const { root, view } = await mount(book);
    await wait();
    const widths = [...root.querySelectorAll<HTMLElement>('col')].map((c) => c.style.width);
    const heights = [...root.querySelectorAll<HTMLElement>('tbody tr')].map(
      (tr) => tr.style.height,
    );
    const { matches } = await findText(await book.openSheet(1), '見積');
    view.setHighlights(matches);
    expect(pieces(cellAt(root, 1, 1))).toEqual(['0:見積']);
    expect(cellAt(root, 1, 1)!.textContent).toBe('見積 A');
    expect(pieces(cellAt(root, 2, 1))).toEqual(['1:見積']);
    expect(cellAt(root, 1, 2)!.querySelector('.xs-hits')).toBeNull();
    expect([...root.querySelectorAll<HTMLElement>('col')].map((c) => c.style.width)).toEqual(
      widths,
    );
    expect(
      [...root.querySelectorAll<HTMLElement>('tbody tr')].map((tr) => tr.style.height),
    ).toEqual(heights);
    view.setHighlights(null);
    expect(root.querySelector('.xs-hits, .xs-hit')).toBeNull();
    view.setHighlights([m(2, 1, 0, 1)]);
    expect(pieces(cellAt(root, 2, 1))).toEqual(['0:再']);
    view.setHighlights([]);
    expect(root.querySelector('.xs-hit')).toBeNull();
  });

  it('スクロールで表を描き直しても残る（描いていなかった行の一致も付く）', async () => {
    const rows: CellSpec[][] = Array.from({ length: 300 }, (_, i) => [`r${i + 1} key`]);
    const book = await openBook({ sheets: [{ name: 'S', rows }] });
    const { root, view } = await mount(book);
    await wait();
    const { matches } = await findText(await book.openSheet(1), 'key');
    view.setHighlights(matches);
    expect(cellAt(root, 280, 1)).toBeUndefined(); // まだ描いていない
    const table0 = root.querySelector('table.xs-grid');
    scrollBody(root, { top: 280 * 20 });
    await wait(60);
    expect(root.querySelector('table.xs-grid')).not.toBe(table0);
    expect(pieces(cellAt(root, 280, 1))).toEqual(['279:key']);
    scrollBody(root, { top: 0 });
    await wait(60);
    expect(pieces(cellAt(root, 1, 1))).toEqual(['0:key']);
  });

  it('はみ出す文字・結合のセル・### のセル', async () => {
    const book = await openBook({
      sheets: [
        {
          name: 'S',
          rows: [
            ['とても長い文字がはみ出すセル', null, null],
            ['結合の左上', null],
            [12345678901234],
          ],
          merges: ['A2:B2'],
          cols: [{ min: 1, max: 1, width: 3 }],
        },
      ],
    });
    const { root, view } = await mount(book);
    await wait();
    view.setHighlights([m(1, 1, 3, 5), m(2, 1, 0, 2), m(3, 1, 0, 1)]);
    // はみ出し：span.xs-text の中を包む
    const spill = cellAt(root, 1, 1)!.querySelector('.xs-cell.is-spill > .xs-text > .xs-hits');
    expect(spill).not.toBeNull();
    expect(pieces(cellAt(root, 1, 1))).toEqual(['0:長い']);
    expect(pieces(cellAt(root, 2, 1))).toEqual(['1:結合']);
    const hashBox = cellAt(root, 3, 1)!.querySelector<HTMLElement>('.xs-cell')!;
    expect(hashBox.textContent).toMatch(/^#+$/);
    expect(hashBox.classList.contains('xs-hit')).toBe(true);
    expect(hashBox.dataset.xsHit).toBe('2');
  });

  it('塗りをまたぐ自動の色のはみ出し（background-clip: text）は、片に is-ink', async () => {
    const book: Partial<XlsxSpec> = {
      fonts: [fontXml('<sz val="11"/>')],
      fills: [patternFillXml('none'), patternFillXml('solid', 'rgb="FFFF0000"')],
      xfs: [0, 0],
      xfFontIds: [0, 0],
      xfFillIds: [0, 1],
    };
    const opened = await openBook({
      sheets: [
        { name: 'S', rows: [[{ v: '赤い塗りから出る文字', t: 'inlineStr', s: 1 }, null, null]] },
      ],
      ...book,
    });
    const { root, view } = await mount(opened);
    await wait();
    view.setHighlights([m(1, 1, 0, 2)]);
    const label = cellAt(root, 1, 1)!.querySelector<HTMLElement>('.xs-text')!;
    expect(label.style.getPropertyValue('background-clip')).toBe('text');
    const piece = label.querySelector<HTMLElement>('.xs-hit')!;
    expect(piece.classList.contains('is-ink')).toBe(true);
  });
});

describe('setHighlights（renderWorkbook）', () => {
  it('中のシートの表示に渡り、シートを替えると消える', async () => {
    const { root, view } = await mountWorkbook({
      sheets: [
        { name: 'A', rows: [['abc']] },
        { name: 'B', rows: [['abc']] },
      ],
    });
    await wait();
    view.setHighlights([m(1, 1, 1, 2)]);
    expect(pieces(cellAt(root, 1, 1))).toEqual(['0:b']);
    await view.showSheet(2);
    await wait();
    expect(root.querySelector('.xs-hit')).toBeNull();
  });

  it('開けなかったブックでも呼べる（何もしない）', async () => {
    const { view } = await mountWorkbook(new Uint8Array([1, 2, 3]));
    expect(() => view.setHighlights([m(1, 1, 0, 1)])).not.toThrow();
  });
});
