// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createGeometry, MESSAGES_JA } from '../../src/index';
import { copyRange, quoteCell, rangeText } from '../../src/render/selection-copy';
import type { Selection } from '../../src/render/selection';
import type { CellSpec } from '../../src/testing/make-xlsx';
import { cleanupViews, mount, openBook, stubClientSize } from './render-sheet.helpers';

/**
 * 選択のコピー（M7-26・ADR-0011 §5・Issue #57）。文字の形（区切り・囲み・結合・非表示・上限）を確かめ、
 * 表示では copy のイベントを送って clipboardData に書いた文字と、既定の動きを止めたかを確かめる。
 * Excel・メモ帳に貼った見え方は実機で確かめる（ユーザーの確認）。
 */

const MAX_C = 16_384;
const MAX_R = 1_048_576;

describe('quoteCell', () => {
  it('タブ・改行・" を含むセルだけ " で囲み、" は "" にする', () => {
    expect(quoteCell('abc')).toBe('abc');
    expect(quoteCell('')).toBe('');
    expect(quoteCell('a\tb')).toBe('"a\tb"');
    expect(quoteCell('a\nb')).toBe('"a\nb"');
    expect(quoteCell('a\r\nb')).toBe('"a\r\nb"');
    expect(quoteCell('say "hi"')).toBe('"say ""hi"""');
  });
});

describe('copyRange', () => {
  const used = { rows: 30, cols: 8 };
  it('行の全体は値のある最後の列まで、列の全体は値のある最後の行まで、セルはそのまま', () => {
    const row: Selection = {
      type: 'row',
      range: { r1: 3, c1: 1, r2: 4, c2: MAX_C },
      active: { r: 3, c: 1 },
    };
    const col: Selection = {
      type: 'column',
      range: { r1: 1, c1: 2, r2: MAX_R, c2: 2 },
      active: { r: 1, c: 2 },
    };
    const cell: Selection = {
      type: 'cell',
      range: { r1: 40, c1: 10, r2: 40, c2: 10 },
      active: { r: 40, c: 10 },
    };
    expect(copyRange(row, used)).toEqual({ r1: 3, c1: 1, r2: 4, c2: 8 });
    expect(copyRange(col, used)).toEqual({ r1: 1, c1: 2, r2: 30, c2: 2 });
    expect(copyRange(cell, used)).toEqual({ r1: 40, c1: 10, r2: 40, c2: 10 });
  });
});

describe('rangeText', () => {
  async function sheetOf(rows: CellSpec[][], extra: object = {}) {
    const book = await openBook({ sheets: [{ name: 'S', rows, ...extra }] });
    const sh = await book.openSheet(1);
    return { sh, geo: createGeometry(sh.layout, 100, 20) };
  }

  it('タブと \\r\\n で区切り、最後の行にも \\r\\n。空のセルは空', async () => {
    const { sh, geo } = await sheetOf([
      ['a', 1, null],
      ['x\ty', true, 'z'],
    ]);
    expect(rangeText(sh, { r1: 1, c1: 1, r2: 2, c2: 3 }, geo)).toBe(
      'a\t1\t\r\n"x\ty"\tTRUE\tz\r\n',
    );
  });

  it('結合したセルは左上に文字、残りは空（結合の左上が範囲より上なら空）', async () => {
    const { sh, geo } = await sheetOf(
      [
        ['m', null, 'c'],
        [null, null, 'd'],
        ['e', 'f', 'g'],
      ],
      { merges: ['A1:B2'] },
    );
    expect(rangeText(sh, { r1: 1, c1: 1, r2: 3, c2: 3 }, geo)).toBe(
      'm\t\tc\r\n\t\td\r\ne\tf\tg\r\n',
    );
    expect(rangeText(sh, { r1: 2, c1: 1, r2: 2, c2: 3 }, geo)).toBe('\t\td\r\n');
  });

  it('非表示の行・列は入れない', async () => {
    const { sh, geo } = await sheetOf(
      [
        ['a', 'b', 'c'],
        ['d', 'e', 'f'],
        ['g', 'h', 'i'],
      ],
      { cols: [{ min: 2, max: 2, hidden: true }], rowAttrs: { 2: { hidden: true } } },
    );
    expect(rangeText(sh, { r1: 1, c1: 1, r2: 3, c2: 3 }, geo)).toBe('a\tc\r\ng\ti\r\n');
  });

  it('見える行 × 見える列が上限を超えたら null、範囲が空なら空の文字', async () => {
    const { sh, geo } = await sheetOf([['a', 'b']]);
    expect(rangeText(sh, { r1: 1, c1: 1, r2: 3, c2: 2 }, geo, 5)).toBeNull();
    expect(rangeText(sh, { r1: 1, c1: 1, r2: 3, c2: 2 }, geo, 6)).toBe('a\tb\r\n\t\r\n\t\r\n');
    expect(rangeText(sh, { r1: 1, c1: 3, r2: 1, c2: 2 }, geo)).toBe('');
  });

  it('読む単位（512 行）の境をまたいでも行が抜けない・重ならない', async () => {
    const rows = Array.from({ length: 1100 }, (_, i) => [`r${i + 1}`]);
    const { sh } = await sheetOf(rows);
    const geo = createGeometry(sh.layout, 1100, 1);
    const text = rangeText(sh, { r1: 1, c1: 1, r2: 1100, c2: 1 }, geo)!;
    expect(text.split('\r\n').slice(0, -1)).toEqual(rows.map((r) => r[0]));
  });
});

describe('表示での Ctrl+C（copy のイベント）', () => {
  let restore: () => void;
  beforeEach(() => {
    restore = stubClientSize(800, 400);
  });
  afterEach(() => {
    cleanupViews();
    restore();
    document.body.replaceChildren();
    document.getSelection()?.removeAllRanges();
  });

  /** copy のイベントを送り、書いた文字と既定の動きを止めたかを返す */
  function copy(body: HTMLElement): { text: string | undefined; prevented: boolean } {
    const data = new Map<string, string>();
    const e = new Event('copy', { bubbles: true, cancelable: true });
    Object.defineProperty(e, 'clipboardData', {
      value: { setData: (type: string, v: string) => data.set(type, v) },
    });
    body.dispatchEvent(e);
    return { text: data.get('text/plain'), prevented: e.defaultPrevented };
  }

  async function setup(rows: CellSpec[][]) {
    const book = await openBook({ sheets: [{ name: 'S', rows }] });
    const m = await mount(book, 1, { selection: 'cell' });
    return { ...m, body: m.root.querySelector<HTMLElement>('.xs-body')! };
  }

  it('選択をタブ区切りの文字にして書き、selectionText と同じ', async () => {
    const { view, body } = await setup([
      ['a', 'b', 'c'],
      ['d', 'e', 'f'],
    ]);
    view.setSelection({ type: 'row', r: 2 });
    expect(copy(body)).toEqual({ text: 'd\te\tf\r\n', prevented: true });
    view.setSelection({ type: 'column', c: 2 });
    expect(copy(body).text).toBe('b\r\ne\r\n');
    expect(view.selectionText()).toBe('b\r\ne\r\n');
  });

  it('選択が無い・文字を選んでいるときは何もせず、ブラウザの既定のコピーに任せる', async () => {
    const { root, view, body } = await setup([['a', 'b']]);
    expect(copy(body)).toEqual({ text: undefined, prevented: false });
    expect(view.selectionText()).toBeNull();
    view.setSelection({ type: 'cell', r: 1, c: 1 });
    const range = document.createRange();
    range.selectNodeContents(root.querySelector('td:not(.xs-rownum) .xs-cell')!);
    document.getSelection()!.removeAllRanges();
    document.getSelection()!.addRange(range);
    expect(copy(body)).toEqual({ text: undefined, prevented: false });
  });

  it('上限を超えたらコピーせず（既定の動きも止める）、通知の帯に出す', async () => {
    const { root, view, body } = await setup([['a']]);
    // 列の全体は値のある最後の行（1 行）までなので、行の全体を 100 万行分選ぶ
    view.setSelection({ type: 'row', r: 1, r2: 1_000_001 });
    expect(copy(body)).toEqual({ text: undefined, prevented: true });
    const bar = root.querySelector<HTMLElement>('.xs-notice')!;
    expect(bar.hidden).toBe(false);
    expect(bar.textContent).toBe(MESSAGES_JA.noticeCopyTooLarge);
  });
});
