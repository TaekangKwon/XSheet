// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { CellSpec, XlsxSpec } from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  mount,
  mountWorkbook,
  openBook,
  sheetTabs,
  stubClientSize,
  wait,
} from './render-sheet.helpers';

/** 選んでいるシート `activeSheet` と、シートが替わった知らせ `onSheetChange`（M7-10・ADR-0007 §7 版 1.1） */

let restoreSize: () => void;
beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
});
afterEach(() => {
  cleanupViews();
  restoreSize();
  vi.useRealTimers();
});

const rows = (prefix: string): CellSpec[][] =>
  Array.from({ length: 5 }, (_, i) => [`${prefix}${i + 1}`]);

/** 1 枚目が非表示の 3 シート（最初に出るのは 2 枚目） */
const threeSheets: XlsxSpec = {
  sheets: [
    { name: 'H', rows: rows('h'), state: 'hidden' },
    { name: 'A', rows: rows('a') },
    { name: 'B', rows: rows('b') },
  ],
};

describe('renderWorkbook の activeSheet・onSheetChange', () => {
  it('最初は最初に出したシート。最初の表示では onSheetChange を呼ばない', async () => {
    const onSheetChange = vi.fn();
    const { view } = await mountWorkbook(threeSheets, { onSheetChange });
    expect(view.activeSheet).toBe(2);
    await wait();
    expect(onSheetChange).not.toHaveBeenCalled();
  });

  it('タブを押すと替わり、表示し終えた後に番号を知らせる', async () => {
    const onSheetChange = vi.fn();
    const { root, view } = await mountWorkbook(threeSheets, { onSheetChange });
    sheetTabs(root)[2]!.click();
    expect(view.activeSheet).toBe(3); // 押した時点で替わる
    await wait();
    expect(onSheetChange).toHaveBeenCalledTimes(1);
    expect(onSheetChange).toHaveBeenCalledWith(3);
  });

  it('showSheet でも知らせる。表示中のシートの showSheet では呼ばない', async () => {
    const onSheetChange = vi.fn();
    const { view } = await mountWorkbook(threeSheets, { onSheetChange });
    await view.showSheet(3);
    expect(onSheetChange).toHaveBeenCalledWith(3);
    await view.showSheet(3);
    expect(onSheetChange).toHaveBeenCalledTimes(1);
    expect(view.activeSheet).toBe(3);
  });

  it('後の切り替えに追い越された切り替えは知らせない', async () => {
    const onSheetChange = vi.fn();
    const { view } = await mountWorkbook(threeSheets, { onSheetChange });
    const first = view.showSheet(1);
    const second = view.showSheet(3);
    await Promise.all([first, second]);
    expect(view.activeSheet).toBe(3);
    expect(onSheetChange.mock.calls).toEqual([[3]]);
  });

  it('ブックに無い番号は RangeError で、選んでいるシートは替えない', async () => {
    const { view } = await mountWorkbook(threeSheets);
    await expect(view.showSheet(9)).rejects.toBeInstanceOf(RangeError);
    expect(view.activeSheet).toBe(2);
  });

  it('呼ぶ側の例外は showSheet を失敗にせず、setTimeout で投げ直す', async () => {
    const error = new Error('呼ぶ側の誤り');
    let calls = 0;
    const { view } = await mountWorkbook(threeSheets, {
      onSheetChange: () => {
        calls++;
        throw error;
      },
    });
    // 投げ直しの setTimeout だけを捕まえる（本物のタイマーで投げると vitest が未処理の例外として拾う）
    const later: (() => void)[] = [];
    const realSetTimeout = globalThis.setTimeout;
    const spy = vi.spyOn(globalThis, 'setTimeout').mockImplementation(((
      cb: () => void,
      ms?: number,
    ) => {
      if (ms === undefined) {
        later.push(cb);
        return 0;
      }
      return realSetTimeout(cb, ms);
    }) as typeof setTimeout);
    try {
      await expect(view.showSheet(3)).resolves.toBeUndefined();
    } finally {
      spy.mockRestore();
    }
    expect(calls).toBe(1);
    expect(later).toHaveLength(1);
    expect(() => later[0]!()).toThrow(error);
  });

  it('destroy の後は知らせない', async () => {
    const onSheetChange = vi.fn();
    const { view } = await mountWorkbook(threeSheets, { onSheetChange });
    const pending = view.showSheet(3);
    view.destroy();
    await pending;
    expect(onSheetChange).not.toHaveBeenCalled();
  });

  it('開けなかったブックでは null', async () => {
    const { view } = await mountWorkbook(new TextEncoder().encode('not a zip'));
    expect(view.activeSheet).toBeNull();
  });
});

describe('renderSheet の activeSheet', () => {
  it('表示したシート。showSheet で替わり、ブックに無い番号では替えない', async () => {
    const book = await openBook(threeSheets);
    const { view } = await mount(book, 2);
    expect(view.activeSheet).toBe(2);
    await view.showSheet(3);
    expect(view.activeSheet).toBe(3);
    await expect(view.showSheet(9)).rejects.toBeInstanceOf(RangeError);
    expect(view.activeSheet).toBe(3);
  });
});
