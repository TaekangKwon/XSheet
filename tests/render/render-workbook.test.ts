// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { MESSAGES_EN, MESSAGES_JA, formatMessage } from '../../src/render/messages';
import { renderWorkbook } from '../../src/render/render-workbook';
import { buildXlsx, buildZip } from '../../src/testing/make-xlsx';
import { cleanupViews, dataRows, mountWorkbook, sheetTabs, wait } from './render-sheet.helpers';

/**
 * `renderWorkbook` の DOM 描画の振る舞いテスト（InterHub `spreadsheet.render.test.ts` のうち、タブ・シートタブ列・
 * シートが無いブック・開けないファイルの表示を写したもの。ISSUE-82・ADR-0025）。
 * jsdom はレイアウトを持たない（幅 0）ため、タブ列のあふれ・スクロールは目視で確かめ、ここでは構造・文言・切替を確かめる。
 * 文言は既定（`MESSAGES_JA`）。
 */

afterEach(() => {
  cleanupViews();
});

/** 表示中のシートの A1 の文字 */
const a1 = (root: HTMLElement): string | null | undefined =>
  dataRows(root)[0]?.cells[1]?.textContent;

describe('renderWorkbook（タブ）', () => {
  it('TD-82-2: シートタブを作り、非表示シートは ⌀ 付きで最初の表示シートを出す。タブを押すと開き直して差し替える', async () => {
    const { root } = await mountWorkbook({
      sheets: [
        { name: 'Hid', state: 'hidden', rows: [['h']] },
        { name: 'One', rows: [['first']] },
        { name: 'Two', rows: [['second']] },
      ],
    });
    const tabs = sheetTabs(root);
    expect(tabs.map((b) => b.textContent)).toEqual(['Hid ⌀', 'One', 'Two']);
    expect(tabs.map((b) => b.type)).toEqual(['button', 'button', 'button']);
    expect(tabs.map((b) => b.classList.contains('is-hidden-sheet'))).toEqual([true, false, false]);
    expect(tabs.map((b) => b.classList.contains('is-active'))).toEqual([false, true, false]);
    expect(a1(root)).toBe('first');
    tabs[2]!.click();
    await wait();
    expect(a1(root)).toBe('second');
    expect(tabs.map((b) => b.classList.contains('is-active'))).toEqual([false, false, true]);
    // 非表示のシートもタブで開ける
    tabs[0]!.click();
    await wait();
    expect(a1(root)).toBe('h');
  });

  it('全部非表示のブックは先頭のシートを出す', async () => {
    const { root } = await mountWorkbook({
      sheets: [
        { name: 'H1', state: 'hidden', rows: [['h1']] },
        { name: 'H2', state: 'veryHidden', rows: [['h2']] },
      ],
    });
    expect(a1(root)).toBe('h1');
    expect(sheetTabs(root).map((b) => b.classList.contains('is-active'))).toEqual([true, false]);
  });

  it('根は container の子の div.xs-workbook（container には何も付けない）。子はタブ列・シートの表示の順', async () => {
    const { root, container } = await mountWorkbook({ sheets: [{ name: 'S', rows: [['a']] }] });
    expect(container.getAttribute('class')).toBeNull();
    expect(root.parentElement).toBe(container);
    expect([...root.children].map((c) => c.className)).toEqual(['xs-sheet-strip', 'xs-sheet']);
  });

  it('シートタブ列は ◀ ▶ ボタンとスクロール領域を持ち、シート数ぶんのタブがスクロール領域の中に並ぶ', async () => {
    const { root, view } = await mountWorkbook({
      sheets: Array.from({ length: 10 }, (_, i) => ({ name: `S${i + 1}`, rows: [[`v${i + 1}`]] })),
    });
    const strip = root.querySelector<HTMLElement>('.xs-sheet-strip')!;
    // 子は ◀・スクロール領域・▶ の順
    const kids = [...strip.children] as HTMLElement[];
    expect(kids.map((k) => k.className)).toEqual(['xs-sheet-nav', 'xs-sheet-tabs', 'xs-sheet-nav']);
    expect(kids[0]?.textContent).toBe('◀');
    expect(kids[2]?.textContent).toBe('▶');
    expect(kids[0]?.getAttribute('aria-label')).toBe(MESSAGES_JA.tabsPrev);
    expect(kids[2]?.getAttribute('aria-label')).toBe(MESSAGES_JA.tabsNext);
    expect(kids[0]?.title).toBe(MESSAGES_JA.tabsPrev);
    // タブは全部スクロール領域の中（矢印は xs-sheet-tab ではない）
    expect(kids[1]!.querySelectorAll('.xs-sheet-tab')).toHaveLength(10);
    expect(strip.querySelectorAll('.xs-sheet-tab')).toHaveLength(10);
    // jsdom は幅 0（あふれ無し）なので矢印は隠れる。destroy で例外なく後始末できる
    expect((kids[0] as HTMLButtonElement).hidden).toBe(true);
    expect((kids[2] as HTMLButtonElement).hidden).toBe(true);
    expect(() => view.destroy()).not.toThrow();
  });

  it('文言を替えると ◀ ▶ の aria-label・非表示の注記も替わる', async () => {
    const { root } = await mountWorkbook(
      {
        sheets: [
          { name: 'A', rows: [['a']] },
          { name: 'H', state: 'hidden', rows: [['h']] },
        ],
      },
      { messages: MESSAGES_EN },
    );
    const navs = [...root.querySelectorAll<HTMLButtonElement>('.xs-sheet-nav')];
    expect(navs.map((n) => n.getAttribute('aria-label'))).toEqual([
      MESSAGES_EN.tabsPrev,
      MESSAGES_EN.tabsNext,
    ]);
    expect(sheetTabs(root)[1]!.title).toBe(`H\n${MESSAGES_EN.hiddenSheet}`);
  });

  it('シートタブは title にシート名全文を持つ（長い名前が省略表示されても分かる）', async () => {
    const longName = 'とても長いシート名_0123456789_abcdefghijklmnopqrstuvwxyz';
    const { root } = await mountWorkbook({
      sheets: [
        { name: longName, rows: [['a']] },
        { name: 'Hid', state: 'hidden', rows: [['h']] },
      ],
    });
    const tabs = sheetTabs(root);
    expect(tabs[0]?.title).toBe(longName);
    // 非表示シートは名前に加えて注記も title に残る
    expect(tabs[1]?.title).toBe(`Hid\n${MESSAGES_JA.hiddenSheet}`);
  });

  it('シート名は textContent・title だけ（HTML 風の名前が要素にならない）', async () => {
    const name = '<img src=x onerror=1>';
    const { root } = await mountWorkbook({ sheets: [{ name, rows: [['a']] }] });
    expect(root.querySelector('img')).toBeNull();
    expect(sheetTabs(root)[0]!.textContent).toBe(name);
  });
});

describe('renderWorkbook（showSheet）', () => {
  it('showSheet は SheetInfo.index でシートを替え、タブも合わせる。ブックに無い番号は RangeError', async () => {
    const { root, view } = await mountWorkbook({
      sheets: [
        { name: 'A', rows: [['a']] },
        { name: 'B', rows: [['b']] },
      ],
    });
    await view.showSheet(2);
    expect(a1(root)).toBe('b');
    expect(sheetTabs(root).map((b) => b.classList.contains('is-active'))).toEqual([false, true]);
    await expect(view.showSheet(3)).rejects.toBeInstanceOf(RangeError);
    await expect(view.showSheet(0)).rejects.toBeInstanceOf(RangeError);
    expect(a1(root)).toBe('b');
  });

  it('表示中のシートを押し直しても開き直さない（表はそのまま）', async () => {
    const { root, view } = await mountWorkbook({
      sheets: [
        { name: 'A', rows: [['a']] },
        { name: 'B', rows: [['b']] },
      ],
    });
    await wait();
    const table = root.querySelector('table');
    sheetTabs(root)[0]!.click();
    await view.showSheet(1);
    await wait();
    expect(root.querySelector('table')).toBe(table);
  });

  it('destroy の後の showSheet は何もしない', async () => {
    const { root, view, container } = await mountWorkbook({
      sheets: [
        { name: 'A', rows: [['a']] },
        { name: 'B', rows: [['b']] },
      ],
    });
    view.destroy();
    view.destroy(); // 2 回目は何もしない
    await view.showSheet(2);
    expect(container.childElementCount).toBe(0);
    expect(root.childElementCount).toBe(0);
  });
});

describe('renderWorkbook（開けない・シートが無い）', () => {
  it('TD-82-2: シートが 0 枚なら noSheets（showSheet は RangeError・destroy で根が外れる）', async () => {
    const { root, view, container } = await mountWorkbook({ sheets: [] });
    expect(root.querySelector('.xs-empty')?.textContent).toBe(MESSAGES_JA.noSheets);
    expect(root.querySelector('.xs-sheet-strip')).toBeNull();
    await expect(view.showSheet(1)).rejects.toBeInstanceOf(RangeError);
    view.destroy();
    expect(container.childElementCount).toBe(0);
  });

  it('TD-82-4: 暗号化（OLE 先頭）・壊れた zip・Strict 形式・上限超えを、例外にせず種類ごとの文言の div.xs-banner にする', async () => {
    const banner = async (bytes: Uint8Array, limits?: { maxEntries: number }): Promise<string> => {
      const { root } = await mountWorkbook(bytes, limits ? { limits } : undefined);
      expect(root.querySelector('.xs-sheet')).toBeNull();
      expect(root.querySelector('.xs-sheet-strip')).toBeNull();
      return root.querySelector(':scope > .xs-banner')?.textContent ?? '';
    };
    expect(await banner(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0, 0, 0]))).toBe(
      MESSAGES_JA.encrypted,
    );
    const parseFailed = formatMessage(MESSAGES_JA.parseFailed, { detail: '' });
    expect(await banner(new TextEncoder().encode('plain text'))).toBe(
      `${parseFailed}xlsx: not-zip`,
    );
    const strict = await buildXlsx({ sheets: [{ name: 'S', rows: [['a']] }], strict: true });
    expect(await banner(strict)).toBe(MESSAGES_JA.unsupported);
    const brokenZip = await buildZip({ 'hello.txt': 'no workbook here' });
    expect((await banner(brokenZip)).startsWith(`${parseFailed}xlsx: broken`)).toBe(true);
    // WorkbookOptions.limits は openXlsx に渡る
    const plain = await buildXlsx({ sheets: [{ name: 'S', rows: [['a']] }] });
    expect(await banner(plain, { maxEntries: 2 })).toBe(MESSAGES_JA.tooLarge);
  });

  it('ArrayBuffer も受け取る', async () => {
    const bytes = await buildXlsx({ sheets: [{ name: 'S', rows: [['ab']] }] });
    const container = document.createElement('div');
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    const view = await renderWorkbook(container, buffer as ArrayBuffer);
    expect(container.querySelector('.xs-cell')?.textContent).toBe('1'); // 行番号
    expect(dataRows(container.querySelector('.xs-sheet')!)[0]?.cells[1]?.textContent).toBe('ab');
    view.destroy();
    expect(container.childElementCount).toBe(0);
  });
});

describe('シートタブ列の ResizeObserver 通知（次フレームへ送る）', () => {
  /** テストから呼べる ResizeObserver のスタブ。観察する要素ごとにコールバックを保持する。 */
  let observed: { cb: () => void; target: Element }[];
  /** rAF のスタブ。予約されたコールバックを id 付きで溜め、手動で実行する。 */
  let rafQueue: Array<{ id: number; cb: FrameRequestCallback }>;
  let rafSpy: ReturnType<typeof vi.fn>;
  let cancelSpy: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    observed = [];
    rafQueue = [];
    let nextId = 1;
    class FakeResizeObserver {
      constructor(private readonly cb: () => void) {}
      observe(target: Element): void {
        observed.push({ cb: this.cb, target });
      }
      unobserve(): void {}
      disconnect(): void {}
    }
    rafSpy = vi.fn((cb: FrameRequestCallback) => {
      const id = nextId++;
      rafQueue.push({ id, cb });
      return id;
    });
    cancelSpy = vi.fn((id: number) => {
      rafQueue = rafQueue.filter((f) => f.id !== id);
    });
    vi.stubGlobal('ResizeObserver', FakeResizeObserver);
    vi.stubGlobal('requestAnimationFrame', rafSpy);
    vi.stubGlobal('cancelAnimationFrame', cancelSpy);
  });

  afterEach(() => {
    cleanupViews();
    vi.unstubAllGlobals();
  });

  /**
   * 2 枚シートの .xlsx を描画し、タブ列をあふれ状態にしてタブ列の RO コールバックと ◀▶ を返す。
   * シートの表示もスクロール枠を観察するため、タブ列（`.xs-sheet-tabs`）を観察するものだけを呼ぶ。
   */
  async function renderOverflowing(): Promise<{
    notify: () => void;
    navs: HTMLButtonElement[];
    destroy: () => void;
  }> {
    const { root, view } = await mountWorkbook({
      sheets: [
        { name: 'A', rows: [['a']] },
        { name: 'B', rows: [['b']] },
      ],
    });
    const tabs = root.querySelector<HTMLElement>('.xs-sheet-tabs')!;
    const navs = [...root.querySelectorAll<HTMLButtonElement>('.xs-sheet-nav')];
    // 生成直後は幅 0（あふれ無し）なので矢印は隠れている
    expect(navs.map((n) => n.hidden)).toEqual([true, true]);
    Object.defineProperty(tabs, 'scrollWidth', { configurable: true, value: 500 });
    Object.defineProperty(tabs, 'clientWidth', { configurable: true, value: 100 });
    const mine = observed.filter((o) => o.target === tabs);
    expect(mine).toHaveLength(1);
    return { notify: () => mine.forEach((o) => o.cb()), navs, destroy: () => view.destroy() };
  }

  it('通知では ◀▶ を同期更新せず rAF へ 1 回だけ予約し、次フレームで表示を切り替える', async () => {
    const { notify, navs } = await renderOverflowing();
    // 描画時に予約された rAF（グリッドの描画など）は対象外。通知後に増えた分だけを見る
    const before = rafSpy.mock.calls.length;
    notify();
    expect(navs.map((n) => n.hidden)).toEqual([true, true]);
    expect(rafSpy.mock.calls.length - before).toBe(1);
    notify();
    expect(rafSpy.mock.calls.length - before).toBe(1);
    // rAF の id は呼び出し順の連番なので、before より大きい id が通知後の予約
    const pendingNow = rafQueue.filter((f) => f.id > before);
    expect(pendingNow).toHaveLength(1);
    pendingNow.forEach((f) => f.cb(0));
    expect(navs.map((n) => n.hidden)).toEqual([false, false]);
  });

  it('予約が残ったまま後始末すると cancelAnimationFrame で取り消す', async () => {
    const { notify, destroy } = await renderOverflowing();
    notify();
    const pendingId = rafQueue[rafQueue.length - 1]!.id;
    cancelSpy.mockClear();
    destroy();
    expect(cancelSpy).toHaveBeenCalledWith(pendingId);
  });
});
