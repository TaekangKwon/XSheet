// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { DrawingNotice, XlsxBook } from '../../src/core/index';
import { MESSAGES_JA } from '../../src/render/messages';
import { makePng } from '../../src/testing/make-image';
import {
  borderXml,
  fontXml,
  oneCellAnchor,
  patternFillXml,
  picXml,
  shapeXml,
  twoCellAnchor,
  wsDr,
  type CellSpec,
  type SheetSpec,
  type XlsxSpec,
} from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  mountWorkbook,
  sheetTabs,
  stubClientSize,
  wait,
} from './render-sheet.helpers';

/**
 * `renderWorkbook` のシートの切替（タブ）と後始末の DOM テスト。InterHub `spreadsheet.*.render.test.ts` のうち、
 * タブの切替を含むもの（罫線の作り直し・書式の `<style>` を作り直さない・blob URL の revoke・通知がシートごと）を写し、
 * 表示を壊したときにリスナ・ResizeObserver・rAF・blob URL が残らず、ブックが閉じることを足した。
 * `openXlsx` を包んで、開いたブックを捕まえ、図の通知・画像の結果を差し替える（InterHub の `vi.mock('../xlsx')` と同じ形）。
 */

const hooks = vi.hoisted(() => ({
  notices: null as DrawingNotice[] | null,
  readImage: null as XlsxBook['readImage'] | null,
  /** 開いたブック（閉じたかを確かめる） */
  books: [] as XlsxBook[],
}));

vi.mock('../../src/core/index', async (importOriginal) => {
  const orig = await importOriginal<typeof import('../../src/core/index')>();
  return {
    ...orig,
    openXlsx: async (...args: Parameters<typeof orig.openXlsx>) => {
      const book = await orig.openXlsx(...args);
      hooks.books.push(book);
      if (!hooks.notices && !hooks.readImage) {
        return book;
      }
      const notices = hooks.notices;
      const readImage = hooks.readImage;
      const wrapped: XlsxBook = {
        sheets: book.sheets,
        openSheet: (i) => book.openSheet(i),
        openDrawings: async (i) => {
          const r = await book.openDrawings(i);
          return notices ? { ...r, notices } : r;
        },
        readImage: readImage ?? ((i, p) => book.readImage(i, p)),
        readChart: (i, p) => book.readChart(i, p),
        close: () => book.close(),
      };
      return wrapped;
    },
  };
});

const PNG = makePng(40, 30);
const A1 = { col: 0, row: 0 };

/** 画像（image1.png）を指す drawing を持つシート */
function sheetWith(name: string, anchors: string[], rows: SheetSpec['rows'] = [['a']]): SheetSpec {
  return {
    name,
    rows,
    drawing: { xml: wsDr(anchors), rels: [{ id: 'rId1', target: '../media/image1.png' }] },
  };
}

/** A1 に画像 1 枚（名前は alt） */
const pic = (n: string): string =>
  oneCellAnchor(A1, 9525 * 40, 9525 * 30, picXml({ name: n, embed: 'rId1' }));

/** 画像付きのブックを描く */
async function renderImages(sheets: SheetSpec[]): Promise<HTMLElement> {
  return (
    await mountWorkbook({
      font: { name: 'Calibri', size: 11 },
      sheets,
      extraFiles: { 'xl/media/image1.png': PNG },
    })
  ).root;
}

/** blob URL の作成・解放の記録 */
const urls = { created: [] as string[], revoked: [] as string[] };
const savedUrl = { create: URL.createObjectURL, revoke: URL.revokeObjectURL };
let restoreSize: () => void;

beforeEach(() => {
  hooks.notices = null;
  hooks.readImage = null;
  hooks.books = [];
  urls.created = [];
  urls.revoked = [];
  let n = 0;
  URL.createObjectURL = vi.fn(() => {
    const u = `blob:r/${++n}`;
    urls.created.push(u);
    return u;
  });
  URL.revokeObjectURL = vi.fn((u: string) => {
    urls.revoked.push(u);
  });
  restoreSize = stubClientSize(800, 400);
});

afterEach(() => {
  cleanupViews();
  restoreSize();
  vi.unstubAllGlobals();
  URL.createObjectURL = savedUrl.create;
  URL.revokeObjectURL = savedUrl.revoke;
});

const figures = (root: HTMLElement): HTMLElement[] => [
  ...root.querySelectorAll<HTMLElement>('.xs-layer .xs-figure'),
];
const noticeOf = (root: HTMLElement): HTMLElement => root.querySelector<HTMLElement>('.xs-notice')!;
const alts = (root: HTMLElement): (string | undefined)[] =>
  figures(root).map((f) => f.querySelector('img')?.alt);

/** タブを押して、描き終わるまで待つ */
async function clickTab(root: HTMLElement, i: number): Promise<void> {
  sheetTabs(root)[i]!.click();
  await wait();
}

describe('renderWorkbook（タブの切替）', () => {
  it('TD-98-11: 罫線の無いシートでは中身が空（要素は残る）。シートを切り替えると作り直す', async () => {
    const cell = (v: string, s?: number): CellSpec => ({
      v,
      t: 'inlineStr',
      ...(s !== undefined ? { s } : {}),
    });
    const { root } = await mountWorkbook({
      font: { name: 'Calibri', size: 11 },
      sheets: [
        { name: 'S', rows: [0, 1, 2].map((i) => [cell(`r${i + 1}`, 1), cell('x', 2)]) },
        { name: 'T', rows: [[cell('plain')]] },
      ],
      xfs: [0, 0, 0],
      borders: [
        borderXml({ left: {} }),
        borderXml({ bottom: { style: 'thin', color: 'auto="1"' } }),
        borderXml({ right: { style: 'medium', color: 'rgb="FFFF0000"' } }),
      ],
      xfBorderIds: [0, 1, 2],
    });
    await wait();
    const paths = (): number => root.querySelectorAll('svg.xs-borders path').length;
    expect(paths()).toBeGreaterThan(0);
    await clickTab(root, 1);
    expect(root.querySelectorAll('svg.xs-borders')).toHaveLength(1);
    expect(paths()).toBe(0);
    await clickTab(root, 0);
    expect(paths()).toBeGreaterThan(0);
  });

  it('TD-100-14: シートを切り替えても <style> は作り直さない（同じ要素・同じ中身）。td のクラスは新しいシートの分が付く', async () => {
    const text = (v: string, s?: number): CellSpec => ({
      v,
      t: 'inlineStr',
      ...(s !== undefined ? { s } : {}),
    });
    const book: Partial<XlsxSpec> = {
      fonts: [fontXml('<sz val="11"/>'), fontXml('<b/>')],
      fills: [patternFillXml('none'), patternFillXml('solid', 'rgb="FFFF0000"')],
      xfs: [0, 0, 0],
      xfFontIds: [0, 0, 1],
      xfFillIds: [0, 1, 0],
    };
    const { root } = await mountWorkbook({
      sheets: [
        { name: 'S', rows: [[text('a', 1)]] },
        { name: 'T', rows: [[text('plain'), text('b', 2)]] },
      ],
      ...book,
    });
    await wait();
    const styles = (): HTMLStyleElement[] => [...root.querySelectorAll('style')];
    const style0 = styles()[0]!;
    const css0 = style0.textContent;
    await clickTab(root, 1);
    expect(styles()).toHaveLength(1);
    expect(styles()[0]).toBe(style0);
    expect(style0.textContent).toBe(css0);
    const tds = [
      ...root
        .querySelectorAll('table.xs-grid tbody tr:not(.xs-spacer)')[0]!
        .querySelectorAll('td:not(.xs-rownum)'),
    ];
    const xf = (td: Element): string[] => [...td.classList].filter((c) => c.startsWith('xs-xf-'));
    expect(xf(tds[0]!)).toEqual([]);
    expect(xf(tds[1]!)).toHaveLength(1);
  });

  it('TD-86-11: タブ切替で前のシートの blob URL を revoke する。戻ると読み直す（別のシートの画像が残らない）', async () => {
    const root = await renderImages([sheetWith('A', [pic('a')]), sheetWith('B', [pic('b')])]);
    await vi.waitFor(() => expect(urls.created).toEqual(['blob:r/1']));
    await clickTab(root, 1);
    await vi.waitFor(() => expect(urls.created).toEqual(['blob:r/1', 'blob:r/2']));
    expect(urls.revoked).toEqual(['blob:r/1']);
    expect(alts(root)).toEqual(['b']);
    await clickTab(root, 0);
    await vi.waitFor(() => expect(urls.created).toHaveLength(3));
    expect(urls.revoked).toEqual(['blob:r/1', 'blob:r/2']);
    expect(alts(root)).toEqual(['a']);
  });

  it('TD-86-12: 読み込み中にタブを素早く切り替えても、古いシートの URL が残らない（作った URL はすべて revoke 済みか、表示中の 1 つだけ）', async () => {
    const root = await renderImages([sheetWith('A', [pic('a')]), sheetWith('B', [pic('b')])]);
    const tabs = sheetTabs(root);
    // 待たずに連続で切り替える
    for (let i = 0; i < 6; i++) {
      tabs[(i + 1) % 2]!.click();
    }
    await wait(80);
    await vi.waitFor(() => expect(root.querySelectorAll('.xs-layer img')).toHaveLength(1));
    const live = urls.created.filter((u) => !urls.revoked.includes(u));
    expect(live.length).toBeLessThanOrEqual(1);
    const shown = [...root.querySelectorAll('.xs-layer img')].map((i) => i.getAttribute('src'));
    expect(shown.every((s) => live.includes(s!))).toBe(true);
    cleanupViews();
    expect(urls.revoked.slice().sort()).toEqual(urls.created.slice().sort());
  });

  it('TD-86-13: 画像の合計の上限超過（over-budget）の通知はタブ切替で消える', async () => {
    hooks.readImage = async () => ({ ok: false, reason: 'over-budget', label: null });
    const root = await renderImages([sheetWith('A', [pic('a')]), { name: 'B', rows: [['x']] }]);
    const bar = noticeOf(root);
    await vi.waitFor(() => expect(bar.hidden).toBe(false));
    expect(bar.textContent).toBe(MESSAGES_JA.noticeOverBudget);
    await clickTab(root, 1);
    expect(bar.hidden).toBe(true);
  });

  it('TD-88-20: text-truncated の通知はシートごと（別のシートで消え、戻ると再表示）', async () => {
    const big = 'あ'.repeat(10001);
    const root = await renderImages([
      sheetWith('A', [
        twoCellAnchor(
          { col: 0, row: 0 },
          { col: 2, row: 2 },
          shapeXml({ name: 'a', paras: [big] }),
        ),
      ]),
      { name: 'B', rows: [['x']] },
    ]);
    const bar = noticeOf(root);
    expect(bar.textContent).toBe(MESSAGES_JA.noticeTextTruncated);
    await clickTab(root, 1);
    expect(bar.hidden).toBe(true);
    await clickTab(root, 0);
    expect(bar.hidden).toBe(false);
    expect(bar.textContent).toBe(MESSAGES_JA.noticeTextTruncated);
  });
});

describe('renderWorkbook（後始末）', () => {
  it('TD-82-5: 画像を読み込んでいる間に destroy しても、URL を残さず、ブックを閉じ、根を外す', async () => {
    const { root, view, container } = await mountWorkbook({
      font: { name: 'Calibri', size: 11 },
      sheets: [sheetWith('A', [pic('a')])],
      extraFiles: { 'xl/media/image1.png': PNG },
    });
    view.destroy(); // 画像の読み込みはまだ終わっていない
    await wait(60);
    expect(urls.revoked.slice().sort()).toEqual(urls.created.slice().sort());
    expect(container.childElementCount).toBe(0);
    expect(root.querySelector('img')).toBeNull();
    await expect(hooks.books[0]!.openSheet(1)).rejects.toThrow();
  });

  it('TD-82-5: タブの切替の途中で destroy しても、後から結果を描かない', async () => {
    const { root, view } = await mountWorkbook({
      sheets: [
        { name: 'A', rows: [['a']] },
        { name: 'B', rows: [['b']] },
      ],
    });
    const pending = view.showSheet(2);
    view.destroy();
    await pending;
    await wait();
    expect(root.childElementCount).toBe(0);
    expect(root.querySelector('table')).toBeNull();
  });

  it('タブを替えてから destroy すると、blob URL・リスナ・ResizeObserver・rAF が残らず、根は外れ、ブックは閉じる', async () => {
    // 付けたリスナ（target・type・listener）。外したものは消す
    let listeners: { target: EventTarget; type: string; listener: unknown }[] = [];
    const origAdd = EventTarget.prototype.addEventListener;
    const origRemove = EventTarget.prototype.removeEventListener;
    EventTarget.prototype.addEventListener = function (
      this: EventTarget,
      type: string,
      listener: EventListenerOrEventListenerObject | null,
      options?: AddEventListenerOptions | boolean,
    ) {
      listeners.push({ target: this, type, listener });
      origAdd.call(this, type, listener, options);
    };
    EventTarget.prototype.removeEventListener = function (
      this: EventTarget,
      type: string,
      listener: EventListenerOrEventListenerObject | null,
      options?: EventListenerOptions | boolean,
    ) {
      listeners = listeners.filter(
        (l) => !(l.target === this && l.type === type && l.listener === listener),
      );
      origRemove.call(this, type, listener, options);
    };
    // ResizeObserver の偽物（observe・disconnect を数える）
    const ro = { created: 0, observing: 0 };
    class FakeResizeObserver {
      private count = 0;
      constructor(_cb: ResizeObserverCallback) {
        ro.created++;
      }
      observe(): void {
        this.count++;
        ro.observing++;
      }
      unobserve(): void {}
      disconnect(): void {
        ro.observing -= this.count;
        this.count = 0;
      }
    }
    vi.stubGlobal('ResizeObserver', FakeResizeObserver);
    // 予約中の rAF（実行・取り消しで消える）
    const pendingFrames = new Set<number>();
    const raf = window.requestAnimationFrame.bind(window);
    const caf = window.cancelAnimationFrame.bind(window);
    vi.stubGlobal('requestAnimationFrame', (cb: FrameRequestCallback) => {
      const id = raf((t) => {
        pendingFrames.delete(id);
        cb(t);
      });
      pendingFrames.add(id);
      return id;
    });
    vi.stubGlobal('cancelAnimationFrame', (id: number) => {
      pendingFrames.delete(id);
      caf(id);
    });
    try {
      const { root, view, container } = await mountWorkbook({
        font: { name: 'Calibri', size: 11 },
        sheets: [sheetWith('A', [pic('a')]), sheetWith('B', [pic('b')])],
        extraFiles: { 'xl/media/image1.png': PNG },
      });
      await vi.waitFor(() => expect(urls.created).toHaveLength(1));
      sheetTabs(root)[1]!.click();
      await vi.waitFor(() => expect(urls.created).toHaveLength(2));
      await wait();
      // タブ列とスクロール枠を観察する。window・document にはリスナを付けない
      expect(ro).toEqual({ created: 2, observing: 2 });
      expect(listeners.filter((l) => l.target === window || l.target === document)).toEqual([]);
      expect(listeners.length).toBeGreaterThan(0);

      view.destroy();
      expect(container.childElementCount).toBe(0);
      expect(urls.revoked.slice().sort()).toEqual(urls.created.slice().sort());
      expect(listeners).toEqual([]);
      expect(ro.observing).toBe(0);
      expect(pendingFrames.size).toBe(0);
      expect(hooks.books).toHaveLength(1);
      await expect(hooks.books[0]!.openSheet(1)).rejects.toThrow();
    } finally {
      EventTarget.prototype.addEventListener = origAdd;
      EventTarget.prototype.removeEventListener = origRemove;
    }
  });
});
