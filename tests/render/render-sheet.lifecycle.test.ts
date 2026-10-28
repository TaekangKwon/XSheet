// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { openXlsx, type XlsxBook } from '../../src/core/index';
import { MESSAGES_JA } from '../../src/render/messages';
import { renderSheet } from '../../src/render/render-sheet';
import { createSheetView } from '../../src/render/sheet-view';
import { makePng } from '../../src/testing/make-image';
import {
  buildXlsx,
  oneCellAnchor,
  picXml,
  wsDr,
  type SheetSpec,
} from '../../src/testing/make-xlsx';
import { scrollBody, stubClientSize, wait } from './render-sheet.helpers';

/**
 * 表示のインスタンスの寿命（M2-3 の完了条件）：同じページに 2 つの表示を同時に作って壊すと、
 * blob URL・リスナ・ResizeObserver・rAF の漏れが 0 で、根は container から外れ、ブックは使えるまま。
 * InterHub はモジュール変数で表示を 1 つだけ持っていたため、ここは新しく足したテスト。
 */

const PNG = makePng(40, 30);

/** 画像 1 枚と 300 行のシート */
function picSheet(name: string): SheetSpec {
  return {
    name,
    rows: [[name], ...Array.from({ length: 299 }, (_, i) => [`${name}${i + 2}`])],
    drawing: {
      xml: wsDr([
        oneCellAnchor({ col: 0, row: 1 }, 9525 * 40, 9525 * 30, picXml({ name, embed: 'rId1' })),
      ]),
      rels: [{ id: 'rId1', target: '../media/image1.png' }],
    },
  };
}

async function twoSheetBook(): Promise<XlsxBook> {
  const bytes = await buildXlsx({
    font: { name: 'Calibri', size: 11 },
    sheets: [picSheet('A'), picSheet('B')],
    extraFiles: { 'xl/media/image1.png': PNG },
  });
  return openXlsx(bytes);
}

/** blob URL の作成・解放の記録 */
const urls = { created: [] as string[], revoked: [] as string[] };
const savedUrl = { create: URL.createObjectURL, revoke: URL.revokeObjectURL };

/** 付けたリスナ（target・type・listener）。外したものは消す */
let listeners: { target: EventTarget; type: string; listener: unknown }[] = [];
const origAdd = EventTarget.prototype.addEventListener;
const origRemove = EventTarget.prototype.removeEventListener;

/** ResizeObserver の偽物（作った数・observe・disconnect を数える） */
const ro = { created: 0, observing: 0, disconnected: 0 };
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
    ro.disconnected++;
  }
}

/** 予約中の rAF（実行・取り消しで消える） */
const pendingFrames = new Set<number>();

let restoreSize: () => void;

beforeEach(() => {
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
  listeners = [];
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
  ro.created = 0;
  ro.observing = 0;
  ro.disconnected = 0;
  vi.stubGlobal('ResizeObserver', FakeResizeObserver);
  pendingFrames.clear();
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
  restoreSize = stubClientSize(800, 400);
});

afterEach(() => {
  restoreSize();
  vi.unstubAllGlobals();
  EventTarget.prototype.addEventListener = origAdd;
  EventTarget.prototype.removeEventListener = origRemove;
  URL.createObjectURL = savedUrl.create;
  URL.revokeObjectURL = savedUrl.revoke;
  document.body.replaceChildren();
});

describe('renderSheet（表示のインスタンスの寿命）', () => {
  it('2 つの表示を同時に作って壊すと、blob URL・リスナ・ResizeObserver・rAF が残らず、根は外れ、ブックは使えるまま', async () => {
    const book = await twoSheetBook();
    const ca = document.createElement('div');
    const cb = document.createElement('div');
    document.body.append(ca, cb);
    const [va, vb] = await Promise.all([renderSheet(ca, book, 1), renderSheet(cb, book, 2)]);
    // 画像の読み込み・取り付け後の rAF（並行して走る他のテストで遅れることがあるため、待つ条件で待つ）
    await vi.waitFor(() => expect(urls.created).toHaveLength(2));
    await wait();
    const ra = ca.querySelector<HTMLElement>('.xs-sheet')!;
    const rb = cb.querySelector<HTMLElement>('.xs-sheet')!;
    // 表示は互いに独立（別の連番・別のシート・それぞれの画像）
    expect(ra.getAttribute('data-xs-fmt')).not.toBe(rb.getAttribute('data-xs-fmt'));
    expect(ra.querySelector('tbody tr:not(.xs-spacer) td:nth-child(2)')?.textContent).toBe('A');
    expect(rb.querySelector('tbody tr:not(.xs-spacer) td:nth-child(2)')?.textContent).toBe('B');
    const urlA = ra.querySelector('.xs-layer img')!.getAttribute('src');
    expect(urls.created).toContain(urlA);
    expect(ro).toEqual({ created: 2, observing: 2, disconnected: 0 });
    // window・document にはリスナを付けない。付けたのはスクロール枠の scroll だけ
    expect(listeners.filter((l) => l.target === window || l.target === document)).toEqual([]);
    expect(listeners.map((l) => l.type).filter((t) => t === 'scroll')).toHaveLength(2);

    // 片方を壊しても、もう片方は動く
    va.destroy();
    expect(ca.childElementCount).toBe(0);
    expect(urls.revoked).toEqual([urlA]);
    const table = rb.querySelector('table');
    scrollBody(rb, { top: 4000 });
    await vi.waitFor(() => expect(rb.querySelector('table')).not.toBe(table));
    expect(rb.querySelectorAll('.xs-layer img').length).toBe(0); // 画像は可視範囲の外

    vb.destroy();
    vb.destroy();
    expect(cb.childElementCount).toBe(0);
    expect(urls.revoked.slice().sort()).toEqual(urls.created.slice().sort());
    expect(listeners.filter((l) => l.type === 'scroll')).toEqual([]);
    expect(ro).toEqual({ created: 2, observing: 0, disconnected: 2 });
    expect(pendingFrames.size).toBe(0);

    // ブックは閉じていない（renderSheet の destroy は book を閉じない）
    const sh = await book.openSheet(1);
    expect(sh.getRows(1, 1)[0]?.cells[0]?.text).toBe('A');
    const again = await renderSheet(ca, book, 2);
    expect(
      ca.querySelector('.xs-sheet tbody tr:not(.xs-spacer) td:nth-child(2)')?.textContent,
    ).toBe('B');
    again.destroy();
    book.close();
  });

  it('同じ表示でシートを替えても書式の <style> は同じ（ブックで 1 回）。前のシートの blob URL は revoke（renderWorkbook の下地）', async () => {
    const book = await twoSheetBook();
    const container = document.createElement('div');
    document.body.append(container);
    const view = createSheetView(container, book, { messages: MESSAGES_JA, drawings: true });
    await view.showSheet(1);
    await vi.waitFor(() => expect(urls.created).toEqual(['blob:r/1']));
    const style = view.root.querySelector('style')!;
    const css = style.textContent;
    expect(view.activeIndex).toBe(1);
    expect(urls.created).toEqual(['blob:r/1']);
    await view.showSheet(2);
    await vi.waitFor(() => expect(urls.created).toEqual(['blob:r/1', 'blob:r/2']));
    expect(view.activeIndex).toBe(2);
    expect(view.root.querySelectorAll('style')).toHaveLength(1);
    expect(view.root.querySelector('style')).toBe(style);
    expect(style.textContent).toBe(css);
    expect(urls.revoked).toEqual(['blob:r/1']);
    expect(urls.created).toEqual(['blob:r/1', 'blob:r/2']);
    expect(view.root.querySelector('.xs-layer img')?.getAttribute('alt')).toBe('B');
    await expect(view.showSheet(3)).rejects.toBeInstanceOf(RangeError);
    view.destroy();
    expect(urls.revoked).toEqual(['blob:r/1', 'blob:r/2']);
    await view.showSheet(1); // 壊した後は何もしない
    expect(container.childElementCount).toBe(0);
    book.close();
  });
});
