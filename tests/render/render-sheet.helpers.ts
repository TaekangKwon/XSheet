// renderSheet の DOM テスト（render-sheet*.test.ts）の共通の部品。jsdom の環境で使う。
import { openXlsx, type XlsxBook } from '../../src/core/index';
import { renderSheet, type RenderOptions, type SheetView } from '../../src/render/render-sheet';
import { buildXlsx, type XlsxSpec } from '../../src/testing/make-xlsx';

/** Uint8Array → ArrayBuffer */
export function toBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

/** 表示 1 つ（後始末の対象） */
export interface Mounted {
  /** 表示の根（`div.xs-sheet`） */
  root: HTMLElement;
  /** 表示を置いた要素（document.body の子） */
  container: HTMLElement;
  view: SheetView;
  book: XlsxBook;
}

const live: Mounted[] = [];

/** 合成のブックを開く */
export async function openBook(spec: XlsxSpec): Promise<XlsxBook> {
  return openXlsx(await buildXlsx(spec));
}

/** 新しい container（document.body の子）に、ブックのシートを表示する（後始末は {@link cleanupViews}） */
export async function mount(book: XlsxBook, index = 1, options?: RenderOptions): Promise<Mounted> {
  const container = document.createElement('div');
  document.body.append(container);
  const view = await renderSheet(container, book, index, options);
  const root = container.querySelector<HTMLElement>(':scope > .xs-sheet')!;
  const m = { root, container, view, book };
  live.push(m);
  return m;
}

/** 合成のブックを開いて 1 シートを表示し、表示の根を返す */
export async function renderSpec(
  spec: XlsxSpec,
  index = 1,
  options?: RenderOptions,
): Promise<HTMLElement> {
  return (await mount(await openBook(spec), index, options)).root;
}

/** 作った表示をすべて壊し、ブックを閉じ、container を外す（afterEach で呼ぶ） */
export function cleanupViews(): void {
  const books = new Set<XlsxBook>();
  for (const m of live.splice(0)) {
    m.view.destroy();
    books.add(m.book);
    m.container.remove();
  }
  books.forEach((b) => b.close());
}

/**
 * すべての要素の clientWidth・clientHeight を固定する（jsdom はレイアウトを持たない）。
 * @returns 元に戻す関数
 */
export function stubClientSize(width: number, height: number): () => void {
  const proto = HTMLElement.prototype;
  const w = Object.getOwnPropertyDescriptor(proto, 'clientWidth');
  const h = Object.getOwnPropertyDescriptor(proto, 'clientHeight');
  Object.defineProperty(proto, 'clientWidth', { configurable: true, get: () => width });
  Object.defineProperty(proto, 'clientHeight', { configurable: true, get: () => height });
  return () => {
    if (w) Object.defineProperty(proto, 'clientWidth', w);
    else Reflect.deleteProperty(proto, 'clientWidth');
    if (h) Object.defineProperty(proto, 'clientHeight', h);
    else Reflect.deleteProperty(proto, 'clientHeight');
  };
}

/** 次のフレームまで待つ（それまでに予約された rAF は先に動く） */
export const nextFrame = (): Promise<void> => new Promise((r) => requestAnimationFrame(() => r()));

/**
 * 待つ（rAF・画像の読み込み）。先に次のフレームを待つので、呼ぶ前に予約された描画（スクロール・取り付け後の描き直し）は
 * 並行して走るほかのテストで遅れても必ず済んでいる。
 */
export async function wait(ms = 30): Promise<void> {
  await nextFrame();
  await new Promise((r) => setTimeout(r, ms));
}

/** px 文字列（`145px`）を数値にする */
export const px = (v: string): number => Number.parseFloat(v);

/** 1 行目の見出しセル（th）の文字を並べる */
export function headTexts(root: HTMLElement): string[] {
  return [...root.querySelectorAll('thead th')].map((th) => th.textContent ?? '');
}

/** 本文の実データ行（余白行を除く） */
export function dataRows(root: HTMLElement): HTMLTableRowElement[] {
  return [...root.querySelectorAll<HTMLTableRowElement>('tbody tr:not(.xs-spacer)')];
}

/** スクロール枠を動かし、scroll を送る */
export function scrollBody(root: HTMLElement, pos: { top?: number; left?: number }): void {
  const body = root.querySelector<HTMLElement>('.xs-body')!;
  if (pos.top !== undefined) body.scrollTop = pos.top;
  if (pos.left !== undefined) body.scrollLeft = pos.left;
  body.dispatchEvent(new Event('scroll'));
}
