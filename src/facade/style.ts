// 正: docs/adr/0003-universal-distribution.md §2
/**
 * `view` の表示の CSS・テーマ・高さ（M6-2）。
 * - CSS：`style.css` の文面を `<style data-xsheet-style>` で 1 回だけ入れる。`<link>`・import で読み込み済みなら入れない
 *   （試しの要素に `.xs-workbook` の規則（display: flex）が効くかで見る。変数は継承され、使う側が `:root` に書くこともあるため、変数では見ない）。
 *   container が Shadow DOM の中なら、その ShadowRoot に入れる（文書の CSS は Shadow DOM の中に届かない）。
 *   書式の `<style>` と同じく nonce を付ける（CSP の要件は今と変わらない）。
 * - テーマ：表示の根に `data-xs-theme` を付ける。ダークの値は style.css が持つ。
 * - 高さ：container の高さが 0 のときだけ、表示の根に既定の高さを付けて console.warn を 1 回出す（P5）。
 *   高さは表示の根（`.xs-viewer`）を隠して測る（前の表示の高さを container の高さと取り違えない）。
 *   container が描かれていない（display: none の中）ときは、描かれたときに ResizeObserver で測る。
 *   高さは CSSOM（`el.style.height`）で付ける（style 属性の文字列は CSP の unsafe-inline が要るため使わない）。
 */
import css from '../style.css?inline';

/** テーマ。`auto` は `prefers-color-scheme` に従う */
export type XSheetTheme = 'light' | 'dark' | 'auto';

/** container の高さが 0 のときに付ける高さ */
export const DEFAULT_HEIGHT = '480px';

/** 自動で入れた `<style>` の印 */
const STYLE_MARK = 'data-xsheet-style';

let warnedHeight = false;

/** CSS を入れる先：container が Shadow DOM の中ならその ShadowRoot、そうでなければ文書の head */
function styleHost(container: HTMLElement): ParentNode {
  const root = container.getRootNode();
  // ShadowRoot（nodeType 11 で host を持つ）。別の realm でも動くよう instanceof で見ない
  if (root.nodeType === 11 && 'host' in root) return root as ShadowRoot;
  const doc = container.ownerDocument;
  return doc.head ?? doc.documentElement;
}

/** 表示の CSS が container に効いているか（試しの要素を足して、すぐ外す。同期なので描画されない） */
function styleApplied(container: HTMLElement): boolean {
  const view = container.ownerDocument.defaultView;
  if (!view) return false;
  const probe = container.ownerDocument.createElement('div');
  probe.className = 'xs-workbook';
  container.append(probe);
  const display = view.getComputedStyle(probe).display;
  probe.remove();
  return display === 'flex';
}

/** 表示の CSS が無ければ入れる（読み込み済み・入れ済みなら何もしない） */
export function ensureStyle(container: HTMLElement, nonce?: string): void {
  const host = styleHost(container);
  if (host.querySelector(`style[${STYLE_MARK}]`) !== null) return;
  // ShadowRoot の中は、入れ済みの印だけを見る（文書の CSS は届かないが、jsdom などは届いたように見せるため試さない）
  if (host === container.ownerDocument.head || host === container.ownerDocument.documentElement) {
    if (styleApplied(container)) return;
  }
  const style = container.ownerDocument.createElement('style');
  style.setAttribute(STYLE_MARK, '');
  if (nonce !== undefined) style.setAttribute('nonce', nonce);
  style.textContent = css;
  host.append(style);
}

/** テーマを表示の根に付ける（`light` は何も付けず、使う側の `--xs-*` の上書きがそのまま効く） */
export function applyTheme(root: HTMLElement, theme: XSheetTheme | undefined): void {
  if (theme === 'dark' || theme === 'auto') root.setAttribute('data-xs-theme', theme);
}

/** container の高さ（中の表示の根を隠して測る） */
function containerHeight(container: HTMLElement): number {
  const roots = [...container.children].filter(
    (el): el is HTMLElement => el.classList.contains('xs-viewer') && 'style' in el,
  );
  const saved = roots.map((el) => el.style.display);
  roots.forEach((el) => (el.style.display = 'none'));
  const height = container.clientHeight;
  roots.forEach((el, i) => (el.style.display = saved[i]!));
  return height;
}

/** 高さが無いときの既定の高さを付ける */
function useDefaultHeight(root: HTMLElement): void {
  root.style.height = DEFAULT_HEIGHT;
  if (!warnedHeight) {
    warnedHeight = true;
    console.warn(
      `XSheet: the container has no height, so the viewer uses ${DEFAULT_HEIGHT}. Give the container a height (or pass the height option) to change it`,
    );
  }
}

/**
 * 表示の根の高さを決める。
 * - 数値は px、文字列は CSS の長さとしてそのまま付ける。
 * - `auto`（既定）は container の高さが 0 のときだけ {@link DEFAULT_HEIGHT} を付ける（それ以外は style.css の 100%）。
 *   container が描かれていないときは、描かれたときに決める。
 * @returns 後始末（描かれるのを待つ ResizeObserver を外す）
 */
export function applyHeight(
  container: HTMLElement,
  root: HTMLElement,
  height: number | string | undefined,
): () => void {
  const none = (): void => undefined;
  if (typeof height === 'number') {
    root.style.height = `${height}px`;
    return none;
  }
  if (height !== undefined && height !== 'auto') {
    root.style.height = height;
    return none;
  }
  const decide = (): void => {
    if (containerHeight(container) === 0) useDefaultHeight(root);
  };
  const Observer = container.ownerDocument.defaultView?.ResizeObserver;
  if (container.getClientRects().length > 0 || Observer === undefined) {
    decide();
    return none;
  }
  // display: none の中（隠れたタブなど）：描かれたときに 1 回だけ決める
  const observer = new Observer(() => {
    if (container.getClientRects().length === 0) return;
    observer.disconnect();
    decide();
  });
  observer.observe(container);
  return () => observer.disconnect();
}
