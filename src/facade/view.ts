// 正: docs/adr/0003-universal-distribution.md §2
/**
 * 1 行で使える入口 `view`・`read`（M6-2）。core（`openXlsx`）と render（`renderWorkbook`）の上の薄い層で、
 * 写したファイル（`src/core/`・`src/render/`）の作りには立ち入らない。
 * - `view`：xlsx をタブ付きで出す。表示の根 `div.xs-viewer` を container の子に足し、その中に `renderWorkbook` の表示を置く。
 *   言語・テーマ・CSS・高さを語で選べる。開けないファイルは今までどおり例外にせず、帯（`div.xs-banner`）に文言を出す。
 * - `read`：DOM なしで開く（`openXlsx` と同じ。受け取れる形を `view` と揃える）。
 */
import { openXlsx, type SheetInfo, type XlsxBook, type XlsxLimits } from '../core/index';
import {
  MESSAGES_EN,
  MESSAGES_JA,
  MESSAGES_KO,
  renderWorkbook,
  type XSheetMessages,
} from '../render/index';
import { resolveMessages } from '../render/messages';
import { toSourceBytes, type XSheetSource } from './source';
import { applyHeight, applyTheme, ensureStyle, type XSheetTheme } from './style';

/** 文言の言語。`auto` は `<html lang>`、無ければブラウザの言語。合うものが無ければ英語 */
export type XSheetLang = 'ja' | 'en' | 'ko' | 'auto';

/** `read` の指定 */
export interface ReadOptions {
  /** 読み取りの上限を替える（既定は `DEFAULT_LIMITS`） */
  limits?: Partial<XlsxLimits>;
}

/** `view` の指定（どれも省略できる） */
export interface ViewOptions extends ReadOptions {
  /** 文言の言語（既定 `auto`） */
  lang?: XSheetLang;
  /** 文言を一部だけ替える（`lang` で選んだ表の上に重ねる） */
  messages?: Partial<XSheetMessages>;
  /** テーマ（既定 `light`）。`auto` は `prefers-color-scheme` に従う */
  theme?: XSheetTheme;
  /** 表示の CSS を自動で入れるか（既定 `auto`＝読み込み済みでなければ入れる。`none` は入れない） */
  css?: 'auto' | 'none';
  /** 表示の高さ（既定 `auto`＝container の高さが 0 のときだけ 480px）。数値は px、文字列は CSS の長さ */
  height?: number | string;
  /** 図の層（画像・図形・線）を出すか（既定 true） */
  drawings?: boolean;
  /** シートタブの列を置く場所（既定 `'top'`＝表の上。`'bottom'` は表の下） */
  tabPosition?: 'top' | 'bottom';
  /** 自動で入れる `<style>`・書式の `<style>` に付ける nonce（CSP の `style-src 'nonce-…'` 用） */
  nonce?: string;
}

/** `view` の戻り値 */
export interface Viewer {
  /** シートの一覧（タブの並び。非表示のシートも入る。開けなかったファイルでは空） */
  readonly sheets: readonly SheetInfo[];
  /**
   * シートを替える。
   * @param sheet シートの番号（`SheetInfo.index`・1 始まり）か名前
   * @throws RangeError ブックに無い番号・名前
   */
  showSheet(sheet: number | string): Promise<void>;
  /** 表示を外し、ブック・リスナ・blob URL を放す。何度呼んでもよい */
  destroy(): void;
}

const TABLES: Readonly<Record<'ja' | 'en' | 'ko', Readonly<XSheetMessages>>> = {
  ja: MESSAGES_JA,
  en: MESSAGES_EN,
  ko: MESSAGES_KO,
};

/** 言語の名前（`ko-KR` など）から表を選ぶ。合わなければ null */
function tableFor(tag: string | null | undefined): Readonly<XSheetMessages> | null {
  const primary = (tag ?? '').trim().toLowerCase().split(/[-_]/)[0];
  return primary === 'ja' || primary === 'en' || primary === 'ko' ? TABLES[primary] : null;
}

/**
 * 使う文言の表（`lang` で選んだ表に `messages` を重ねる）。
 * `auto` は `<html lang>` があればそれだけで決め（合わなければ英語）、無いときだけブラウザの言語を見る。
 * 型を通らない呼び出し（`<script>` から `'ko-KR'`・`'zh'` など）も、名前から選ぶか `auto` と同じにする。
 */
function pickMessages(doc: Document, options: ViewOptions): Readonly<XSheetMessages> {
  const lang: unknown = options.lang ?? 'auto';
  let base = lang === 'auto' || typeof lang !== 'string' ? null : tableFor(lang);
  if (base === null) {
    const pageLang = doc.documentElement.lang.trim();
    const nav = doc.defaultView?.navigator;
    base =
      pageLang !== ''
        ? (tableFor(pageLang) ?? MESSAGES_EN)
        : (tableFor(nav?.languages?.[0]) ?? tableFor(nav?.language) ?? MESSAGES_EN);
  }
  return resolveMessages(options.messages, base);
}

/** 同じ container の今の表示（新しい `view` が前の表示を外す。読み込み中に次の `view` が来たら、前の読み込みは捨てる） */
const current = new WeakMap<HTMLElement, { token: object; viewer: Viewer | null }>();

/** 捨てた読み込みの戻り値（表示は無い） */
const DISCARDED: Viewer = Object.freeze({
  sheets: Object.freeze([]),
  showSheet: () => Promise.reject(new RangeError('XSheet: the viewer was replaced or destroyed')),
  destroy: () => undefined,
});

/** 要素か CSS セレクタを要素にする */
function resolveTarget(target: HTMLElement | string): HTMLElement {
  if (typeof target !== 'string') {
    if (typeof target !== 'object' || target === null || target.nodeType !== 1) {
      throw new TypeError('XSheet: the target must be an element or a CSS selector');
    }
    return target;
  }
  const el = document.querySelector<HTMLElement>(target);
  if (el === null) throw new TypeError(`XSheet: no element matches the selector ${target}`);
  return el;
}

/**
 * xlsx をタブ付きで表示する。
 * 同じ container に続けて呼ぶと、前の表示を外してから出す（読み込み中の前の呼び出しは、表示の無い Viewer を返す）。
 * @param target 表示を置く要素か CSS セレクタ（子として `div.xs-viewer` を足す）
 * @param source File・Blob・ArrayBuffer・型付き配列・Response、またはその Promise（`fetch(url)` をそのまま渡せる）
 * @throws TypeError 要素が見つからない・受けない形の source・失敗した Response
 */
export async function view(
  target: HTMLElement | string,
  source: XSheetSource,
  options: ViewOptions = {},
): Promise<Viewer> {
  const container = resolveTarget(target);
  const token = {};
  current.get(container)?.viewer?.destroy();
  current.set(container, { token, viewer: null });
  const isCurrent = (): boolean => current.get(container)?.token === token;
  let bytes;
  try {
    bytes = await toSourceBytes(source);
  } catch (e) {
    if (isCurrent()) current.delete(container);
    throw e;
  }
  if (!isCurrent()) return DISCARDED;
  if (options.css !== 'none') ensureStyle(container, options.nonce);
  const root = container.ownerDocument.createElement('div');
  root.className = 'xs-viewer';
  applyTheme(root, options.theme);
  const disposeHeight = applyHeight(container, root, options.height);
  container.append(root);

  let workbook;
  try {
    workbook = await renderWorkbook(root, bytes, {
      messages: pickMessages(container.ownerDocument, options),
      drawings: options.drawings,
      tabPosition: options.tabPosition,
      nonce: options.nonce,
      limits: options.limits,
    });
  } catch (e) {
    disposeHeight();
    root.remove();
    if (isCurrent()) current.delete(container);
    throw e;
  }
  const wb = workbook;
  if (!isCurrent()) {
    disposeHeight();
    wb.destroy();
    root.remove();
    return DISCARDED;
  }
  const destroy = (): void => {
    disposeHeight();
    wb.destroy();
    root.remove();
    if (current.get(container)?.viewer === viewer) current.delete(container);
  };
  const viewer: Viewer = {
    sheets: wb.sheets,
    showSheet(sheet: number | string): Promise<void> {
      const info = wb.sheets.find((s) => (typeof sheet === 'number' ? s.index : s.name) === sheet);
      if (info === undefined) return Promise.reject(new RangeError('XSheet: no such sheet'));
      return wb.showSheet(info.index);
    },
    destroy,
  };
  current.set(container, { token, viewer });
  return viewer;
}

/**
 * xlsx を DOM なしで開く（`openXlsx` と同じ。閉じるのは呼び出し側の `close()`）。
 * @throws TypeError 受けない形の source・失敗した Response
 * @throws XlsxError 開けないファイル（`kind` で種類が分かる）
 */
export async function read(source: XSheetSource, options: ReadOptions = {}): Promise<XlsxBook> {
  return openXlsx(await toSourceBytes(source), { limits: options.limits });
}
