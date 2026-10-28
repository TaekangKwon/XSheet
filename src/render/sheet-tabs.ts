// 正: InterHub docs/plan/adr/0025-xlsx-own-reader.md（§決定 4）
/**
 * シートタブの列（InterHub `spreadsheet.ts` の `makeToggleButton`・`buildSheetTabStrip`。M2-4）。
 * Excel のシート見出しのように 1 行に並べ、入りきらない分は横スクロールさせる。`renderWorkbook` が使う。
 * シート名はユーザーのファイル内容（信頼できない入力）なので、`textContent` と `title` にだけ入れる。
 */
import type { XSheetMessages } from './messages';

/** シートタブ列（{@link buildSheetTabStrip} の戻り値）。 */
export interface SheetTabStrip {
  /** `div.xs-sheet-strip` 要素（◀ ＋ タブのスクロール領域 ＋ ▶）。表示へ取り付ける。 */
  element: HTMLElement;
  /** 指定のタブが見える位置までスクロールする。まだレイアウト前（幅 0）なら、幅が確定してから行う。 */
  reveal: (btn: HTMLElement | undefined) => void;
  /** ResizeObserver・イベントリスナ・rAF を外す。 */
  dispose: () => void;
}

/**
 * シートタブのボタンを作る（クリックの処理は呼び出し側で付ける）。
 * 長いシート名は CSS で省略表示されるため、全文を title で見られるようにする。
 * 非表示のシートは名前に ` ⌀` を足し、`is-hidden-sheet` と、title に注記（`hiddenSheet`）を付ける。
 */
export function makeSheetTab(
  name: string,
  hidden: boolean,
  messages: Readonly<XSheetMessages>,
): HTMLButtonElement {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'xs-sheet-tab';
  btn.textContent = name + (hidden ? ' ⌀' : '');
  btn.title = name;
  if (hidden) {
    btn.classList.add('is-hidden-sheet');
    // 非表示の注記にシート名全文も添える（省略表示されても名前が分かるように）。
    btn.title = `${name}\n${messages.hiddenSheet}`;
  }
  return btn;
}

/**
 * シートタブを 1 行に並べ、入りきらない分は横スクロールさせる列を作る。
 * 構造は `div.xs-sheet-strip` ＞ `[◀ button.xs-sheet-nav][div.xs-sheet-tabs（overflow-x:auto）＞ タブ…][▶ button.xs-sheet-nav]`。
 * - ◀▶ は押すとスクロール領域の可視幅の約 8 割ぶん滑らかに動かす。端に着いた側は disabled、全タブが収まるときは両方隠す。
 * - 状態はスクロール・リサイズ（ResizeObserver。無い環境ではガードして使わない）のたびに更新する。
 * - 縦のホイール回転は、あふれているときだけ横スクロールに変換する。
 *
 * @param buttons シート名のタブ（{@link makeSheetTab}。クリック処理は呼び出し側で設定済みのもの）
 */
export function buildSheetTabStrip(
  buttons: HTMLButtonElement[],
  messages: Readonly<XSheetMessages>,
): SheetTabStrip {
  const element = document.createElement('div');
  element.className = 'xs-sheet-strip';

  const tabs = document.createElement('div');
  tabs.className = 'xs-sheet-tabs';
  tabs.append(...buttons);

  const makeNav = (text: string, label: string): HTMLButtonElement => {
    const nav = document.createElement('button');
    nav.type = 'button';
    nav.className = 'xs-sheet-nav';
    nav.textContent = text;
    nav.title = label;
    nav.setAttribute('aria-label', label);
    return nav;
  };
  const prev = makeNav('◀', messages.tabsPrev);
  const next = makeNav('▶', messages.tabsNext);
  element.append(prev, tabs, next);

  /** スクロール位置を動かす（smooth 指定は scrollTo 未実装の環境では即時に落とす）。 */
  const scrollTabsTo = (left: number, smooth: boolean): void => {
    if (typeof tabs.scrollTo === 'function') {
      tabs.scrollTo({ left, behavior: smooth ? 'smooth' : 'auto' });
    } else {
      tabs.scrollLeft = left;
    }
  };

  /** ◀▶ の表示・disabled を現在のスクロール位置に合わせる。 */
  const updateNav = (): void => {
    const max = tabs.scrollWidth - tabs.clientWidth;
    const overflow = max > 1;
    prev.hidden = !overflow;
    next.hidden = !overflow;
    prev.disabled = tabs.scrollLeft <= 0;
    next.disabled = tabs.scrollLeft >= max - 1;
  };

  let pending: HTMLElement | undefined;
  /** 選択中タブが見える範囲へ寄せる（tabs が offsetParent なので offsetLeft はスクロール領域内の座標）。 */
  const applyReveal = (btn: HTMLElement, smooth: boolean): void => {
    const left = btn.offsetLeft;
    const right = left + btn.offsetWidth;
    if (left < tabs.scrollLeft) {
      scrollTabsTo(left, smooth);
    } else if (right > tabs.scrollLeft + tabs.clientWidth) {
      scrollTabsTo(right - tabs.clientWidth, smooth);
    }
  };
  const reveal = (btn: HTMLElement | undefined): void => {
    if (!btn) {
      return;
    }
    if (tabs.clientWidth === 0) {
      pending = btn;
      return;
    }
    pending = undefined;
    applyReveal(btn, true);
  };

  const onLayout = (): void => {
    if (pending && tabs.clientWidth > 0) {
      const target = pending;
      pending = undefined;
      applyReveal(target, false);
    }
    updateNav();
  };

  const onPrev = (): void => scrollTabsTo(tabs.scrollLeft - tabs.clientWidth * 0.8, true);
  const onNext = (): void => scrollTabsTo(tabs.scrollLeft + tabs.clientWidth * 0.8, true);
  prev.addEventListener('click', onPrev);
  next.addEventListener('click', onNext);

  /** 縦ホイールを横スクロールへ変換する（あふれていて縦成分が主のときだけ。preventDefault のため非 passive）。 */
  const onWheel = (e: WheelEvent): void => {
    if (tabs.scrollWidth - tabs.clientWidth <= 1 || Math.abs(e.deltaY) <= Math.abs(e.deltaX)) {
      return;
    }
    e.preventDefault();
    tabs.scrollLeft += e.deltaY;
  };
  tabs.addEventListener('scroll', updateNav, { passive: true });
  tabs.addEventListener('wheel', onWheel, { passive: false });

  // ◀▶ の表示切替は tabs 自身の幅を変えるため、通知の中で直接行うと同じフレームで再通知が出て
  // 「ResizeObserver loop completed with undelivered notifications」になる。次のフレームへ送って 1 回にまとめる。
  let frame = 0;
  const onResize = (): void => {
    if (frame === 0) {
      frame = requestAnimationFrame(() => {
        frame = 0;
        onLayout();
      });
    }
  };
  // jsdom 等には ResizeObserver が無いためガードする。
  const observer = typeof ResizeObserver === 'function' ? new ResizeObserver(onResize) : null;
  observer?.observe(tabs);
  updateNav();

  const dispose = (): void => {
    observer?.disconnect();
    if (frame !== 0) {
      cancelAnimationFrame(frame);
      frame = 0;
    }
    prev.removeEventListener('click', onPrev);
    next.removeEventListener('click', onNext);
    tabs.removeEventListener('scroll', updateNav);
    tabs.removeEventListener('wheel', onWheel);
    pending = undefined;
  };
  return { element, reveal, dispose };
}
