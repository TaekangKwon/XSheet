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
export declare function makeSheetTab(name: string, hidden: boolean, messages: Readonly<XSheetMessages>): HTMLButtonElement;
/**
 * シートタブを 1 行に並べ、入りきらない分は横スクロールさせる列を作る。
 * 構造は `div.xs-sheet-strip` ＞ `[◀ button.xs-sheet-nav][div.xs-sheet-tabs（overflow-x:auto）＞ タブ…][▶ button.xs-sheet-nav]`。
 * - ◀▶ は押すとスクロール領域の可視幅の約 8 割ぶん滑らかに動かす。端に着いた側は disabled、全タブが収まるときは両方隠す。
 * - 状態はスクロール・リサイズ（ResizeObserver。無い環境ではガードして使わない）のたびに更新する。
 * - 縦のホイール回転は、あふれているときだけ横スクロールに変換する。
 *
 * @param buttons シート名のタブ（{@link makeSheetTab}。クリック処理は呼び出し側で設定済みのもの）
 */
export declare function buildSheetTabStrip(buttons: HTMLButtonElement[], messages: Readonly<XSheetMessages>): SheetTabStrip;
