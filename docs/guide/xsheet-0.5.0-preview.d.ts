// XSheet 0.5.0 で足す公開 API の型の先出し（Issue #26・ADR-0007 §9）。説明は xsheet-0.5.0-preview.md。
// 単独版 0.4.1 の xsheet.d.ts と同じフォルダに置くと、0.4.1 の型に 0.5.0 の追加が重なる（大域の XSheet にも出る）。
// 実装はまだ無い（0.4.1 で呼ぶと undefined）。0.5.0 を出したら、その xsheet.d.ts に置き換えてこのファイルは消す。
// 0.5.0 までに形が変わるときは、このファイルと ADR-0007 を直してから知らせる。

declare module 'xsheet' {
  import type { XlsxSheet } from 'xsheet/_/core/types';

  // ---- 検索（R1・ADR-0007 §1） ----

  /** `findText` の指定 */
  export interface FindTextOptions {
    /** `'partial'`＝文字の一部に一致・`'cell'`＝セルの文字全体が一致（既定 `'partial'`） */
    match?: 'partial' | 'cell';
    /** 英字の大文字と小文字を区別しない（既定 true）。全角と半角はいつも区別する */
    ignoreCase?: boolean;
    /** セルの中の改行（`\r\n`・`\n`・`\r`）を照合のときだけ除く（既定 true）。位置は元の文字で返す */
    ignoreNewlines?: boolean;
    /** 非表示の行・列（高さ・幅 0 の行・列も）を調べない（既定 true） */
    skipHidden?: boolean;
    /** 中断。中断されたら `AbortError`（`DOMException`）で reject する */
    signal?: AbortSignal;
    /** 返す一致の数の上限（既定 10,000。100,000 より大きい値は 100,000 にする） */
    maxMatches?: number;
  }

  /** `findText` の結果 */
  export interface FindTextResult {
    /** 一致（行 → 列 → セルの中の位置の順） */
    matches: TextMatch[];
    /** `maxMatches` で打ち切ったか */
    truncated: boolean;
  }

  /** 1 つの一致 */
  export interface TextMatch {
    /** 行（1 始まり） */
    r: number;
    /** 列（1 始まり） */
    c: number;
    /** 一致の始まり（`CellView.text` 上の UTF-16 の位置） */
    start: number;
    /** 一致の終わり（含まない） */
    end: number;
  }

  /**
   * シートの中の文字を探す。表示形式を当てた後の文字（`CellView.text`）で照らす。
   * - 1 セルの中の複数の一致をすべて返す。結合したセルは左上のセルだけ。空の検索語は空の結果。
   * - 図形・グラフの中の文字は探さない。
   * - 約 10ms ごとに休むので、大きなシートでも画面は固まらない。
   * - ブック全体を探すときは、呼ぶ側がシートを 1 つずつ `openSheet` で開いて呼ぶ。
   * @param sheet `XlsxBook.openSheet` で開いたシート
   * @param query 検索語
   * @throws DOMException `signal` で中断された（`name` は `'AbortError'`）
   */
  export function findText(
    sheet: XlsxSheet,
    query: string,
    options?: FindTextOptions,
  ): Promise<FindTextResult>;

  // ---- 拡大縮小の指定（R5・ADR-0007 §5） ----

  /**
   * 倍率。数は 0.5〜2.0（範囲の外は端の値にする。`NaN` などは `RangeError`）。
   * `{ fit: 'width' }` はデータのある列と図の右端までを表示枠の幅に合わせる（`min`・`max` の既定は 0.5・2.0）。
   */
  export type Zoom = number | { fit: 'width'; min?: number; max?: number };

  /** `focusHighlight` の指定 */
  export interface FocusHighlightOptions {
    /** そのヒットが見える所までスクロールする（既定 true） */
    scroll?: boolean;
    /** なめらかにスクロールする（既定 false） */
    smooth?: boolean;
  }

  /** `scrollToCell` の指定 */
  export interface ScrollToCellOptions {
    /** なめらかにスクロールする（既定 false） */
    smooth?: boolean;
  }

  /** 画面の位置にあるセル（1 始まり） */
  export interface CellPosition {
    r: number;
    c: number;
  }

  /** 組み込みの日付の書式 27〜36・50〜58 の見せ方（`'japanese-era'`＝和暦。既定 `'default'`＝西暦） */
  export type BuiltinDates = 'default' | 'japanese-era';
}

// ---- 読み取りの指定（R6・ADR-0007 §6） ----

declare module 'xsheet/_/core/index' {
  import type { BuiltinDates } from 'xsheet';
  interface OpenXlsxOptions {
    /**
     * 組み込みの日付の書式 27〜36・50〜58 を和暦にするか（既定 `'default'`＝西暦のまま）。
     * 書式に直接書かれた和暦（`ggg`・`e` など）と曜日（`aaa`・`aaaa`）は、指定しなくてもいつも出す。
     */
    builtinDates?: BuiltinDates;
  }
}

declare module 'xsheet/_/facade/view' {
  import type { BuiltinDates } from 'xsheet';
  interface ReadOptions {
    /** `OpenXlsxOptions.builtinDates` と同じ（`read`・`view` で使える） */
    builtinDates?: BuiltinDates;
  }
}

declare module 'xsheet/_/render/render-workbook' {
  import type { BuiltinDates } from 'xsheet';
  interface WorkbookOptions {
    /** `OpenXlsxOptions.builtinDates` と同じ（`renderWorkbook` はブックを自分で開くため） */
    builtinDates?: BuiltinDates;
  }
}

// ---- 表示の指定と表示の操作（R2〜R5・R7・ADR-0007 §2〜§5・§7） ----
// SheetView に足すので、WorkbookView（renderWorkbook の戻り値）も同じものを持つ。

declare module 'xsheet/_/render/render-sheet' {
  import type {
    CellPosition,
    FocusHighlightOptions,
    ScrollToCellOptions,
    TextMatch,
    Zoom,
  } from 'xsheet';

  interface RenderOptions {
    /** 最初の倍率（既定 1）。最初から指定するとちらつかない */
    zoom?: Zoom;
    /** 行番号の列と列見出しを出すか（既定 true） */
    headers?: boolean;
  }

  interface SheetView {
    /**
     * 一致を光らせる（`findText` の `matches` をそのまま渡す）。`null`・空の配列で消す。
     * スクロールで描き直しても残る。シートを替えると消える。
     * 色は CSS 変数 `--xs-hit-bg`（既定 `rgba(255, 255, 0, 0.4)`）・`--xs-hit-current-bg`（既定 `#ffa500`）。
     */
    setHighlights(matches: readonly TextMatch[] | null): void;
    /**
     * 今のヒットを替える。
     * @param index `setHighlights` に渡した配列の番号（0 始まり）。-1 で外す
     */
    focusHighlight(index: number, options?: FocusHighlightOptions): void;
    /** セルが見える所までスクロールする（縦は表示枠の上から約 1/3。横は見えていないときだけ動かす） */
    scrollToCell(r: number, c: number, options?: ScrollToCellOptions): void;
    /**
     * 画面の座標（`MouseEvent.clientX`・`clientY`）にあるセル。見出し・余白・表の外は null。
     * 結合したセルは左上、はみ出した文字の上ははみ出し元のセルを返す。
     */
    cellAt(clientX: number, clientY: number): CellPosition | null;
    /** 画面の座標にあるハイライトの番号（`setHighlights` の配列の番号）。無ければ -1 */
    highlightAt(clientX: number, clientY: number): number;
    /** 倍率を替える。見ていたセルが同じ所に残るようにスクロールする。見出しには倍率を掛けない */
    setZoom(zoom: Zoom): void;
    /** 今の倍率（fit のときは計算した値） */
    readonly zoom: number;
    /**
     * シートを替える（倍率は保ち、ハイライトと今のヒットは消す）。表示中のシートなら何もしない。
     * @param index シートの番号（`SheetInfo.index`。1 始まり・非表示のシートも数える）
     * @throws RangeError ブックに無い番号
     */
    showSheet(index: number): Promise<void>;
  }
}
