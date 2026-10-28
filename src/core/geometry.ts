// 正: InterHub docs/plan/adr/0026-xlsx-image-overlay.md（§決定 3）
/**
 * シートの寸法（layout）を px の座標へ換算する純関数（DOM 非依存）。
 * - 表（colgroup・行高）・仮想化の上下余白・画像レイヤーが、同じ座標（Excel の寸法）を唯一の幾何データとして使う。
 * - 累積は「既定値×番号＋疎な差分の累積和」で出す。前計算は指定のある数に比例し、
 *   1,048,576 行・16,384 列でも全行・全列の配列は作らない。`rowAt`／`colAt` は二分探索。
 * - 非表示の行・列は寸法 0。
 */
import { MAX_COLS, MAX_COL_WIDTH, MAX_ROWS, MAX_ROW_PT } from './limits';
import type { SheetLayout } from './types';

/** 行高 pt → px。`round(pt × 96 / 72)`（15pt＝20px、18.75pt＝25px） */
export function ptToPx(pt: number): number {
  return Math.round((pt * 96) / 72);
}

/**
 * 列幅（`<col width>` の文字数・余白を含む値）→ px。
 * ECMA-376 Part 1 の `col` 要素の換算式 `trunc(((256·w + trunc(128/MDW)) / 256) · MDW)`。
 * @param w 列幅（文字数）
 * @param mdw 最大桁幅 MDW（px）
 */
export function colWidthToPx(w: number, mdw: number): number {
  return Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw);
}

/**
 * 既定の列幅（`defaultColWidth` が無いとき）→ px。`ceil((baseColWidth·MDW + 余白) / 8) · 8`
 * （MDW 7 で 64px、MDW 8 で 72px）。余白は `2·ceil(MDW/4) + 1`（MDW 5〜8 では 5px）。
 * 根拠: Excel 16 の実ファイルの `<col width>` の端数から余白を逆算（2026-10-27）。
 */
export function defaultColWidthToPx(baseColWidth: number, mdw: number): number {
  const padding = 2 * Math.ceil(mdw / 4) + 1;
  return Math.ceil((baseColWidth * mdw + padding) / 8) * 8;
}

/** 番号の範囲 [min, max]（1 始まり・両端を含む）に共通の寸法 size（px）を割り当てる 1 件 */
interface Span {
  min: number;
  max: number;
  size: number;
}

/**
 * 1 本の軸（行または列）。番号は 1 始まり。
 * 既定の寸法 `def` と、既定と違う範囲 `spans`（昇順・重なりなし）から、位置の累積と逆引きを出す。
 */
interface Axis {
  /** 番号 i の手前までの寸法の合計（i ≤ 1 は 0。要素数の上限は無い） */
  offset(i: number): number;
  /** 番号 i の寸法（i < 1 は 0） */
  size(i: number): number;
  /** 位置 pos を含む（pos 以下で最大の offset を持つ）番号を [1, n] の範囲で返す。n < 1 は 0 */
  at(pos: number, n: number): number;
}

/**
 * 軸を作る。`spans` は並べ替え・重なりの除去（min の小さい範囲を優先し、後の範囲は重なりを削る）を行う。
 * 前計算（累積和）は spans の数に比例し、番号の上限までの配列は作らない。
 */
function createAxis(def: number, input: Span[]): Axis {
  const sorted = input.filter((s) => s.min <= s.max).sort((a, b) => a.min - b.min || a.max - b.max);
  const spans: Span[] = [];
  let lastMax = 0;
  for (const s of sorted) {
    const min = Math.max(s.min, lastMax + 1);
    if (min > s.max) continue; // 前の範囲に完全に含まれる
    const prev = spans[spans.length - 1];
    if (prev !== undefined && prev.max + 1 === min && prev.size === s.size)
      prev.max = s.max; // 隣り合う同じ寸法の範囲は 1 件にまとめる
    else spans.push({ min, max: s.max, size: s.size });
    lastMax = s.max;
  }
  // cum[k] = spans[0..k-1] の「既定との差分」の合計
  const cum: number[] = new Array<number>(spans.length + 1).fill(0);
  for (let k = 0; k < spans.length; k++) {
    const s = spans[k]!;
    cum[k + 1] = cum[k]! + (s.max - s.min + 1) * (s.size - def);
  }

  /** min ≤ i を満たす範囲の数（二分探索） */
  const countStarting = (i: number): number => {
    let lo = 0;
    let hi = spans.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (spans[mid]!.min <= i) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };

  const offset = (i: number): number => {
    if (!(i > 1)) return 0;
    const last = i - 1; // 手前までの最後の番号
    const k = countStarting(last);
    let delta = 0;
    if (k > 0) {
      const s = spans[k - 1]!;
      delta = cum[k - 1]! + (Math.min(s.max, last) - s.min + 1) * (s.size - def);
    }
    return def * last + delta;
  };

  const size = (i: number): number => {
    if (!(i >= 1)) return 0;
    const k = countStarting(i);
    if (k > 0) {
      const s = spans[k - 1]!;
      if (i <= s.max) return s.size;
    }
    return def;
  };

  const at = (pos: number, n: number): number => {
    if (!(n >= 1)) return 0;
    const p = pos > 0 ? pos : 0; // 負・NaN は先頭として扱う
    // offset(i) ≤ p を満たす最大の i（offset は番号について単調非減少）
    let lo = 1;
    let hi = n;
    while (lo < hi) {
      const mid = Math.ceil((lo + hi) / 2);
      if (offset(mid) <= p) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  };

  return { offset, size, at };
}

/** 換算した座標（px・すべて整数）。番号は 1 始まり */
export interface Geometry {
  /** 列 c の左端の x（c ≤ 1 は 0）。cols より右の列も既定の幅で延ばして返す */
  colLeft(c: number): number;
  /** 列 c の幅（非表示は 0。c < 1 は 0） */
  colWidth(c: number): number;
  /** 行 r の上端の y（r ≤ 1 は 0）。rows より下の行も既定の高さで延ばして返す */
  rowTop(r: number): number;
  /** 行 r の高さ（非表示は 0。r < 1 は 0） */
  rowHeight(r: number): number;
  /**
   * y を含む行（1〜rows）。負・NaN は先頭、下端以降は最後の行へ丸める。rows が 0 なら 0。
   * 非表示の行（高さ 0）は、同じ y の次の表示行より前になるため返らない（末尾だけが非表示のときを除く）。
   */
  rowAt(y: number): number;
  /** x を含む列（1〜cols）。丸め方は `rowAt` と同じ */
  colAt(x: number): number;
  /** 表の幅（1〜cols の列幅の合計） */
  readonly totalWidth: number;
  /** 表の高さ（1〜rows の行高の合計） */
  readonly totalHeight: number;
}

/** 数が有限で 0 以上のときだけ採り、上限 max に丸める（それ以外は既定へ。巨大な値で座標が壊れないように） */
const clampNum = (v: number, max: number, fallback: number): number =>
  Number.isFinite(v) && v >= 0 ? Math.min(v, max) : fallback;

/**
 * layout から座標（px）の計算器を作る。
 * - 列幅は `<col width>` を `colWidthToPx` で換算し、指定の無い列は既定の列幅（`defaultColWidth` があれば同じ式で換算、
 *   無ければ `defaultColWidthToPx`）。非表示の列は 0。
 * - 行高：シートの既定の行高は、`customHeight="1"` なら `defaultRowHeight`（pt）の換算値、
 *   そうでなければフォントの既定の行高（`fontRowPx`。実測表）を優先し、無ければ `defaultRowHeight` の換算値。
 *   （Excel は `customHeight` の無い `defaultRowHeight`・`<row ht>` を、その PC のフォントから計算した行高で置き換えるため）
 *   行は、`customHeight="1"` の `ht` は `ptToPx(ht)`。`customHeight` の無い `ht` は、ファイルの
 *   `defaultRowHeight` と ±0.01 で等しければシートの既定、違えば保存時の自動調整値として `ptToPx(ht)` とする。
 *   `ht` が無ければシートの既定。非表示の行は 0。
 *   `zeroHeight` のときは、`<row>` の無い行を既定で高さ 0 とする（`<row>` のある行は表示）。
 * - 列幅・行高は Excel の上限（255 文字・409.5pt）に丸める。
 * @param layout `XlsxSheet.layout`
 * @param rows 表の行数（`XlsxSheet.rowCount`）。`totalHeight`・`rowAt` の範囲に使う
 * @param cols 表の列数（`XlsxSheet.colCount`）。`totalWidth`・`colAt` の範囲に使う
 */
export function createGeometry(layout: SheetLayout, rows: number, cols: number): Geometry {
  const mdw = Number.isFinite(layout.mdw) && layout.mdw >= 1 ? Math.min(layout.mdw, 100) : 7;

  const defColPx =
    layout.defaultColWidth !== null
      ? colWidthToPx(clampNum(layout.defaultColWidth, MAX_COL_WIDTH, 0), mdw)
      : defaultColWidthToPx(clampNum(layout.baseColWidth, MAX_COL_WIDTH, 8), mdw);
  const colSpans: Span[] = [];
  for (const c of layout.cols) {
    const min = Math.max(1, Math.trunc(c.min));
    const max = Math.min(MAX_COLS, Math.trunc(c.max));
    const size = c.hidden
      ? 0
      : c.width === null
        ? defColPx
        : colWidthToPx(clampNum(c.width, MAX_COL_WIDTH, 0), mdw);
    colSpans.push({ min, max, size });
  }
  const colAxis = createAxis(defColPx, colSpans);

  const defaultPt = clampNum(layout.defaultRowHeight, MAX_ROW_PT, 15);
  const fontRowPx =
    layout.fontRowPx !== null && Number.isFinite(layout.fontRowPx) && layout.fontRowPx >= 0
      ? layout.fontRowPx
      : null;
  /** シートの既定の行高（px）。zeroHeight のとき `<row>` のある行（表示）の高さにも使う */
  const sheetRowPx = layout.customHeight ? ptToPx(defaultPt) : (fontRowPx ?? ptToPx(defaultPt));
  const defRowPx = layout.zeroHeight ? 0 : sheetRowPx;
  const rowSpans: Span[] = [];
  for (const r of layout.rows) {
    const min = Math.max(1, Math.trunc(r.min));
    const max = Math.min(MAX_ROWS, Math.trunc(r.max));
    let size: number;
    if (r.hidden) size = 0;
    else if (r.ht === null) size = sheetRowPx;
    else {
      const ht = clampNum(r.ht, MAX_ROW_PT, defaultPt);
      size = r.custom || Math.abs(ht - defaultPt) > 0.01 ? ptToPx(ht) : sheetRowPx;
    }
    if (size === defRowPx) continue; // 既定と同じなら持たない
    rowSpans.push({ min, max, size });
  }
  const rowAxis = createAxis(defRowPx, rowSpans);

  const nRows = Math.max(0, Math.trunc(rows));
  const nCols = Math.max(0, Math.trunc(cols));
  return {
    colLeft: (c) => colAxis.offset(c),
    colWidth: (c) => colAxis.size(c),
    rowTop: (r) => rowAxis.offset(r),
    rowHeight: (r) => rowAxis.size(r),
    rowAt: (y) => rowAxis.at(y, nRows),
    colAt: (x) => colAxis.at(x, nCols),
    totalWidth: colAxis.offset(nCols + 1),
    totalHeight: rowAxis.offset(nRows + 1),
  };
}
