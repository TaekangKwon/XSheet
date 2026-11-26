// 正: docs/adr/0011-selection.md §4
/**
 * キーボードで選択を移す先の計算（XSheet の追加。M7-25・Issue #56）。DOM を使わない。
 * - ↑・↓・←・→：前・次の行・列。行の全体の選択では ←・→ が今のセルの列を、列の全体の選択では ↑・↓ が今のセルの行を動かす。
 *   結合・複数の行の選択は範囲の端から動く（範囲を 1 つとして飛ぶ）。
 * - PageUp・PageDown：表示枠のセルの部分の高さだけ動く。Home：行の先頭の列。Ctrl+Home：左上。Ctrl+End：値のある最後のセル。
 * - 高さ 0（非表示）の行・幅 0 の列は飛ばす。表の範囲（`extent`）の端より先へは動かない。
 * `selection.ts` と同じく遅れて読む部品の中にあり、表示の部品を import しない。
 */
import type { Geometry } from '../core/index';
import type { Selection } from './selection';

/** 選択を動かすキー（Ctrl・Shift・Alt の扱いは {@link keyTarget}） */
export const SELECTION_KEYS: ReadonlySet<string> = new Set([
  'ArrowUp',
  'ArrowDown',
  'ArrowLeft',
  'ArrowRight',
  'PageUp',
  'PageDown',
  'Home',
  'End',
]);

/**
 * from の次（dir 方向）の、大きさが 0 でない番号（1〜max）。無ければ null。
 * from が範囲の外（max より先）でも、範囲の中へは戻らない（null）。
 */
export function nextVisible(
  from: number,
  dir: 1 | -1,
  max: number,
  size: (i: number) => number,
): number | null {
  for (let i = from + dir; i >= 1 && i <= max; i += dir) {
    if (size(i) > 0) return i;
  }
  return null;
}

/** 動かす先の計算に使う表示の状態 */
export interface KeyContext {
  geo: Pick<Geometry, 'rowTop' | 'rowHeight' | 'colWidth' | 'rowAt' | 'totalHeight'>;
  /** 表の行数・列数（動ける範囲） */
  extent: { rows: number; cols: number };
  /** 値のある範囲（`rowCount`・`colCount`。Ctrl+End の行き先） */
  used: { rows: number; cols: number };
  /** 1 ページの高さ（表示枠のセルの部分の高さ・px） */
  pageHeight: number;
}

/**
 * キーで動かした先の今のセル。扱わないキーは undefined、端で動けないときは null。
 * @param ctrl Ctrl（Mac の ⌘ も呼ぶ側で渡す）を押しているか
 */
export function keyTarget(
  key: string,
  ctrl: boolean,
  sel: Selection,
  ctx: KeyContext,
): { r: number; c: number } | null | undefined {
  const { geo: g, extent: ext } = ctx;
  const rowH = (r: number): number => g.rowHeight(r);
  const colW = (c: number): number => g.colWidth(c);
  const { type, range, active } = sel;
  const at = (r: number | null, c: number | null) => (r === null || c === null ? null : { r, c });
  switch (key) {
    case 'ArrowDown':
      return at(nextVisible(type === 'column' ? active.r : range.r2, 1, ext.rows, rowH), active.c);
    case 'ArrowUp':
      return at(nextVisible(type === 'column' ? active.r : range.r1, -1, ext.rows, rowH), active.c);
    case 'ArrowRight':
      return at(active.r, nextVisible(type === 'row' ? active.c : range.c2, 1, ext.cols, colW));
    case 'ArrowLeft':
      return at(active.r, nextVisible(type === 'row' ? active.c : range.c1, -1, ext.cols, colW));
    case 'PageDown':
    case 'PageUp': {
      const dir: 1 | -1 = key === 'PageDown' ? 1 : -1;
      const y = g.rowTop(active.r) + dir * Math.max(1, ctx.pageHeight);
      let r = Math.min(ext.rows, Math.max(1, g.rowAt(Math.min(g.totalHeight - 1, Math.max(0, y)))));
      if (rowH(r) === 0)
        r =
          nextVisible(r, dir, ext.rows, rowH) ??
          nextVisible(r, dir === 1 ? -1 : 1, ext.rows, rowH) ??
          r;
      if (r === active.r) return at(nextVisible(active.r, dir, ext.rows, rowH), active.c);
      return { r, c: active.c };
    }
    case 'Home':
      return ctrl
        ? at(nextVisible(0, 1, ext.rows, rowH), nextVisible(0, 1, ext.cols, colW))
        : at(active.r, nextVisible(0, 1, ext.cols, colW));
    case 'End': {
      if (!ctrl) return undefined; // End だけ（Excel の End モード）は扱わない
      const rows = Math.min(ext.rows, Math.max(1, ctx.used.rows));
      const cols = Math.min(ext.cols, Math.max(1, ctx.used.cols));
      return at(nextVisible(rows + 1, -1, rows, rowH), nextVisible(cols + 1, -1, cols, colW));
    }
    default:
      return undefined;
  }
}
