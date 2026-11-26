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
export declare const SELECTION_KEYS: ReadonlySet<string>;
/**
 * from の次（dir 方向）の、大きさが 0 でない番号（1〜max）。無ければ null。
 * from が範囲の外（max より先）でも、範囲の中へは戻らない（null）。
 */
export declare function nextVisible(from: number, dir: 1 | -1, max: number, size: (i: number) => number): number | null;
/** 動かす先の計算に使う表示の状態 */
export interface KeyContext {
    geo: Pick<Geometry, 'rowTop' | 'rowHeight' | 'colWidth' | 'rowAt' | 'totalHeight'>;
    /** 表の行数・列数（動ける範囲） */
    extent: {
        rows: number;
        cols: number;
    };
    /** 値のある範囲（`rowCount`・`colCount`。Ctrl+End の行き先） */
    used: {
        rows: number;
        cols: number;
    };
    /** 1 ページの高さ（表示枠のセルの部分の高さ・px） */
    pageHeight: number;
}
/**
 * キーで動かした先の今のセル。扱わないキーは undefined、端で動けないときは null。
 * @param ctrl Ctrl（Mac の ⌘ も呼ぶ側で渡す）を押しているか
 */
export declare function keyTarget(key: string, ctrl: boolean, sel: Selection, ctx: KeyContext): {
    r: number;
    c: number;
} | null | undefined;
