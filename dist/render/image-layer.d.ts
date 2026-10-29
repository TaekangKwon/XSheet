/**
 * `.xlsx` グリッドの上に重ねる画像レイヤー（ISSUE-86）。
 * - 構造：`position: absolute` のレイヤー要素（原点＝A1 の左上）の中に、図ごとの要素を `left/top/width/height`（px）で置く。
 *   表とは独立で、スクロールは枠ごと動くため位置の再計算は要らない。可視範囲（上下左右 1 画面分）と交わる図だけを DOM に出し入れする。
 * - 画像は `<img>`（`decoding="async"`・`alt` は説明か名前）。画像以外・描けない画像・壊れた画像は種類名入りの破線枠。
 *   種類名は文言（`messages`）を `textContent` で入れる（図の名前・パスは画面にも console にも出さない。名前は `alt` のみ）。
 * - 図形・テキストボックスは、形を評価できれば形どおりに描く（ISSUE-92）：`div.xs-figure`（外接矩形）＞ `<svg>`（path 1 つにつき `<path>` 1 つ。枠線・塗り・面の濃淡はクラスと CSS）
 *   ＋ `div.xs-shape-box`（回転前の矩形・回転・図形の矩形で切る）＞ `div.xs-shape-text`（文字領域）。枠も塗りも文字も無い図形は描かない（配置の時点で除く）。
 * - 形を評価できない図形（代替）で文字のあるものは種類名ではなく文字を描く（ISSUE-88。段落＝`div`・ラン＝`span`・改行＝`br`、すべて `textContent`）。
 *   枠線・塗りの有無はクラス、回転・縦書きは列挙値の対応表から `style` に入れる（ファイルの文字列は style に入れない）。
 * - 線・矢印（ISSUE-88）は 1 本につき `<svg>` 1 つ（外接矩形の位置）。`<marker>`・`innerHTML` は使わず、矢じりは `<path>`・`<ellipse>`。
 * - 色（ISSUE-96）：lib の整数 RGB・不透明度を検証して `#rrggbb`・数値にし、要素の style（CSSOM）で入れる（図形の fill・stroke、線・矢じり、文字の color、代替の文字枠の背景・枠線）。
 *   面の濃淡は lib の `applyFillMode`。色が無い（省略）ときは style を入れず、暫定のテーマ変数（CSS）のまま。
 * - ダーク表示の下地（ISSUE-96）：図形・線・代替の文字枠は、図の要素の直前に独立した `svg.xs-backdrop`（`z-index: -1`・片側 3px 広げる）を兄弟として置く。
 *   形の塗りの path と線の縁取り（太さ＋6px・実線・丸）だけを持ち、色は CSS のテーマ変数（ライトは透明）で決める＝style に色を入れず、テーマの切り替えで描き直さない。
 *   レイヤーが `isolation: isolate` で重ね順の文脈を作り、下地は全図の下・セルの上に描かれる。図と一緒に出し入れする。
 * - blob URL は {@link createImageStore}（シート 1 枚ぶん）が持つ。シートの差し替え・{@link ImageLayer.dispose} で revoke する。
 * 位置の計算は DOM 非依存の `drawings.ts`、blob の管理は `image-store.ts` に切り出している。
 */
import type { ImageResult, SheetDrawing } from '../core/index';
import { type DrawingBox, type PxRect } from './drawings';
import { type ImageEntry } from './image-store';
import { type XSheetMessages } from './messages';
/** 画像レイヤーに渡すシートの図 */
export interface LayerSheet {
    /** 図の一覧（drawing XML の順） */
    drawings: readonly SheetDrawing[];
    /** 配置済みの図（`layoutDrawings` の結果。`index` は `drawings` の位置） */
    boxes: readonly DrawingBox[];
    /** 画像を読む（`XlsxBook.readImage` をシート番号込みで束ねたもの） */
    readImage: (path: string) => Promise<ImageResult>;
    /** 画像の合計の上限を超えたとき（バナー用）。1 シートにつき 1 回 */
    onOverBudget?: () => void;
}
/** 画像レイヤー */
export interface ImageLayer {
    /** スクロール枠の中の wrapper に取り付けるレイヤー要素 */
    readonly element: HTMLElement;
    /** 描くシートを差し替える（前のシートの blob URL はすべて revoke し、DOM の図も外す） */
    setSheet(sheet: LayerSheet | null): void;
    /** 原点（A1 の左上）の位置とレイヤーの大きさ（px。wrapper の左上基準） */
    place(left: number, top: number, width: number, height: number): void;
    /** 可視範囲（A1 基準・余白込み。`viewWindow`）と交わる図だけを DOM に出し、外れた図を外す */
    update(view: PxRect): void;
    /** すべての blob URL を revoke し、DOM の図を外す。以後 `setSheet`・`update` は何もしない */
    dispose(): void;
}
/**
 * 破線枠に入れる種類名。画像（`kind: 'image'`）は、描けない形式で形式名が分かれば「画像（EMF）」、それ以外は「画像（表示できない）」。
 *
 * @param entry 画像の読み込み失敗（画像以外・読み込み前は不要）
 * @param messages 文言の表（既定は日本語）
 */
export declare function frameLabel(drawing: SheetDrawing, entry?: ImageEntry, messages?: Readonly<XSheetMessages>): string;
/**
 * 文字を描く図形か（kind が shape・textbox で、表示する文字が 1 文字以上ある。改行だけ・空のランは含めない）。
 * 該当しない図形は種類名入りの破線枠になる。
 */
export declare function hasShapeText(drawing: SheetDrawing): boolean;
/**
 * 画像レイヤーを作る。
 *
 * @param deps テスト用に差し替える依存（blob URL の作成・解放）。通常は省略（`URL.createObjectURL`・`revokeObjectURL`）
 */
export declare function createImageLayer(deps?: {
    createUrl?: (blob: Blob) => string;
    revokeUrl?: (url: string) => void;
    /** 破線枠の文言（既定は日本語） */
    messages?: Readonly<XSheetMessages>;
}): ImageLayer;
