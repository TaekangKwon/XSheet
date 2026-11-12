/**
 * 図形（`xdr:sp`）の文字（`xdr:txBody`）の読み取り：段落（`a:p`）ごとに `a:r`・`a:br`・`a:fld` を出現順に並べ、書式（大きさ・太字・斜体・下線・取り消し線・揃え・箇条書き）と
 * 本文の枠（縦位置・余白・折り返し・縮小・縦書き）を読む。lib は pt・EMU・列挙値のまま持ち、px への換算は表示側。
 * 強調表示・フォント名・上付き／下付き・行間・段落の間隔・字下げ・`a:lstStyle`・ハイパーリンクは読まない。
 * 文字の色は `a:rPr` の塗りと図形の `fontRef` から解いて持たせる（ISSUE-95。描くのは ISSUE-96）。
 * DOM に依存せず、開きタグは `[^<>]*>`、閉じは `indexOf` で探して線形時間に走査する（`a:p`・`a:r`・`a:t` は入れ子にならない）。
 * 非信頼の入力として、文字数・段落数・ランの数に上限を掛ける（超えた分は捨てて `truncated` で知らせる）。文字はログに出さない（console を使わない）。
 */
import type { DrawingText } from './types';
import { type Theme } from './theme';
/** 読み取りの上限（`limits.ts` の値。テストでは小さい値を渡す） */
export interface TextCaps {
    maxChars: number;
    maxParagraphs: number;
    maxRuns: number;
    /** 1 シートの合計の残り（文字数・ラン・段落を 1 つずつ数えた数。この図形の分がここを超えたら以後を捨てる） */
    maxUnits: number;
}
/** `readShapeText` の結果 */
export interface ShapeTextResult {
    /** 文字（ランの文字が 1 つも無い図形・`txBody` の無い図形は null） */
    text: DrawingText | null;
    /** 読み取った文字数（末尾に足した「…」を含む） */
    chars: number;
    /** 1 シートの合計に足す数＝文字数＋ラン（`a:br` を含む）の数＋段落の数 */
    units: number;
    /** 文字を切った・段落やランを捨てた（文字の無い図形では、上限で失った文字が無ければ立てない） */
    truncated: boolean;
}
/**
 * 図形（`xdr:sp`）の中身から文字（`xdr:txBody`）を読む。
 * - `a:t` は `decodeEntities` で戻し、タブと改行以外の制御文字を除く。
 * - 上限（文字数 `maxChars`・段落 `maxParagraphs`・ラン `maxRuns`・1 シートの残り `maxUnits`）を超えたら、文字は切って末尾に「…」を付け、段落とランは以降を捨てて `truncated` にする。
 * - 空のラン（文字も `a:br` も無いもの）は一覧に入れず、ランの数にも数えない。
 * - ランの文字（長さ 1 以上）が 1 つも無い図形は `text` を null にする（文字の無い図形）。このとき `truncated` は、本文に空でない `a:t` があって上限で失ったときだけ立てる。
 * - ランの色（ISSUE-95）は `a:rPr` の塗り → `xdr:style` の `a:fontRef` の子の色 → テーマの `tx1`（`drawing-color.ts`）。
 * @param spBody `xdr:sp` の中身
 * @param theme ブックのテーマ（色の解決に使う。省くと既定の 12 色）
 */
export declare function readShapeText(spBody: string, caps: TextCaps, theme?: Theme): ShapeTextResult;
