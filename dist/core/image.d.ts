/**
 * 画像の読み取り：先頭バイトで形式（PNG・JPEG・GIF・BMP・WebP）を判定し、ヘッダから縦横の px を読む。
 * 拡張子と `[Content_Types].xml` の MIME は信じない（画像は細工できる非信頼のバイナリ）。
 * 切れたバイト列・不正な値で例外を出さない（読む前に必ず長さを確かめる）。SVG は先頭バイトで判定できないので扱わない。
 * 画像のバイト・パスはログに出さない（console を使わない）。
 */
import type { XlsxLimits } from './limits';
import type { XlsxPackage } from './open';
import { type ImageFormat, type ImageResult } from './types';
/** `sniffImage` の結果。描けない形式は、分かれば形式名（EMF など）を label に入れる */
export type ImageSniff = {
    ok: true;
    format: ImageFormat;
    width: number;
    height: number;
} | {
    ok: false;
    label: string | null;
};
/**
 * 先頭バイトで画像の形式を判定し、ヘッダから縦横（px）を読む。
 * 描けない形式（EMF・WMF・TIFF）は形式名つきで、判定できないもの（SVG・不正・切れた・空）は名前なしで `ok: false` を返す。
 * どんなバイト列でも例外を出さない。
 */
export declare function sniffImage(b: Uint8Array): ImageSniff;
/** 1 シートの画像の合計の数え方（readImage の呼び出しをまたいで持つ） */
export interface ImageBudget {
    /** これまでに読んだ画像の展開後の合計（バイト） */
    used: number;
    /** 合計の上限を一度でも超えた（以後、新しい画像はすべて over-budget。ADR-0026 §決定 7 の「以後の画像」） */
    exhausted: boolean;
    /** すでに数えた画像のパス（同じ画像を何度読んでも 1 回だけ数える） */
    counted: Set<string>;
}
/** 空の ImageBudget */
export declare function createImageBudget(): ImageBudget;
/**
 * zip の中の画像を読み、形式と縦横を返す。出せないときも例外にせず、理由つきの `ok: false` を返す。
 * - 1 枚の上限（maxImageBytes）で展開しながら止める（超えたら too-large。これを先に判定するので、シートの残りが少なくても理由は取り違えない）。
 * - 1 枚に収まる画像でも、シートの合計の残りを超えるなら over-budget。一度 over-budget になったら、以後そのシートの新しい画像はすべて over-budget（読まずに断る）。
 *   すでに数えた画像（同じ画像の再読み込み）は、合計を使い切った後も読める。
 * - 形式を判定したあと、1 辺と画素数の上限を確かめる。
 */
export declare function readImageFrom(pkg: XlsxPackage, path: string, limits: XlsxLimits, budget: ImageBudget): Promise<ImageResult>;
