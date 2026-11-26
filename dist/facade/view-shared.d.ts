/**
 * `view` と確認用の `view2` が共有する中の関数（公開しない）。
 * `view.ts` から export すると単独版の型（xsheet.d.ts）に出るので、公開の型から参照されないこのファイルに置く。
 */
import { type XSheetMessages } from '../render/index';
import type { ViewOptions } from './view';
/**
 * 使う文言の表（`lang` で選んだ表に `messages` を重ねる）。
 * `auto` は `<html lang>` があればそれだけで決め（合わなければ英語）、無いときだけブラウザの言語を見る。
 * 型を通らない呼び出し（`<script>` から `'ko-KR'`・`'zh'` など）も、名前から選ぶか `auto` と同じにする。
 */
export declare function pickMessages(doc: Document, options: ViewOptions): Readonly<XSheetMessages>;
/** 要素か CSS セレクタを要素にする */
export declare function resolveTarget(target: HTMLElement | string): HTMLElement;
