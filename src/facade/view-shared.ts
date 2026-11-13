// 正: docs/adr/0003-universal-distribution.md §2
/**
 * `view` と確認用の `view2` が共有する中の関数（公開しない）。
 * `view.ts` から export すると単独版の型（xsheet.d.ts）に出るので、公開の型から参照されないこのファイルに置く。
 */
import { MESSAGES_EN, MESSAGES_JA, MESSAGES_KO, type XSheetMessages } from '../render/index';
import { resolveMessages } from '../render/messages';
import type { ViewOptions } from './view';

const TABLES: Readonly<Record<'ja' | 'en' | 'ko', Readonly<XSheetMessages>>> = {
  ja: MESSAGES_JA,
  en: MESSAGES_EN,
  ko: MESSAGES_KO,
};

/** 言語の名前（`ko-KR` など）から表を選ぶ。合わなければ null */
function tableFor(tag: string | null | undefined): Readonly<XSheetMessages> | null {
  const primary = (tag ?? '').trim().toLowerCase().split(/[-_]/)[0];
  return primary === 'ja' || primary === 'en' || primary === 'ko' ? TABLES[primary] : null;
}

/**
 * 使う文言の表（`lang` で選んだ表に `messages` を重ねる）。
 * `auto` は `<html lang>` があればそれだけで決め（合わなければ英語）、無いときだけブラウザの言語を見る。
 * 型を通らない呼び出し（`<script>` から `'ko-KR'`・`'zh'` など）も、名前から選ぶか `auto` と同じにする。
 */
export function pickMessages(doc: Document, options: ViewOptions): Readonly<XSheetMessages> {
  const lang: unknown = options.lang ?? 'auto';
  let base = lang === 'auto' || typeof lang !== 'string' ? null : tableFor(lang);
  if (base === null) {
    const pageLang = doc.documentElement.lang.trim();
    const nav = doc.defaultView?.navigator;
    base =
      pageLang !== ''
        ? (tableFor(pageLang) ?? MESSAGES_EN)
        : (tableFor(nav?.languages?.[0]) ?? tableFor(nav?.language) ?? MESSAGES_EN);
  }
  return resolveMessages(options.messages, base);
}

/** 要素か CSS セレクタを要素にする */
export function resolveTarget(target: HTMLElement | string): HTMLElement {
  if (typeof target !== 'string') {
    if (typeof target !== 'object' || target === null || target.nodeType !== 1) {
      throw new TypeError('XSheet: the target must be an element or a CSS selector');
    }
    return target;
  }
  const el = document.querySelector<HTMLElement>(target);
  if (el === null) throw new TypeError(`XSheet: no element matches the selector ${target}`);
  return el;
}
