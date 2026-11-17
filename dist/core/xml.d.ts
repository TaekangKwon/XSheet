/** 要素名の前に付いてよい名前空間の接頭辞（正規表現の断片。例 `x:`） */
export declare const NS: string;
/** XML の実体参照（`&amp;` `&#x41;` `&#65;`）を戻す。属性値と、`_xHHHH_` を使わない値に使う */
export declare function decodeEntities(s: string): string;
/**
 * 文字列セルの本文を戻す。実体参照のあと、Excel の `_xHHHH_`（制御文字などの退避）を戻す。
 * `_x005F_` は `_` を表す（本物の `_x0041_` という文字列を守るため）。
 */
export declare function decodeText(s: string): string;
/** `findElement` の結果 */
export interface ElementSpan {
    /** 開きタグ全体（属性を含む。`parseAttrs` に渡せる） */
    open: string;
    /** 中身。自己閉じ（`<cols/>`）・閉じタグの無いときは null */
    body: string | null;
    /** 要素の終わりの次の位置（次の検索の開始位置）。閉じタグの無いときは xml の長さ（検索を打ち切る） */
    end: number;
}
/**
 * `from` 以降で最初の `<name ...>…</name>`（名前空間の接頭辞可）を、線形時間で切り出す。
 * 開きタグを探し、閉じタグは `indexOf` で探す（`[\s\S]*?` の正規表現は、閉じが無いと入れ子の走査で二乗になる）。
 */
export declare function findElement(xml: string, name: string, from?: number): ElementSpan | null;
/**
 * `<name …>…</name>`（名前空間の接頭辞可）を出現順に callback へ渡す。同じ名前の入れ子は扱わない（`c`・`si`・`t` など入れ子にならない要素用）。
 * 開きタグは `[^<>]*>`、閉じは `indexOf` で探すので、`>` や閉じの無い要素を大量に並べても時間は線形。
 * 自己閉じ（`<c/>`）は body を null で渡す。閉じタグの無い要素に当たったら、以後に閉じが無いので打ち切る。
 * 正規表現は名前ごとに 1 度だけ作る（セル 1 つごとに作らない）。callback の中で同じ `xml` を再び走査してよい。
 */
export declare function forEachElement(xml: string, name: string, cb: (open: string, body: string | null) => void): void;
/**
 * `from` 以降で最初の `<name …>` の開きタグだけを返す（閉じは探さない。名前空間の接頭辞可。無ければ null）。
 * 開きタグは `[^<>]*>` で打ち切るので、`>` の無いタグを大量に並べても時間は線形。正規表現は名前ごとに 1 度だけ作る。
 */
export declare function findOpenTag(xml: string, name: string, from?: number): string | null;
/** 属性の一覧。キーに `__proto__` などが来てもプロトタイプを汚さないよう、プロトタイプの無いオブジェクトで返す */
export type Attrs = Record<string, string | undefined>;
/**
 * タグ（`<c r="A1" s="2">` の属性の部分）から属性を取り出す。値は実体参照を戻す。
 * 名前の先頭は後読みで「前が名前の文字でない位置」に限る（長い名前の途中から照合し直して二乗時間になるのを防ぐ。線形時間）。
 */
export declare function parseAttrs(tag: string): Attrs;
/** 名前空間の接頭辞を除いた属性の値（`r:id` と `rel:id` を同じに扱う）。`id` を探すとき `sheetId` は拾わない */
export declare function attrByLocalName(attrs: Attrs, local: string): string | undefined;
/** 列番号（1 始まり）を列の記号にする。1→A、27→AA。1 未満・整数でなければ空文字 */
export declare function colName(c: number): string;
/** 番地 `AB12` の位置（1 始まり）。範囲外・不正なら null */
export declare function parseCellRef(ref: string): {
    r: number;
    c: number;
} | null;
/** 列だけの番地（`AB`）。セルの r 属性（`AB12`）から列だけ欲しいとき用。範囲外・不正なら 0 */
export declare function colOfRef(ref: string): number;
/** 開きタグの名前の直後の文字として許す文字（`<c` の後が `o`（`<cols`）などでないことの確認用） */
export declare function isTagDelimiter(ch: string | undefined): boolean;
/** Strict 形式（ISO 29500 Strict）の名前空間の目印。1 次では読まない（unsupported） */
export declare const STRICT_NS = "purl.oclc.org/ooxml";
/** ルート要素の宣言（先頭の 2,000 文字）に Strict の名前空間があるか */
export declare function looksStrict(xml: string): boolean;
