/**
 * 書式コードを小さな字句に分け、ssf が知らない日本語の字句（曜日）を、計算した文字を引用符で包んだ形に置き換える。
 * 置き換えるのは、引用符の中・`\` の後・`_`／`*` の後・`[...]` の中・`AM/PM`・`A/P` 以外にある `a` の並びだけ。
 * 曜日は `Date` を使わず、日付の通し番号から整数の計算で出す（時間帯のずれが入らない）。
 * M7-8 の和暦（`g`・`e`）もこの字句の分解を使う。
 */
/** 字句の種類 */
export type FormatTokenKind = 
/** `"..."` の文字（text は引用符を含む元の形） */
'quoted'
/** `\x`・`_x`・`*x`（text は記号を含む 2 文字） */
 | 'escaped'
/** `[...]`（色・条件・`$-411` など。text は角かっこを含む） */
 | 'bracket'
/** `AM/PM`・`A/P`（大文字小文字を問わない） */
 | 'ampm'
/** 曜日 `aaa`・`aaaa`（text は元の並び） */
 | 'weekday'
/** 上のどれでもない文字（1 文字ずつ） */
 | 'char';
export interface FormatToken {
    kind: FormatTokenKind;
    text: string;
}
/** `;` で分けた区画。tokens は区画の字句（`;` 自体は含まない） */
export type FormatSection = readonly FormatToken[];
/**
 * 書式コードを `;` の区画ごとの字句に分ける。
 * 閉じていない引用符・角かっこは末尾までを 1 つの字句にする（壊れた書式でも止まらない）。
 * 1 文字ずつ前へ進むだけなので、長さに比例した時間で終わる。
 */
export declare function tokenizeFormat(code: string): FormatSection[];
/**
 * Excel の日付の通し番号から曜日（0＝日曜〜6＝土曜）を出す。Excel の `WEEKDAY(値)-1` と同じ。
 * 1900 年基準：Excel は 1900-02-29（通し番号 60）を実在の日として数えるので、通し番号 1（1900-01-01）を日曜とする
 * 連続した数え方になる（61 以降は実際の暦と合う）。1904 年基準：通し番号 0 が 1904-01-01（金曜）。
 * 時刻の部分（小数）は切り捨てる。
 */
export declare function weekdayOfSerial(serial: number, date1904: boolean): number;
/** 曜日の字句を含む書式の準備（区画ごとの字句）。含まなければ null */
export interface JaFormat {
    sections: FormatSection[];
}
/** 書式コードに曜日の字句があれば字句に分けたものを返す。無ければ null（そのまま ssf に渡せる） */
export declare function prepareJaFormat(code: string): JaFormat | null;
/**
 * 曜日の字句を、値から計算した曜日を引用符で包んだ形に置き換えた書式コードを返す。
 * 区画の選び方は ssf に任せるため、全部の区画を同じ曜日で置き換える。負の値の区画（第 2 区画）は ssf が
 * 絶対値で日付を出すので、曜日も絶対値から出す。値が有限でなければ何もしない。
 * 日付の範囲の外（Excel が #### にする値）で、使う区画に曜日があれば、セル全体を空にする書式 `""` を返す
 * （曜日の無い日付の書式も ssf が空にするのにそろえる）。
 */
export declare function applyJaFormat(prepared: JaFormat, value: number, date1904: boolean): string;
/** ssf の使う部分（`format.ts` の SsfApi と同じ形） */
interface SsfLike {
    format(fmt: string | number, val: unknown, opts?: object): string;
    is_date(fmt: string): boolean;
}
/**
 * ssf を包み、曜日の字句を含む書式コードだけ、ssf に渡す前に置き換える。
 * `is_date` は曜日の字句があれば真にする（`aaa` だけの書式は ssf では日付にならないため）。
 * 字句の分解は書式コードごとに 1 回だけ（上限付きで覚える）。
 */
export declare function withJaTokens<T extends SsfLike>(inner: T): SsfLike;
export {};
