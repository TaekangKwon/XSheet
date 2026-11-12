// XSheet: 日本語の書式の字句（曜日 aaa・aaaa）。InterHub に無い処理なので写しの外に置く（ADR-0007 §6・ADR-0003 §4-1）
/**
 * 書式コードを小さな字句に分け、ssf が知らない日本語の字句（曜日）を、計算した文字を引用符で包んだ形に置き換える。
 * 置き換えるのは、引用符の中・`\` の後・`_`／`*` の後・`[...]` の中・`AM/PM`・`A/P` 以外にある `a` の並びだけ。
 * 曜日は `Date` を使わず、日付の通し番号から整数の計算で出す（時間帯のずれが入らない）。
 * M7-8 の和暦（`g`・`e`）もこの字句の分解を使う。
 */

/** 字句の種類 */
export type FormatTokenKind =
  /** `"..."` の文字（text は引用符を含む元の形） */
  | 'quoted'
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

/** 曜日になる `a` の並びの長さの下限（`aaa`＝短い形。それより短い並びは文字のまま） */
const WEEKDAY_MIN = 3;
/** これ以上の長さは長い形（`aaaa`＝「月曜日」） */
const WEEKDAY_LONG = 4;

const WEEKDAY_SHORT = ['日', '月', '火', '水', '木', '金', '土'] as const;
const WEEKDAY_FULL = [
  '日曜日',
  '月曜日',
  '火曜日',
  '水曜日',
  '木曜日',
  '金曜日',
  '土曜日',
] as const;

/**
 * 曜日の字句の文字か。Excel は同じ文字（大文字小文字も区別）の並びで数える：`AAA` も曜日（保存すると `aaa` に直る）、
 * `aAa` は 1 文字ずつの文字のまま、`Aaaa` は `A` と `aaa` に分かれる（ref-weekday.xlsx で確かめた）。
 */
const isWeekdayChar = (ch: string | undefined) => ch === 'a' || ch === 'A';

/** 表示する文字の無い言語の指定 `[$-411]`・`[$-ja-JP]`（`[$€-2]` のような通貨の記号付きは含まない） */
const LOCALE_ONLY = /^\[\$-[^\]]*\]$/;

/** code の i 文字目から s（大文字小文字を問わない）が続くか */
function startsWithCi(code: string, i: number, s: string): boolean {
  return code.slice(i, i + s.length).toUpperCase() === s;
}

/**
 * 書式コードを `;` の区画ごとの字句に分ける。
 * 閉じていない引用符・角かっこは末尾までを 1 つの字句にする（壊れた書式でも止まらない）。
 * 1 文字ずつ前へ進むだけなので、長さに比例した時間で終わる。
 */
export function tokenizeFormat(code: string): FormatSection[] {
  const sections: FormatToken[][] = [[]];
  let cur = sections[0] as FormatToken[];
  let i = 0;
  while (i < code.length) {
    const ch = code[i] as string;
    if (ch === ';') {
      cur = [];
      sections.push(cur);
      i++;
    } else if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      const stop = end < 0 ? code.length : end + 1;
      cur.push({ kind: 'quoted', text: code.slice(i, stop) });
      i = stop;
    } else if (ch === '[') {
      const end = code.indexOf(']', i + 1);
      const stop = end < 0 ? code.length : end + 1;
      cur.push({ kind: 'bracket', text: code.slice(i, stop) });
      i = stop;
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      cur.push({ kind: 'escaped', text: code.slice(i, i + 2) });
      i += 2;
    } else if (startsWithCi(code, i, 'AM/PM')) {
      cur.push({ kind: 'ampm', text: code.slice(i, i + 5) });
      i += 5;
    } else if (startsWithCi(code, i, 'A/P')) {
      cur.push({ kind: 'ampm', text: code.slice(i, i + 3) });
      i += 3;
    } else if (isWeekdayChar(ch)) {
      let j = i;
      while (code[j] === ch && !startsWithCi(code, j, 'A/P')) j++;
      const run = code.slice(i, j);
      if (run.length >= WEEKDAY_MIN) cur.push({ kind: 'weekday', text: run });
      else for (const c of run) cur.push({ kind: 'char', text: c });
      i = j;
    } else {
      cur.push({ kind: 'char', text: ch });
      i++;
    }
  }
  return sections;
}

/**
 * Excel の日付の通し番号から曜日（0＝日曜〜6＝土曜）を出す。Excel の `WEEKDAY(値)-1` と同じ。
 * 1900 年基準：Excel は 1900-02-29（通し番号 60）を実在の日として数えるので、通し番号 1（1900-01-01）を日曜とする
 * 連続した数え方になる（61 以降は実際の暦と合う）。1904 年基準：通し番号 0 が 1904-01-01（金曜）。
 * 時刻の部分（小数）は切り捨てる。
 */
export function weekdayOfSerial(serial: number, date1904: boolean): number {
  const day = Math.floor(serial);
  const offset = date1904 ? 5 : 6;
  return (((day + offset) % 7) + 7) % 7;
}

/** 曜日の字句を含む書式の準備（区画ごとの字句）。含まなければ null */
export interface JaFormat {
  sections: FormatSection[];
}

/** 書式コードに曜日の字句があれば字句に分けたものを返す。無ければ null（そのまま ssf に渡せる） */
export function prepareJaFormat(code: string): JaFormat | null {
  if (!code.includes('a') && !code.includes('A')) return null;
  const sections = tokenizeFormat(code);
  return sections.some((s) => s.some((t) => t.kind === 'weekday')) ? { sections } : null;
}

/**
 * 曜日の字句を、値から計算した曜日を引用符で包んだ形に置き換えた書式コードを返す。
 * 区画の選び方は ssf に任せるため、全部の区画を同じ曜日で置き換える。負の値の区画（第 2 区画）は ssf が
 * 絶対値で日付を出すので、曜日も絶対値から出す。値が有限でなければ何もしない。
 * 日付の範囲の外（Excel が #### にする値）で、使う区画に曜日があれば、セル全体を空にする書式 `""` を返す
 * （曜日の無い日付の書式も ssf が空にするのにそろえる）。
 */
export function applyJaFormat(prepared: JaFormat, value: number, date1904: boolean): string {
  if (!Number.isFinite(value)) return joinSections(prepared.sections, null);
  if (!inDateRange(value, date1904)) {
    if (usedSectionHasWeekday(prepared.sections, value)) return '""';
    return joinSections(prepared.sections, OUT_OF_RANGE);
  }
  return joinSections(prepared.sections, weekdayOfSerial(Math.abs(value), date1904));
}

/** 条件 `[<=45000]` などの区画の印 */
const CONDITION = /^\[[<>=]/;

/**
 * 値に使う区画に曜日の字句があるか（日付の範囲の外のときだけ使う）。
 * 条件付きの区画があるときは選び方を真似ず、どこかに曜日があれば真にする。
 * 条件が無ければ Excel の既定（負の値は第 2 区画があればそれ、0 は第 3 区画があればそれ、ほかは第 1 区画）。
 */
function usedSectionHasWeekday(sections: readonly FormatSection[], value: number): boolean {
  const hasWeekday = (s: FormatSection | undefined) =>
    s?.some((t) => t.kind === 'weekday') ?? false;
  if (sections.some((s) => s.some((t) => t.kind === 'bracket' && CONDITION.test(t.text))))
    return sections.some(hasWeekday);
  const idx = value < 0 && sections.length >= 2 ? 1 : value === 0 && sections.length >= 3 ? 2 : 0;
  return hasWeekday(sections[idx]);
}

/** 9999/12/31 の次の日の通し番号（1900 年基準・1904 年基準） */
const SERIAL_END_1900 = 2_958_466;
const SERIAL_END_1904 = 2_957_004;
/** 日付の範囲の外の印（曜日を空にする） */
const OUT_OF_RANGE = -1;

/**
 * Excel が日付として出せる値か。1900 年基準は 0〜9999/12/31（負の値は ####）、1904 年基準は負の値も
 * 「-」付きで出せるので絶対値で 9999/12/31 まで（ref-weekday.xlsx で確かめた）。
 * 範囲の外は Excel が #### にし、ssf は日付の書式を空にするので、曜日も空にしてそろえる。
 */
function inDateRange(value: number, date1904: boolean): boolean {
  return date1904 ? Math.abs(value) < SERIAL_END_1904 : value >= 0 && value < SERIAL_END_1900;
}

/** ssf の使う部分（`format.ts` の SsfApi と同じ形） */
interface SsfLike {
  format(fmt: string | number, val: unknown, opts?: object): string;
  is_date(fmt: string): boolean;
}

/** 準備の結果を覚えておく書式コードの数の上限（グラフの書式など、ファイルから来る数に上限を置く） */
const CACHE_MAX = 256;

/**
 * ssf を包み、曜日の字句を含む書式コードだけ、ssf に渡す前に置き換える。
 * `is_date` は曜日の字句があれば真にする（`aaa` だけの書式は ssf では日付にならないため）。
 * 字句の分解は書式コードごとに 1 回だけ（上限付きで覚える）。
 */
export function withJaTokens<T extends SsfLike>(inner: T): SsfLike {
  const cache = new Map<string, JaFormat | null>();
  const prepare = (code: string): JaFormat | null => {
    let p = cache.get(code);
    if (p === undefined) {
      p = prepareJaFormat(code);
      if (cache.size >= CACHE_MAX) cache.clear();
      cache.set(code, p);
    }
    return p;
  };
  return {
    format(fmt, val, opts) {
      if (typeof fmt === 'string' && typeof val === 'number') {
        const p = prepare(fmt);
        if (p) {
          const date1904 = (opts as { date1904?: boolean } | undefined)?.date1904 === true;
          return inner.format(applyJaFormat(p, val, date1904), val, opts);
        }
      }
      return inner.format(fmt, val, opts);
    },
    is_date(fmt) {
      return inner.is_date(fmt) || prepare(fmt) !== null;
    },
  };
}

/**
 * 区画をつなぎ直す。wd があれば曜日の字句を置き換え、表示する文字の無い言語の指定（`[$-411]`）を外す
 * （置き換えで日付の字句が無くなると、ssf は `[$-411]` を通貨の記号と読んで「$」を出すため）。
 */
function joinSections(sections: readonly FormatSection[], wd: number | null): string {
  if (wd === null) return sections.map((s) => s.map((t) => t.text).join('')).join(';');
  return sections
    .map((s) =>
      s
        .map((t) => {
          if (t.kind === 'weekday') {
            if (wd === OUT_OF_RANGE) return '""';
            return `"${(t.text.length >= WEEKDAY_LONG ? WEEKDAY_FULL : WEEKDAY_SHORT)[wd] as string}"`;
          }
          if (t.kind === 'bracket' && LOCALE_ONLY.test(t.text)) return '';
          return t.text;
        })
        .join(''),
    )
    .join(';');
}
