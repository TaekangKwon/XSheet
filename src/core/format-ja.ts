// XSheet: 日本語の書式の字句（曜日 aaa・aaaa・和暦 g・e・r）。InterHub に無い処理なので写しの外に置く（ADR-0007 §6・ADR-0003 §4-1）
/**
 * 書式コードを小さな字句に分け、ssf が知らない日本語の字句（曜日・元号・和暦の年）を、計算した文字を引用符で包んだ形に置き換える。
 * 置き換えるのは、引用符の中・`\` の後・`_`／`*` の後・`[...]` の中・`AM/PM`・`A/P` 以外にある `a` の並びだけ。
 * 曜日は `Date` を使わず、日付の通し番号から整数の計算で出す（時間帯のずれが入らない）。
 * 和暦（`g`・`e`・`r`）は M7-8。正解は Excel 16（ja-JP）で作った `fixtures/excel/ref-weekday.xlsx`・`ref-era.xlsx`。
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
  /** 元号 `g`・`gg`・`ggg`（大文字も。text は元の並び） */
  | 'era'
  /** 和暦の年 `e`・`ee`（大文字も） */
  | 'eraYear'
  /** 和暦の `r`（＝`ee`）・`rr`（＝`gggee`）（大文字も） */
  | 'eraR'
  /** `General`（大文字小文字を問わない。中の e・G を字句にしない） */
  | 'general'
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
    } else if (startsWithCi(code, i, 'GENERAL')) {
      cur.push({ kind: 'general', text: code.slice(i, i + 7) });
      i += 7;
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
    } else if (ERA_KIND[ch.toLowerCase()] !== undefined && !isExponent(code, i)) {
      // 同じ文字の並び（大文字小文字は問わない。Excel は G・E・R も g・e・r と読む）
      const lower = ch.toLowerCase();
      let j = i;
      while (j < code.length && (code[j] as string).toLowerCase() === lower && !isExponent(code, j))
        j++;
      cur.push({ kind: ERA_KIND[lower] as FormatTokenKind, text: code.slice(i, j) });
      i = j;
    } else {
      cur.push({ kind: 'char', text: ch });
      i++;
    }
  }
  return sections;
}

/** 和暦の字句になる文字（小文字） */
const ERA_KIND: Readonly<Record<string, FormatTokenKind>> = { g: 'era', e: 'eraYear', r: 'eraR' };

/** 指数の `E+`・`E-`・`e+`・`e-`（数値の書式）か */
function isExponent(code: string, i: number): boolean {
  const ch = code[i];
  return (ch === 'e' || ch === 'E') && (code[i + 1] === '+' || code[i + 1] === '-');
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

/**
 * 区画の和暦の扱い（ref-era.xlsx で確かめた）。
 * - `ja`：和暦で出す（言語の指定が日本語、または指定が無く元号 `g`・`rr` がある）
 * - `gannen`：`ja` に加え、2 文字以上の元号の後の年が 1 なら「元」（`[$-ja-JP-x-gannen]`）
 * - `other`：日本語でない言語の指定（`[$-409]` など）。元号は空、年（`e`）は ssf の西暦のまま
 * - `none`：和暦の字句が無い、または指定が無く `e` だけ（Excel は開く言語環境で変わるので西暦のまま。ADR-0007 §6）
 */
type EraMode = 'ja' | 'gannen' | 'other' | 'none';

/** 曜日・和暦の字句を含む書式の準備（区画ごとの字句と和暦の扱い）。含まなければ null */
export interface JaFormat {
  sections: FormatSection[];
  /** 区画ごとの和暦の扱い（sections と同じ長さ） */
  eras: EraMode[];
}

/** 言語の指定（`[$-411]`）が日本語か。16 進の LCID は下位 16 ビットが 0x411、タグは `ja`・`ja-…`（`-x-gannen` は元年） */
function localeKind(bracket: string): 'ja' | 'gannen' | 'other' {
  const tag = bracket.slice(3, -1);
  if (/^[0-9a-f]+$/i.test(tag)) return (parseInt(tag, 16) & 0xffff) === 0x411 ? 'ja' : 'other';
  const lower = tag.toLowerCase();
  if (lower !== 'ja' && !lower.startsWith('ja-')) return 'other';
  return lower.includes('-x-gannen') ? 'gannen' : 'ja';
}

const isEraToken = (t: FormatToken) =>
  t.kind === 'era' || t.kind === 'eraYear' || t.kind === 'eraR';

/** 区画の和暦の扱いを決める */
function eraModeOf(section: FormatSection): EraMode {
  if (!section.some(isEraToken)) return 'none';
  const locale = section.find((t) => t.kind === 'bracket' && LOCALE_ONLY.test(t.text));
  if (locale !== undefined) return localeKind(locale.text);
  const named = section.some((t) => t.kind === 'era' || (t.kind === 'eraR' && t.text.length >= 2));
  return named ? 'ja' : 'none';
}

/** 区画に置き換える字句があるか */
function sectionIsActive(section: FormatSection | undefined, mode: EraMode | undefined): boolean {
  if (section === undefined) return false;
  return section.some(
    (t) =>
      t.kind === 'weekday' ||
      ((mode === 'ja' || mode === 'gannen') && isEraToken(t)) ||
      (mode === 'other' && t.kind === 'era'),
  );
}

/** 書式コードに曜日・和暦の字句があれば字句に分けたものを返す。無ければ null（そのまま ssf に渡せる） */
export function prepareJaFormat(code: string): JaFormat | null {
  if (!/[aAgGeErR]/.test(code)) return null;
  const sections = tokenizeFormat(code);
  const eras = sections.map(eraModeOf);
  return sections.some((s, i) => sectionIsActive(s, eras[i])) ? { sections, eras } : null;
}

/**
 * 曜日の字句を、値から計算した曜日を引用符で包んだ形に置き換えた書式コードを返す。
 * 区画の選び方は ssf に任せるため、全部の区画を同じ曜日で置き換える。負の値の区画（第 2 区画）は ssf が
 * 絶対値で日付を出すので、曜日も絶対値から出す。値が有限でなければ何もしない。
 * 日付の範囲の外（Excel が #### にする値）で、使う区画に曜日があれば、セル全体を空にする書式 `""` を返す
 * （曜日の無い日付の書式も ssf が空にするのにそろえる）。
 */
export function applyJaFormat(prepared: JaFormat, value: number, date1904: boolean): string {
  if (!Number.isFinite(value)) return joinSections(prepared, null);
  if (!inDateRange(value, date1904)) {
    if (usedSectionIsActive(prepared, value)) return '""';
    return joinSections(prepared, OUT_OF_RANGE);
  }
  const abs = Math.abs(value);
  return joinSections(prepared, {
    weekday: weekdayOfSerial(abs, date1904),
    era: eraOfSerial(abs, date1904),
  });
}

/** 元号（始まりの日の 1900 年基準の通し番号・名前・略称・記号・西暦から引く数） */
const ERAS = [
  { start: 43_586, name: '令和', short: '令', letter: 'R', base: 2018 }, // 2019-05-01
  { start: 32_516, name: '平成', short: '平', letter: 'H', base: 1988 }, // 1989-01-08
  { start: 9_856, name: '昭和', short: '昭', letter: 'S', base: 1925 }, // 1926-12-25
  { start: 4_595, name: '大正', short: '大', letter: 'T', base: 1911 }, // 1912-07-30
  { start: -Infinity, name: '明治', short: '明', letter: 'M', base: 1867 },
] as const;

/** 通し番号の元号と和暦の年 */
export interface EraOfSerial {
  /** 元号の名前（`令和`） */
  name: string;
  /** 略称（`令`） */
  short: string;
  /** 記号（`R`） */
  letter: string;
  /** 和暦の年（1 始まり） */
  year: number;
}

/**
 * Excel の日付の通し番号（0 以上。時刻の部分は切り捨て）から元号と和暦の年を出す。
 * 1900 年基準の 0〜60（1900-01-00〜Excel の 1900-02-29）は 1900 年、61 以降は 1899-12-30 からの日数で年を出す
 * （UTC で数えるので時間帯に依らない）。1904 年基準は 1462 を足して 1900 年基準にする。
 * Excel の範囲（1900 年〜）では明治 33 年から。
 */
export function eraOfSerial(serial: number, date1904: boolean): EraOfSerial {
  const day = Math.floor(serial) + (date1904 ? 1462 : 0);
  const year =
    day < 61 ? 1900 : new Date(Date.UTC(1899, 11, 30) + day * 86_400_000).getUTCFullYear();
  const era = ERAS.find((e) => day >= e.start) ?? ERAS[ERAS.length - 1]!;
  return { name: era.name, short: era.short, letter: era.letter, year: year - era.base };
}

/** 条件 `[<=45000]` などの区画の印 */
const CONDITION = /^\[[<>=]/;

/**
 * 値に使う区画に置き換える字句（曜日・和暦）があるか（日付の範囲の外のときだけ使う）。
 * 条件付きの区画があるときは選び方を真似ず、どこかにあれば真にする。
 * 条件が無ければ Excel の既定（負の値は第 2 区画があればそれ、0 は第 3 区画があればそれ、ほかは第 1 区画）。
 */
function usedSectionIsActive(prepared: JaFormat, value: number): boolean {
  const { sections, eras } = prepared;
  if (sections.some((s) => s.some((t) => t.kind === 'bracket' && CONDITION.test(t.text))))
    return sections.some((s, i) => sectionIsActive(s, eras[i]));
  // 「@」のある最後の区画は文字の区画で、数値には使わない（`ggge"年";@` の負の値は第 1 区画。ref-era.xlsx で確かめた）
  const last = sections[sections.length - 1];
  const textSection =
    sections.length >= 2 && last?.some((t) => t.kind === 'char' && t.text === '@');
  const count = textSection ? sections.length - 1 : sections.length;
  const idx = value < 0 && count >= 2 ? 1 : value === 0 && count >= 3 ? 2 : 0;
  return sectionIsActive(sections[idx], eras[idx]);
}

/** 9999/12/31 の次の日の通し番号（1900 年基準・1904 年基準） */
const SERIAL_END_1900 = 2_958_466;
const SERIAL_END_1904 = 2_957_004;
/** 日付の範囲の外の印（曜日・和暦を空にする） */
const OUT_OF_RANGE = -1;

/** 値から計算した曜日と和暦 */
interface DateParts {
  weekday: number;
  era: EraOfSerial;
}

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
 * ssf を包み、曜日・和暦の字句を含む書式コードだけ、ssf に渡す前に置き換える。
 * `is_date` は置き換える字句があれば真にする（`aaa`・`[$-411]g` だけの書式は ssf では日付にならないため）。
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

/** 元号の字句の文字（1 文字＝記号 `R`・2 文字＝略称 `令`・3 文字以上＝名前 `令和`） */
function eraName(era: EraOfSerial, length: number): string {
  return length === 1 ? era.letter : length === 2 ? era.short : era.name;
}

/** 和暦の年の文字（1 文字は桁を詰め、2 文字以上は 2 桁。元年の条件が立つときの 1 年は「元」） */
function eraYearText(era: EraOfSerial, twoDigits: boolean, gannen: boolean): string {
  if (gannen && era.year === 1) return '元';
  return twoDigits ? String(era.year).padStart(2, '0') : String(era.year);
}

/**
 * 区画をつなぎ直す。parts があれば曜日・和暦の字句を置き換え、表示する文字の無い言語の指定（`[$-411]`）を外す
 * （置き換えで日付の字句が無くなると、ssf は `[$-411]` を通貨の記号と読んで「$」を出すため）。
 * 日本語でない言語の指定の区画（`[$-409]ggge`）は、元号を空にし、年は ssf の西暦に任せる。
 * 元年（`[$-ja-JP-x-gannen]`）は、同じ区画で 2 文字以上の元号の後にある年だけ（`ge`・`e"年"` は 1。ref-era.xlsx で確かめた）。
 */
function joinSections(prepared: JaFormat, parts: DateParts | typeof OUT_OF_RANGE | null): string {
  const { sections, eras } = prepared;
  if (parts === null) return sections.map((s) => s.map((t) => t.text).join('')).join(';');
  return sections
    .map((s, i) => {
      const mode = eras[i];
      const japanese = mode === 'ja' || mode === 'gannen';
      const gannen = mode === 'gannen';
      let named = false; // この区画で 2 文字以上の元号を前に出したか（元年の条件）
      return s
        .map((t, k) => {
          // ssf は日付の書式の「.」（`ge.m.d` の m.d）を読めずに例外にするので、和暦の区画では引用符で包む。
          // 秒の小数（`ss.0`）の「.」は後ろが 0 なので残す
          if (japanese && t.kind === 'char' && t.text === '.' && s[k + 1]?.text !== '0')
            return '"."';
          if (t.kind === 'weekday') {
            if (parts === OUT_OF_RANGE) return '""';
            return `"${(t.text.length >= WEEKDAY_LONG ? WEEKDAY_FULL : WEEKDAY_SHORT)[parts.weekday] as string}"`;
          }
          if (t.kind === 'era' && mode === 'other') return '';
          if (japanese && isEraToken(t)) {
            if (parts === OUT_OF_RANGE) return '""';
            const era = parts.era;
            if (t.kind === 'era') {
              if (t.text.length >= 2) named = true;
              return `"${eraName(era, t.text.length)}"`;
            }
            if (t.kind === 'eraYear')
              return `"${eraYearText(era, t.text.length >= 2, gannen && named)}"`;
            // r＝ee、rr 以上＝gggee（Excel は保存のときにこう読み替える）
            if (t.text.length === 1) return `"${eraYearText(era, true, gannen && named)}"`;
            named = true;
            return `"${era.name}${eraYearText(era, true, gannen)}"`;
          }
          if (t.kind === 'bracket' && LOCALE_ONLY.test(t.text)) return '';
          return t.text;
        })
        .join('');
    })
    .join(';');
}

/**
 * 組み込みの書式 27〜36・50〜58 の、日本語版 Excel での書式コード
 * （ref-era.xlsx の builtin のシートで、Excel 16（ja-JP）の NumberFormat を読んで確かめた）。
 * `builtinDates: 'japanese-era'` のときだけ使う（既定は `format.ts` の表のまま。ADR-0007 §6）。
 */
export const JA_BUILTIN_DATES: ReadonlyMap<number, string> = new Map<number, string>([
  [27, '[$-411]ge.m.d'],
  [28, '[$-411]ggge"年"m"月"d"日"'],
  [29, '[$-411]ggge"年"m"月"d"日"'],
  [30, 'm/d/yy'],
  [31, 'yyyy"年"m"月"d"日"'],
  [32, 'h"時"mm"分"'],
  [33, 'h"時"mm"分"ss"秒"'],
  [34, 'yyyy"年"m"月"'],
  [35, 'm"月"d"日"'],
  [36, '[$-411]ge.m.d'],
  [50, '[$-411]ge.m.d'],
  [51, '[$-411]ggge"年"m"月"d"日"'],
  [52, 'yyyy"年"m"月"'],
  [53, 'm"月"d"日"'],
  [54, '[$-411]ggge"年"m"月"d"日"'],
  [55, 'yyyy"年"m"月"'],
  [56, 'm"月"d"日"'],
  [57, '[$-411]ge.m.d'],
  [58, '[$-411]ggge"年"m"月"d"日"'],
]);

/**
 * 組み込みの日付の書式 27〜36・50〜58 の見せ方。
 * `'japanese-era'`＝日本語版 Excel と同じ（和暦を含む）。`'default'`（既定）＝今までの西暦の表。
 */
export type BuiltinDates = 'default' | 'japanese-era';
