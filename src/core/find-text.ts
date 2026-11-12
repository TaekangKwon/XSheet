// 正: docs/adr/0007-moobius-viewer-api.md §1
/**
 * シートの中の文字の検索 `findText`（M7-4・Moobius の依頼 R1）。
 * 表示形式を当てた後の文字（`CellView.text`）を、`getRows` で数百行ずつ読んで照らす。
 * - 非同期で、約 10ms 調べるごとに `setTimeout` で 1 回休む（大きなシートでも画面を止めない）。
 * - `signal` で中断できる（`AbortError` の `DOMException` で reject）。一致の数は `maxMatches` で打ち切る。
 * - 非表示の行・列は、描く側と同じく `createGeometry` の高さ・幅 0 で決める。
 * - 結合は左上のセルだけを調べる。`getRows` は範囲の上で切れた結合の左上の文字を範囲の最初の行に置くので、
 *   2 回目からは 1 行上から読んで、その行は数えない（切れた結合の文字を 2 回数えないため）。
 * core なので DOM・Node 固有の API を使わない（`setTimeout`・`DOMException` は大域から引き、型は自前で持つ）。
 */
import { createGeometry } from './geometry';
import type { XlsxSheet } from './types';

/** `findText` の指定 */
export interface FindTextOptions {
  /** `'partial'`＝文字の一部に一致・`'cell'`＝セルの文字全体が一致（既定 `'partial'`） */
  match?: 'partial' | 'cell';
  /** 英字などの大文字と小文字を区別しない（既定 true）。全角と半角はいつも区別する */
  ignoreCase?: boolean;
  /** セルの中の改行（`\r\n`・`\n`・`\r`）を照合のときだけ除く（既定 true）。位置は元の文字で返す */
  ignoreNewlines?: boolean;
  /** 非表示の行・列（高さ・幅 0 の行・列も）を調べない（既定 true） */
  skipHidden?: boolean;
  /** 中断（`AbortController` の `signal`）。中断されたら `AbortError`（`DOMException`）で reject する */
  signal?: { readonly aborted: boolean };
  /** 返す一致の数の上限（既定 10,000。100,000 より大きい値は 100,000 にする） */
  maxMatches?: number;
}

/** 1 つの一致 */
export interface TextMatch {
  /** 行（1 始まり） */
  r: number;
  /** 列（1 始まり） */
  c: number;
  /** 一致の始まり（`CellView.text` 上の UTF-16 の位置） */
  start: number;
  /** 一致の終わり（含まない） */
  end: number;
}

/** `findText` の結果 */
export interface FindTextResult {
  /** 一致（行 → 列 → セルの中の位置の順） */
  matches: TextMatch[];
  /** `maxMatches` で打ち切ったか（上限を超える一致があった） */
  truncated: boolean;
}

/** `maxMatches` の既定と上限 */
const DEFAULT_MAX_MATCHES = 10_000;
const MAX_MATCHES_LIMIT = 100_000;
/** 1 回に続けて調べる時間の目安（ms）。次の区切りを読むと超えそうなら休む */
const SLICE_MS = 10;
/**
 * 1 回の `getRows` で読むセルの数の目安（行数は列数で割って 1〜500 行）。実機（Chrome）で 1 区切りが約 8ms。
 * 1 区切りが SLICE_MS を超えたら、行数を半分にしていく（遅い端末でも休む間隔を保つ）
 */
const CELLS_PER_CHUNK = 5_000;

/** 大域の `setTimeout`・`DOMException`（core の型には DOM が無いので自前で書く） */
const host = globalThis as unknown as {
  setTimeout?: (cb: () => void, ms: number) => unknown;
  DOMException?: new (message: string, name: string) => Error;
};

/** 次のタスクまで休む（描画・スクロール・入力を先に通す） */
const pause = (): Promise<void> =>
  new Promise((resolve) => {
    if (host.setTimeout !== undefined) host.setTimeout(resolve, 0);
    else resolve();
  });

/** 中断の例外（`DOMException` が無い環境では name だけを合わせた Error） */
function abortError(): Error {
  const message = '検索を中断した';
  if (host.DOMException !== undefined) return new host.DOMException(message, 'AbortError');
  const e = new Error(message);
  e.name = 'AbortError';
  return e;
}

/** 照合に使う文字と、その各位置（UTF-16）から元の文字の位置への対応。対応が null なら位置は同じ */
interface Folded {
  text: string;
  /** 照合の文字の位置 i を作った元の文字の始まり */
  starts: number[] | null;
  /** 同じく元の文字の終わり（含まない） */
  ends: number[] | null;
}

const NEWLINE = /[\r\n]/;
/** 語末のシグマ（ς）。`toLowerCase` が前後の文字で σ と ς を出し分けるので σ にそろえる */
const FINAL_SIGMA = /ς/g;

/** 1 文字（コードポイント）を小文字にする（前後の文字に依らない形） */
const lower1 = (ch: string): string => ch.toLowerCase().replace(FINAL_SIGMA, 'σ');

/** 照合の文字を作る（大文字と小文字・改行の扱い）。長さが変わるときだけ位置の対応を作る */
function fold(s: string, ignoreCase: boolean, ignoreNewlines: boolean): Folded {
  const lowered = ignoreCase ? s.toLowerCase().replace(FINAL_SIGMA, 'σ') : s;
  const dropNewlines = ignoreNewlines && NEWLINE.test(s);
  // 速い道：長さが同じ（各位置が元の同じ位置）で、除く改行も無い
  if (lowered.length === s.length && !dropNewlines)
    return { text: lowered, starts: null, ends: null };

  // 遅い道：1 文字ずつ（İ のように小文字で長くなる文字・改行を除くとき）
  let text = '';
  const starts: number[] = [];
  const ends: number[] = [];
  for (let i = 0; i < s.length;) {
    const cp = s.codePointAt(i)!;
    const w = cp > 0xffff ? 2 : 1;
    const ch = s.slice(i, i + w);
    if (!(ignoreNewlines && (ch === '\r' || ch === '\n'))) {
      const f = ignoreCase ? lower1(ch) : ch;
      for (let k = 0; k < f.length; k++) {
        starts.push(i);
        ends.push(i + w);
      }
      text += f;
    }
    i += w;
  }
  return { text, starts, ends };
}

/** `maxMatches` を整える（数でない・無い値は既定。0 未満は 0、上限より大きい値は上限） */
function maxMatchesOf(v: number | undefined): number {
  if (typeof v !== 'number' || Number.isNaN(v)) return DEFAULT_MAX_MATCHES;
  return Math.min(MAX_MATCHES_LIMIT, Math.max(0, Math.floor(v)));
}

/**
 * シートの中の文字を探す。表示形式を当てた後の文字（`CellView.text`）で照らす。
 * - 1 セルの中の重ならない一致をすべて返す。結合したセルは左上のセルだけ。空の検索語は空の結果。
 * - 図形・グラフの中の文字は探さない。
 * - 約 10ms ごとに休むので、大きなシートでも画面は固まらない。
 * - ブック全体を探すときは、呼ぶ側がシートを 1 つずつ `openSheet` で開いて呼ぶ。
 * @param sheet `XlsxBook.openSheet` で開いたシート
 * @param query 検索語
 * @throws DOMException `signal` で中断された（`name` は `'AbortError'`）
 */
export async function findText(
  sheet: XlsxSheet,
  query: string,
  options: FindTextOptions = {},
): Promise<FindTextResult> {
  const signal = options.signal;
  // 関数で読む（休む間に変わる値なので、型の絞り込みを効かせない）
  const aborted = (): boolean => signal?.aborted === true;
  if (aborted()) throw abortError();

  const cellMatch = options.match === 'cell';
  const ignoreCase = options.ignoreCase !== false;
  const ignoreNewlines = options.ignoreNewlines !== false;
  const skipHidden = options.skipHidden !== false;
  const max = maxMatchesOf(options.maxMatches);

  const matches: TextMatch[] = [];
  const q = fold(String(query), ignoreCase, ignoreNewlines).text;
  const rows = sheet.rowCount;
  const cols = sheet.colCount;
  if (q === '' || rows <= 0 || cols <= 0) return { matches, truncated: false };

  // 非表示の行・列（描く側と同じく、高さ・幅 0 の行・列）
  const geom = skipHidden ? createGeometry(sheet.layout, rows, cols) : null;
  const hiddenCol = new Uint8Array(cols);
  if (geom !== null)
    for (let c = 1; c <= cols; c++) if (geom.colWidth(c) === 0) hiddenCol[c - 1] = 1;

  let chunk = Math.max(1, Math.min(500, Math.floor(CELLS_PER_CHUNK / cols)));
  let sliceStart = Date.now();

  // 次の from は読んだ最後の行の次（区切りの行数を途中で変えても、同じ行を 2 回読まない）
  for (let from = 1; from <= rows;) {
    const chunkStart = Date.now();
    const to = Math.min(rows, from + chunk - 1);
    // 2 回目からは 1 行上から読む（その行で切れた結合の左上を、from の行に置かせないため）
    const got = sheet.getRows(from === 1 ? 1 : from - 1, to);
    for (const row of got) {
      if (row.r < from) continue;
      if (geom !== null && geom.rowHeight(row.r) === 0) continue;
      const cells = row.cells;
      for (let c = 1; c <= cells.length; c++) {
        const cell = cells[c - 1];
        if (cell === null || cell === undefined || cell.text === '' || hiddenCol[c - 1] === 1)
          continue;
        const t = fold(cell.text, ignoreCase, ignoreNewlines);
        if (cellMatch) {
          if (t.text !== q) continue;
          if (matches.length >= max) return { matches, truncated: true };
          matches.push({ r: row.r, c, start: 0, end: cell.text.length });
          continue;
        }
        for (let at = t.text.indexOf(q); at >= 0; at = t.text.indexOf(q, at + q.length)) {
          if (matches.length >= max) return { matches, truncated: true };
          const last = at + q.length - 1;
          matches.push({
            r: row.r,
            c,
            start: t.starts === null ? at : t.starts[at]!,
            end: t.ends === null ? last + 1 : t.ends[last]!,
          });
        }
      }
    }
    const now = Date.now();
    const chunkMs = now - chunkStart;
    // 次の区切りを読んだら目安を超えそうなら、先に休む
    if (to < rows && now - sliceStart + chunkMs >= SLICE_MS) {
      if (chunkMs > SLICE_MS && chunk > 1) chunk = Math.ceil(chunk / 2);
      await pause();
      if (aborted()) throw abortError();
      sliceStart = Date.now();
    }
    from = to + 1;
  }
  if (aborted()) throw abortError();
  return { matches, truncated: false };
}
