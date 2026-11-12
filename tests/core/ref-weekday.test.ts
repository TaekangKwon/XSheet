/// <reference types="node" />
/**
 * 実際の Excel で作った曜日の書式の基準ファイルとの照合（Issue #25・ADR-0007 §6）。
 * `fixtures/excel/ref-weekday.xlsx`（1900 年基準）・`ref-weekday-1904.xlsx`（1904 年基準）は
 * `make-ref-weekday.ps1` が Excel 16（ja-JP）の COM で作り、各セルの表示（Range.Text）を
 * `ref-weekday.expected.json` に書いた。
 */
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { tokenizeFormat, weekdayOfSerial } from '../../src/core/format-ja';
import { openXlsx, type XlsxSheet } from '../../src/core/index';

/** Excel の式 `=WEEKDAY(値)`・`=TEXT(値,"aaa")` の結果（エラーなら value2 は null） */
interface ExpectedWeekday {
  date1904: boolean;
  serial: number;
  kind: string;
  value2: number | string | null;
}

/**
 * skip に当たるセルの数（全 1026 のうち。埋め 23・負の値の「-」78・小文字の am/pm 42・曜日の無い区画の #### 28）。
 * #### の残り 100 は「空」として照合する。
 */
const SKIPPED = 171;

interface ExpectedCell {
  book: string;
  sheet: string;
  cell: string;
  serial: number;
  formatCode: string;
  text: string;
  hashAtMaxWidth?: boolean;
}

const expected = JSON.parse(
  readFileSync(new URL('../../fixtures/excel/ref-weekday.expected.json', import.meta.url), 'utf8'),
) as { cells: ExpectedCell[]; weekday: ExpectedWeekday[] };

/** A1 形式 → [行, 列]（1 始まり） */
function rc(a1: string): [number, number] {
  const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
  let c = 0;
  for (const ch of m[1]!) c = c * 26 + ch.charCodeAt(0) - 64;
  return [Number(m[2]), c];
}

/** ブックを開いて、シート名 → シート */
async function openBook(file: string): Promise<Map<string, XlsxSheet>> {
  const bytes = readFileSync(new URL(`../../fixtures/excel/${file}`, import.meta.url));
  const book = await openXlsx(bytes);
  const sheets = new Map<string, XlsxSheet>();
  for (const info of book.sheets) sheets.set(info.name, await book.openSheet(info.index));
  return sheets;
}

/** セルの表示文字 */
function textAt(sheet: XlsxSheet, a1: string): string | undefined {
  const [r, c] = rc(a1);
  return sheet.getRows(r, r)[0]?.cells[c - 1]?.text;
}

/**
 * XSheet で期待する文字。Excel が日付の範囲外（1900 年基準の負の値・9999/12/31 の先）で #### にするセル
 * （`hashAtMaxWidth`）は空にする（曜日の無い日付の書式も ssf が空にするのにそろえる）。
 */
const want = (c: ExpectedCell) => (c.hashAtMaxWidth === true ? '' : c.text);

/**
 * Excel と合わせない所。曜日の置き換えではなく ssf の受け持ちの違いで、曜日の無い書式（`m/d`・`"x"`）でも同じになる。
 * - `*a`：Excel は列幅いっぱいまで a を繰り返す（幅で変わる）。ssf は埋めの文字を出さない。
 * - 負の値で Excel が先頭に「-」を付けるもの：ssf は文字だけの書式（`\a`・置き換えた後の `"土"`）では「-」を付けず、
 *   1904 年基準の負の日付は空にする。
 * - 小文字の `am/pm`・`a/p`：Excel は書いたまま小文字で出し、ssf は大文字にする。
 * - 使う区画に曜日の無い #### のセル：ssf は 1904 年基準の 9999/12/31 の先も日付で出す（`ddd`・`m/d`）。
 */
const skip = (c: ExpectedCell) =>
  c.formatCode.includes('*') ||
  (c.serial < 0 && c.text.startsWith('-')) ||
  /am\/pm|a\/p/.test(c.formatCode) ||
  (c.hashAtMaxWidth === true && !usedSectionHasWeekday(c.formatCode, c.serial));

/** 値に使う区画（負の値は第 2 区画があればそれ）に曜日の字句があるか */
function usedSectionHasWeekday(code: string, value: number): boolean {
  const sections = tokenizeFormat(code);
  const section = sections[value < 0 && sections.length >= 2 ? 1 : 0] ?? [];
  return section.some((t) => t.kind === 'weekday');
}

describe('Excel 実ファイル：曜日の書式（ref-weekday）', () => {
  for (const file of ['ref-weekday.xlsx', 'ref-weekday-1904.xlsx']) {
    it(`${file}：セルの表示が Excel と同じ`, async () => {
      const sheets = await openBook(file);
      const cells = expected.cells.filter((c) => c.book === file && !skip(c));
      expect(cells.length).toBeGreaterThan(100);
      const diffs = cells
        .map((c) => ({
          at: `${c.sheet}!${c.cell}`,
          code: c.formatCode,
          serial: c.serial,
          excel: want(c),
          // 空の文字のセルは XSheet では undefined（空のセル）になる
          xsheet: textAt(sheets.get(c.sheet)!, c.cell) ?? '',
        }))
        .filter((d) => d.xsheet !== d.excel);
      expect(diffs).toEqual([]);
    });
  }

  it('除いたセルの数（ssf の受け持ちの違い）が増えていない', () => {
    // 増えたら、曜日の置き換えの不具合が skip に紛れていないかを見る
    expect(expected.cells.filter(skip).length).toBe(SKIPPED);
  });

  it('weekdayOfSerial は Excel の WEEKDAY と同じ（1900 年基準・1904 年基準）', () => {
    const rows = expected.weekday.filter(
      (w) => w.kind === 'WEEKDAY' && typeof w.value2 === 'number',
    );
    expect(rows.length).toBeGreaterThan(20);
    for (const w of rows) {
      expect({
        serial: w.serial,
        date1904: w.date1904,
        wd: weekdayOfSerial(w.serial, w.date1904) + 1,
      }).toEqual({
        serial: w.serial,
        date1904: w.date1904,
        wd: w.value2,
      });
    }
  });
});
