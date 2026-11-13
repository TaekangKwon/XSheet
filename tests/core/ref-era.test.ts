/// <reference types="node" />
/**
 * 実際の Excel で確かめた和暦の書式・組み込みの日付の書式との照合（Issue #32・ADR-0007 §6）。
 * `fixtures/excel/ref-era.xlsx`（1900 年基準）・`ref-era-1904.xlsx`（1904 年基準）は `make-ref-era.ps1` が XML から組み立てた合成のファイルで、
 * Excel 16（ja-JP）の COM で開いて（保存しない）各セルの表示（Range.Text）を `ref-era.expected.json` に書いた。
 * シート custom は独自の書式（既定の読み取りで照合）、builtin は組み込みの番号（`builtinDates: 'japanese-era'` で照合）。
 */
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { eraOfSerial, tokenizeFormat } from '../../src/core/format-ja';
import { openXlsx, type OpenXlsxOptions, type XlsxSheet } from '../../src/core/index';

interface ExpectedCell {
  book: string;
  date1904: boolean;
  sheet: 'custom' | 'builtin';
  cell: string;
  serial: number;
  numFmtId: number;
  formatCode: string | null;
  formatRead: string;
  text: string;
  hashAtMaxWidth?: boolean;
}

const expected = JSON.parse(
  readFileSync(new URL('../../fixtures/excel/ref-era.expected.json', import.meta.url), 'utf8'),
) as { cells: ExpectedCell[] };

/** A1 形式 → [行, 列]（1 始まり） */
function rc(a1: string): [number, number] {
  const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
  let c = 0;
  for (const ch of m[1]!) c = c * 26 + ch.charCodeAt(0) - 64;
  return [Number(m[2]), c];
}

/** ブックを開いて、シート名 → シート */
async function openBook(file: string, options?: OpenXlsxOptions): Promise<Map<string, XlsxSheet>> {
  const bytes = readFileSync(new URL(`../../fixtures/excel/${file}`, import.meta.url));
  const book = await openXlsx(bytes, options);
  const sheets = new Map<string, XlsxSheet>();
  for (const info of book.sheets) sheets.set(info.name, await book.openSheet(info.index));
  return sheets;
}

function textAt(sheet: XlsxSheet, a1: string): string {
  const [r, c] = rc(a1);
  return sheet.getRows(r, r)[0]?.cells[c - 1]?.text ?? '';
}

/** XSheet で期待する文字。Excel が #### にする日付の範囲外のセルは空（ref-weekday と同じ扱い） */
const want = (c: ExpectedCell) => (c.hashAtMaxWidth === true ? '' : c.text);

/** 書式コード（組み込みは Excel の読み戻し）に、言語の指定の無い区画で元号の無い e だけがあるか */
function bareEraYear(code: string): boolean {
  return tokenizeFormat(code).some(
    (s) =>
      !s.some((t) => t.kind === 'bracket' && t.text.startsWith('[$-')) &&
      !s.some((t) => t.kind === 'era') &&
      s.some((t) => t.kind === 'eraYear'),
  );
}

/**
 * Excel と合わせない所。
 * - 1904 年基準の負の日付（Excel は「-」を付けて出す）：ssf は空にする（ref-weekday と同じ）。
 * - 言語の指定が無く元号の無い `e`・`ee`（`"ggg"e` も）：日本語版 Excel は和暦の年（33）だが、Excel の言語環境で変わるので
 *   西暦のまま（ADR-0007 §6。組み込みの番号と同じ理由）。
 */
const skip = (c: ExpectedCell) =>
  (c.serial < 0 && c.text.startsWith('-')) || (c.formatCode !== null && bareEraYear(c.formatCode));

/**
 * 除いたセルの数（全 2952 のうち。1904 年基準の負の日付 144・元号の無い e（`e`・`ee`・`"ggg"e`）117）。
 * 増えたら、和暦の不具合が skip に紛れていないかを見る。
 */
const SKIPPED = 261;

describe('Excel 実ファイル：和暦と組み込みの日付の書式（ref-era）', () => {
  for (const file of ['ref-era.xlsx', 'ref-era-1904.xlsx']) {
    it(`${file}：独自の書式（custom）が Excel と同じ`, async () => {
      const sheets = await openBook(file);
      const cells = expected.cells.filter(
        (c) => c.book === file && c.sheet === 'custom' && !skip(c),
      );
      expect(cells.length).toBeGreaterThan(300);
      const diffs = cells
        .map((c) => ({
          at: c.cell,
          code: c.formatCode,
          serial: c.serial,
          excel: want(c),
          xsheet: textAt(sheets.get('custom')!, c.cell),
        }))
        .filter((d) => d.xsheet !== d.excel);
      expect(diffs).toEqual([]);
    });

    it(`${file}：builtinDates: 'japanese-era' の組み込みの番号（builtin）が Excel と同じ`, async () => {
      const sheets = await openBook(file, { builtinDates: 'japanese-era' });
      const cells = expected.cells.filter(
        (c) => c.book === file && c.sheet === 'builtin' && !skip(c),
      );
      expect(cells.length).toBeGreaterThan(300);
      const diffs = cells
        .map((c) => ({
          at: c.cell,
          id: c.numFmtId,
          serial: c.serial,
          excel: want(c),
          xsheet: textAt(sheets.get('builtin')!, c.cell),
        }))
        .filter((d) => d.xsheet !== d.excel);
      expect(diffs).toEqual([]);
    });
  }

  it('除いたセルの数が増えていない', () => {
    expect(expected.cells.filter(skip).length).toBe(SKIPPED);
  });

  it('既定（builtinDates を指定しない）では組み込みの 28 は西暦のまま', async () => {
    const sheets = await openBook('ref-era.xlsx');
    const c = expected.cells.find(
      (x) => x.book === 'ref-era.xlsx' && x.numFmtId === 28 && x.serial === 46337,
    )!;
    expect(c.text).toBe('令和8年11月11日');
    expect(textAt(sheets.get('builtin')!, c.cell)).toBe('2026/11/11');
  });

  it('eraOfSerial：元号の境目（1900 年基準・1904 年基準）', () => {
    const at = (s: number, d1904 = false) => {
      const e = eraOfSerial(s, d1904);
      return `${e.name}${e.year}`;
    };
    expect([0, 60, 4594, 4595, 9855, 9856, 32515, 32516, 43585, 43586].map((s) => at(s))).toEqual([
      '明治33',
      '明治33',
      '明治45',
      '大正1',
      '大正15',
      '昭和1',
      '昭和64',
      '平成1',
      '平成31',
      '令和1',
    ]);
    expect(at(0, true)).toBe('明治37');
    expect(at(42124, true)).toBe('令和1');
  });
});
