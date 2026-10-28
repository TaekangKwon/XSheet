// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { oneCellAnchor, spXml, wsDr, type CellSpec } from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  dataRows,
  headTexts,
  renderSpec,
  scrollBody,
  stubClientSize,
  wait as waitMs,
} from './render-sheet.helpers';

/**
 * `renderSheet` の列の仮想化（ISSUE-86 の修正・ADR-0026 §決定 8 の前提変更）の DOM テスト（InterHub `spreadsheet.cols.render.test.ts` を写したもの）。
 * 描く列は可視範囲の左右 1 画面分だけ。左右の描かない列は幅つきの余白の列（`<col>` 1 本＋各行に空の td 1 個）で置き換え、
 * 表の全幅は変えない。jsdom はレイアウトを持たない（clientWidth／Height はテストで固定）ため、実寸の位置は
 * 実測（Chrome）で確認し、ここでは DOM の構造・列数・余白の幅・結合の縮め・横スクロールでの描き直しを検証する。
 */

let restoreSize: () => void;

beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
});

afterEach(() => {
  cleanupViews();
  restoreSize();
});

const wait = (ms = 40): Promise<void> => waitMs(ms);

/** 列幅（px）の合計（colgroup の col の width）。表の全幅（行番号の列を含む） */
function colWidthSum(root: HTMLElement): number {
  return [...root.querySelectorAll<HTMLElement>('colgroup > col')]
    .map((c) => parseFloat(c.style.width))
    .reduce((a, b) => a + b, 0);
}

/** XFD（16,384 列目）にだけ値があるシート */
function xfdBook(): Promise<HTMLElement> {
  const row: CellSpec[] = ['a', 1];
  while (row.length < 16383) row.push(null);
  row.push('XFD-value');
  const rows: CellSpec[][] = [['r1'], ['r2'], row, ['r4']];
  return renderSpec({ font: { name: 'Calibri', size: 11 }, sheets: [{ name: 'S', rows }] });
}

describe('renderSheet（.xlsx・列の仮想化）', () => {
  it('TD-86-18: XFD にだけ値があっても、描く列は可視範囲の左右 1 画面分だけ（右は余白の列）。表の全幅は変わらない', async () => {
    const root = await xfdBook();
    const cols = root.querySelectorAll('colgroup > col').length;
    // 行番号の列＋描く列（幅 800 → 約 3 画面ぶん＝40 列前後）＋右の余白 1 列。16,385 列を描かない
    expect(cols).toBeLessThan(80);
    expect(cols).toBeGreaterThan(10);
    expect(root.querySelectorAll('td').length).toBeLessThan(80 * 60);
    // 全幅＝行番号の列＋16,384 列×64px（geometry の totalWidth）
    const table = root.querySelector('table')!;
    expect(colWidthSum(root)).toBe(parseFloat(table.style.width));
    const rowCount = Number(table.getAttribute('aria-rowcount')) - 1;
    const numWidth = Math.max(32, String(rowCount).length * 8 + 16);
    expect(parseFloat(table.style.width)).toBe(numWidth + 16384 * 64);
    // 右の余白の列：最後の col は広い
    const last = [...root.querySelectorAll<HTMLElement>('colgroup > col')].pop()!;
    expect(parseFloat(last.style.width)).toBeGreaterThan(1_000_000 - 80 * 64);
    // どの行も td の数は同じ（行番号＋描く列＋右の余白）
    const counts = new Set(dataRows(root).map((tr) => tr.cells.length));
    expect(counts.size).toBe(1);
    expect([...counts][0]).toBe(cols);
    // 見出しは描く列だけ（A から。左の余白は無い）
    expect(headTexts(root).slice(0, 3)).toEqual(['', 'A', 'B']);
  });

  it('TD-86-18: 横スクロールで描く範囲が変わったら描き直す（左右に余白の列・遠い列の値が出る）', async () => {
    const root = await xfdBook();
    expect(root.textContent).not.toContain('XFD-value');
    // 右端までスクロール
    scrollBody(root, { left: 16384 * 64 });
    await wait();
    expect(root.textContent).toContain('XFD-value');
    const heads = headTexts(root);
    expect(heads).toContain('XFD');
    expect(heads).not.toContain('A');
    // 左の余白の列（先頭が行番号の列、次が余白）。全幅は同じ
    const cols = [...root.querySelectorAll<HTMLElement>('colgroup > col')];
    expect(cols.length).toBeLessThan(80);
    expect(parseFloat(cols[1]!.style.width)).toBeGreaterThan(1_000_000 - 80 * 64);
    const table = root.querySelector('table')!;
    expect(colWidthSum(root)).toBe(parseFloat(table.style.width));
    // 戻ると A 列から描き直す
    scrollBody(root, { left: 0 });
    await wait();
    expect(headTexts(root)).toContain('A');
    expect(root.textContent).not.toContain('XFD-value');
  });

  it('TD-86-18: 範囲の中を少し動いただけでは描き直さない（同じ table のまま）', async () => {
    const root = await xfdBook();
    await wait(); // 最初の描き直し（取り付け後の rAF）を待つ
    const table = root.querySelector('table');
    scrollBody(root, { left: 64 * 3 });
    await wait();
    expect(root.querySelector('table')).toBe(table);
  });

  it('TD-86-18: 遠い列の図（oneCell・XFD）でも、表は描く列だけ。図は図のレイヤーにあり、列の仮想化と無関係に位置が出る', async () => {
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      sheets: [
        {
          name: 'S',
          rows: [['a']],
          drawing: {
            xml: wsDr([oneCellAnchor({ col: 16383, row: 2 }, 9525 * 40, 9525 * 20, spXml('far'))]),
          },
        },
      ],
    });
    expect(root.querySelectorAll('colgroup > col').length).toBeLessThan(80);
    expect(root.querySelectorAll('.xs-figure')).toHaveLength(0); // 可視範囲の外
    scrollBody(root, { left: 16383 * 64 });
    await wait();
    const fig = root.querySelector<HTMLElement>('.xs-figure')!;
    // 形で描く（ISSUE-92）：外接矩形は枠線 1px の miter の張り出し（miterlimit 4 の半分 × 1px = 2px）だけ外へ広がる
    expect(fig.style.left).toBe(`${16383 * 64 - 2}px`);
    expect(fig.style.top).toBe('38px');
    expect(root.querySelectorAll('colgroup > col').length).toBeLessThan(80);
  });

  it('TD-86-19: 結合の起点が描く範囲の左外なら、起点のセルを範囲の左端へ移し colSpan を縮める（右端で切れる結合も縮める）', async () => {
    // B3:BZ4（列 2〜78）の結合。値のある列は 130 まで（範囲は横スクロールで動かす）
    const rows: CellSpec[][] = [];
    for (let r = 1; r <= 6; r++) {
      const row: CellSpec[] = [];
      for (let c = 1; c <= 130; c++) row.push(`${c}.${r}`);
      rows.push(row);
    }
    rows[2]![1] = '結合タイトル';
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      sheets: [{ name: 'S', rows, merges: ['B3:BZ4'] }],
    });
    // 初期：起点（B）が範囲内。右端（描く範囲の最後の列）で切れて縮んでいる
    const first = dataRows(root)[2]!;
    const mergedTd = [...first.cells].find((td) => td.textContent === '結合タイトル')!;
    const spanAtStart = mergedTd.colSpan;
    expect(spanAtStart).toBeGreaterThan(10);
    expect(spanAtStart).toBeLessThan(77); // 全幅 77 列より縮んでいる
    // 範囲が結合の途中（列 60〜 付近）に来るまでスクロール。範囲の左外（B 側）に起点がある
    scrollBody(root, { left: 60 * 64 + 40 + 800 + 10 });
    await wait();
    const heads = headTexts(root);
    expect(heads).not.toContain('B');
    const row3 = dataRows(root).find((tr) => tr.getAttribute('aria-rowindex') === '4')!; // 行 3
    const clipped = [...row3.cells].find((td) => td.textContent === '結合タイトル');
    expect(clipped).toBeDefined();
    expect(clipped!.colSpan).toBeGreaterThanOrEqual(1);
    expect(clipped!.rowSpan).toBe(2);
    // 行 4（rowSpan で覆われる）には td を作らない：行 3 と同じ列数から結合の分だけ少ない
    const row4 = dataRows(root).find((tr) => tr.getAttribute('aria-rowindex') === '5')!;
    expect(row4.cells.length).toBeLessThan(row3.cells.length - 1 + clipped!.colSpan);
    // 表の行ごとの列の幅が合うこと（結合が列を食い違わせない）：td の colSpan の合計＝列数
    const total = (tr: HTMLTableRowElement, covered = 0): number =>
      [...tr.cells].reduce((a, td) => a + td.colSpan, 0) + covered;
    const cols = root.querySelectorAll('colgroup > col').length;
    expect(total(row3)).toBe(cols);
    expect(total(row4, clipped!.colSpan)).toBe(cols);
  });

  it('TD-86-18: 非表示の列は畳んだまま（余白に数えない）。広い範囲でも幅の合計が表の全幅に一致', async () => {
    const rows: CellSpec[][] = [Array.from({ length: 200 }, (_, i) => `v${i + 1}`)];
    const root = await renderSpec({
      font: { name: 'Calibri', size: 11 },
      sheets: [{ name: 'S', rows, cols: [{ min: 2, max: 4, hidden: true }] }],
    });
    const cols = [...root.querySelectorAll<HTMLElement>('colgroup > col')];
    const collapsed = cols.filter((c) => c.classList.contains('xs-collapsed'));
    expect(collapsed).toHaveLength(3);
    const table = root.querySelector('table')!;
    expect(colWidthSum(root)).toBe(parseFloat(table.style.width));
  });
});
