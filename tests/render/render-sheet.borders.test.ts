// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { borderXml, type CellSpec, type SheetSpec } from '../../src/testing/make-xlsx';
import { cleanupViews, renderSpec, scrollBody, stubClientSize, wait } from './render-sheet.helpers';

/**
 * `renderSheet` の罫線層（ISSUE-98・ADR-0029 §決定 1・8）の DOM 描画テスト（InterHub `spreadsheet.borders.render.test.ts` を写したもの）。
 * jsdom はレイアウトを持たない（clientWidth／Height はテストで固定する）ため、見た目は目視（TV-43）で確かめ、
 * ここでは層の位置（DOM の順）・作り直しの時機を確かめる。シートの切替（タブ）のテストは M2-4（renderWorkbook）で写す。
 */

/** 罫線の表：0＝なし、1＝下の細線（自動）、2＝右の赤い太線 */
const BORDERS = [
  borderXml({ left: {} }),
  borderXml({ bottom: { style: 'thin', color: 'auto="1"' } }),
  borderXml({ right: { style: 'medium', color: 'rgb="FFFF0000"' } }),
];

const cell = (v: string, s?: number): CellSpec => ({
  v,
  t: 'inlineStr',
  ...(s !== undefined ? { s } : {}),
});

function render(sheets: SheetSpec[]): Promise<HTMLElement> {
  return renderSpec({
    font: { name: 'Calibri', size: 11 },
    sheets,
    xfs: [0, 0, 0],
    borders: BORDERS,
    xfBorderIds: [0, 1, 2],
  });
}

let restoreSize: () => void;

beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
});

afterEach(() => {
  cleanupViews();
  restoreSize();
});

const svgOf = (root: HTMLElement): SVGSVGElement =>
  root.querySelector<SVGSVGElement>('svg.xs-borders')!;
const pathsOf = (root: HTMLElement): SVGPathElement[] => [...svgOf(root).querySelectorAll('path')];

/** 下の細線のセルを n 行並べたシート（行 1〜n の A 列） */
const manyRows = (n: number): SheetSpec => ({
  name: 'S',
  rows: Array.from({ length: n }, (_, i) => [cell(`r${i + 1}`, 1), cell('x', 2)]),
});

describe('renderSheet（.xlsx・罫線層／ISSUE-98）', () => {
  it('TD-98-10: svg.xs-borders が 1 枚だけ、wrapper の中の table の直後・図の層の直前にある', async () => {
    const root = await render([manyRows(3)]);
    await wait();
    const wrap = root.querySelector('.xs-wrap')!;
    expect(root.querySelectorAll('svg.xs-borders')).toHaveLength(1);
    expect(
      [...wrap.children].map(
        (c) => c.tagName.toLowerCase() + '.' + (c.getAttribute('class') ?? ''),
      ),
    ).toEqual(['table.xs-grid', 'svg.xs-borders', 'div.xs-layer']);
  });

  it('TD-98-10: 罫線が path になる（下の細線は墨色・右の太線は赤を CSSOM で）。原点は A1（行番号の列・見出しの分ずらす）', async () => {
    const root = await render([manyRows(3)]);
    await wait();
    const ps = pathsOf(root);
    expect(ps.length).toBeGreaterThanOrEqual(2);
    const ink = ps.find((p) => p.classList.contains('is-ink'))!;
    expect(ink.getAttribute('style')).toBeNull();
    // 3 行ぶんの下の線：既定の列幅 64px・行高 20px。A 列の下は y＝19・39・59 の画素
    expect(ink.getAttribute('d')).toBe('M0 19.5H64M0 39.5H64M0 59.5H64');
    const red = ps.find((p) => !p.classList.contains('is-ink'))!;
    expect(red.style.getPropertyValue('stroke')).toBe('rgb(255, 0, 0)');
    expect(red.getAttribute('stroke-width')).toBe('2');
    // B 列の右の辺（x＝colLeft(3)−1＝127）に 3 行つながった 1 本
    expect(red.getAttribute('d')).toBe('M127 0V60');
    const svg = svgOf(root);
    expect(svg.style.top).toBe('24px'); // 見出し行の高さ（A1 原点）
    expect(Number.parseFloat(svg.style.left)).toBeGreaterThanOrEqual(32); // 行番号の列の幅
  });

  it('TD-98-11: 表を作り直さないスクロールでは作り直さない（path の要素が同じ）。作り直す時（範囲の外へ出た時）は作り直す', async () => {
    const root = await render([manyRows(300)]);
    await wait();
    const table0 = root.querySelector('table.xs-grid');
    const first = pathsOf(root)[0]!;
    scrollBody(root, { top: 20 });
    await wait(60);
    expect(root.querySelector('table.xs-grid')).toBe(table0);
    expect(pathsOf(root)[0]).toBe(first);
    scrollBody(root, { top: 4000 });
    await wait(60);
    expect(root.querySelector('table.xs-grid')).not.toBe(table0);
    expect(pathsOf(root)[0]).not.toBe(first);
    // 作り直した後も順序は保たれる
    const wrap = root.querySelector('.xs-wrap')!;
    expect([...wrap.children].map((c) => c.tagName.toLowerCase())).toEqual(['table', 'svg', 'div']);
  });

  it('TD-98-11: 罫線の無いシートでは中身が空（要素は残る）', async () => {
    const root = await render([manyRows(3), { name: 'T', rows: [[cell('plain')]] }]);
    await wait();
    expect(pathsOf(root).length).toBeGreaterThan(0);
    const plain = await renderSpec(
      {
        font: { name: 'Calibri', size: 11 },
        sheets: [manyRows(3), { name: 'T', rows: [[cell('plain')]] }],
        xfs: [0, 0, 0],
        borders: BORDERS,
        xfBorderIds: [0, 1, 2],
      },
      2,
    );
    await wait();
    expect(plain.querySelectorAll('svg.xs-borders')).toHaveLength(1);
    expect(pathsOf(plain)).toHaveLength(0);
  });

  it('TD-98-11: 罫線を持たないブック（罫線の表なし）でも空の svg が 1 枚あり、表は変わらない', async () => {
    const root = await renderSpec({ sheets: [{ name: 'S', rows: [[cell('a')]] }] });
    await wait();
    expect(svgOf(root).childElementCount).toBe(0);
    expect(root.querySelector('table.xs-grid')).not.toBeNull();
  });
});
