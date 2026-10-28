// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  fontXml,
  patternFillXml,
  type CellSpec,
  type SheetSpec,
  type XlsxSpec,
} from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  mount,
  openBook,
  renderSpec,
  scrollBody,
  stubClientSize,
  wait,
} from './render-sheet.helpers';

/**
 * `renderSheet` の書式（塗り・文字の装飾。ISSUE-100・ADR-0030 §決定 1・9・10）の DOM 描画テスト（InterHub `spreadsheet.format.render.test.ts` を写したもの）。
 * jsdom は CSS を適用しない（レイアウトも持たない。clientWidth／Height はテストで固定する）ため、見た目は目視（TV-44）で確かめ、
 * ここでは `<style>` の置き場所と寿命・td のクラス・はみ出しの `is-over-fill`・`###` の判定・作り直しの時機を確かめる。
 * シートの切替（タブ）で `<style>` を作り直さないテスト（TD-100-14）は M2-4（renderWorkbook）で写す。
 */

/**
 * 書式の表：xf 0＝書式なし、1＝赤の塗り、2＝太字、3＝白の塗り、4＝文字サイズ 22（比 2）、5＝赤の塗り＋太字、6＝赤の塗り＋赤い文字。
 * fonts[0] は既定のフォント（sz 11）
 */
const BOOK: Partial<XlsxSpec> = {
  fonts: [
    fontXml('<sz val="11"/>'),
    fontXml('<b/>'),
    fontXml('<sz val="22"/>'),
    fontXml('<color rgb="FFFF0000"/>'),
  ],
  fills: [
    patternFillXml('none'),
    patternFillXml('solid', 'rgb="FFFF0000"'),
    patternFillXml('solid', 'rgb="FFFFFFFF"'),
  ],
  xfs: [0, 0, 0, 0, 0, 0, 0],
  xfFontIds: [0, 0, 1, 0, 2, 1, 3],
  xfFillIds: [0, 1, 0, 2, 0, 1, 1],
};

const text = (v: string, s?: number): CellSpec => ({
  v,
  t: 'inlineStr',
  ...(s !== undefined ? { s } : {}),
});

function render(sheets: SheetSpec[], extra: Partial<XlsxSpec> = BOOK): Promise<HTMLElement> {
  return renderSpec({ sheets, ...extra });
}

let restoreSize: () => void;

beforeEach(() => {
  restoreSize = stubClientSize(800, 400);
});

afterEach(() => {
  cleanupViews();
  restoreSize();
});

const styleOf = (root: HTMLElement): HTMLStyleElement[] => [...root.querySelectorAll('style')];
/** 描いた table の td（行番号の td・余白の行を除く）。行ごとの配列 */
function dataRows(root: HTMLElement): HTMLTableCellElement[][] {
  return [
    ...root.querySelectorAll<HTMLTableRowElement>('table.xs-grid tbody tr:not(.xs-spacer)'),
  ].map((tr) => [...tr.querySelectorAll<HTMLTableCellElement>('td:not(.xs-rownum)')]);
}
const xfClasses = (td: Element): string[] =>
  [...td.classList].filter((c) => c.startsWith('xs-xf-'));

describe('renderSheet（.xlsx・書式／ISSUE-100）', () => {
  it('TD-100-10: 表示の根に <style> が 1 枚だけあり、連番の属性が付く。規則は連番の属性で絞られる', async () => {
    const root = await render([{ name: 'S', rows: [[text('a', 1), text('b', 2), text('c')]] }]);
    await wait();
    const styles = styleOf(root);
    expect(styles).toHaveLength(1);
    expect(styles[0]!.parentElement).toBe(root); // document.head ではなく表示の根の子
    expect(styles[0]!.hasAttribute('nonce')).toBe(false);
    const seq = root.getAttribute('data-xs-fmt');
    expect(seq).toMatch(/^\d+$/);
    const css = styles[0]!.textContent!;
    expect(css).toContain(`.xs-sheet[data-xs-fmt="${seq}"] td.xs-xf-0 {`);
    expect(css).toContain('background-color: #ff0000');
    expect(css).toContain('font-weight: 700');
    expect(document.head.querySelector('style')).toBeNull();
  });

  it('TD-100-10: nonce を指定すると書式の <style> に付く（CSP 用）', async () => {
    const root = await renderSpec({ sheets: [{ name: 'S', rows: [[text('a', 1)]] }], ...BOOK }, 1, {
      nonce: 'abc123',
    });
    expect(styleOf(root)[0]!.getAttribute('nonce')).toBe('abc123');
  });

  it('TD-100-11: 書式のあるセルの td にだけ xs-xf-<番号> が付く。style 属性には何も書かない', async () => {
    const root = await render([
      { name: 'S', rows: [[text('a', 1), text('b', 2), text('c'), { s: 1 }]] },
    ]);
    await wait();
    const tds = dataRows(root)[0]!;
    const red = xfClasses(tds[0]!);
    const bold = xfClasses(tds[1]!);
    expect(red).toHaveLength(1);
    expect(bold).toHaveLength(1);
    expect(red[0]).not.toBe(bold[0]);
    expect(xfClasses(tds[2]!)).toEqual([]);
    expect(xfClasses(tds[3]!)).toEqual(red); // 値の無い塗りのセルにも付く（同じ書式＝同じクラス）
    // 書式の無いセルの className は空。色は style に書かない
    expect(tds[2]!.className).toBe('');
    for (const td of tds) {
      expect(td.getAttribute('style')).toBeNull();
    }
  });

  it('TD-100-12: 結合セルは左上のセルの書式（結合で隠れるセルの書式は見ない）', async () => {
    const root = await render([
      { name: 'S', rows: [[text('m', 1), { s: 2 }, text('z')]], merges: ['A1:B1'] },
    ]);
    await wait();
    const tds = dataRows(root)[0]!;
    expect(tds[0]!.colSpan).toBe(2);
    expect(xfClasses(tds[0]!)).toHaveLength(1);
    const css = styleOf(root)[0]!.textContent!;
    const cls = xfClasses(tds[0]!)[0]!;
    const rule = css.split('\n').find((l) => l.includes(`td.${cls} `))!;
    expect(rule).toContain('background-color: #ff0000'); // xf 1（赤の塗り）
    expect(rule).not.toContain('font-weight'); // 隠れる B1 の xf 2（太字）ではない
  });

  it('TD-100-13: 書式の無いブックは <style> が空で、td に xs-xf- クラスが付かない', async () => {
    const root = await render([{ name: 'S', rows: [[text('a'), text('b')]] }], {});
    await wait();
    expect(styleOf(root)).toHaveLength(1);
    expect(styleOf(root)[0]!.textContent).toBe('');
    for (const td of dataRows(root).flat()) {
      expect(td.className).toBe('');
    }
  });

  it('TD-100-15: 別の表示は自分の <style>・別の連番を持つ（同時に置ける）。destroy すると <style> は消える', async () => {
    const m1 = await mount(
      await openBook({ sheets: [{ name: 'S', rows: [[text('a', 1)]] }], ...BOOK }),
    );
    await wait();
    const seq1 = m1.root.getAttribute('data-xs-fmt');
    expect(styleOf(m1.root)).toHaveLength(1);
    const m2 = await mount(
      await openBook({ sheets: [{ name: 'S', rows: [[text('b', 2)]] }], ...BOOK }),
    );
    await wait();
    const seq2 = m2.root.getAttribute('data-xs-fmt');
    expect(seq2).not.toBe(seq1);
    expect(styleOf(m2.root)).toHaveLength(1);
    expect(styleOf(m2.root)[0]!.textContent).toContain(`data-xs-fmt="${seq2}"`);
    expect(styleOf(m2.root)[0]!.textContent).not.toContain(`data-xs-fmt="${seq1}"`);
    // 前の表示はそのまま（InterHub は別のファイルを開くと前の表示を後始末したが、ここでは表示ごとに独立）
    expect(styleOf(m1.root)).toHaveLength(1);
    m1.view.destroy();
    expect(styleOf(m1.root)).toHaveLength(0); // 規則が残らない
    expect(document.querySelectorAll('style')).toHaveLength(1);
    m2.view.destroy();
    expect(styleOf(m2.root)).toHaveLength(0);
    expect(document.querySelectorAll('style')).toHaveLength(0);
  });

  it('TD-100-16: 白以外の塗りが絡むはみ出しに is-over-fill。絡まない・白の塗りだけが絡む・塗りなしの装飾だけのはみ出しには付かない', async () => {
    const root = await render([
      {
        name: 'S',
        rows: [
          [text('元が赤', 1), null, null], // 1: 元のセルが白以外の塗り
          [text('空が赤'), { s: 1 }, null], // 2: 右の空セルが白以外の塗り
          [text('空が白'), { s: 3 }, null], // 3: 右の空セルが白の塗りだけ
          [text('太字だけ', 2), null, null], // 4: 塗りなしの装飾
          [text('素のまま'), null, null], // 5: 書式なし
          [text('先が赤'), text('値'), { s: 1 }], // 6: はみ出さない（右隣が値）。赤は数えた列の外
        ],
      },
    ]);
    await wait();
    // 6 行ぶん（その先は画面を埋める空の行）
    const boxes = dataRows(root)
      .slice(0, 6)
      .map((tds) => tds[0]!.querySelector('.xs-cell')!);
    const over = boxes.map((b) => b.classList.contains('is-over-fill'));
    const spill = boxes.map((b) => b.classList.contains('is-spill'));
    expect(spill).toEqual([true, true, true, true, true, false]);
    expect(over).toEqual([true, true, false, false, false, false]);
  });

  it('TD-100-n: 塗りの状態の違う列をまたぐはみ出しの文字は、元の自動色のとき列ごとに色を塗り分ける（background-clip: text）。明示色・塗りの絡まないはみ出しは単色', async () => {
    const root = await render([
      {
        name: 'S',
        rows: [
          [text('元が赤', 1), null, null], // 1: 黒 → 画面の文字色
          [text('空が赤'), { s: 1 }, null], // 2: 画面の文字色 → 黒 → 画面の文字色
          [text('赤い文字', 6), null, null], // 3: 元が明示色＝単色
          [text('空が白'), { s: 3 }, null], // 4: 白は塗りなし扱い（is-over-fill なし）
          [text('素のまま'), null, null], // 5: 書式なし
        ],
      },
    ]);
    await wait();
    const spans = dataRows(root)
      .slice(0, 5)
      .map((tds) => tds[0]!.querySelector<HTMLElement>('.xs-text')!);
    const image = (el: HTMLElement): string => el.style.getPropertyValue('background-image');
    expect(image(spans[0]!)).toMatch(
      /^linear-gradient\(to right, #000000 0px \d+px, var\(--xs-fg\) \d+px \d+px\)$/,
    );
    expect(image(spans[1]!)).toMatch(
      /^linear-gradient\(to right, var\(--xs-fg\) 0px \d+px, #000000 \d+px \d+px, var\(--xs-fg\) \d+px \d+px\)$/,
    );
    // 塗り分けた span は文字の形で切り抜く（color は透明）
    expect(spans[0]!.style.getPropertyValue('background-clip')).toBe('text');
    expect(spans[0]!.style.getPropertyValue('color')).toBe('transparent');
    for (const i of [2, 3, 4]) {
      expect(image(spans[i]!), String(i)).toBe('');
      expect(spans[i]!.hasAttribute('style'), String(i)).toBe(false);
    }
  });

  it('TD-100-n: 表の範囲の外の余白（データの先の列・行）にも、行の書式 → 列の書式の順で塗りの書式の番号が付く。表の範囲は広がらない', async () => {
    const root = await render([
      {
        name: 'S',
        rows: [[text('a'), text('b')], [null], [null]],
        cols: [{ min: 5, max: 5, style: 1 }], // 列 E は赤（データは A〜B だけ）
        rowAttrs: { 3: { s: 3, customFormat: true } }, // 行 3 は白の塗り（行の書式が列の書式に勝つ）
      },
    ]);
    await wait();
    const rows = dataRows(root);
    const red = xfClasses(rows[0]![4]!);
    expect(red).toHaveLength(1);
    expect(xfClasses(rows[0]![3]!)).toEqual([]); // 列 D は指定なし
    expect(xfClasses(rows[1]![4]!)).toEqual(red); // 行 2（データの行・書式の無い行）の余白の列 E
    expect(xfClasses(rows[9]!.at(4)!)).toEqual(red); // 行 10（データの先の余白の行）の列 E
    expect(xfClasses(rows[9]![3]!)).toEqual([]);
    expect(xfClasses(rows[0]![1]!)).toEqual([]); // 表の中の列 B には何も付かない
    // 行 3 の行の書式（白＝別の番号）は、余白の全列（列 E も）に当たる
    const white = xfClasses(rows[2]![4]!);
    expect(white).toHaveLength(1);
    expect(white).not.toEqual(red);
    expect(xfClasses(rows[2]![3]!)).toEqual(white);
    expect(xfClasses(rows[2]![5]!)).toEqual(white);
  });

  it('TD-100-17: 大きな文字の数字は推定幅に比を掛けて ###、同じ数字でも書式なしなら収まる（元の値は title）', async () => {
    const root = await render([{ name: 'S', rows: [[{ v: 12345, s: 4 }, { v: 12345 }]] }]);
    await wait();
    const [big, plain] = dataRows(root)[0]!;
    const bigBox = big!.querySelector<HTMLElement>('.xs-cell')!;
    expect(bigBox.textContent).toMatch(/^#+$/);
    expect(bigBox.title).toBe('12345');
    expect(plain!.querySelector('.xs-cell')!.textContent).toBe('12345');
  });

  it('TD-100-18: 表を作り直さないスクロールでは td に触らない。作り直す時（範囲の外へ出た時）に新しい table へクラスが付き、<style> は同じ', async () => {
    const rows: CellSpec[][] = Array.from({ length: 300 }, (_, i) => [
      text(`r${i + 1}`, i % 2 === 0 ? 1 : 2),
    ]);
    const root = await render([{ name: 'S', rows }]);
    await wait();
    const style0 = styleOf(root)[0]!;
    const table0 = root.querySelector('table.xs-grid')!;
    const td0 = dataRows(root)[0]![0]!;
    const cls0 = td0.className;
    expect(cls0).toMatch(/^xs-xf-\d+$/);
    scrollBody(root, { top: 20 });
    await wait(60);
    expect(root.querySelector('table.xs-grid')).toBe(table0);
    expect(dataRows(root)[0]![0]).toBe(td0);
    expect(td0.className).toBe(cls0);
    scrollBody(root, { top: 4000 });
    await wait(60);
    const table1 = root.querySelector('table.xs-grid')!;
    expect(table1).not.toBe(table0);
    const firstShown = dataRows(root)[0]![0]!;
    expect(firstShown.className).toMatch(/^xs-xf-\d+$/);
    expect(styleOf(root)).toHaveLength(1);
    expect(styleOf(root)[0]).toBe(style0);
  });
});
