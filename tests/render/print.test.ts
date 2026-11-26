// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createGeometry, type XlsxBook } from '../../src/core/index';
import {
  planPrint,
  renderPrint,
  resolvePrintOptions,
  splitBands,
  type PrintResult,
} from '../../src/render/print';
import { PRINT_KIT } from '../../src/render/print-gate';
import { makePng } from '../../src/testing/make-image';
import {
  picXml,
  twoCellAnchor,
  wsDr,
  type CellSpec,
  type SheetSpec,
} from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  mount,
  mountWorkbook,
  openBook,
  stubClientSize,
  wait,
} from './render-sheet.helpers';
import { css, rule } from './style-css.helpers';

/**
 * 印刷の描画 `renderForPrint`（M7-20・ADR-0010 §2）。jsdom はレイアウトも印刷も持たないので、
 * ページへの切り分け（範囲・順・大きさ）・描いたセル・上限・中断・片付け・見た目のクラスを DOM で確かめる。
 * 印刷のプレビューで抜けが無いことは examples（`print.html`）で Chrome・Edge・Firefox を目視する。
 */

const PNG = makePng(40, 30);

/** blob URL の作成・解放の記録 */
const urls = { created: [] as string[], revoked: [] as string[] };
const savedUrl = { create: URL.createObjectURL, revoke: URL.revokeObjectURL };
let restoreSize: () => void;
const results: PrintResult[] = [];

beforeEach(() => {
  urls.created = [];
  urls.revoked = [];
  let n = 0;
  URL.createObjectURL = vi.fn(() => {
    const u = `blob:p/${++n}`;
    urls.created.push(u);
    return u;
  });
  URL.revokeObjectURL = vi.fn((u: string) => {
    urls.revoked.push(u);
  });
  restoreSize = stubClientSize(800, 400);
});

afterEach(() => {
  results.splice(0).forEach((r) => r.destroy());
  cleanupViews();
  restoreSize();
  URL.createObjectURL = savedUrl.create;
  URL.revokeObjectURL = savedUrl.revoke;
  document.body.replaceChildren();
});

/** 印刷の描画先（document.body の子） */
function newTarget(): HTMLElement {
  const t = document.createElement('div');
  document.body.append(t);
  return t;
}

/** 描いたページ */
const pagesOf = (t: HTMLElement): HTMLElement[] => [
  ...t.querySelectorAll<HTMLElement>(':scope > .xs-print > .xs-print-page'),
];

/** ページのセルの文字（行番号・空のセルを除く。上から、左から） */
const textsOf = (page: HTMLElement): string[] =>
  [...page.querySelectorAll('tbody td:not(.xs-rownum)')]
    .map((td) => td.textContent ?? '')
    .filter((s) => s !== '');

/** ページの行番号（tr の aria-rowindex − 1） */
const rowsOf = (page: HTMLElement): number[] =>
  [...page.querySelectorAll('tbody tr[aria-rowindex]')].map(
    (tr) => Number(tr.getAttribute('aria-rowindex')) - 1,
  );

/** 行 r（1 始まり）・列 c（1 始まり）にセルを置いた行の並び */
function grid(cells: Record<string, string>): CellSpec[][] {
  const rows: CellSpec[][] = [];
  for (const [key, v] of Object.entries(cells)) {
    const [r, c] = key.split(',').map(Number) as [number, number];
    const row = (rows[r - 1] ??= []);
    while (row.length < c - 1) row.push(null);
    row[c - 1] = v;
  }
  for (let i = 0; i < rows.length; i++) rows[i] ??= [];
  return rows;
}

/** 1 シートのブックを開いて表示し、表示とシートの寸法（100%）を返す */
async function setup(sheet: Omit<SheetSpec, 'name'>, extra?: Record<string, Uint8Array>) {
  const book = await openBook({
    font: { name: 'Calibri', size: 11 },
    sheets: [{ name: 'S', ...sheet }],
    extraFiles: extra,
  });
  const m = await mount(book);
  const sh = await book.openSheet(1);
  return { ...m, geo: createGeometry(sh.layout, 1000, 100), sheet: sh };
}

/** renderForPrint を呼び、結果を後始末の対象にする */
async function print(
  view: { renderForPrint: (t: HTMLElement, o?: object) => Promise<PrintResult> },
  target: HTMLElement,
  options?: object,
): Promise<PrintResult> {
  const r = await view.renderForPrint(target, options);
  results.push(r);
  return r;
}

describe('splitBands（ページの切れ目）', () => {
  const edge = (sizes: number[]) => (i: number) => sizes.slice(0, i - 1).reduce((a, b) => a + b, 0);
  const at = (sizes: number[]) => (v: number) => {
    let y = 0;
    for (let i = 0; i < sizes.length; i++) {
      if (sizes[i]! > 0 && v < y + sizes[i]!) return i + 1;
      y += sizes[i]!;
    }
    return sizes.length;
  };
  const split = (sizes: number[], limit: number) =>
    splitBands(1, sizes.length, edge(sizes), at(sizes), limit);

  it('収まるところで行の境で切る（ちょうど収まる行は前のページ）', () => {
    expect(split([20, 20, 20, 20, 20], 50)).toEqual([
      [1, 2],
      [3, 4],
      [5, 5],
    ]);
    expect(split([25, 25, 25, 25], 50)).toEqual([
      [1, 2],
      [3, 4],
    ]);
  });

  it('1 つだけでページを超える行は、その行だけのページ', () => {
    expect(split([20, 120, 20], 50)).toEqual([
      [1, 1],
      [2, 2],
      [3, 3],
    ]);
  });

  it('非表示（大きさ 0）だけの区切りは返さない', () => {
    expect(split([50, 0, 0], 40)).toEqual([[1, 1]]);
    // 区切りの後ろに続く非表示は、その区切りに含める（描く行は変わらない）
    expect(split([20, 20, 0, 0], 40)).toEqual([[1, 4]]);
    expect(split([20, 0, 20], 30)).toEqual([
      [1, 2],
      [3, 3],
    ]);
  });
});

describe('resolvePrintOptions', () => {
  it('既定は A4 縦・標準の余白（659×978）・縮めない・見出しと目盛り線なし・200000 セル', () => {
    expect(resolvePrintOptions()).toEqual({
      page: { width: 659, height: 978 },
      fit: false,
      headers: false,
      gridlines: false,
      maxCells: 200000,
    });
  });

  it.each([
    { page: { width: 0, height: 100 } },
    { page: { width: 100, height: Number.NaN } },
    { page: { width: Infinity, height: 100 } },
    { fit: 'height' },
    { maxCells: 0 },
    { maxCells: Number.NaN },
  ])('正しくない指定は RangeError：%o', (o) => {
    expect(() => resolvePrintOptions(o as never)).toThrow(RangeError);
  });
});

describe('renderForPrint：ページへの切り分け', () => {
  it('下へ、それから右へ。何も無いページは出さず、全部のセルが抜けずに 1 回ずつ出る', async () => {
    const cells: Record<string, string> = {};
    for (let r = 1; r <= 12; r++) {
      cells[`${r},1`] = `a${r}`;
      cells[`${r},10`] = `j${r}`;
    }
    const { view, geo } = await setup({ rows: grid(cells) });
    const rowH = geo.rowHeight(1);
    const colW = geo.colWidth(1);
    const t = newTarget();
    const r = await print(view, t, { page: { width: colW * 3, height: rowH * 5 } });
    const pages = pagesOf(t);
    // 列の区切り：A〜C・D〜F・G〜I・J。D〜I は空なので出さない。行の区切り：1〜5・6〜10・11〜12
    expect(r.pages).toBe(6);
    expect(pages).toHaveLength(6);
    expect(pages.map((p) => textsOf(p)[0])).toEqual(['a1', 'a6', 'a11', 'j1', 'j6', 'j11']);
    expect(pages.map(rowsOf)).toEqual([
      [1, 2, 3, 4, 5],
      [6, 7, 8, 9, 10],
      [11, 12],
      [1, 2, 3, 4, 5],
      [6, 7, 8, 9, 10],
      [11, 12],
    ]);
    const all = pages.flatMap(textsOf).sort();
    expect(all).toEqual(Object.values(cells).sort());
    expect(r.truncated).toBe(false);
    // ページは用紙の中身の大きさ
    expect(pages[0]!.style.width).toBe(`${colW * 3}px`);
    expect(pages[0]!.style.height).toBe(`${rowH * 5}px`);
  });

  it('表は範囲の外の余白を作らない（ページの幅・高さだけの表）', async () => {
    const { view, geo } = await setup({ rows: grid({ '1,1': 'a', '9,9': 'z' }) });
    const t = newTarget();
    await print(view, t, { page: { width: geo.colWidth(1) * 4, height: geo.rowHeight(1) * 4 } });
    const last = pagesOf(t).at(-1)!;
    const table = last.querySelector('table')!;
    expect(textsOf(last)).toEqual(['z']);
    expect(table.style.width).toBe(`${geo.colLeft(10) - geo.colLeft(9)}px`);
    expect(table.querySelectorAll('col')).toHaveLength(2); // 行番号の列（畳む）＋I
    expect(table.querySelectorAll('tr.xs-spacer td[style*="height: 0px"]').length).toBe(2);
  });

  it('非表示の行・列は印刷しない', async () => {
    const { view } = await setup({
      rows: grid({ '1,1': 'a1', '2,1': 'a2', '3,1': 'a3', '1,2': 'b1', '1,3': 'c1' }),
      rowAttrs: { 2: { hidden: true } },
      cols: [{ min: 2, max: 2, width: 10, hidden: true }],
    });
    const t = newTarget();
    await print(view, t);
    const [page] = pagesOf(t);
    expect(rowsOf(page!)).toEqual([1, 3]);
    const cols = page!.querySelectorAll('col');
    expect(cols[2]!.classList.contains('xs-collapsed')).toBe(true); // B は幅 0 で畳む（0 は行番号の列）
    expect(textsOf(page!)).not.toContain('a2');
  });

  it('1 行でページを超える行は、その行だけのページにして切る', async () => {
    const { view, geo } = await setup({
      rows: grid({ '1,1': 'a', '2,1': 'tall', '3,1': 'c' }),
      rowAttrs: { 2: { ht: 300 } },
    });
    const t = newTarget();
    const r = await print(view, t, { page: { width: 300, height: geo.rowHeight(1) * 3 } });
    expect(r.pages).toBe(3);
    expect(pagesOf(t).map(rowsOf)).toEqual([[1], [2], [3]]);
  });

  it('ページの境をまたぐ結合セルは、ページごとに範囲の中へ縮め、続きのページにも左上の文字を出す', async () => {
    const { view, geo } = await setup({
      rows: grid({ '1,1': 'merged', '10,2': 'z' }),
      merges: ['A1:A8'],
    });
    const t = newTarget();
    await print(view, t, { page: { width: 400, height: geo.rowHeight(1) * 5 } });
    const pages = pagesOf(t);
    expect(pages).toHaveLength(2);
    const span = (p: HTMLElement) =>
      [...p.querySelectorAll<HTMLTableCellElement>('tbody td:not(.xs-rownum)')].find(
        (td) => td.textContent === 'merged',
      )?.rowSpan;
    expect(span(pages[0]!)).toBe(5);
    expect(span(pages[1]!)).toBe(3);
  });

  it("fit: 'width' は全部の列を 1 ページの幅に収める（縦だけ分ける）。倍率は --xs-zoom", async () => {
    const cells: Record<string, string> = {};
    for (let c = 1; c <= 20; c++) cells[`1,${c}`] = `c${c}`;
    const { view, geo } = await setup({ rows: grid(cells) });
    const t = newTarget();
    const width = Math.floor(geo.colLeft(21) / 2);
    const r = await print(view, t, { page: { width, height: 500 }, fit: 'width' });
    expect(r.pages).toBe(1);
    const root = t.querySelector<HTMLElement>('.xs-print')!;
    const z = Number(root.style.getPropertyValue('--xs-zoom'));
    expect(z).toBeGreaterThan(0.45);
    expect(z).toBeLessThan(0.55);
    expect(parseFloat(root.querySelector('table')!.style.width)).toBeLessThanOrEqual(width);
    expect(textsOf(pagesOf(t)[0]!)).toHaveLength(20);
  });

  it('はみ出す文字の届く列まで印刷の範囲を広げる', async () => {
    const { view, sheet } = await setup({
      rows: grid({ '1,1': 'x'.repeat(60) }),
    });
    const plan = planPrint(PRINT_KIT, sheet, [], resolvePrintOptions());
    expect(plan.cols).toBeGreaterThan(3);
    const t = newTarget();
    await print(view, t);
    expect(pagesOf(t)[0]!.querySelectorAll('col').length).toBe(plan.cols + 1);
  });
});

describe('renderForPrint：見た目', () => {
  it('根は div.xs-sheet.xs-print（負の連番の data-xs-fmt・nonce 付きの書式の <style>）。target 自身には何も付けない', async () => {
    const book = await openBook({ sheets: [{ name: 'S', rows: [['a']] }] });
    const { view } = await mount(book, 1, { nonce: 'n0' });
    const t = newTarget();
    await print(view, t);
    expect(t.attributes).toHaveLength(0);
    expect(t.children).toHaveLength(1);
    const root = t.firstElementChild as HTMLElement;
    expect(root.className).toBe('xs-sheet xs-print');
    expect(Number(root.getAttribute('data-xs-fmt'))).toBeLessThan(0);
    const style = root.querySelector(':scope > style')!;
    expect(style.getAttribute('nonce')).toBe('n0');
  });

  it('見出しと目盛り線は既定で出さず、指定で出す（各ページに、そのページの列記号）', async () => {
    const { view, geo } = await setup({ rows: grid({ '1,1': 'a', '1,5': 'e' }) });
    const t1 = newTarget();
    await print(view, t1);
    const root1 = t1.querySelector('.xs-print')!;
    expect(root1.classList.contains('is-gridlines')).toBe(false);
    expect(root1.querySelector('thead')).toBeNull();

    const t2 = newTarget();
    await print(view, t2, {
      headers: true,
      gridlines: true,
      page: { width: 60 + geo.colWidth(1) * 2, height: 500 },
    });
    const root2 = t2.querySelector('.xs-print')!;
    expect(root2.classList.contains('is-gridlines')).toBe(true);
    const heads = pagesOf(t2 as HTMLElement).map((p) =>
      [...p.querySelectorAll('thead th:not(.xs-rownum)')].map((th) => th.textContent),
    );
    expect(heads[0]).toEqual(['A', 'B']);
    expect(heads.at(-1)).toContain('E');
  });

  it('ページは文書の外で組み、描き終えてから target に 1 回だけ取り付ける', async () => {
    const rows = Array.from({ length: 300 }, (_, i) => [`r${i + 1}`]);
    const { view } = await setup({ rows });
    const t = newTarget();
    const added: Node[] = [];
    const mo = new MutationObserver((list) => list.forEach((m) => added.push(...m.addedNodes)));
    mo.observe(t, { childList: true, subtree: true });
    const r = await print(view, t);
    await wait(0);
    mo.disconnect();
    expect(r.pages).toBeGreaterThan(1);
    expect(added).toHaveLength(1);
    expect(added[0]).toBe(t.firstElementChild);
  });

  it('今の表示（位置・表）には触らない', async () => {
    const rows = Array.from({ length: 200 }, (_, i) => [`r${i + 1}`]);
    const { view, root } = await setup({ rows });
    const body = view.scrollElement!;
    body.scrollTop = 120;
    body.dispatchEvent(new Event('scroll'));
    await wait(60);
    const before = root.querySelector('table')!;
    const t = newTarget();
    await print(view, t);
    expect(body.scrollTop).toBe(120);
    expect(root.querySelector('table')).toBe(before);
    expect(view.zoom).toBe(1);
  });
});

describe('renderForPrint：上限・空のシート', () => {
  it('maxCells を超える手前の行で打ち切り、truncated を返す', async () => {
    const rows = Array.from({ length: 100 }, (_, i) => [`a${i + 1}`, `b${i + 1}`]);
    const { view } = await setup({ rows });
    const t = newTarget();
    const r = await print(view, t, { maxCells: 50 });
    expect(r.truncated).toBe(true);
    const printed = pagesOf(t).flatMap(rowsOf);
    expect(printed.at(-1)).toBe(25);
    expect(printed).toHaveLength(25);
  });

  it('上限に届かなければ truncated は false', async () => {
    const rows = Array.from({ length: 25 }, (_, i) => [`a${i + 1}`, `b${i + 1}`]);
    const { view } = await setup({ rows });
    const r = await print(view, newTarget(), { maxCells: 50 });
    expect(r.truncated).toBe(false);
  });

  it('空のシート・開けなかったブックは 0 ページで、何も足さない', async () => {
    const { view } = await setup({ rows: [] });
    const t = newTarget();
    const r = await print(view, t);
    expect(r.pages).toBe(0);
    expect(t.children).toHaveLength(0);

    const { view: broken } = await mountWorkbook(new Uint8Array([1, 2, 3]));
    const t2 = newTarget();
    expect((await print(broken, t2)).pages).toBe(0);
    expect(t2.children).toHaveLength(0);
    await expect(broken.renderForPrint(t2, { page: { width: -1, height: 1 } })).rejects.toThrow(
      RangeError,
    );
  });

  it('renderPrint は表の無い表示（null）で 0 ページ', async () => {
    const r = await renderPrint(PRINT_KIT, newTarget(), null);
    expect(r.pages).toBe(0);
  });
});

/** 画像 1 枚（B2〜D6）を持つシート */
function imageSheet(): Omit<SheetSpec, 'name'> {
  return {
    rows: [['a']],
    drawing: {
      xml: wsDr([
        twoCellAnchor(
          { col: 1, row: 1 },
          { col: 3, row: 5 },
          picXml({ name: 'p1', embed: 'rId1' }),
        ),
      ]),
      rels: [{ id: 'rId1', target: '../media/image1.png' }],
    },
  };
}

describe('renderForPrint：図・待つ・中断・片付け', () => {
  it('画像を読み終えてから解決し、destroy で外して blob URL を放す', async () => {
    const { view } = await setup(imageSheet(), { 'xl/media/image1.png': PNG });
    await wait();
    const t = newTarget();
    const r = await print(view, t);
    expect(r.pages).toBe(1);
    const img = t.querySelector<HTMLImageElement>('.xs-print-clip .xs-layer img.xs-img');
    expect(img).not.toBeNull();
    expect(img!.src).toMatch(/^blob:p\//);
    r.destroy();
    expect(t.children).toHaveLength(0);
    expect(urls.revoked).toContain(img!.getAttribute('src'));
    r.destroy(); // 何度呼んでもよい
  });

  it('図の層はセルの範囲で切る入れ物の中（見出しの幅・高さだけずらす）', async () => {
    const { view } = await setup(imageSheet(), { 'xl/media/image1.png': PNG });
    const t = newTarget();
    await print(view, t, { headers: true });
    const clip = t.querySelector<HTMLElement>('.xs-print-clip')!;
    expect(clip.style.left).not.toBe('0px');
    expect(clip.style.top).toBe('24px');
  });

  it('始める前に中断していれば AbortError で、何も足さない', async () => {
    const { view } = await setup({ rows: [['a']] });
    const t = newTarget();
    const ctl = new AbortController();
    ctl.abort();
    await expect(view.renderForPrint(t, { signal: ctl.signal })).rejects.toMatchObject({
      name: 'AbortError',
    });
    expect(t.children).toHaveLength(0);
  });

  /** 画像の読み込みを止めたブック（release で続ける） */
  async function stalledImage() {
    const book = await openBook({
      font: { name: 'Calibri', size: 11 },
      sheets: [{ name: 'S', ...imageSheet() }],
      extraFiles: { 'xl/media/image1.png': PNG },
    });
    let release!: () => void;
    const gate = new Promise<void>((r) => (release = r));
    let calls = 0;
    const wrapped: XlsxBook = {
      sheets: book.sheets,
      openSheet: (i) => book.openSheet(i),
      openDrawings: (i) => book.openDrawings(i),
      readImage: async (i, p) => {
        calls++;
        if (calls > 1) await gate; // 1 回目は表示、2 回目からが印刷
        return book.readImage(i, p);
      },
      readChart: (i, p) => book.readChart(i, p),
      close: () => book.close(),
    };
    const m = await mount(wrapped);
    await wait();
    return { ...m, release };
  }

  it('画像を待つ間に signal で中断すると、足した物を外して AbortError', async () => {
    const { view, release } = await stalledImage();
    const t = newTarget();
    const ctl = new AbortController();
    const p = view.renderForPrint(t, { signal: ctl.signal });
    await wait(10);
    expect(t.querySelector('.xs-print')).not.toBeNull(); // 描いて待っている
    ctl.abort();
    await expect(p).rejects.toMatchObject({ name: 'AbortError' });
    expect(t.children).toHaveLength(0);
    release();
  });

  it('画像を待つ間に表示を destroy・シートを切り替えると AbortError', async () => {
    const a = await stalledImage();
    const t = newTarget();
    const p = a.view.renderForPrint(t);
    await wait(10);
    a.view.destroy();
    await expect(p).rejects.toMatchObject({ name: 'AbortError' });
    expect(t.children).toHaveLength(0);
    a.release();

    const { view } = await mountWorkbook({
      sheets: [
        { name: 'A', ...imageSheet() },
        { name: 'B', rows: [['b']] },
      ],
      extraFiles: { 'xl/media/image1.png': PNG },
    });
    const t2 = newTarget();
    const ctl = new AbortController();
    // 切り替えは描いている途中（待つ間）に来る
    const p2 = view.renderForPrint(t2, { signal: ctl.signal });
    void view.showSheet(2);
    await expect(p2).rejects.toMatchObject({ name: 'AbortError' });
    expect(t2.children).toHaveLength(0);
  });

  it('renderWorkbook の表示でも描ける（表示の destroy の後も印刷の DOM は残り、destroy で外れる）', async () => {
    const { view } = await mountWorkbook({ sheets: [{ name: 'A', rows: [['x', 'y']] }] });
    const t = newTarget();
    const r = await print(view, t);
    expect(r.pages).toBe(1);
    view.destroy();
    expect(textsOf(pagesOf(t)[0]!)).toEqual(['x', 'y']);
    r.destroy();
    expect(t.children).toHaveLength(0);
  });
});

describe('印刷の CSS（style.css の文面。M7-21 で 3 つのブラウザの PDF で確かめた形）', () => {
  it('ページは固定の窓：overflow: hidden・contain: strict（図の層の中身で文書全体を縮めさせない）・後ろで改ページ', () => {
    const page = rule('.xs-print-page');
    expect(page).toContain('overflow: hidden;');
    expect(page).toContain('contain: strict;');
    expect(page).toContain('break-after: page;');
    expect(rule('.xs-print-page:last-child')).toContain('break-after: auto;');
  });

  it('根は色を落とさず（print-color-adjust: exact）、変数の既定値が直に当たる', () => {
    expect(rule('.xs-sheet.xs-print')).toContain('print-color-adjust: exact;');
    expect(css).toContain(':where(:root, :host, .xs-print) {');
  });

  it('縮めた印刷では列記号も行番号も倍率で小さくする（5px より小さくしない）', () => {
    expect(rule('.xs-print .xs-grid thead th,\n.xs-print .xs-grid .xs-rownum .xs-cell')).toContain(
      'font-size: max(5px, calc(13px * min(1, var(--xs-zoom, 1))));',
    );
  });
});
