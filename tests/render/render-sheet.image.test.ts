// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { DrawingNotice, XlsxBook } from '../../src/core/index';
import { MESSAGES_JA } from '../../src/render/messages';
import { createSheetView } from '../../src/render/sheet-view';
import { makePng } from '../../src/testing/make-image';
import {
  custGeomXml,
  cxnSpXml,
  graphicFrameXml,
  grpSpXml,
  lineShapeXml,
  oneCellAnchor,
  picXml,
  shapeXml,
  spXml,
  twoCellAnchor,
  wsDr,
  type SheetSpec,
} from '../../src/testing/make-xlsx';
import {
  cleanupViews,
  mount,
  openBook,
  scrollBody,
  stubClientSize,
  wait,
} from './render-sheet.helpers';

/**
 * `renderSheet` の画像レイヤー（ISSUE-86・ADR-0026 §決定 8・9）の DOM 描画テスト（InterHub `spreadsheet.image.render.test.ts` を写したもの）。
 * jsdom は画像をデコードせずレイアウトも持たない（clientWidth／Height はテストで固定する）ため、
 * 実寸・描画（`naturalWidth`・スクロール揺れ）は目視（TV-38）と実測（Chrome）で確認し、ここでは
 * レイヤーの構造・位置（style）・出し入れ・blob URL の revoke・通知の行を検証する。文言は既定（`MESSAGES_JA`）。
 * InterHub の `vi.mock('../xlsx')` は、図の通知・画像の結果を差し替えたブック（{@link wrapBook}）を renderSheet に渡す形にした。
 * タブの切替を含むテスト（TD-86-11 の 2 つ目・TD-86-12 の 1 つ目・TD-86-13 の over-budget の後半・TD-88-20 の 2 つ目の前半）は M2-4 で写す。
 */

/** ブックを差し替える（図の通知・画像の失敗を作る試験用）。null なら素通し */
const hooks: {
  notices: DrawingNotice[] | null;
  readImage: XlsxBook['readImage'] | null;
} = { notices: null, readImage: null };

/** 読み込み中の画像の数（{@link settle} が 0 になるまで待つ） */
let inflight = 0;

/**
 * 画像の読み込み・rAF が落ち着くまで待つ。並行して走るほかのテストで遅れることがあるため、
 * 一定の時間（InterHub の wait）だけでなく、読み込み中の画像が無くなるまで待つ。
 */
async function settle(ms = 30): Promise<void> {
  await wait(ms);
  await vi.waitFor(() => expect(inflight).toBe(0));
  await wait(0);
}

/** hooks に合わせて openDrawings・readImage を差し替えたブック */
function wrapBook(book: XlsxBook): XlsxBook {
  const notices = hooks.notices;
  const readImage = hooks.readImage;
  return {
    sheets: book.sheets,
    openSheet: (i) => book.openSheet(i),
    openDrawings: async (i) => {
      const r = await book.openDrawings(i);
      return notices ? { ...r, notices } : r;
    },
    readImage: (i, p) => {
      inflight++;
      return (readImage ?? ((j: number, q: string) => book.readImage(j, q)))(i, p).finally(() => {
        inflight--;
      });
    },
    close: () => book.close(),
  };
}

const PNG = makePng(40, 30);
const A1 = { col: 0, row: 0 };

/** 画像（image1.png）を指す drawing を持つシート */
function sheetWith(name: string, anchors: string[], rows: SheetSpec['rows'] = [['a']]): SheetSpec {
  return {
    name,
    rows,
    drawing: { xml: wsDr(anchors), rels: [{ id: 'rId1', target: '../media/image1.png' }] },
  };
}

/** ブックを作って index のシートを描く */
async function render(sheets: SheetSpec[], index = 1): Promise<HTMLElement> {
  const book = await openBook({
    font: { name: 'Calibri', size: 11 },
    sheets,
    extraFiles: { 'xl/media/image1.png': PNG },
  });
  return (await mount(wrapBook(book), index)).root;
}

/** blob URL の作成・解放の記録 */
const urls = { created: [] as string[], revoked: [] as string[] };
const savedUrl = { create: URL.createObjectURL, revoke: URL.revokeObjectURL };
let restoreSize: () => void;

beforeEach(() => {
  hooks.notices = null;
  hooks.readImage = null;
  inflight = 0;
  urls.created = [];
  urls.revoked = [];
  let n = 0;
  URL.createObjectURL = vi.fn(() => {
    const u = `blob:r/${++n}`;
    urls.created.push(u);
    return u;
  });
  URL.revokeObjectURL = vi.fn((u: string) => {
    urls.revoked.push(u);
  });
  // スクロール枠の大きさ（jsdom はレイアウトを持たない）
  restoreSize = stubClientSize(800, 400);
});

afterEach(() => {
  cleanupViews();
  restoreSize();
  URL.createObjectURL = savedUrl.create;
  URL.revokeObjectURL = savedUrl.revoke;
});

/** レイヤーの図（DOM の順） */
const figures = (root: HTMLElement): HTMLElement[] => [
  ...root.querySelectorAll<HTMLElement>('.xs-layer .xs-figure'),
];
const px = (v: string): number => Number.parseFloat(v);
const rectOf = (el: HTMLElement) => [el.style.left, el.style.top, el.style.width, el.style.height];
const noticeOf = (root: HTMLElement): HTMLElement => root.querySelector<HTMLElement>('.xs-notice')!;

describe('renderSheet（.xlsx・画像レイヤー／ISSUE-86）', () => {
  it('TD-86-8: 画像（<img>）・種類名入りの破線枠が、アンカーの位置（geometry）に載る。線・矢印は SVG（svg.xs-line）で出る', async () => {
    // 既定の列幅 64px・行高 20px。col・row は 0 始まり
    const root = await render([
      sheetWith('S', [
        twoCellAnchor(
          { col: 1, row: 1 },
          { col: 3, colOff: 9525 * 10, row: 3, rowOff: 9525 * 5 },
          picXml({ name: 'p1', descr: '画面', embed: 'rId1' }),
        ),
        twoCellAnchor({ col: 5, row: 2 }, { col: 7, row: 6 }, graphicFrameXml('c1')),
        oneCellAnchor(
          { col: 0, colOff: 9525 * 4, row: 8 },
          9525 * 100,
          9525 * 50,
          spXml('t1', true),
        ),
        twoCellAnchor({ col: 0, row: 12 }, { col: 2, row: 14 }, grpSpXml('g1', [spXml('in')])),
        twoCellAnchor({ col: 0, row: 20 }, { col: 4, row: 21 }, cxnSpXml('line1')),
      ]),
    ]);
    await settle();
    const figs = figures(root);
    expect(figs).toHaveLength(5); // 線（line1）も出る（ISSUE-88）
    // 画像：B2（64,20）〜 D 列左端(192)+10 / 行 4 上端(60)+5
    expect(rectOf(figs[0]!)).toEqual(['64px', '20px', `${192 + 10 - 64}px`, `${60 + 5 - 20}px`]);
    const img = figs[0]!.querySelector('img')!;
    expect(img.alt).toBe('画面');
    expect(img.getAttribute('src')).toBe('blob:r/1');
    expect(img.decoding).toBe('async');
    // グラフ：F3〜H7
    expect(rectOf(figs[1]!)).toEqual(['320px', '40px', `${2 * 64}px`, `${4 * 20}px`]);
    expect(figs[1]!.querySelector('.xs-frame')?.textContent).toBe(MESSAGES_JA.chart);
    // テキストボックス：oneCell A9＋4px・100×50。形（矩形）で描く（ISSUE-92）ので、外接矩形は枠線 1px の miter の張り出し（2px）だけ広がり、種類名の破線枠は出ない
    expect(rectOf(figs[2]!)).toEqual(['2px', '158px', '104px', '54px']);
    expect(figs[2]!.querySelector('svg.xs-shape-svg path')).not.toBeNull();
    expect(figs[2]!.querySelector('.xs-frame')).toBeNull();
    // グループ（座標系の無い grpSpPr は展開できず、今までどおり破線枠。展開は TD-89-10）
    expect(figs[3]!.querySelector('.xs-frame')?.textContent).toBe(MESSAGES_JA.group);
    // 線・矢印：A21〜E22 の外接矩形に載る SVG（種類名の破線枠ではない）
    const line = figs[4]!;
    expect(line.matches('svg.xs-line')).toBe(true);
    expect(line.querySelector('path')).not.toBeNull();
    expect(line.querySelector('.xs-frame')).toBeNull();
    // 外接矩形は線の太さの分だけアンカーより外へ広がる（数 px）
    expect(Math.abs(px(line.style.left) - 0)).toBeLessThanOrEqual(3);
    expect(Math.abs(px(line.style.top) - 400)).toBeLessThanOrEqual(3);
    expect(root.querySelectorAll('svg.xs-line')).toHaveLength(1);
  });

  it('TD-89-10: グループは中の子を展開して描く（文字の枠・SVG の線・<img>）。座標系の無いグループ・深さ 9 のグループは「グループ」の破線枠のまま', async () => {
    // 深さ n のグループ（外側の座標系は 1 単位 = 1 EMU の恒等）。いちばん内側は空
    let deep = shapeXml({ name: 'leaf', box: [0, 0, 10, 10], paras: ['x'] });
    for (let d = 9; d >= 2; d--) {
      deep = grpSpXml(`d${d}`, [deep], 20 + d, { box: [0, 0, 100, 100] });
    }
    const root = await render([
      sheetWith('S', [
        // A13〜E17（256×80 px）に、2560×800 の座標系の 0.1 倍で置く
        twoCellAnchor(
          { col: 0, row: 12 },
          { col: 4, row: 16 },
          grpSpXml(
            'g',
            [
              shapeXml({ name: 'tb', txBox: true, box: [256, 100, 640, 200], paras: ['中の文字'] }),
              lineShapeXml({ name: 'ln', box: [0, 400, 2560, 0] }),
              picXml({ name: 'pic', embed: 'rId1', box: [1280, 300, 640, 400] }),
              deep,
            ],
            6,
            { box: [0, 0, 2560, 800] },
          ),
        ),
        // 座標系の無いグループ（grpSpPr が空）は展開できない
        twoCellAnchor({ col: 0, row: 30 }, { col: 2, row: 32 }, grpSpXml('plain', [spXml('in')])),
      ]),
    ]);
    await settle();
    const figs = figures(root);
    // 子（文字の枠・線・画像）＋ 深さ 9 の破線枠 ＋ 座標系の無いグループの破線枠
    expect(figs).toHaveLength(5);
    // 文字の枠（形の文字領域。ISSUE-92）：(256,100)×0.1 = (26, 10) → A13 の上端 240 を足す。外接矩形は枠線の分だけ広がる
    expect(figs[0]!.querySelector('.xs-shape-text')).not.toBeNull();
    expect(rectOf(figs[0]!)).toEqual(['24px', '248px', '68px', '24px']);
    expect(figs[0]!.textContent).toBe('中の文字');
    // 線：SVG
    expect(figs[1]!.matches('svg.xs-line')).toBe(true);
    // 画像：<img>（blob URL は 1 つ）
    expect(rectOf(figs[2]!)).toEqual(['128px', '270px', '64px', '40px']);
    expect(figs[2]!.querySelector('img')?.getAttribute('src')).toBe('blob:r/1');
    expect(urls.created).toHaveLength(1);
    // 破線枠は「グループ」だけ（展開したグループ自身の枠は出ない）
    const frames = [...root.querySelectorAll('.xs-layer .xs-frame')].map((f) => f.textContent);
    expect(frames).toEqual([MESSAGES_JA.group, MESSAGES_JA.group]);
  });

  it('TD-86-8: 構造は スクロール枠 ＞ wrapper ＞ [表・罫線層・画像レイヤー]。レイヤーの原点は A1 の左上（行番号の列の幅・見出し行 24px）', async () => {
    const root = await render([
      sheetWith('S', [twoCellAnchor(A1, { col: 1, row: 1 }, picXml({ name: 'p', embed: 'rId1' }))]),
    ]);
    const body = root.querySelector<HTMLElement>('.xs-body')!;
    const wrap = body.firstElementChild as HTMLElement;
    expect(wrap.className).toBe('xs-wrap');
    // 罫線層（ISSUE-98）は表の直後・画像レイヤーの直前（svg の className は文字列でないため属性で読む）
    expect([...wrap.children].map((c) => c.tagName + '.' + c.getAttribute('class'))).toEqual([
      'TABLE.xs-grid',
      'svg.xs-borders',
      'DIV.xs-layer',
    ]);
    const layer = wrap.querySelector<HTMLElement>('.xs-layer')!;
    const rowCount = Number(wrap.querySelector('table')!.getAttribute('aria-rowcount')) - 1;
    const numWidth = Math.max(32, String(rowCount).length * 8 + 16);
    expect(layer.style.left).toBe(`${numWidth}px`);
    expect(layer.style.top).toBe('24px');
  });

  it('TD-86-9: 値が無く図だけのシートも、空シート扱いにせずグリッドを出し、図のある範囲まで表を広げる', async () => {
    const root = await render([
      // セルの値なし・G 列（col 6）の先まで届く図
      sheetWith(
        'S',
        [
          twoCellAnchor(
            { col: 1, row: 1 },
            { col: 9, row: 3 },
            picXml({ name: 'p', embed: 'rId1' }),
          ),
        ],
        [],
      ),
    ]);
    expect(root.querySelector('.xs-empty')).toBeNull();
    expect(root.querySelector('table.xs-grid')).not.toBeNull();
    // 列は少なくとも J 列（図の最右列＝col 9 の先頭 → 10 列目）まで
    expect(root.querySelectorAll('colgroup > col').length).toBeGreaterThanOrEqual(1 + 10);
    expect(figures(root)).toHaveLength(1);
  });

  it('TD-86-9: 値も図も無いシートは従来どおり空シート表示', async () => {
    const root = await render([{ name: 'S', rows: [] }]);
    expect(root.querySelector('.xs-empty')).not.toBeNull();
    expect(root.querySelector('.xs-grid')).toBeNull();
  });

  it('TD-86-9: 値が無く、置ける図が無い（大きさ 0 の図だけ）シートも空シート表示（置けた図の数で判定する）', async () => {
    const root = await render([sheetWith('S', [oneCellAnchor(A1, 0, 0, spXml('zero'))], [])]);
    expect(root.querySelector('.xs-empty')).not.toBeNull();
    expect(root.querySelector('.xs-grid')).toBeNull();
  });

  it('TD-86-10: 可視範囲の外の図は DOM に出さず、スクロールで出し入れする（図は 1 画面分の余白まで先読み）', async () => {
    const anchors = [0, 100, 600].map((row) =>
      oneCellAnchor(
        { col: 0, row },
        9525 * 40,
        9525 * 30,
        picXml({ name: `p${row}`, embed: 'rId1' }),
      ),
    );
    const root = await render([sheetWith('S', anchors)]);
    await settle();
    expect(figures(root).map((f) => f.style.top)).toEqual(['0px']); // 行 100（y=2000）は遠い
    scrollBody(root, { top: 2000 });
    await settle(60);
    expect(figures(root).map((f) => f.style.top)).toEqual(['2000px']);
    scrollBody(root, { top: 12000 });
    await settle(60);
    expect(figures(root).map((f) => f.style.top)).toEqual(['12000px']);
  });

  it('TD-86-11: 同じ画像を複数のアンカーで使っても blob URL は 1 つ。表示を壊す（destroy）とすべて revoke', async () => {
    const anchors = [0, 2, 4].map((row) =>
      oneCellAnchor(
        { col: 0, row },
        9525 * 40,
        9525 * 30,
        picXml({ name: `p${row}`, embed: 'rId1' }),
      ),
    );
    const root = await render([sheetWith('S', anchors)]);
    await settle();
    expect(urls.created).toHaveLength(1);
    expect(root.querySelectorAll('.xs-layer img')).toHaveLength(3);
    cleanupViews();
    expect(urls.revoked).toEqual(urls.created);
  });

  it('TD-86-12: 画像を待っている間に表示を壊して次の表示を作っても、URL を作って残さない', async () => {
    const pic = oneCellAnchor(
      { col: 0, row: 0 },
      9525 * 40,
      9525 * 30,
      picXml({ name: 'a', embed: 'rId1' }),
    );
    await render([sheetWith('S', [pic])]);
    // 画像の読み込みが終わる前に壊して、次のファイルを開く（InterHub は次の表示が前の表示を後始末した）
    cleanupViews();
    const next = await render([{ name: 'T', rows: [['x']] }]);
    await settle(60);
    expect(urls.revoked.slice().sort()).toEqual(urls.created.slice().sort());
    expect(next.querySelector('img')).toBeNull();
  });

  it('TD-86-12: シートを開いている途中で destroy しても、URL を作って残さない（図の層も空）', async () => {
    const pic = oneCellAnchor(
      { col: 0, row: 0 },
      9525 * 40,
      9525 * 30,
      picXml({ name: 'a', embed: 'rId1' }),
    );
    const book = await openBook({
      font: { name: 'Calibri', size: 11 },
      sheets: [sheetWith('S', [pic])],
      extraFiles: { 'xl/media/image1.png': PNG },
    });
    const container = document.createElement('div');
    const view = createSheetView(container, book, { messages: MESSAGES_JA, drawings: true });
    const pending = view.showSheet(1);
    await Promise.resolve();
    view.destroy();
    await pending;
    await settle(60);
    expect(urls.created).toEqual([]);
    expect(view.root.querySelector('.xs-figure')).toBeNull();
    book.close();
  });

  it('TD-86-13: 図が 1,000 を超えると先頭 1,000 個だけ（reader の上限）と通知 1 行。表は出る。DOM の図は可視範囲のぶんだけ', async () => {
    const anchors: string[] = [];
    for (let i = 0; i < 1100; i++) {
      anchors.push(oneCellAnchor({ col: 0, row: i }, 9525 * 40, 9525 * 18, spXml(`s${i}`)));
    }
    const root = await render([sheetWith('S', anchors)]);
    await settle();
    const bar = noticeOf(root);
    expect(bar.hidden).toBe(false);
    expect(bar.textContent).toBe(MESSAGES_JA.noticeTooMany);
    expect(root.querySelector('table.xs-grid')).not.toBeNull();
    // 行 1,000 の位置まで表が広がる（先頭 1,000 個の最下行）
    const rowCount = Number(root.querySelector('table')!.getAttribute('aria-rowcount')) - 1;
    expect(rowCount).toBe(1000);
    expect(figures(root).length).toBeLessThan(100);
    expect(figures(root).length).toBeGreaterThan(0);
  });

  it('TD-86-13: 通知（図が大きすぎる・読み取れない）：該当の文言を 1 行で出し、表は出る。通知が無ければ hidden', async () => {
    hooks.notices = ['part-too-large', 'broken'];
    const root = await render([
      sheetWith('S', [twoCellAnchor(A1, { col: 1, row: 1 }, spXml('s'))]),
    ]);
    const bar = noticeOf(root);
    expect(bar.hidden).toBe(false);
    expect(bar.textContent).toBe(`${MESSAGES_JA.noticeTooLarge} ${MESSAGES_JA.noticeBroken}`);
    expect(root.querySelector('table.xs-grid')).not.toBeNull();

    hooks.notices = null;
    const plain = await render([
      sheetWith('S', [twoCellAnchor(A1, { col: 1, row: 1 }, spXml('s'))]),
    ]);
    expect(noticeOf(plain).hidden).toBe(true);
  });

  it('TD-86-13: 画像の合計の上限超過（over-budget）は通知を出し、その画像は破線枠', async () => {
    hooks.readImage = async () => ({ ok: false, reason: 'over-budget', label: null });
    const root = await render([
      sheetWith('A', [
        oneCellAnchor(A1, 9525 * 40, 9525 * 30, picXml({ name: 'a', embed: 'rId1' })),
      ]),
      { name: 'B', rows: [['x']] },
    ]);
    await settle();
    const bar = noticeOf(root);
    expect(bar.hidden).toBe(false);
    expect(bar.textContent).toBe(MESSAGES_JA.noticeOverBudget);
    expect(figures(root)[0]!.querySelector('.xs-frame')?.textContent).toBe(MESSAGES_JA.imageBroken);
  });

  it('TD-86-14: 外部リンクの画像（r:link のみ）・描けない形式は、表示できない枠になる（外部取得・img は作らない）', async () => {
    const root = await render([
      sheetWith('S', [
        twoCellAnchor(A1, { col: 2, row: 2 }, picXml({ name: 'ext', link: 'rId9' })),
        oneCellAnchor({ col: 4, row: 4 }, 9525 * 40, 9525 * 30, picXml({ name: 'noembed' })),
      ]),
    ]);
    await settle();
    expect(root.querySelector('img')).toBeNull();
    expect(urls.created).toHaveLength(0);
    expect([...root.querySelectorAll('.xs-frame')].map((f) => f.textContent)).toEqual([
      MESSAGES_JA.imageBroken,
      MESSAGES_JA.imageBroken,
    ]);
  });

  it('drawings: false なら図を読まず（openDrawings を呼ばない）、図も通知も出さない。表と層の構造は同じ', async () => {
    const book = await openBook({
      font: { name: 'Calibri', size: 11 },
      sheets: [
        sheetWith('S', [
          oneCellAnchor(A1, 9525 * 40, 9525 * 30, picXml({ name: 'a', embed: 'rId1' })),
        ]),
      ],
      extraFiles: { 'xl/media/image1.png': PNG },
    });
    const spy = vi.spyOn(book, 'openDrawings');
    const { root } = await mount(book, 1, { drawings: false });
    await settle();
    expect(spy).not.toHaveBeenCalled();
    expect(figures(root)).toHaveLength(0);
    expect(urls.created).toHaveLength(0);
    expect(noticeOf(root).hidden).toBe(true);
    expect(root.querySelector('.xs-wrap > .xs-layer')).not.toBeNull();
    expect(root.querySelector('table.xs-grid')).not.toBeNull();
  });
});

describe('renderSheet（.xlsx・線と図形の文字／ISSUE-88）', () => {
  /** 文字のある図形（段落 1 つ） */
  const textShape = (name: string, text: string): string => shapeXml({ name, paras: [text] });

  it('TD-88-17: 文字のある図形は種類名でなく文字を出し、線は SVG で出す（種類名の破線枠は文字の無い図形だけ）', async () => {
    const root = await render([
      sheetWith('S', [
        twoCellAnchor({ col: 1, row: 1 }, { col: 4, row: 4 }, textShape('t', '処理A')),
        twoCellAnchor({ col: 1, row: 6 }, { col: 4, row: 8 }, spXml('plain')),
        twoCellAnchor(
          { col: 6, row: 1 },
          { col: 9, row: 4 },
          lineShapeXml({
            name: 'l',
            ln: '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
          }),
        ),
      ]),
    ]);
    await settle();
    const figs = figures(root);
    expect(figs).toHaveLength(3);
    // 形で描く（ISSUE-92）：文字は文字領域（.xs-shape-text）に入り、文字の無い図形も形（svg）で描く（種類名の破線枠は代替だけ）
    expect(figs[0]!.querySelector('.xs-shape-text')).not.toBeNull();
    expect(figs[0]!.textContent).toBe('処理A');
    expect(figs[0]!.querySelector('.xs-frame')).toBeNull();
    expect(figs[1]!.querySelector('svg.xs-shape-svg path')).not.toBeNull();
    expect(figs[1]!.querySelector('.xs-frame')).toBeNull();
    expect(figs[2]!.matches('svg.xs-line')).toBe(true);
    expect(figs[2]!.querySelector('path')?.getAttribute('d')).toBeTruthy();
  });

  it('TD-88-18: 水平の線（アンカーの高さ 0）も SVG で出る。置けた図として数え、値の無いシートも空シートにしない', async () => {
    const root = await render([
      sheetWith(
        'S',
        [
          oneCellAnchor(
            { col: 1, row: 2 },
            9525 * 100,
            0,
            lineShapeXml({ name: 'h', ln: '<a:ln w="12700"/>' }),
          ),
        ],
        [],
      ),
    ]);
    await settle();
    const lines = root.querySelectorAll<SVGElement>('svg.xs-line');
    expect(lines).toHaveLength(1);
    // 高さ 0 でも落とさない。外接矩形は線の太さの分だけ広がる（数 px）
    expect(Math.abs(px(lines[0]!.style.left) - 64)).toBeLessThanOrEqual(3);
    expect(Math.abs(px(lines[0]!.style.width) - 100)).toBeLessThanOrEqual(6);
    expect(Math.abs(px(lines[0]!.style.top) - 40)).toBeLessThanOrEqual(3);
    expect(px(lines[0]!.style.height)).toBeGreaterThan(0);
    expect(root.querySelector('.xs-empty')).toBeNull();
    expect(root.querySelector('table.xs-grid')).not.toBeNull();
  });

  it('TD-88-19: 垂直の線（アンカーの幅 0）も SVG で出る', async () => {
    const root = await render([
      sheetWith('S', [
        oneCellAnchor(
          { col: 2, row: 1 },
          0,
          9525 * 60,
          lineShapeXml({ name: 'v', ln: '<a:ln w="12700"/>' }),
        ),
      ]),
    ]);
    await settle();
    const lines = root.querySelectorAll<SVGElement>('svg.xs-line');
    expect(lines).toHaveLength(1);
    expect(Math.abs(px(lines[0]!.style.height) - 60)).toBeLessThanOrEqual(6);
    expect(px(lines[0]!.style.width)).toBeGreaterThan(0);
  });

  it('TD-88-20: text-truncated の通知：複数の図形が切り詰められても 1 回・1 行。表も図も出る', async () => {
    const big = 'あ'.repeat(10001); // 図形 1 つの上限（10,000 文字）を超える
    const root = await render([
      sheetWith('S', [
        twoCellAnchor({ col: 0, row: 0 }, { col: 2, row: 2 }, textShape('a', big)),
        twoCellAnchor({ col: 0, row: 4 }, { col: 2, row: 6 }, textShape('b', big)),
        twoCellAnchor({ col: 0, row: 8 }, { col: 2, row: 10 }, textShape('c', big)),
      ]),
    ]);
    await settle();
    const bar = noticeOf(root);
    expect(bar.hidden).toBe(false);
    expect(bar.textContent).toBe(MESSAGES_JA.noticeTextTruncated);
    expect(root.querySelector('table.xs-grid')).not.toBeNull();
    expect(figures(root)).toHaveLength(3);
  });

  it('TD-88-20: text-truncated の通知はシートごと（通知の無いシートでは出ない）。too-many と並べて 1 行にする', async () => {
    const big = 'あ'.repeat(10001);
    const sheets: SheetSpec[] = [
      sheetWith('A', [twoCellAnchor({ col: 0, row: 0 }, { col: 2, row: 2 }, textShape('a', big))]),
      { name: 'B', rows: [['x']] },
    ];
    expect(noticeOf(await render(sheets, 1)).textContent).toBe(MESSAGES_JA.noticeTextTruncated);
    expect(noticeOf(await render(sheets, 2)).hidden).toBe(true);

    hooks.notices = ['too-many', 'text-truncated'];
    const both = await render([
      sheetWith('S', [twoCellAnchor(A1, { col: 1, row: 1 }, spXml('s'))]),
    ]);
    expect(noticeOf(both).textContent).toBe(
      `${MESSAGES_JA.noticeTooMany} ${MESSAGES_JA.noticeTextTruncated}`,
    );
  });

  it('TD-88-21: U 字線の張り出し（外接矩形）の分だけ表の列が広がる（drawingExtent に線が入る）', async () => {
    const uLine = (adj1: number): string =>
      lineShapeXml({
        name: 'u',
        prst: 'bentConnector3',
        avLst: `<a:gd name="adj1" fmla="val ${adj1}"/>`,
        ln: '<a:ln w="12700"/>',
      });
    const cols = async (adj1: number): Promise<number> => {
      // 幅 20px・高さ 60px のアンカーを U 列（col 20＝x 1280）に置く。画面幅（800px）より右なので範囲は図で決まる
      const root = await render([
        sheetWith('S', [oneCellAnchor({ col: 20, row: 1 }, 9525 * 20, 9525 * 60, uLine(adj1))], []),
      ]);
      await settle();
      return root.querySelectorAll('colgroup > col').length;
    };
    const plain = await cols(50000);
    const wide = await cols(500000); // 右へ幅の 5 倍（100px）張り出す
    expect(wide).toBeGreaterThan(plain);
  });
});

describe('renderSheet（.xlsx・図形の形／ISSUE-92）', () => {
  /** 自由図形（moveTo ＋ lnTo を `n` 個）。命令数が 1 図形の上限（4,096）を超えるかを `n` で決める */
  const polyXml = (name: string, n: number): string =>
    shapeXml({
      name,
      custGeom: custGeomXml({
        pathLst: `<a:path w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>${'<a:lnTo><a:pt x="5" y="5"/></a:lnTo>'.repeat(n)}</a:path>`,
      }),
    });

  it('TD-92-12: shape-simplified の通知：1 行・1 回。ほかの通知と並べても 1 行', async () => {
    hooks.notices = ['shape-simplified'];
    const one = await render([sheetWith('S', [twoCellAnchor(A1, { col: 1, row: 1 }, spXml('s'))])]);
    const bar = noticeOf(one);
    expect(bar.hidden).toBe(false);
    expect(bar.textContent).toBe(MESSAGES_JA.noticeShapeSimplified);
    expect(bar.querySelectorAll('div, br')).toHaveLength(0);
    hooks.notices = ['too-many', 'text-truncated', 'shape-simplified'];
    const all = await render([sheetWith('S', [twoCellAnchor(A1, { col: 1, row: 1 }, spXml('s'))])]);
    expect(noticeOf(all).textContent).toBe(
      `${MESSAGES_JA.noticeTooMany} ${MESSAGES_JA.noticeTextTruncated} ${MESSAGES_JA.noticeShapeSimplified}`,
    );
  });

  it('TD-92-12: 上限を超えた自由図形は代替（種類名の破線枠）で、通知は複数あっても 1 回。未知の prst は知らせない', async () => {
    const root = await render([
      sheetWith('S', [
        twoCellAnchor({ col: 0, row: 0 }, { col: 2, row: 2 }, polyXml('a', 4097)),
        twoCellAnchor({ col: 0, row: 4 }, { col: 2, row: 6 }, polyXml('b', 4097)),
        twoCellAnchor({ col: 4, row: 0 }, { col: 6, row: 2 }, polyXml('ok', 10)),
      ]),
    ]);
    await settle();
    const bar = noticeOf(root);
    expect(bar.hidden).toBe(false);
    expect(bar.textContent).toBe(MESSAGES_JA.noticeShapeSimplified);
    const figs = figures(root);
    expect(figs).toHaveLength(3);
    expect(figs[0]!.querySelector('.xs-frame')?.textContent).toBe(MESSAGES_JA.shape);
    expect(figs[1]!.querySelector('.xs-frame')?.textContent).toBe(MESSAGES_JA.shape);
    expect(figs[2]!.querySelector('svg.xs-shape-svg path')).not.toBeNull();

    const unknown = await render([
      sheetWith('S', [
        twoCellAnchor(
          { col: 0, row: 0 },
          { col: 2, row: 2 },
          shapeXml({ name: 'u', prst: 'noSuchShape' }),
        ),
      ]),
    ]);
    await settle();
    expect(noticeOf(unknown).hidden).toBe(true);
    expect(figures(unknown)[0]!.querySelector('.xs-frame')).not.toBeNull();
  });

  it('TD-92-13: 形で描く（ブロック矢印・円柱）。見えない図形（枠線なし・塗りなし・文字なし）は DOM に出ない。回転した図形は box を回す', async () => {
    const root = await render([
      sheetWith('S', [
        twoCellAnchor(
          { col: 0, row: 0 },
          { col: 3, row: 3 },
          shapeXml({ name: 'arrow', prst: 'rightArrow' }),
        ),
        twoCellAnchor(
          { col: 4, row: 0 },
          { col: 6, row: 4 },
          shapeXml({ name: 'can', prst: 'can' }),
        ),
        twoCellAnchor(
          { col: 0, row: 6 },
          { col: 3, row: 8 },
          shapeXml({ name: 'none', ln: '<a:ln><a:noFill/></a:ln>', fill: '<a:noFill/>' }),
        ),
        twoCellAnchor(
          { col: 0, row: 10 },
          { col: 3, row: 12 },
          shapeXml({ name: 'rot', xfrm: 'rot="2700000"', paras: ['回転'] }),
        ),
      ]),
    ]);
    await settle();
    const figs = figures(root);
    expect(figs).toHaveLength(3); // 見えない図形は出ない
    expect(figs[0]!.querySelectorAll('svg path')).toHaveLength(1);
    expect(figs[1]!.querySelectorAll('svg path').length).toBeGreaterThanOrEqual(2); // 円柱は path が複数
    expect(figs[1]!.querySelector('.is-lighten')).not.toBeNull(); // 上面の濃淡
    expect(root.querySelectorAll('.xs-frame')).toHaveLength(0);
    const box = figs[2]!.querySelector<HTMLElement>('.xs-shape-box')!;
    expect(box.style.transform).toBe('rotate(45deg)');
    expect(box.querySelector('.xs-shape-text')!.textContent).toBe('回転');
  });
});
