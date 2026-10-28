// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ImageResult, SheetDrawing } from '../../src/core/index';
import { layoutDrawings, viewWindow, EMU_PER_PX } from '../../src/render/drawings';
import { createImageLayer, frameLabel, type ImageLayer } from '../../src/render/image-layer';
import { createGeometry } from '../../src/core/index';
import {
  emptyLayout,
  type DrawingLine,
  type DrawingText,
  type DrawingTextParagraph,
  type DrawingTextRun,
} from '../../src/core/types';
import { MESSAGES_JA } from '../../src/render/messages';

/**
 * 画像レイヤー（DOM）の振る舞いテスト（ISSUE-86・ADR-0026 §決定 8・9）。
 * jsdom は画像をデコードしない・レイアウトを持たないため、`<img>` の読み込み・実寸は目視（TV-38）で確認し、
 * ここでは DOM の構造（位置・出し入れ・破線枠・差し替え）と blob URL の revoke を検証する。
 * 破線枠の文字は既定の文言（日本語の `MESSAGES_JA`。`{format}` は差し込む）。
 */

const px = (n: number): number => n * EMU_PER_PX;
const geo = createGeometry(emptyLayout(), 1000, 100);

/** oneCell の図（A1 から cx×cy px。col・row は 0 始まり） */
function pic(
  col: number,
  row: number,
  w = 40,
  h = 30,
  over: Partial<SheetDrawing> = {},
): SheetDrawing {
  return {
    anchor: { type: 'oneCell', from: { col, colOff: 0, row, rowOff: 0 }, cx: px(w), cy: px(h) },
    kind: 'image',
    name: `pic${col}_${row}`,
    descr: '',
    mediaPath: 'xl/media/image1.png',
    ...over,
  };
}

const okResult = (): ImageResult => ({
  ok: true,
  format: 'png',
  mime: 'image/png',
  width: 40,
  height: 30,
  bytes: new Uint8Array([1]),
});

/** 画面の全体が見える可視範囲 */
const ALL = viewWindow(0, 0, 3000, 3000, 0, 0);

/** blob URL の作成・解放を数える依存を持つレイヤー */
function makeLayer() {
  let n = 0;
  const created: string[] = [];
  const revoked: string[] = [];
  const layer = createImageLayer({
    createUrl: () => {
      const u = `blob:l/${++n}`;
      created.push(u);
      return u;
    },
    revokeUrl: (u) => revoked.push(u),
  });
  document.body.append(layer.element);
  return { layer, created, revoked };
}

/** Promise のチェーンを流す */
const flush = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

/** シートを渡す */
function setSheet(
  layer: ImageLayer,
  drawings: SheetDrawing[],
  readImage: (p: string) => Promise<ImageResult> = async () => okResult(),
  onOverBudget?: () => void,
) {
  layer.setSheet({ drawings, boxes: layoutDrawings(geo, drawings), readImage, onOverBudget });
}

afterEach(() => {
  document.body.replaceChildren();
});

describe('TD-86-5 画像レイヤー：画像・破線枠・差し替え', () => {
  it('画像は <img>（decoding=async・alt は descr、無ければ name・src は blob URL）。位置と大きさは px で固定', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [
      pic(1, 2, 40, 30, { name: 'n1', descr: '説明' }),
      pic(0, 0, 10, 10, { name: 'only-name' }),
    ]);
    layer.place(32, 24, 6400, 20000);
    layer.update(ALL);
    await flush();
    const figs = [...layer.element.querySelectorAll<HTMLElement>('.xs-figure')];
    expect(figs).toHaveLength(2);
    expect([
      figs[0]!.style.left,
      figs[0]!.style.top,
      figs[0]!.style.width,
      figs[0]!.style.height,
    ]).toEqual(['64px', '40px', '40px', '30px']);
    const img0 = figs[0]!.querySelector('img')!;
    expect(img0.decoding).toBe('async');
    expect(img0.alt).toBe('説明');
    expect(img0.getAttribute('src')).toBe('blob:l/1');
    expect(figs[1]!.querySelector('img')!.alt).toBe('only-name');
    // レイヤーの原点（A1 の左上）と大きさ
    expect([layer.element.style.left, layer.element.style.top]).toEqual(['32px', '24px']);
    expect(layer.element.style.width).toBe('6400px');
  });

  it('同じ画像を複数のアンカーで使っても blob URL は 1 つ（共有）', async () => {
    const { layer, created } = makeLayer();
    const read = vi.fn(async () => okResult());
    setSheet(layer, [pic(0, 0), pic(3, 3), pic(6, 6)], read);
    layer.update(ALL);
    await flush();
    expect(read).toHaveBeenCalledTimes(1);
    expect(created).toHaveLength(1);
    const srcs = [...layer.element.querySelectorAll('img')].map((i) => i.getAttribute('src'));
    expect(srcs).toEqual(['blob:l/1', 'blob:l/1', 'blob:l/1']);
  });

  it('画像以外は種類名入りの破線枠（textContent・HTML は解釈しない）。画像の読み込みは行わない', async () => {
    const { layer } = makeLayer();
    const read = vi.fn(async () => okResult());
    const kinds = ['chart', 'smartart', 'shape', 'textbox', 'group'] as const;
    setSheet(
      layer,
      kinds.map((kind, i) => pic(0, i * 3, 40, 30, { kind, mediaPath: null, name: '<b>x</b>' })),
      read,
    );
    layer.update(ALL);
    await flush();
    const frames = [...layer.element.querySelectorAll('.xs-frame')];
    expect(frames.map((f) => f.textContent)).toEqual(kinds.map((k) => MESSAGES_JA[k]));
    expect(layer.element.querySelector('img, b')).toBeNull();
    expect(read).not.toHaveBeenCalled();
  });

  it('描けない画像：形式名があれば「画像（EMF）」、それ以外の失敗・外部リンク（mediaPath null）は「表示できない」の枠', async () => {
    const { layer } = makeLayer();
    const reasons: Record<string, ImageResult> = {
      emf: { ok: false, reason: 'unsupported', label: 'EMF' },
      none: { ok: false, reason: 'unsupported', label: null },
      big: { ok: false, reason: 'too-large', label: null },
      px: { ok: false, reason: 'oversize-px', label: null },
      miss: { ok: false, reason: 'missing', label: null },
    };
    const names = Object.keys(reasons);
    const drawings = names.map((n, i) => pic(0, i * 3, 40, 30, { mediaPath: n }));
    drawings.push(pic(0, 30, 40, 30, { mediaPath: null }));
    setSheet(layer, drawings, async (p) => reasons[p]!);
    layer.update(ALL);
    await flush();
    const texts = [...layer.element.querySelectorAll('.xs-frame')].map((f) => f.textContent);
    expect(texts).toEqual([
      '画像（EMF）',
      MESSAGES_JA.imageBroken,
      MESSAGES_JA.imageBroken,
      MESSAGES_JA.imageBroken,
      MESSAGES_JA.imageBroken,
      MESSAGES_JA.imageBroken,
    ]);
    expect(layer.element.querySelector('img')).toBeNull();
  });

  it('frameLabel：EMF などの形式名は {format} に入る（画像以外は種類名）', () => {
    const d = pic(0, 0);
    expect(frameLabel(d, { ok: false, reason: 'unsupported', label: 'EMF' })).toBe('画像（EMF）');
    expect(frameLabel(d, { ok: false, reason: 'unsupported', label: null })).toBe(
      MESSAGES_JA.imageBroken,
    );
    expect(frameLabel(pic(0, 0, 1, 1, { kind: 'chart' }))).toBe(MESSAGES_JA.chart);
  });

  it('壊れた画像（img の onerror）は破線枠「表示できない」に差し替える', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [pic(0, 0)]);
    layer.update(ALL);
    await flush();
    const img = layer.element.querySelector('img')!;
    img.dispatchEvent(new Event('error'));
    expect(layer.element.querySelector('img')).toBeNull();
    expect(layer.element.querySelector('.xs-frame')?.textContent).toBe(MESSAGES_JA.imageBroken);
  });

  it('重ね順は drawing の順（後のものが上＝DOM の後ろ）。出し入れしても保たれる', async () => {
    const { layer } = makeLayer();
    const list = [
      pic(0, 0, 10, 10, { name: 'a' }),
      pic(0, 0, 10, 10, { name: 'b' }),
      pic(0, 0, 10, 10, { name: 'c' }),
    ];
    setSheet(layer, list);
    layer.update(ALL);
    await flush();
    const names = () => [...layer.element.querySelectorAll('img')].map((i) => i.alt);
    expect(names()).toEqual(['a', 'b', 'c']);
    layer.update(viewWindow(0, 0, 0, 0, 0, 0)); // すべて外す
    expect(names()).toEqual([]);
    layer.update(ALL);
    expect(names()).toEqual(['a', 'b', 'c']); // 読み込み済みなので、すぐ <img> が入る
  });
});

describe('TD-86-6 画像レイヤー：出し入れ（可視範囲 ± 1 画面）', () => {
  it('可視範囲の外の図は DOM に出さず、スクロールで出し入れする。外れた図も blob URL は保持（再表示で再読み込みしない）', async () => {
    const { layer, created } = makeLayer();
    const read = vi.fn(async () => okResult());
    // 行 1（上）・行 500（y=9980）・行 900（y=17980）
    setSheet(layer, [pic(0, 0), pic(0, 499), pic(0, 899)], read);
    const viewAt = (top: number) => viewWindow(0, top, 800, 400, 32, 24);
    layer.update(viewAt(0));
    await flush();
    const figCount = () => layer.element.querySelectorAll('.xs-figure').length;
    expect(figCount()).toBe(1);
    layer.update(viewAt(9980));
    await flush();
    expect(figCount()).toBe(1);
    expect(layer.element.querySelector<HTMLElement>('.xs-figure')!.style.top).toBe('9980px');
    layer.update(viewAt(0));
    expect(figCount()).toBe(1);
    expect(read).toHaveBeenCalledTimes(1);
    expect(created).toHaveLength(1);
  });

  it('可視範囲が変わらなければ DOM を触らない（同じ要素のまま）', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [pic(0, 0)]);
    layer.update(ALL);
    await flush();
    const el = layer.element.querySelector('.xs-figure');
    layer.update(viewWindow(1, 1, 3000, 3000, 0, 0));
    expect(layer.element.querySelector('.xs-figure')).toBe(el);
  });

  it('読み込み中に外れた図は、読み込みが終わっても要素を触らず、次に出す時に結果を反映する', async () => {
    const { layer } = makeLayer();
    let resolve!: (r: ImageResult) => void;
    setSheet(layer, [pic(0, 0)], () => new Promise<ImageResult>((r) => (resolve = r)));
    layer.update(ALL);
    layer.update(viewWindow(0, 0, 0, 0, 0, 0)); // 外す
    resolve(okResult());
    await flush();
    expect(layer.element.querySelector('img')).toBeNull();
    layer.update(ALL);
    expect(layer.element.querySelector('img')?.getAttribute('src')).toBe('blob:l/1');
  });
});

describe('TD-86-7 画像レイヤー：blob URL の revoke と世代', () => {
  it('シートの差し替えで、前のシートの blob URL をすべて revoke し、DOM の図を外す', async () => {
    const { layer, created, revoked } = makeLayer();
    setSheet(layer, [pic(0, 0), pic(0, 3, 10, 10, { mediaPath: 'xl/media/image2.png' })]);
    layer.update(ALL);
    await flush();
    expect(created).toHaveLength(2);
    setSheet(layer, []);
    expect(revoked.sort()).toEqual(created.slice().sort());
    expect(layer.element.children).toHaveLength(0);
    layer.setSheet(null);
    expect(revoked).toHaveLength(2); // 二重に revoke しない
  });

  it('dispose：すべて revoke し、以後 setSheet・update は何もしない', async () => {
    const { layer, created, revoked } = makeLayer();
    setSheet(layer, [pic(0, 0)]);
    layer.update(ALL);
    await flush();
    layer.dispose();
    expect(revoked).toEqual(created);
    expect(layer.element.children).toHaveLength(0);
    setSheet(layer, [pic(0, 0)]);
    layer.update(ALL);
    await flush();
    expect(created).toHaveLength(1); // 作り直さない
    expect(layer.element.children).toHaveLength(0);
  });

  it('読み込み中にシートが切り替わったら、古い結果の URL を作らない（世代の抑止）', async () => {
    const { layer, created, revoked } = makeLayer();
    let resolveOld!: (r: ImageResult) => void;
    setSheet(layer, [pic(0, 0)], () => new Promise<ImageResult>((r) => (resolveOld = r)));
    layer.update(ALL);
    setSheet(layer, [pic(0, 0, 10, 10, { name: 'new' })]); // 切替（古い読み込みは未完）
    layer.update(ALL);
    await flush();
    resolveOld(okResult()); // 古い読み込みがあとから終わる
    await flush();
    // 作られた URL は新しいシートの 1 つだけ。古い結果の URL は作られず、新しい図に古い画像が入らない
    expect(created).toHaveLength(1);
    expect(revoked).toHaveLength(0);
    const imgs = [...layer.element.querySelectorAll('img')];
    expect(imgs).toHaveLength(1);
    expect(imgs[0]!.alt).toBe('new');
  });

  it('読み込み中に dispose されたら、終わっても URL を作らない', async () => {
    const { layer, created } = makeLayer();
    let resolve!: (r: ImageResult) => void;
    setSheet(layer, [pic(0, 0)], () => new Promise<ImageResult>((r) => (resolve = r)));
    layer.update(ALL);
    layer.dispose();
    resolve(okResult());
    await flush();
    expect(created).toHaveLength(0);
  });

  it('画像の合計の上限超過（over-budget）は onOverBudget を 1 回呼び、その画像は破線枠', async () => {
    const { layer } = makeLayer();
    const onOverBudget = vi.fn();
    setSheet(
      layer,
      [pic(0, 0), pic(0, 3, 10, 10, { mediaPath: 'xl/media/image2.png' })],
      async () => ({ ok: false, reason: 'over-budget', label: null }),
      onOverBudget,
    );
    layer.update(ALL);
    await flush();
    expect(onOverBudget).toHaveBeenCalledTimes(1);
    expect(layer.element.querySelectorAll('.xs-frame')).toHaveLength(2);
  });
});

/* ---- ISSUE-88：文字のある図形・線の DOM ---- */

const run = (text: string, over: Partial<DrawingTextRun> = {}): DrawingTextRun => ({
  br: false,
  text,
  sizePt: 11,
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  ...over,
});
const para = (
  runs: DrawingTextRun[],
  over: Partial<DrawingTextParagraph> = {},
): DrawingTextParagraph => ({ runs, align: 'l', bullet: null, endSizePt: 11, ...over });
const textOf = (
  paragraphs: DrawingTextParagraph[],
  over: Partial<DrawingText> = {},
): DrawingText => ({
  paragraphs,
  anchor: 't',
  insets: { l: 91440, t: 45720, r: 91440, b: 45720 },
  noWrap: false,
  fontScale: 1,
  vert: 'horz',
  ...over,
});

/** 文字のある図形（absolute 配置。x・y・w・h は px） */
function shapeWithText(
  text: DrawingText,
  over: Partial<SheetDrawing> = {},
  rect = { x: 100, y: 50, w: 200, h: 80 },
): SheetDrawing {
  return {
    anchor: { type: 'absolute', x: px(rect.x), y: px(rect.y), cx: px(rect.w), cy: px(rect.h) },
    kind: 'shape',
    name: '<b>name</b>',
    descr: '',
    mediaPath: null,
    text,
    ...over,
  };
}

/** 線（absolute 配置） */
function lineDrawing(
  over: Partial<SheetDrawing> = {},
  line: Partial<DrawingLine> = {},
  rect = { x: 100, y: 50, w: 200, h: 80 },
): SheetDrawing {
  return {
    anchor: { type: 'absolute', x: px(rect.x), y: px(rect.y), cx: px(rect.w), cy: px(rect.h) },
    kind: 'line',
    name: 'line',
    descr: '',
    mediaPath: null,
    line: {
      geom: 'straightConnector1',
      adj: [50000, 50000, 50000],
      widthEmu: 12700,
      dash: 'solid',
      head: null,
      tail: null,
      ...line,
    },
    ...over,
  };
}

/** レイヤーの直下の図の要素（ダーク表示の下地 `.xs-backdrop` は除く。ISSUE-96） */
const figEls = (layer: ImageLayer): Element[] =>
  [...layer.element.children].filter((c) => !c.classList.contains('xs-backdrop'));

/** 図を 1 つ出して、その要素を返す */
function mountOne(drawing: SheetDrawing): { layer: ImageLayer; el: HTMLElement } {
  const { layer } = makeLayer();
  setSheet(layer, [drawing]);
  layer.update(ALL);
  return { layer, el: figEls(layer)[0] as HTMLElement };
}

describe('TD-88-12 画像レイヤー：文字のある図形の DOM', () => {
  it('段落は div・ラン は span・改行は br。HTML 風の文字列は要素にならず textContent で入る。種類名の破線枠は出ない', () => {
    const { layer, el } = mountOne(
      shapeWithText(
        textOf([
          para([
            run('<b>x</b>'),
            run('<img src=x onerror=alert(1)>'),
            run('', { br: true }),
            run('y'),
          ]),
          para([run('2 段落目')]),
        ]),
      ),
    );
    expect(el.classList.contains('xs-shape-text')).toBe(true);
    const paras = [...el.children];
    expect(paras.map((p) => p.tagName)).toEqual(['DIV', 'DIV']);
    const first = paras[0]!;
    expect([...first.children].map((c) => c.tagName)).toEqual(['SPAN', 'SPAN', 'BR', 'SPAN']);
    expect(first.children[0]!.textContent).toBe('<b>x</b>');
    expect(first.children[1]!.textContent).toBe('<img src=x onerror=alert(1)>');
    expect(layer.element.querySelector('b, img, .xs-frame')).toBeNull();
    expect(el.textContent).toContain('2 段落目');
    // 図形の名前は画面に出さない
    expect(el.textContent).not.toContain('name');
  });

  it('枠と背景：hasOutline・hasFill の 4 通りがクラスになる（style に色は入れない）', () => {
    for (const [outline, fill] of [
      [false, false],
      [true, false],
      [false, true],
      [true, true],
    ] as const) {
      const { el } = mountOne(
        shapeWithText(textOf([para([run('a')])]), { hasOutline: outline, hasFill: fill }),
      );
      expect(el.classList.contains('has-outline')).toBe(outline);
      expect(el.classList.contains('has-fill')).toBe(fill);
      expect(el.style.background).toBe('');
      expect(el.style.border).toBe('');
      document.body.replaceChildren();
    }
  });

  it('文字の無い図形・テキストなしのテキストボックス・改行だけの図形は種類名の破線枠のまま', () => {
    const { layer } = makeLayer();
    setSheet(layer, [
      shapeWithText(textOf([para([run('', { br: true })])]), { kind: 'textbox' }),
      shapeWithText(textOf([para([run('')])]), { text: undefined }, { x: 0, y: 200, w: 50, h: 40 }),
      { ...pic(0, 20), kind: 'shape', mediaPath: null },
    ]);
    layer.update(ALL);
    const frames = [...layer.element.querySelectorAll('.xs-frame')];
    expect(frames.map((f) => f.textContent)).toEqual([
      MESSAGES_JA.textbox,
      MESSAGES_JA.shape,
      MESSAGES_JA.shape,
    ]);
    expect(layer.element.querySelector('.xs-shape-text')).toBeNull();
  });

  it('グラフ・SmartArt・グループは text があっても種類名の破線枠（文字は描かない）', () => {
    const { layer } = makeLayer();
    setSheet(
      layer,
      (['chart', 'smartart', 'group'] as const).map((kind, i) =>
        shapeWithText(textOf([para([run('文字')])]), { kind }, { x: 0, y: i * 100, w: 50, h: 40 }),
      ),
    );
    layer.update(ALL);
    expect(layer.element.querySelectorAll('.xs-frame')).toHaveLength(3);
    expect(layer.element.querySelector('.xs-shape-text')).toBeNull();
  });
});

describe('TD-88-13 画像レイヤー：文字のスタイル', () => {
  it('文字の大きさ（pt→px・fontScale）・太字・斜体・下線と取り消し線・余白（EMU→px・既定）', () => {
    const { el } = mountOne(
      shapeWithText(
        textOf(
          [
            para([
              run('a', { sizePt: 12, bold: true, italic: true, underline: true, strike: true }),
              run('b', { sizePt: 18 }),
            ]),
          ],
          { fontScale: 0.5, insets: { l: 95250, t: 19050, r: 0, b: 9525 } },
        ),
      ),
    );
    const [a, b] = [...el.querySelectorAll('span')] as HTMLElement[];
    expect(a!.style.fontSize).toBe('8px'); // 12pt = 16px × 0.5
    expect(b!.style.fontSize).toBe('12px'); // 18pt = 24px × 0.5
    expect(a!.style.fontWeight).toBe('bold');
    expect(a!.style.fontStyle).toBe('italic');
    expect(a!.style.textDecoration).toBe('underline line-through');
    expect(b!.style.fontWeight).toBe('');
    expect(b!.style.textDecoration).toBe('');
    expect([
      el.style.paddingLeft,
      el.style.paddingTop,
      el.style.paddingRight,
      el.style.paddingBottom,
    ]).toEqual(['10px', '2px', '0px', '1px']);
  });

  it('余白が無いとき Excel の既定（l/r 91440・t/b 45720 EMU＝9.6px・4.8px）', () => {
    const text = textOf([para([run('a')])]);
    delete (text as { insets?: unknown }).insets;
    const { el } = mountOne(shapeWithText(text));
    expect([el.style.paddingLeft, el.style.paddingTop]).toEqual(['9.6px', '4.8px']);
  });

  it('折り返し：既定は pre-wrap・noWrap は pre', () => {
    const wrap = mountOne(shapeWithText(textOf([para([run('a')])]))).el;
    expect(wrap.style.whiteSpace).toBe('pre-wrap');
    document.body.replaceChildren();
    const none = mountOne(shapeWithText(textOf([para([run('a')])], { noWrap: true }))).el;
    expect(none.style.whiteSpace).toBe('pre');
  });

  it('段落の揃え（論理値）と縦位置（flex の justify-content）', () => {
    const aligns = {
      l: 'start',
      ctr: 'center',
      r: 'end',
      just: 'justify',
      dist: 'justify',
    } as const;
    const paragraphs = Object.keys(aligns).map((k) =>
      para([run(k)], { align: k as keyof typeof aligns }),
    );
    const { el } = mountOne(shapeWithText(textOf(paragraphs)));
    expect([...el.children].map((p) => (p as HTMLElement).style.textAlign)).toEqual(
      Object.values(aligns),
    );
    for (const [anchor, jc] of [
      ['t', 'flex-start'],
      ['ctr', 'center'],
      ['b', 'flex-end'],
    ] as const) {
      document.body.replaceChildren();
      const one = mountOne(shapeWithText(textOf([para([run('a')])], { anchor }))).el;
      expect(one.style.justifyContent).toBe(jc);
    }
  });

  it('箇条書きは段落の先頭に文字として足す。空段落は endSizePt の行の高さを保つ', () => {
    const { el } = mountOne(
      shapeWithText(
        textOf([para([run('項目', { sizePt: 10 })], { bullet: '•' }), para([], { endSizePt: 20 })]),
      ),
    );
    const [p1, p2] = [...el.children] as HTMLElement[];
    expect(p1!.firstElementChild!.textContent).toBe('• ');
    expect(p1!.textContent).toBe('• 項目');
    expect(p2!.textContent).toBe('');
    // 20pt = 26.67px × 行の高さ 1.2 = 32px
    expect(p2!.style.getPropertyValue('min-block-size')).toBe('32px');
  });

  it('文字のある段落の行は最大のランの大きさ。endSizePt も最小の高さも付けない', () => {
    const { el } = mountOne(
      shapeWithText(
        textOf([
          para(
            [
              run('a', { sizePt: 8 }),
              { ...run('', { sizePt: 8 }), br: true },
              run('b', { sizePt: 8 }),
            ],
            {
              endSizePt: 11,
            },
          ),
          para([run('x', { sizePt: 8 }), run('y', { sizePt: 20 })]),
        ]),
      ),
    );
    const [p1, p2] = [...el.children] as HTMLElement[];
    expect(p1!.style.fontSize).toBe('10.67px'); // 8pt（endSizePt 11pt ではない）
    expect(p1!.style.getPropertyValue('min-block-size')).toBe('');
    expect(p2!.style.fontSize).toBe('26.67px'); // 最大の 20pt
  });

  it('余白は各辺とも枠の半分まで（巨大な insets で要素が膨らまない）', () => {
    const { el } = mountOne(
      shapeWithText(
        textOf([para([run('a')])], {
          insets: { l: 2147483647, t: 2147483647, r: 2147483647, b: 0 },
        }),
        {},
        { x: 0, y: 0, w: 200, h: 80 },
      ),
    );
    expect([el.style.paddingLeft, el.style.paddingRight]).toEqual(['100px', '100px']);
    expect([el.style.paddingTop, el.style.paddingBottom]).toEqual(['40px', '0px']);
  });

  it('末尾が改行のラン列は br を 1 つ足す（HTML で空行が出る）。途中の改行は足さない', () => {
    const br = { ...run(''), br: true };
    const { el } = mountOne(
      shapeWithText(textOf([para([run('a'), br]), para([run('a'), br, run('b')])])),
    );
    const [p1, p2] = [...el.children] as HTMLElement[];
    expect(p1!.querySelectorAll('br')).toHaveLength(2);
    expect(p2!.querySelectorAll('br')).toHaveLength(1);
  });
});

describe('TD-88-14 画像レイヤー：回転と縦書き', () => {
  it('回転は transform: rotate(deg)。0 度は付けない。反転は掛けない', () => {
    const rot = mountOne(
      shapeWithText(textOf([para([run('a')])]), { rot: 30 * 60000, flipH: true, flipV: true }),
    ).el;
    expect(rot.style.transform).toBe('rotate(30deg)');
    document.body.replaceChildren();
    const none = mountOne(shapeWithText(textOf([para([run('a')])]))).el;
    expect(none.style.transform).toBe('');
  });

  it('90° の図は回転前の矩形（幅と高さの入れ替え・中心は同じ）に置く。アンカーの矩形はそのまま可視判定に使う', () => {
    // アンカー 200×80 の中心 (200, 90) → 回転前 80×200 の左上 (160, -10)
    const { el } = mountOne(
      shapeWithText(
        textOf([para([run('a')])]),
        { rot: 90 * 60000 },
        { x: 100, y: 50, w: 200, h: 80 },
      ),
    );
    expect([el.style.left, el.style.top, el.style.width, el.style.height]).toEqual([
      '160px',
      '-10px',
      '80px',
      '200px',
    ]);
    expect(el.style.transform).toBe('rotate(90deg)');
    // 回転が 45° 未満なら入れ替えない
    document.body.replaceChildren();
    const small = mountOne(shapeWithText(textOf([para([run('a')])]), { rot: 30 * 60000 })).el;
    expect([small.style.left, small.style.width]).toEqual(['100px', '200px']);
  });

  it('可視判定はアンカーの矩形のまま（アンカーが窓の外なら回転前の矩形が窓に掛かっても出さない）', () => {
    const { layer } = makeLayer();
    // アンカー x 500〜700。窓 x 0〜400 の外、0〜600 なら内（回転前の矩形 x 560〜640 は関係しない）
    setSheet(layer, [
      shapeWithText(
        textOf([para([run('a')])]),
        { rot: 90 * 60000 },
        { x: 500, y: 0, w: 200, h: 80 },
      ),
    ]);
    layer.update({ x: 0, y: 0, w: 400, h: 400 });
    expect(figEls(layer)).toHaveLength(0);
    layer.update({ x: 0, y: 0, w: 600, h: 400 });
    expect(figEls(layer)).toHaveLength(1);
  });

  it('縦書き：vert の列挙値ごとに writing-mode・text-orientation。horz と未知の値は横書き（style に出ない）', () => {
    const table: [string, string, string][] = [
      ['eaVert', 'vertical-rl', ''],
      ['wordArtVertRtl', 'vertical-rl', ''],
      ['mongolianVert', 'vertical-lr', ''],
      ['vert', 'vertical-rl', 'sideways'],
      ['vert270', 'sideways-lr', ''],
      ['wordArtVert', 'vertical-rl', 'upright'],
      ['horz', '', ''],
      ['unknown', '', ''],
    ];
    for (const [vert, mode, orientation] of table) {
      document.body.replaceChildren();
      const { el } = mountOne(
        shapeWithText(textOf([para([run('a')])], { vert: vert as DrawingText['vert'] })),
      );
      expect(el.style.getPropertyValue('writing-mode'), vert).toBe(mode);
      expect(el.style.getPropertyValue('text-orientation'), vert).toBe(orientation);
    }
  });

  it('文字の無い図形の破線枠は回転を無視する（今どおり）', () => {
    const { el } = mountOne({
      ...pic(0, 0, 40, 30),
      kind: 'shape',
      mediaPath: null,
      rot: 45 * 60000,
    });
    expect(el.style.transform).toBe('');
    expect(el.querySelector('.xs-frame')).not.toBeNull();
  });
});

describe('TD-88-15 画像レイヤー：線の SVG', () => {
  const NS = 'http://www.w3.org/2000/svg';
  const arrows = {
    head: { type: 'triangle', w: 'med', len: 'med' },
    tail: { type: 'oval', w: 'med', len: 'med' },
  } as const;

  /** 線 1 本を出して svg を返す */
  function mountLine(d: SheetDrawing): { svg: SVGSVGElement; layer: ImageLayer } {
    const { layer } = makeLayer();
    setSheet(layer, [d]);
    layer.update(ALL);
    return { svg: figEls(layer)[0] as SVGSVGElement, layer };
  }

  it('<svg> は SVG 名前空間・外接矩形の位置と大きさ・viewBox は数値だけ。<marker>・破線枠・ラベルは無い', () => {
    const d = lineDrawing({}, { ...arrows });
    const box = layoutDrawings(geo, [d])[0]!;
    const { svg, layer } = mountLine(d);
    expect(svg.namespaceURI).toBe(NS);
    expect(svg.tagName).toBe('svg');
    expect(svg.classList.contains('xs-line')).toBe(true);
    expect([svg.style.left, svg.style.top, svg.style.width, svg.style.height]).toEqual([
      `${box.rect.x}px`,
      `${box.rect.y}px`,
      `${box.rect.w}px`,
      `${box.rect.h}px`,
    ]);
    expect(svg.getAttribute('viewBox')).toBe(`0 0 ${box.rect.w} ${box.rect.h}`);
    expect(layer.element.querySelector('marker, .xs-frame, foreignObject')).toBeNull();
    expect(svg.textContent).toBe('');
  });

  it('属性は数値・列挙値だけ：path の d は数値と M/L/C/Z のみ。stroke は currentColor・fill は none', () => {
    const { svg } = mountLine(
      lineDrawing({}, { ...arrows, dash: 'dash', geom: 'curvedConnector3' }),
    );
    const main = svg.querySelector('path')!;
    expect(main.getAttribute('d')).toMatch(/^[MLCZ0-9.\-\s]+$/);
    expect(main.getAttribute('fill')).toBe('none');
    expect(main.getAttribute('stroke')).toBe('currentColor');
    expect(main.getAttribute('stroke-width')).toMatch(/^[0-9.]+$/);
    expect(main.getAttribute('stroke-dasharray')).toMatch(/^[0-9.]+( [0-9.]+)*$/);
    expect(main.getAttribute('stroke-linejoin')).toBe('round');
    expect(main.getAttribute('stroke-linecap')).toBe('butt');
    for (const e of svg.querySelectorAll('*')) {
      expect(e.namespaceURI).toBe(NS);
      for (const a of [...e.attributes]) {
        expect(a.value, `${e.tagName}@${a.name}`).toMatch(/^[A-Za-z0-9.\-\s(),]*$/);
      }
    }
  });

  it('実線は stroke-dasharray を付けない', () => {
    const { svg } = mountLine(lineDrawing());
    expect(svg.querySelector('path')!.hasAttribute('stroke-dasharray')).toBe(false);
  });

  it('矢じり：多角形は塗りの path、楕円は ellipse（rotate 変換は数値だけ）', () => {
    const d = lineDrawing({}, { ...arrows });
    const { svg } = mountLine(d);
    const paths = svg.querySelectorAll('path');
    expect(paths).toHaveLength(2); // 本線＋三角形
    expect(paths[1]!.getAttribute('fill')).toBe('currentColor');
    expect(paths[1]!.getAttribute('d')).toMatch(/^M[0-9.\-\s]+(L[0-9.\-\s]+)+Z$/);
    const ell = svg.querySelector('ellipse')!;
    expect(ell.namespaceURI).toBe(NS);
    expect(ell.getAttribute('transform')).toMatch(/^rotate\([0-9.-]+ [0-9.-]+ [0-9.-]+\)$/);
    for (const k of ['cx', 'cy', 'rx', 'ry']) {
      expect(ell.getAttribute(k)).toMatch(/^[0-9.-]+$/);
    }
  });

  it('開いた V の矢じり（arrow）は fill none・stroke currentColor で、閉じない', () => {
    const { svg } = mountLine(lineDrawing({}, { tail: { type: 'arrow', w: 'med', len: 'med' } }));
    const v = svg.querySelectorAll('path')[1]!;
    expect(v.getAttribute('fill')).toBe('none');
    expect(v.getAttribute('stroke')).toBe('currentColor');
    expect(v.getAttribute('d')).not.toContain('Z');
  });

  it('矢じりなしは path 1 つ。水平線（高さ 0 のアンカー）でも太さ分の大きさの svg が出る', () => {
    const { svg } = mountLine(lineDrawing({}, {}, { x: 100, y: 50, w: 200, h: 0 }));
    expect(svg.querySelectorAll('path')).toHaveLength(1);
    expect(svg.querySelector('ellipse')).toBeNull();
    expect(Number.parseFloat(svg.style.height)).toBeGreaterThan(0);
  });

  it('可視範囲の外の線は DOM に無く、窓に入ると出て、外れると外す（外接矩形で判定）', () => {
    const { layer } = makeLayer();
    setSheet(layer, [lineDrawing({}, {}, { x: 100, y: 50, w: 200, h: 80 })]);
    layer.update({ x: 1000, y: 1000, w: 500, h: 500 });
    expect(layer.element.children).toHaveLength(0);
    layer.update({ x: 0, y: 0, w: 500, h: 500 });
    expect(layer.element.querySelectorAll('svg.xs-line')).toHaveLength(1);
    layer.update({ x: 1000, y: 1000, w: 500, h: 500 });
    expect(layer.element.children).toHaveLength(0);
  });

  it('console に図形の文字・名前を出さない', () => {
    const spies = (['log', 'warn', 'error', 'info', 'debug'] as const).map((k) =>
      vi.spyOn(console, k).mockImplementation(() => {}),
    );
    const { layer } = makeLayer();
    setSheet(layer, [
      shapeWithText(textOf([para([run('秘密の文字')])]), { name: '秘密の名前' }),
      lineDrawing({ name: '秘密の線' }),
    ]);
    layer.update(ALL);
    for (const s of spies) {
      expect(s).not.toHaveBeenCalled();
      s.mockRestore();
    }
  });
});

describe('TD-88-16 画像レイヤー：重ね順（画像・文字・線を drawing の順に）', () => {
  it('drawing の順（後が上）で並ぶ。窓の出し入れの後も保たれ、既存の要素は再利用する', async () => {
    const { layer } = makeLayer();
    setSheet(layer, [
      lineDrawing({ name: 'a' }),
      pic(0, 10),
      shapeWithText(textOf([para([run('文字')])]), {}, { x: 0, y: 400, w: 100, h: 40 }),
      lineDrawing({ name: 'b' }, {}, { x: 100, y: 500, w: 200, h: 80 }),
    ]);
    layer.update(ALL);
    await flush();
    const kinds = (): string[] =>
      figEls(layer).map((c) =>
        c.tagName === 'svg' ? 'line' : c.classList.contains('xs-shape-text') ? 'text' : 'image',
      );
    expect(kinds()).toEqual(['line', 'image', 'text', 'line']);
    const keep = figEls(layer)[1];
    // 先頭の線だけ窓の外へ（y 0〜130 → 窓を 200 以降に）
    layer.update({ x: -3000, y: 200, w: 9000, h: 1000 });
    expect(kinds()).toEqual(['image', 'text', 'line']);
    expect(figEls(layer)[0]).toBe(keep);
    layer.update(ALL);
    expect(kinds()).toEqual(['line', 'image', 'text', 'line']);
  });
});
