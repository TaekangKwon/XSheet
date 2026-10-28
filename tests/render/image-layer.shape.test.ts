// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createGeometry, type SheetDrawing } from '../../src/core/index';
import { emptyLayout, type DrawingText } from '../../src/core/types';
import { EMU_PER_PX, layoutDrawings, viewWindow } from '../../src/render/drawings';
import { createImageLayer, type ImageLayer } from '../../src/render/image-layer';

/**
 * 画像レイヤー（DOM）：図形を形で描く（ISSUE-92・ADR-0028 §決定 4・5）。
 * 形の `<svg>`・`<path>`、枠線・塗り・面の濃淡のクラス、文字領域、回転・反転、見えない図形、代替（ADR-0027 の描き方）を確かめる。
 * jsdom はレイアウトを持たないため、見た目（色・実寸）は目視（TV-41）で確認する。
 */

const px = (n: number): number => n * EMU_PER_PX;
const geo = createGeometry(emptyLayout(), 1000, 1000);
const ALL = viewWindow(0, 0, 3000, 3000, 0, 0);
const SVG_NS = 'http://www.w3.org/2000/svg';

function makeLayer(): ImageLayer {
  const layer = createImageLayer({ createUrl: () => 'blob:x', revokeUrl: () => undefined });
  document.body.append(layer.element);
  return layer;
}

afterEach(() => {
  document.body.replaceChildren();
});

const TEXT: DrawingText = {
  paragraphs: [
    {
      runs: [
        {
          text: '<b>文字</b>',
          sizePt: 11,
          bold: false,
          italic: false,
          underline: false,
          strike: false,
        },
      ],
      align: 'l',
      bullet: null,
      endSizePt: 11,
    },
  ],
  anchor: 't',
  insets: { l: 91440, t: 45720, r: 91440, b: 45720 },
  noWrap: false,
  fontScale: 1,
  vert: 'horz',
} as DrawingText;

/** 図形（absolute 配置の 100×50・x=10・y=20）。既定は枠線のある矩形 */
function shape(over: Partial<SheetDrawing> = {}): SheetDrawing {
  return {
    anchor: { type: 'absolute', x: px(10), y: px(20), cx: px(100), cy: px(50) },
    kind: 'shape',
    name: '図形の名前',
    descr: '',
    mediaPath: null,
    geom: { type: 'preset', prst: 'rect', adj: [] },
    hasOutline: true,
    hasFill: false,
    ...over,
  };
}

/** 図を並べて出し、DOM の図（`.xs-figure`）を返す */
function mount(drawings: SheetDrawing[]): { layer: ImageLayer; figs: HTMLElement[] } {
  const layer = makeLayer();
  layer.setSheet({
    drawings,
    boxes: layoutDrawings(geo, drawings),
    readImage: async () => ({ ok: false, reason: 'missing', label: null }),
  });
  layer.update(ALL);
  const figs = [...layer.element.querySelectorAll<HTMLElement>(':scope > .xs-figure')];
  return { layer, figs };
}

describe('TD-92-7 形の <svg>・<path>（createElementNS・数値だけの属性）', () => {
  it('矩形（枠線のみ）：figure ＞ svg（外接矩形いっぱい）＞ path 1 つ。属性は数値・列挙値だけ', () => {
    const { figs } = mount([shape()]);
    expect(figs).toHaveLength(1);
    const fig = figs[0]!;
    expect([fig.style.left, fig.style.top, fig.style.width, fig.style.height]).toEqual([
      '8px',
      '18px',
      '104px',
      '54px',
    ]);
    const svg = fig.querySelector('svg')!;
    expect(svg.namespaceURI).toBe(SVG_NS);
    expect(svg.classList.contains('xs-shape-svg')).toBe(true);
    expect(svg.getAttribute('viewBox')).toBe('0 0 104 54');
    expect(svg.getAttribute('width')).toBe('104');
    expect(svg.getAttribute('aria-hidden')).toBe('true');
    const paths = [...svg.querySelectorAll('path')];
    expect(paths).toHaveLength(1);
    const p = paths[0]!;
    expect(p.namespaceURI).toBe(SVG_NS);
    expect(p.getAttribute('d')).toBe('M2 2 L102 2 L102 52 L2 52 Z');
    expect(p.getAttribute('stroke-width')).toBe('1');
    expect(p.getAttribute('stroke-linejoin')).toBe('miter');
    expect(p.getAttribute('stroke-miterlimit')).toBe('4');
    expect(p.getAttribute('stroke-dasharray')).toBeNull();
    expect(p.classList.contains('has-stroke')).toBe(true);
    expect(p.classList.contains('has-fill')).toBe(false);
    // 図形の名前は画面に出さない・種類名の破線枠は出ない・色は style に入れない
    expect(fig.textContent).toBe('');
    expect(fig.querySelector('.xs-frame')).toBeNull();
    expect(p.getAttribute('style')).toBeNull();
    expect(p.getAttribute('fill')).toBeNull();
    expect(p.getAttribute('stroke')).toBeNull();
  });

  it('属性の値は数値・空白・path の命令文字だけ（ファイルの文字列は入らない）', () => {
    const { layer } = mount([
      shape({
        hasFill: true,
        geom: { type: 'preset', prst: 'rightArrow', adj: [['adj1', 30000]] },
      }),
      shape({ hasFill: true, geom: { type: 'preset', prst: 'can', adj: [] } }),
    ]);
    for (const p of layer.element.querySelectorAll('path')) {
      expect(p.getAttribute('d')).toMatch(/^[MLCZ0-9 .-]+$/);
      expect(p.getAttribute('d')).not.toMatch(/NaN|Infinity/);
    }
    for (const el of layer.element.querySelectorAll('svg, svg *')) {
      for (const a of el.getAttributeNames()) {
        expect(el.getAttribute(a)).not.toMatch(/[<>"']/);
      }
    }
  });

  it('枠線の太さ・線種：太さは stroke-width、線種は stroke-dasharray（点線 [w,2w]・破線 [4w,3w]）', () => {
    const { layer } = mount([
      shape({ outline: { widthEmu: px(3), dash: 'dash' } }),
      shape({ outline: { widthEmu: px(2), dash: 'dot' } }),
      shape({ outline: { widthEmu: 0, dash: 'solid' } }),
    ]);
    const paths = [...layer.element.querySelectorAll('.xs-shape-svg path')];
    expect(paths.map((p) => p.getAttribute('stroke-width'))).toEqual(['3', '2', '1']);
    expect(paths.map((p) => p.getAttribute('stroke-dasharray'))).toEqual(['12 9', '2 4', null]);
  });

  it('塗り：has-fill（半透明の塗りは CSS）。枠線なしなら has-stroke は付かない', () => {
    const { layer } = mount([shape({ hasOutline: false, hasFill: true })]);
    const p = layer.element.querySelector('.xs-shape-svg path')!;
    expect(p.classList.contains('has-fill')).toBe(true);
    expect(p.classList.contains('has-stroke')).toBe(false);
    expect(p.hasAttribute('stroke-width')).toBe(false);
  });

  it('円柱：path ごとに 1 つの <path>。面の濃淡は is-* のクラス、fill="none" の path は塗らず、stroke="false" の path は線を引かない', () => {
    const { layer } = mount([
      shape({ hasFill: true, geom: { type: 'preset', prst: 'can', adj: [] } }),
    ]);
    const paths = [...layer.element.querySelectorAll('.xs-shape-svg path')];
    expect(paths).toHaveLength(3);
    const cls = (p: Element) =>
      ['has-fill', 'has-stroke', 'is-lighten'].filter((c) => p.classList.contains(c));
    expect(cls(paths[0]!)).toEqual(['has-fill']);
    expect(cls(paths[1]!)).toEqual(['has-fill', 'is-lighten']);
    expect(cls(paths[2]!)).toEqual(['has-stroke']);
    // 線を引かない path には stroke 系の属性を付けない
    expect(paths[0]!.hasAttribute('stroke-width')).toBe(false);
    expect(paths[2]!.getAttribute('stroke-width')).toBe('1');
  });

  it('面の濃淡 4 種のクラス（darken・darkenLess・lighten・lightenLess）', () => {
    const def = (fill: string) => ({
      type: 'custom' as const,
      def: {
        av: [],
        gd: [],
        paths: [
          {
            w: 10,
            h: 10,
            fill: fill as 'darken',
            c: [['M', 0, 0], ['L', 10, 0], ['L', 10, 10], ['Z']] as const,
          },
        ],
      },
    });
    const { layer } = mount(
      ['darken', 'darkenLess', 'lighten', 'lightenLess', 'norm'].map((f) =>
        shape({ hasFill: true, hasOutline: false, geom: def(f) }),
      ),
    );
    const classes = [...layer.element.querySelectorAll('.xs-shape-svg path')].map((p) =>
      [...p.classList].filter((c) => c.startsWith('is-')),
    );
    expect(classes).toEqual([
      ['is-darken'],
      ['is-darken-less'],
      ['is-lighten'],
      ['is-lighten-less'],
      [],
    ]);
  });

  it('見えない図形（枠も塗りも文字も無い）は描かない。文字だけの図形は svg 無しで文字だけ', () => {
    const { layer, figs } = mount([
      shape({ hasOutline: false, hasFill: false }),
      shape({ hasOutline: false, hasFill: false, text: TEXT }),
    ]);
    expect(figs).toHaveLength(1);
    expect(figs[0]!.querySelector('svg')).toBeNull();
    expect(figs[0]!.querySelector('.xs-frame')).toBeNull();
    expect(layer.element.querySelectorAll('.xs-frame')).toHaveLength(0);
    expect(figs[0]!.textContent).toBe('<b>文字</b>');
  });
});

describe('TD-92-8 文字領域・回転・反転・図形の矩形で切る', () => {
  it('文字：figure ＞ div.xs-shape-box（回転前の矩形）＞ div.xs-shape-text（文字領域）。本文は textContent', () => {
    const { layer, figs } = mount([shape({ text: TEXT, hasFill: true })]);
    const fig = figs[0]!;
    const box = fig.querySelector<HTMLElement>(':scope > .xs-shape-box')!;
    // 図形の矩形 (10,20,100×50)、figure は (8,18)。回転なしなので transform は付けない
    expect([box.style.left, box.style.top, box.style.width, box.style.height]).toEqual([
      '2px',
      '2px',
      '100px',
      '50px',
    ]);
    expect(box.style.transform).toBe('');
    const text = box.querySelector<HTMLElement>(':scope > .xs-shape-text')!;
    expect([text.style.left, text.style.top, text.style.width, text.style.height]).toEqual([
      '0px',
      '0px',
      '100px',
      '50px',
    ]);
    expect(text.textContent).toBe('<b>文字</b>');
    expect(layer.element.querySelector('b')).toBeNull();
    // 文字の余白（既定 l/r 0.1 インチ = 9.6px、t/b 0.05 インチ = 4.8px）
    expect(text.style.paddingLeft).toBe('9.6px');
    expect(text.style.paddingTop).toBe('4.8px');
    // 枠線・塗りは svg が描く（文字の枠に has-outline・has-fill を付けない）
    expect(text.classList.contains('has-outline')).toBe(false);
    expect(text.classList.contains('has-fill')).toBe(false);
    expect(fig.querySelector('svg')).not.toBeNull();
  });

  it('楕円：文字領域は内接矩形。図形の矩形（box）で切る（overflow は CSS・box は図形の矩形の大きさ）', () => {
    const { figs } = mount([
      shape({ text: TEXT, geom: { type: 'preset', prst: 'ellipse', adj: [] } }),
    ]);
    const box = figs[0]!.querySelector<HTMLElement>('.xs-shape-box')!;
    const text = box.querySelector<HTMLElement>('.xs-shape-text')!;
    expect(box.style.width).toBe('100px');
    const w = Number.parseFloat(text.style.width);
    expect(w).toBeCloseTo(2 * 50 * Math.cos(Math.PI / 4), 1);
    expect(Number.parseFloat(text.style.left)).toBeCloseTo(50 - w / 2, 1);
  });

  it('回転：回転前の矩形の box に rotate() を掛ける。図形の幅・高さは入れ替えない（アンカーが縦横入れ替わりの外接）', () => {
    // 回転前 100×50 の中心 (60,45)→ アンカーは 50×100 で (35,-5)。absolute 配置で y 負は 0 扱いになるので中心を下げる
    const d = shape({
      anchor: { type: 'absolute', x: px(35), y: px(20), cx: px(50), cy: px(100) },
      rot: 90 * 60000,
      text: TEXT,
      hasOutline: false,
      hasFill: true,
    });
    const { figs } = mount([d]);
    const box = figs[0]!.querySelector<HTMLElement>('.xs-shape-box')!;
    expect(box.style.transform).toBe('rotate(90deg)');
    expect([box.style.width, box.style.height]).toEqual(['100px', '50px']);
    // figure はアンカーと経路の外接（回転後 50×100）。box は図形の中心に置かれる：中心 (60,70) → 左上 (10,45)、figure は (35,20)
    expect([box.style.left, box.style.top]).toEqual(['-25px', '25px']);
  });

  it('反転：文字領域の位置は矩形の中で鏡写し。文字そのものは反転しない（scale を使わない）', () => {
    const arrow = { type: 'preset' as const, prst: 'rightArrow', adj: [] };
    const base = mount([shape({ text: TEXT, geom: arrow })]);
    const flipped = mount([shape({ text: TEXT, geom: arrow, flipH: true })]);
    const t0 = base.figs[0]!.querySelector<HTMLElement>('.xs-shape-text')!;
    const t1 = flipped.figs[0]!.querySelector<HTMLElement>('.xs-shape-text')!;
    const l0 = Number.parseFloat(t0.style.left);
    const w0 = Number.parseFloat(t0.style.width);
    expect(Number.parseFloat(t1.style.left)).toBeCloseTo(100 - (l0 + w0), 1);
    expect(flipped.figs[0]!.innerHTML).not.toContain('scale');
  });

  it('上下反転（flipV・flipHV）：文字領域の位置は鏡写しのまま、文字を 180° 回す（Excel と同じ。左右反転・反転なしは回さない。ISSUE-93）', () => {
    const arrow = { type: 'preset' as const, prst: 'rightArrow', adj: [] };
    const textOf = (d: Partial<SheetDrawing>): HTMLElement =>
      mount([shape({ text: TEXT, geom: arrow, ...d })]).figs[0]!.querySelector<HTMLElement>(
        '.xs-shape-text',
      )!;
    expect(textOf({}).style.transform).toBe('');
    expect(textOf({ flipH: true }).style.transform).toBe('');
    expect(textOf({ flipV: true }).style.transform).toBe('rotate(180deg)');
    expect(textOf({ flipH: true, flipV: true }).style.transform).toBe('rotate(180deg)');
    // 位置は rightArrow の文字領域を上下に鏡写ししたもの（矢の軸は上下対称なので top は変わらない）
    const none = textOf({});
    const v = textOf({ flipV: true });
    expect(v.style.top).toBe(none.style.top);
    expect(v.style.left).toBe(none.style.left);
    // 回転図形の中でも文字だけを回す：box の回転は rot のまま
    const fig = mount([shape({ text: TEXT, geom: arrow, flipV: true, rot: 30 * 60000 })]).figs[0]!;
    expect(fig.querySelector<HTMLElement>('.xs-shape-box')!.style.transform).toBe('rotate(30deg)');
    expect(fig.querySelector<HTMLElement>('.xs-shape-text')!.style.transform).toBe(
      'rotate(180deg)',
    );
  });

  it('文字の向きと位置：上下反転の三角形は、反転なしの文字領域（下半分）を上半分へ鏡写しして、その中で 180° 回る', () => {
    const tri = { type: 'preset' as const, prst: 'triangle', adj: [] };
    const none = mount([shape({ text: TEXT, geom: tri })]).figs[0]!.querySelector<HTMLElement>(
      '.xs-shape-text',
    )!;
    const v = mount([
      shape({ text: TEXT, geom: tri, flipV: true }),
    ]).figs[0]!.querySelector<HTMLElement>('.xs-shape-text')!;
    // 反転なし：文字領域は下半分（top=h/2・height=h/2）、反転：上半分（top=0）
    expect(Number.parseFloat(none.style.top)).toBeCloseTo(25, 1);
    expect(Number.parseFloat(v.style.top)).toBeCloseTo(0, 1);
    expect(none.style.height).toBe(v.style.height);
  });
});

describe('TD-92-9 代替（ADR-0027 の描き方）', () => {
  it('表に無い prst：文字があれば矩形の文字枠（枠線のある図形は破線の has-outline）、無ければ種類名入りの破線枠', () => {
    const unknown = { type: 'preset' as const, prst: 'noSuchShape', adj: [] };
    const { layer, figs } = mount([
      shape({ geom: unknown, text: TEXT }),
      shape({
        geom: unknown,
        hasOutline: false,
        anchor: { type: 'absolute', x: 0, y: px(200), cx: px(80), cy: px(40) },
      }),
    ]);
    expect(figs).toHaveLength(2);
    expect(figs[0]!.classList.contains('xs-shape-text')).toBe(true);
    expect(figs[0]!.classList.contains('has-outline')).toBe(true);
    expect(figs[0]!.textContent).toBe('<b>文字</b>');
    expect(figs[1]!.querySelector('.xs-frame')?.textContent).toBe('spreadsheet.drawing.shape');
    // 図の svg は無い（下地の svg は図ではない）
    expect(layer.element.querySelector('svg:not(.xs-backdrop)')).toBeNull();
  });

  it('評価に失敗する自由図形（未知の名前）も代替になる', () => {
    const bad = {
      type: 'custom' as const,
      def: {
        av: [],
        gd: [],
        paths: [
          {
            c: [
              ['M', 'noSuchName', 0],
              ['L', 1, 1],
            ] as const,
          },
        ],
      },
    };
    const { layer, figs } = mount([shape({ geom: bad })]);
    expect(figs[0]!.querySelector('.xs-frame')).not.toBeNull();
    expect(layer.element.querySelector('svg')).toBeNull();
  });

  it('形を持たない図形（geom が無い。上限超え）も代替になる', () => {
    const { figs } = mount([shape({ geom: undefined })]);
    expect(figs[0]!.querySelector('.xs-frame')?.textContent).toBe('spreadsheet.drawing.shape');
  });
});

describe('TD-92-10 可視判定：外接矩形とアンカーの矩形を合わせた矩形で出し入れする', () => {
  it('吹き出しの尾が窓に入れば、アンカーが窓の外でも DOM に出る。窓から外れれば外す', () => {
    const d = shape({
      hasFill: true,
      geom: {
        type: 'preset',
        prst: 'wedgeRectCallout',
        adj: [
          ['adj1', -20833],
          ['adj2', 300000],
        ],
      },
    });
    const boxes = layoutDrawings(geo, [d]);
    const layer = makeLayer();
    layer.setSheet({
      drawings: [d],
      boxes,
      readImage: async () => ({ ok: false, reason: 'missing', label: null }),
    });
    // アンカー（y=20〜70）より 100px 以上下の細い窓
    layer.update({ x: 0, y: 170, w: 500, h: 20 });
    expect(layer.element.querySelectorAll('.xs-figure')).toHaveLength(1);
    layer.update({ x: 0, y: 5000, w: 500, h: 20 });
    expect(layer.element.querySelectorAll('.xs-figure')).toHaveLength(0);
  });
});
