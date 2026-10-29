// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createGeometry, type SheetDrawing } from '../../src/core/index';
import { applyFillMode } from '../../src/core/color';
import {
  emptyLayout,
  type DrawingColor,
  type DrawingLine,
  type DrawingText,
} from '../../src/core/types';
import { EMU_PER_PX, layoutDrawings, viewWindow } from '../../src/render/drawings';
import { createImageLayer, type ImageLayer } from '../../src/render/image-layer';

/**
 * 画像レイヤー（DOM）：Excel の色（ISSUE-96・ADR-0028 §決定 6.6）とダーク表示の下地（§決定 6.7）。
 * 色は整数と不透明度から組み立てた `#rrggbb`・数値だけが style に入る。下地は図ごとの独立した `svg.xs-backdrop`（図の直前の兄弟）で、
 * 色は CSS のテーマ変数（`--xs-backdrop`）が決める（要素の style に色を入れない）。実寸・実際の色は目視（TV-42）で確認する。
 */

const px = (n: number): number => n * EMU_PER_PX;
const geo = createGeometry(emptyLayout(), 1000, 1000);
const ALL = viewWindow(0, 0, 3000, 3000, 0, 0);
const RGB = /^rgb\(\d+, \d+, \d+\)$/;

/** jsdom は style の色を rgb() に正規化して返す（入れる値は #rrggbb） */
const rgbOf = (hex: string): string =>
  `rgb(${parseInt(hex.slice(1, 3), 16)}, ${parseInt(hex.slice(3, 5), 16)}, ${parseInt(hex.slice(5, 7), 16)})`;

const c = (rgb: number, alpha = 1): DrawingColor => ({ rgb, alpha });

function makeLayer(): ImageLayer {
  const layer = createImageLayer({ createUrl: () => 'blob:x', revokeUrl: () => undefined });
  document.body.append(layer.element);
  return layer;
}

afterEach(() => {
  document.body.replaceChildren();
});

function runOf(text: string, color?: DrawingColor) {
  return {
    br: false,
    text,
    sizePt: 11,
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    color,
  };
}

function textOf(runs: ReturnType<typeof runOf>[], bullet: string | null = null): DrawingText {
  return {
    paragraphs: [{ runs, align: 'l', bullet, endSizePt: 11 }],
    anchor: 't',
    insets: { l: 91440, t: 45720, r: 91440, b: 45720 },
    noWrap: false,
    fontScale: 1,
    vert: 'horz',
  } as DrawingText;
}

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

/** 線（absolute 配置の 100×0・x=10・y=20。矢じりなし・太さ 2px） */
function line(over: Partial<DrawingLine> = {}, d: Partial<SheetDrawing> = {}): SheetDrawing {
  return {
    anchor: { type: 'absolute', x: px(10), y: px(20), cx: px(100), cy: px(40) },
    kind: 'line',
    name: '線',
    descr: '',
    mediaPath: null,
    line: {
      geom: 'line',
      adj: [50000, 50000, 50000],
      widthEmu: px(2),
      dash: 'solid',
      head: null,
      tail: null,
      ...over,
    },
    ...d,
  } as SheetDrawing;
}

/** 色を持つ図形・線（下地は Excel の色を持つ図にだけ敷く。色の値そのものは下地に関係しない） */
const cshape = (over: Partial<SheetDrawing> = {}): SheetDrawing =>
  shape({ outlineColor: c(0x000000), fillColor: c(0xffffff), ...over });
const cline = (over: Partial<DrawingLine> = {}): SheetDrawing =>
  line({ color: c(0x000000), ...over });

function mount(drawings: SheetDrawing[]): { layer: ImageLayer; kids: Element[] } {
  const layer = makeLayer();
  layer.setSheet({
    drawings,
    boxes: layoutDrawings(geo, drawings),
    readImage: async () => ({ ok: false, reason: 'missing', label: null }),
  });
  layer.update(ALL);
  return { layer, kids: [...layer.element.children] };
}

describe('TD-96-1 図形の色：塗り・枠線・面の濃淡（style の fill・stroke・不透明度）', () => {
  it('塗りの色は fill（#rrggbb）、枠線の色は stroke。不透明度が 1 なら opacity は入れない', () => {
    const { layer } = mount([
      shape({ hasFill: true, fillColor: c(0xffc000), outlineColor: c(0x172c51) }),
    ]);
    const p = layer.element.querySelector('.xs-shape-svg path')!;
    expect(p.getAttribute('style')).toContain(`fill: ${rgbOf('#ffc000')}`);
    expect(p.getAttribute('style')).toContain(`stroke: ${rgbOf('#172c51')}`);
    expect(p.getAttribute('style')).not.toContain('opacity');
  });

  it('不透明度は fill-opacity・stroke-opacity（数値）', () => {
    const { layer } = mount([
      shape({ hasFill: true, fillColor: c(0x00b050, 0.4), outlineColor: c(0xff0000, 0.25) }),
    ]);
    const p = layer.element.querySelector<SVGPathElement>('.xs-shape-svg path')!;
    expect(p.style.getPropertyValue('fill-opacity')).toBe('0.4');
    expect(p.style.getPropertyValue('stroke-opacity')).toBe('0.25');
  });

  it('塗りだけ・枠線だけ：持たない方の色は style に入らない（枠線が無い path に stroke は付けない）', () => {
    const { layer } = mount([
      shape({
        hasOutline: false,
        hasFill: true,
        fillColor: c(0x112233),
        outlineColor: c(0x445566),
      }),
      shape({
        hasOutline: true,
        hasFill: false,
        fillColor: c(0x112233),
        outlineColor: c(0x445566),
      }),
    ]);
    const [a, b] = [...layer.element.querySelectorAll<SVGPathElement>('.xs-shape-svg path')];
    expect(a!.style.getPropertyValue('fill')).toBe(rgbOf('#112233'));
    expect(a!.style.getPropertyValue('stroke')).toBe('');
    expect(b!.style.getPropertyValue('stroke')).toBe(rgbOf('#445566'));
    expect(b!.style.getPropertyValue('fill')).toBe('');
  });

  it('面の濃淡：path の fill に応じて、塗りの色に lib の applyFillMode を掛けた色（4 種）', () => {
    const base = c(0x4472c4, 0.8);
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
    const modes = ['darken', 'darkenLess', 'lighten', 'lightenLess', 'norm'];
    const { layer } = mount(
      modes.map((m) => shape({ hasFill: true, hasOutline: false, geom: def(m), fillColor: base })),
    );
    const paths = [...layer.element.querySelectorAll<SVGPathElement>('.xs-shape-svg path')];
    const hex = (n: number) => `#${n.toString(16).padStart(6, '0')}`;
    expect(paths.map((p) => p.style.getPropertyValue('fill'))).toEqual(
      modes.map((m) => rgbOf(hex(applyFillMode(base, m).rgb))),
    );
    // 濃淡は色を変える（norm は塗りの色そのもの）
    expect(new Set(paths.map((p) => p.style.getPropertyValue('fill'))).size).toBe(5);
    expect(paths[4]!.style.getPropertyValue('fill')).toBe(rgbOf('#4472c4'));
    // 不透明度は濃淡で変えない
    expect(paths.every((p) => p.style.getPropertyValue('fill-opacity') === '0.8')).toBe(true);
  });

  it('色が無い（省略）図形は style に色を入れない（暫定のテーマ変数のまま）', () => {
    const { layer } = mount([shape({ hasFill: true })]);
    const p = layer.element.querySelector('.xs-shape-svg path')!;
    expect(p.getAttribute('style')).toBeNull();
    expect(p.classList.contains('has-fill')).toBe(true);
  });

  it('不正な色（整数でない・範囲外・不透明度が有限でない）は入れない。入れる値は #rrggbb と数値だけ', () => {
    const { layer } = mount([
      shape({ hasFill: true, fillColor: c(1.5), outlineColor: c(0x1000000) }),
      shape({ hasFill: true, fillColor: c(Number.NaN), outlineColor: c(0x123456, Number.NaN) }),
      shape({ hasFill: true, fillColor: c(-1), outlineColor: c(0x123456, 7) }),
    ]);
    const paths = [...layer.element.querySelectorAll<SVGPathElement>('.xs-shape-svg path')];
    expect(paths[0]!.getAttribute('style')).toBeNull();
    expect(paths[1]!.getAttribute('style')).toBeNull();
    // 不透明度が範囲外（7）は 1 に収める＝opacity は付けない
    expect(paths[2]!.style.getPropertyValue('stroke')).toBe(rgbOf('#123456'));
    expect(paths[2]!.style.getPropertyValue('stroke-opacity')).toBe('');
    for (const p of paths) {
      for (const k of ['fill', 'stroke']) {
        const v = p.style.getPropertyValue(k);
        expect(v === '' || RGB.test(v)).toBe(true);
      }
    }
  });
});

describe('TD-96-2 線・矢じりの色', () => {
  it('線は stroke、矢じりは塗り・線とも同じ色（不透明度も）。色が無ければ今の currentColor のまま', () => {
    const arrows = {
      head: { type: 'triangle', w: 'med', len: 'med' } as const,
      tail: { type: 'arrow', w: 'med', len: 'med' } as const,
    };
    const { layer } = mount([line({ color: c(0xc00000, 0.5), ...arrows }), line({ ...arrows })]);
    const svgs = [...layer.element.querySelectorAll<SVGSVGElement>('svg.xs-line')];
    const [colored, plain] = svgs.map((s) => [...s.querySelectorAll('path')]);
    expect(colored).toHaveLength(3);
    // 本体・塗りの矢じり（fill）・開いた V（stroke）
    expect(colored![0]!.style.getPropertyValue('stroke')).toBe(rgbOf('#c00000'));
    expect(colored![0]!.style.getPropertyValue('stroke-opacity')).toBe('0.5');
    expect(colored![1]!.style.getPropertyValue('fill')).toBe(rgbOf('#c00000'));
    expect(colored![1]!.style.getPropertyValue('fill-opacity')).toBe('0.5');
    expect(colored![2]!.style.getPropertyValue('stroke')).toBe(rgbOf('#c00000'));
    // 色なし：style は無く、属性の currentColor のまま
    expect(plain![0]!.getAttribute('style')).toBeNull();
    expect(plain![0]!.getAttribute('stroke')).toBe('currentColor');
  });

  it('楕円の矢じり（oval）も塗りの色', () => {
    const { layer } = mount([
      line({ color: c(0x0070c0), tail: { type: 'oval', w: 'med', len: 'med' } }),
    ]);
    const e = layer.element.querySelector('svg.xs-line ellipse') as SVGEllipseElement;
    expect(e.style.getPropertyValue('fill')).toBe(rgbOf('#0070c0'));
  });
});

describe('TD-96-3 文字の色・代替の文字枠の色', () => {
  it('ランの color：不透明なら #rrggbb、不透明度があれば #rrggbbaa、noFill（不透明度 0）は透明', () => {
    const { layer } = mount([
      shape({
        text: textOf([
          runOf('赤', c(0xff0000)),
          runOf('半透明', c(0x00ff00, 0.5)),
          runOf('透明', c(0x0000ff, 0)),
          runOf('色なし'),
        ]),
      }),
    ]);
    const spans = [...layer.element.querySelectorAll<HTMLElement>('.xs-shape-text span')];
    expect(spans.map((s) => s.style.color)).toEqual([
      'rgb(255, 0, 0)',
      'rgba(0, 255, 0, 0.5)',
      'rgba(0, 0, 255, 0)',
      '',
    ]);
  });

  it('箇条書きの記号は先頭のランの色', () => {
    const { layer } = mount([shape({ text: textOf([runOf('本文', c(0x123456))], '•') })]);
    const spans = [...layer.element.querySelectorAll<HTMLElement>('.xs-shape-text span')];
    expect(spans[0]!.textContent).toBe('• ');
    expect(spans[0]!.style.color).toBe('rgb(18, 52, 86)');
  });

  it('代替の矩形の文字枠：背景は塗りの色、枠線は枠線の色の破線、文字は文字の色', () => {
    const unknown = { type: 'preset' as const, prst: 'noSuchShape', adj: [] };
    const { kids } = mount([
      shape({
        geom: unknown,
        hasFill: true,
        fillColor: c(0xffff00, 0.5),
        outlineColor: c(0x7030a0),
        text: textOf([runOf('字', c(0x222222))]),
      }),
    ]);
    const fig = kids.find((k) => k.classList.contains('xs-shape-text')) as HTMLElement;
    expect(fig.classList.contains('has-outline')).toBe(true);
    expect(fig.style.backgroundColor).toBe('rgba(255, 255, 0, 0.5)');
    expect(fig.style.outlineColor).toBe('rgb(112, 48, 160)');
    expect((fig.querySelector('span') as HTMLElement).style.color).toBe('rgb(34, 34, 34)');
  });

  it('種類名の破線枠と画像は、図形の色を持たない（テーマ変数のまま）', () => {
    const { layer } = mount([
      shape({
        geom: { type: 'preset', prst: 'noSuchShape', adj: [] },
        hasOutline: false,
        fillColor: c(0xff0000),
      }),
    ]);
    const frame = layer.element.querySelector<HTMLElement>('.xs-frame')!;
    expect(frame.getAttribute('style')).toBeNull();
  });
});

/** 下地の svg（`svg.xs-backdrop`）と、その直後の図の要素の組 */
function pairs(layer: ImageLayer): { back: SVGSVGElement; fig: Element }[] {
  const out: { back: SVGSVGElement; fig: Element }[] = [];
  const kids = [...layer.element.children];
  kids.forEach((k, i) => {
    if (k.classList.contains('xs-backdrop')) {
      out.push({ back: k as SVGSVGElement, fig: kids[i + 1]! });
    }
  });
  return out;
}

describe('TD-96-4 下地：独立した svg を図の直前に置く（図の形・線の縁取り）', () => {
  it('塗りのある図形：図の直前に svg.xs-backdrop。図と同じ位置から片側 3px 広げる（viewBox は図の座標のまま）', () => {
    const { layer, kids } = mount([cshape({ hasFill: true })]);
    // 枠線 1px・矩形：図は (8,18,104×54)
    expect(kids.map((k) => k.getAttribute('class'))).toEqual(['xs-backdrop', 'xs-figure']);
    const [{ back }] = pairs(layer);
    expect(back.namespaceURI).toBe('http://www.w3.org/2000/svg');
    expect([back.style.left, back.style.top, back.style.width, back.style.height]).toEqual([
      '5px',
      '15px',
      '110px',
      '60px',
    ]);
    expect(back.getAttribute('viewBox')).toBe('-3 -3 110 60');
    expect(back.getAttribute('aria-hidden')).toBe('true');
    expect(back.getAttribute('focusable')).toBe('false');
    // 図形の svg と同じ形：塗り（is-fill）と縁取り（is-stroke。太さ＝線の太さ 1 ＋ 6）
    const paths = [...back.querySelectorAll('path')];
    expect(paths).toHaveLength(1);
    const figPath = layer.element.querySelector('.xs-shape-svg path')!;
    expect(paths[0]!.getAttribute('d')).toBe(figPath.getAttribute('d'));
    expect(paths[0]!.classList.contains('is-fill')).toBe(true);
    expect(paths[0]!.classList.contains('is-stroke')).toBe(true);
    expect(paths[0]!.getAttribute('stroke-width')).toBe('7');
    expect(paths[0]!.getAttribute('stroke-linejoin')).toBe('round');
    expect(paths[0]!.getAttribute('stroke-linecap')).toBe('round');
    expect(paths[0]!.hasAttribute('stroke-dasharray')).toBe(false);
  });

  it('下地の色はテーマ変数（CSS）：要素の style に色・不透明度・fill・stroke を入れない（描き直さずテーマが切り替わる）', () => {
    const { layer } = mount([
      cshape({ hasFill: true, fillColor: c(0xff0000), outlineColor: c(0x00ff00) }),
      cline({ color: c(0x0000ff), tail: { type: 'triangle', w: 'med', len: 'med' } }),
      cshape({
        geom: { type: 'preset', prst: 'noSuchShape', adj: [] },
        text: textOf([runOf('字', c(0x123456))]),
        hasFill: true,
        fillColor: c(0x654321),
      }),
    ]);
    const backs = [...layer.element.querySelectorAll('.xs-backdrop')];
    expect(backs).toHaveLength(3);
    for (const b of backs) {
      for (const el of [b, ...b.querySelectorAll('*')]) {
        const s = (el as HTMLElement).getAttribute('style') ?? '';
        expect(s).not.toMatch(/color|fill|stroke|opacity|background/);
        expect(el.getAttribute('fill')).toBeNull();
        expect(el.getAttribute('stroke')).toBeNull();
        expect(el.getAttribute('fill-opacity')).toBeNull();
      }
    }
  });

  it('塗りも文字も無い図形（枠線だけ）：縁取りだけで内側は塗らない（is-fill が付かない）', () => {
    const { layer } = mount([cshape({ outlineColor: c(0xff0000) })]);
    const [{ back }] = pairs(layer);
    const p = back.querySelector('path')!;
    expect(p.classList.contains('is-stroke')).toBe(true);
    expect(p.classList.contains('is-fill')).toBe(false);
    expect(p.getAttribute('stroke-width')).toBe('7');
  });

  it('枠線が破線でも縁取りは実線（stroke-dasharray を付けない）。太い線は太さ＋6', () => {
    const { layer } = mount([cshape({ outline: { widthEmu: px(4), dash: 'dash' } })]);
    const [{ back }] = pairs(layer);
    const p = back.querySelector('path')!;
    expect(p.getAttribute('stroke-width')).toBe('10');
    expect(p.hasAttribute('stroke-dasharray')).toBe(false);
  });

  it('塗りが無く文字のある図形（枠線なし）：形の塗りの path を塗る。文字だけの図形にも下地が付く', () => {
    const { layer } = mount([cshape({ hasOutline: false, text: textOf([runOf('ラベル')]) })]);
    const [{ back }] = pairs(layer);
    const p = back.querySelector('path')!;
    expect(p.classList.contains('is-fill')).toBe(true);
    expect(p.classList.contains('is-stroke')).toBe(false);
    expect(p.hasAttribute('stroke-width')).toBe(false);
  });

  it('見えない図形（枠も塗りも文字も無い）・種類名の破線枠・画像・グラフには下地が無い', () => {
    const { layer, kids } = mount([
      cshape({ hasOutline: false, hasFill: false }),
      cshape({ geom: { type: 'preset', prst: 'noSuchShape', adj: [] }, hasOutline: false }),
      {
        anchor: { type: 'absolute', x: px(0), y: px(200), cx: px(40), cy: px(30) },
        kind: 'chart',
        name: 'c',
        descr: '',
        mediaPath: null,
      },
      {
        anchor: { type: 'absolute', x: px(0), y: px(300), cx: px(40), cy: px(30) },
        kind: 'image',
        name: 'i',
        descr: '',
        mediaPath: 'xl/media/a.png',
      },
    ]);
    expect(layer.element.querySelectorAll('.xs-backdrop')).toHaveLength(0);
    expect(kids.every((k) => k.classList.contains('xs-figure'))).toBe(true);
  });

  it('線：外接矩形から片側 3px 広げた svg に、線の縁取り（太さ＋6・実線・丸）。矢じりは形を塗り、同じ縁取りを付ける', () => {
    const { layer } = mount([
      cline({
        dash: 'dash',
        head: { type: 'triangle', w: 'med', len: 'med' },
        tail: { type: 'arrow', w: 'med', len: 'med' },
      }),
    ]);
    const [{ back, fig }] = pairs(layer);
    const f = fig as SVGSVGElement;
    expect(fig.classList.contains('xs-line')).toBe(true);
    const fl = Number.parseFloat(f.style.left);
    const ft = Number.parseFloat(f.style.top);
    const fw = Number.parseFloat(f.style.width);
    const fh = Number.parseFloat(f.style.height);
    expect([back.style.left, back.style.top, back.style.width, back.style.height]).toEqual([
      `${fl - 3}px`,
      `${ft - 3}px`,
      `${fw + 6}px`,
      `${fh + 6}px`,
    ]);
    expect(back.getAttribute('viewBox')).toBe(`-3 -3 ${fw + 6} ${fh + 6}`);
    const paths = [...back.querySelectorAll('path')];
    // 本体・塗りの矢じり・開いた V
    expect(paths).toHaveLength(3);
    expect(
      paths.map((p) => [p.classList.contains('is-fill'), p.classList.contains('is-stroke')]),
    ).toEqual([
      [false, true],
      [true, true],
      [false, true],
    ]);
    for (const p of paths) {
      expect(p.getAttribute('stroke-width')).toBe('8'); // 太さ 2px ＋ 6
      expect(p.getAttribute('stroke-linejoin')).toBe('round');
      expect(p.getAttribute('stroke-linecap')).toBe('round');
      expect(p.hasAttribute('stroke-dasharray')).toBe(false);
    }
    // 本体の d は図の線と同じ（縮めた経路）
    const figPaths = [...f.querySelectorAll('path')];
    expect(paths.map((p) => p.getAttribute('d'))).toEqual(figPaths.map((p) => p.getAttribute('d')));
  });

  it('楕円の矢じり：ellipse を塗り、同じ縁取りを付ける', () => {
    const { layer } = mount([cline({ tail: { type: 'oval', w: 'med', len: 'med' } })]);
    const [{ back }] = pairs(layer);
    const e = back.querySelector('ellipse')!;
    expect(e.classList.contains('is-fill')).toBe(true);
    expect(e.classList.contains('is-stroke')).toBe(true);
    expect(e.getAttribute('stroke-width')).toBe('8');
    expect(e.getAttribute('transform')).toMatch(/^rotate\(/);
  });

  it('代替の矩形の文字枠：回転前の矩形を回した形（svg に rotate()）を塗る。グループの中は縦横を入れ替えない', () => {
    const unknown = { type: 'preset' as const, prst: 'noSuchShape', adj: [] };
    const { layer } = mount([
      cshape({
        geom: unknown,
        text: textOf([runOf('字')]),
        anchor: { type: 'absolute', x: px(35), y: px(20), cx: px(50), cy: px(100) },
        rot: 90 * 60000,
      }),
    ]);
    const [{ back, fig }] = pairs(layer);
    const f = fig as HTMLElement;
    // 回転前の矩形 100×50（図の要素と同じ位置・大きさ）。余白は付けない
    expect([back.style.left, back.style.top, back.style.width, back.style.height]).toEqual([
      f.style.left,
      f.style.top,
      f.style.width,
      f.style.height,
    ]);
    expect(back.style.transform).toBe('rotate(90deg)');
    const r = back.querySelector('rect')!;
    expect(r.classList.contains('is-fill')).toBe(true);
    expect([r.getAttribute('width'), r.getAttribute('height')]).toEqual(['100', '50']);
  });
});

describe('TD-96-5 下地の出し入れ：図と一緒に DOM に出し入れする。重ね順は CSS（z-index: -1・isolation）', () => {
  it('窓から外れた図は下地ごと外し、戻れば下地ごと出す。図が無い下地は残らない', () => {
    const d = [
      cshape({ hasFill: true }),
      cshape({
        hasFill: true,
        anchor: { type: 'absolute', x: px(10), y: px(900), cx: px(100), cy: px(50) },
      }),
    ];
    const layer = makeLayer();
    layer.setSheet({
      drawings: d,
      boxes: layoutDrawings(geo, d),
      readImage: async () => ({ ok: false, reason: 'missing', label: null }),
    });
    layer.update({ x: 0, y: 0, w: 500, h: 200 });
    expect(layer.element.querySelectorAll('.xs-backdrop')).toHaveLength(1);
    expect(layer.element.querySelectorAll('.xs-figure')).toHaveLength(1);
    const first = layer.element.querySelector('.xs-backdrop');
    layer.update({ x: 0, y: 800, w: 500, h: 200 });
    expect(layer.element.querySelectorAll('.xs-backdrop')).toHaveLength(1);
    expect(layer.element.querySelector('.xs-backdrop')).not.toBe(first);
    layer.update({ x: 0, y: 5000, w: 500, h: 20 });
    expect(layer.element.children).toHaveLength(0);
    layer.update(ALL);
    expect(layer.element.querySelectorAll('.xs-backdrop')).toHaveLength(2);
    // 図の順（drawing の順）に [下地, 図, 下地, 図]
    expect([...layer.element.children].map((k) => k.getAttribute('class'))).toEqual([
      'xs-backdrop',
      'xs-figure',
      'xs-backdrop',
      'xs-figure',
    ]);
  });

  it('図の要素（div.xs-figure・線の svg）に transform・opacity・z-index を付けない（重ね順の文脈を作らない）', () => {
    const { kids } = mount([
      cshape({ hasFill: true, rot: 45 * 60000, text: textOf([runOf('字')]) }),
      cline({ tail: { type: 'triangle', w: 'med', len: 'med' } }),
    ]);
    for (const k of kids.filter((x) => x.classList.contains('xs-figure'))) {
      const s = (k as HTMLElement).style;
      expect([s.transform, s.opacity, s.zIndex]).toEqual(['', '', '']);
    }
  });
});

describe('TD-96-7 下地の条件：見える塗り・見える文字だけ。色を省略した図には敷かない', () => {
  const unknown = { type: 'preset' as const, prst: 'noSuchShape', adj: [] };
  const hasBackdrop = (d: SheetDrawing): boolean =>
    mount([d]).layer.element.querySelector('.xs-backdrop') !== null;
  const fills = (d: SheetDrawing): boolean[] =>
    [...mount([d]).layer.element.querySelectorAll('.xs-backdrop path')].map((p) =>
      p.classList.contains('is-fill'),
    );

  it('塗りの不透明度が 0 の図形は塗らない（枠線もなければ下地なし。枠線があれば縁取りだけ）', () => {
    expect(
      hasBackdrop(cshape({ hasFill: true, hasOutline: false, fillColor: c(0xff0000, 0) })),
    ).toBe(false);
    expect(fills(cshape({ hasFill: true, hasOutline: true, fillColor: c(0xff0000, 0) }))).toEqual([
      false,
    ]);
    expect(
      fills(cshape({ hasFill: true, hasOutline: false, fillColor: c(0xff0000, 0.2) })),
    ).toEqual([true]);
  });

  it('文字の色が noFill（不透明度 0）の文字だけの図形は塗らない。見えるランが 1 つでもあれば塗る', () => {
    const none = textOf([runOf('透明', c(0, 0))]);
    expect(hasBackdrop(cshape({ hasOutline: false, hasFill: false, text: none }))).toBe(false);
    const some = textOf([runOf('透明', c(0, 0)), runOf('見える', c(0x333333))]);
    expect(fills(cshape({ hasOutline: false, hasFill: false, text: some }))).toEqual([true]);
    // 塗りが見えれば、文字が透明でも塗りの形を塗る
    expect(fills(cshape({ hasOutline: false, hasFill: true, text: none }))).toEqual([true]);
  });

  it('代替の文字枠：塗りが無い（または不透明度 0）で文字も透明なら下地なし。見える文字か見える塗りがあれば敷く', () => {
    const none = textOf([runOf('透明', c(0, 0))]);
    expect(hasBackdrop(cshape({ geom: unknown, hasFill: false, text: none }))).toBe(false);
    expect(
      hasBackdrop(cshape({ geom: unknown, hasFill: true, fillColor: c(0xffff00, 0), text: none })),
    ).toBe(false);
    expect(
      hasBackdrop(
        cshape({ geom: unknown, hasFill: false, text: textOf([runOf('字', c(0x111111))]) }),
      ),
    ).toBe(true);
    expect(hasBackdrop(cshape({ geom: unknown, hasFill: true, text: none }))).toBe(true);
  });

  it('色を省略した図（塗り・枠線・線・文字のどの色も無い）は下地を敷かない。どれか 1 つでもあれば敷く', () => {
    // 形・線・代替の文字枠のいずれも、暫定のテーマ変数の色で描かれるので白い下地に乗せない
    expect(hasBackdrop(shape({ hasFill: true }))).toBe(false);
    expect(hasBackdrop(shape({ hasFill: true, text: textOf([runOf('字')]) }))).toBe(false);
    expect(hasBackdrop(line({ tail: { type: 'triangle', w: 'med', len: 'med' } }))).toBe(false);
    expect(hasBackdrop(shape({ geom: unknown, text: textOf([runOf('字')]) }))).toBe(false);
    // 色が 1 つでもあれば敷く（線の color・塗りの色・ランの色）
    expect(hasBackdrop(line({ color: c(0x123456) }))).toBe(true);
    expect(hasBackdrop(shape({ hasFill: true, fillColor: c(0x123456) }))).toBe(true);
    expect(
      hasBackdrop(shape({ hasOutline: false, text: textOf([runOf('字', c(0x123456))]) })),
    ).toBe(true);
  });
});
