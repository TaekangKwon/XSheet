// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { buildChartSvg } from '../../src/render/chart-svg';
import {
  layoutChart,
  type ChartItem,
  type ChartScene,
  type ChartStroke,
} from '../../src/render/chart-layout';
import { barModel, series } from './chart-fixtures';

/**
 * グラフの場面 → `<svg>`（ISSUE-103・ADR-0031 §決定 3・8・10）。jsdom は描画を持たないため、要素・属性・文字の入り方を確かめる。
 */

const stroke = (over: Partial<ChartStroke> = {}): ChartStroke => ({
  color: '#112233',
  opacity: 1,
  width: 2,
  dash: 'solid',
  cap: 'butt',
  join: 'miter',
  ...over,
});

const sceneOf = (items: ChartItem[], over: Partial<ChartScene> = {}): ChartScene => ({
  width: 300,
  height: 200,
  backdrop: false,
  items,
  ...over,
});

const build = (s: ChartScene): SVGSVGElement => buildChartSvg(s, document);
const build2 = (s: ChartScene, title: string): SVGSVGElement => buildChartSvg(s, document, title);

/** 属性に入ってよい文字（数値・`#rrggbb`・固定の語・`d`・`rotate(...)`・`0 0 w h` の空白）。`<`・`"`・`'`・`&`・`;` を含まない */
const SAFE_ATTR = /^[-0-9A-Za-z#.,+\s()]*$/;

describe('TD-103-1 場面の各項目が要素になる', () => {
  it('svg の大きさ・viewBox・クラス・aria-hidden。items が無ければ子は無い', () => {
    const svg = build(sceneOf([]));
    expect(svg.namespaceURI).toBe('http://www.w3.org/2000/svg');
    expect(svg.getAttribute('width')).toBe('300');
    expect(svg.getAttribute('height')).toBe('200');
    expect(svg.getAttribute('viewBox')).toBe('0 0 300 200');
    expect(svg.getAttribute('class')).toBe('xs-chart');
    expect(svg.getAttribute('aria-hidden')).toBe('true');
    expect(svg.children).toHaveLength(0);
  });

  it('rect・path・line・text が順に要素になる（塗り・線・不透明度・太さ）', () => {
    const svg = build(
      sceneOf([
        {
          kind: 'rect',
          x: 1,
          y: 2,
          width: 30,
          height: 40,
          fill: { color: '#ffffff', opacity: 1 },
          stroke: stroke({ color: '#d9d9d9', width: 1 }),
        },
        {
          kind: 'path',
          d: 'M 0 0 L 10 0 L 10 10 Z',
          fill: { color: '#4472c4', opacity: 0.5 },
          stroke: null,
        },
        { kind: 'line', x1: 0, y1: 5, x2: 50, y2: 5, stroke: stroke({ opacity: 0.25 }) },
        {
          kind: 'text',
          x: 10,
          y: 20,
          text: 'ラベル',
          size: 12,
          bold: true,
          color: { color: '#595959', opacity: 1 },
          anchor: 'end',
          baseline: 'middle',
          rotate: 0,
        },
      ]),
    );
    expect([...svg.children].map((e) => e.tagName)).toEqual(['rect', 'path', 'line', 'text']);
    const [rect, path, line, text] = [...svg.children] as SVGElement[];
    expect(rect!.getAttribute('x')).toBe('1');
    expect(rect!.getAttribute('width')).toBe('30');
    expect(rect!.getAttribute('fill')).toBe('#ffffff');
    expect(rect!.hasAttribute('fill-opacity')).toBe(false);
    expect(rect!.getAttribute('stroke')).toBe('#d9d9d9');
    expect(rect!.getAttribute('stroke-width')).toBe('1');
    expect(path!.getAttribute('d')).toBe('M 0 0 L 10 0 L 10 10 Z');
    expect(path!.getAttribute('fill')).toBe('#4472c4');
    expect(path!.getAttribute('fill-opacity')).toBe('0.5');
    expect(path!.getAttribute('stroke')).toBe('none');
    expect(line!.getAttribute('x2')).toBe('50');
    expect(line!.getAttribute('stroke')).toBe('#112233');
    expect(line!.getAttribute('stroke-opacity')).toBe('0.25');
    expect(line!.getAttribute('fill')).toBe('none');
    expect(text!.textContent).toBe('ラベル');
    expect(text!.getAttribute('font-size')).toBe('12');
    expect(text!.getAttribute('font-weight')).toBe('bold');
    expect(text!.getAttribute('text-anchor')).toBe('end');
    expect(text!.getAttribute('dominant-baseline')).toBe('central');
    expect(text!.getAttribute('fill')).toBe('#595959');
    expect(text!.hasAttribute('transform')).toBe(false);
  });

  it('文字の揃え 3 種・縦位置 3 種は固定の語に写る。太字でなければ font-weight を書かない', () => {
    const t = (anchor: 'start' | 'middle' | 'end', baseline: 'alphabetic' | 'middle' | 'hanging') =>
      ({
        kind: 'text',
        x: 0,
        y: 0,
        text: 'x',
        size: 9,
        bold: false,
        color: { color: '#000000', opacity: 1 },
        anchor,
        baseline,
        rotate: 0,
      }) as const;
    const svg = build(
      sceneOf([t('start', 'alphabetic'), t('middle', 'hanging'), t('end', 'middle')]),
    );
    const els = [...svg.querySelectorAll('text')];
    expect(els.map((e) => e.getAttribute('text-anchor'))).toEqual(['start', 'middle', 'end']);
    expect(els.map((e) => e.getAttribute('dominant-baseline'))).toEqual([
      'alphabetic',
      'hanging',
      'central',
    ]);
    expect(els.some((e) => e.hasAttribute('font-weight'))).toBe(false);
  });

  it('回転は (x, y) を中心にした transform の数値だけ（0 は書かない）', () => {
    const svg = build(
      sceneOf([
        {
          kind: 'text',
          x: 12.5,
          y: 80,
          text: '縦',
          size: 10,
          bold: false,
          color: { color: '#000000', opacity: 1 },
          anchor: 'middle',
          baseline: 'middle',
          rotate: -90,
        },
      ]),
    );
    expect(svg.querySelector('text')!.getAttribute('transform')).toBe('rotate(-90 12.5 80)');
  });

  it('線種は固定の表から stroke-dasharray になる（実線は書かない・点線は [w, 2w]・破線は [4w, 3w]）', () => {
    const lines = (['solid', 'dot', 'dash'] as const).map((dash): ChartItem => ({
      kind: 'line',
      x1: 0,
      y1: 0,
      x2: 10,
      y2: 0,
      stroke: stroke({ dash, width: 2 }),
    }));
    const els = [...build(sceneOf(lines)).querySelectorAll('line')];
    expect(els[0]!.hasAttribute('stroke-dasharray')).toBe(false);
    expect(els[1]!.getAttribute('stroke-dasharray')).toBe('2 4');
    expect(els[2]!.getAttribute('stroke-dasharray')).toBe('8 6');
  });

  it('線の端・結びは固定の語（丸・miter）', () => {
    const svg = build(
      sceneOf([
        {
          kind: 'line',
          x1: 0,
          y1: 0,
          x2: 1,
          y2: 1,
          stroke: stroke({ cap: 'round', join: 'round' }),
        },
      ]),
    );
    const l = svg.querySelector('line')!;
    expect(l.getAttribute('stroke-linecap')).toBe('round');
    expect(l.getAttribute('stroke-linejoin')).toBe('round');
  });
});

describe('TD-103-2 値が不正な項目は描かない・属性は数値と固定の語だけ', () => {
  it('不正な色・d に想定外の文字・文字の色が不正な項目は要素にならない。有限でない数は 0', () => {
    const svg = build(
      sceneOf([
        {
          kind: 'path',
          d: 'M 0 0 L <script>',
          fill: { color: '#112233', opacity: 1 },
          stroke: null,
        },
        {
          kind: 'path',
          d: 'M 0 0 L 1 1',
          fill: { color: 'red" onload="x', opacity: 1 },
          stroke: null,
        },
        {
          kind: 'text',
          x: 0,
          y: 0,
          text: 'x',
          size: 9,
          bold: false,
          color: { color: 'url(#a)', opacity: 1 },
          anchor: 'start',
          baseline: 'middle',
          rotate: 0,
        },
        { kind: 'rect', x: Number.NaN, y: 0, width: Infinity, height: 1, fill: null, stroke: null },
      ]),
    );
    const kids = [...svg.children];
    expect(kids.map((e) => e.tagName)).toEqual(['path', 'rect']); // 2 つ目の path は色が不正なので fill none で残る
    expect(kids[0]!.getAttribute('d')).toBe('M 0 0 L 1 1');
    expect(kids[0]!.getAttribute('fill')).toBe('none');
    expect(kids[1]!.getAttribute('x')).toBe('0');
    expect(kids[1]!.getAttribute('width')).toBe('0');
  });

  it('不透明度は 0〜1 に収める', () => {
    const svg = build(
      sceneOf([
        { kind: 'path', d: 'M 0 0', fill: { color: '#000000', opacity: 7 }, stroke: null },
        { kind: 'path', d: 'M 0 0', fill: { color: '#000000', opacity: -1 }, stroke: null },
      ]),
    );
    const [a, b] = [...svg.children];
    expect(a!.hasAttribute('fill-opacity')).toBe(false);
    expect(b!.getAttribute('fill-opacity')).toBe('0');
  });
});

describe('TD-103-3 ファイル由来の文字列は text の textContent にだけ入る', () => {
  const EVIL = '<script>alert(1)</script>"onload="x" \'a\' &amp; </text><rect id="inj"/>';

  it('系列名・項目名・タイトル・軸タイトルに危険な文字列を入れても、要素・属性・style にならない', () => {
    const svg = build(
      layoutChart(
        barModel({
          categories: [EVIL, `${EVIL}2`, 'C'],
          series: [series(0, EVIL, [10, 20, 30]), series(1, `${EVIL}-s2`, [5, 15, 25])],
          title: { lines: [EVIL], text: {}, overlay: false },
          legend: { pos: 'r', overlay: false, deleted: [], text: {} },
          catAxis: {
            kind: 'cat',
            deleted: false,
            orientation: 'minMax',
            text: {},
            title: { lines: [EVIL], text: {}, overlay: false },
          },
        }),
        480,
        320,
      )!,
    );
    // 要素は固定の 4 種だけ（script・injected の rect 以外が無い）
    const tags = new Set([...svg.querySelectorAll('*')].map((e) => e.tagName));
    expect([...tags].every((t) => ['rect', 'path', 'line', 'text'].includes(t))).toBe(true);
    expect(svg.querySelector('script')).toBeNull();
    expect(svg.querySelector('#inj, [id]')).toBeNull();
    // 全属性の値が安全な文字だけ。onload・id・style・font-family などの属性が無い
    for (const e of [svg, ...svg.querySelectorAll('*')]) {
      for (const a of e.attributes) {
        expect(a.name).not.toMatch(/^(on|id$|style$|font-family|href|xlink)/);
        expect(a.value, `${e.tagName}@${a.name}`).toMatch(SAFE_ATTR);
      }
    }
    // 文字は text の中身として（エスケープされずに文字列のまま）入る
    const contents = [...svg.querySelectorAll('text')].map((t) => t.textContent);
    expect(contents).toContain(EVIL);
    expect(contents.filter((c) => c?.includes('<script>')).length).toBeGreaterThanOrEqual(3);
    // text 要素の子は文字（テキストノード）だけ
    for (const t of svg.querySelectorAll('text')) {
      expect(t.children).toHaveLength(0);
    }
  });

  it('font-family も style も書かない（画面の文字を継承する）', () => {
    const svg = build(
      layoutChart(barModel({ title: { lines: ['T'], text: {}, overlay: false } }), 400, 300)!,
    );
    expect(svg.outerHTML).not.toMatch(/font-family|style=/);
  });
});

describe('TD-103-4 下地（backdrop）はグラフの svg の中に置かない（画像レイヤーが兄弟の下地にする）', () => {
  it('backdrop が真でも偽でも、svg の中身は場面の項目だけ（下地の矩形・クラスを足さない）', () => {
    const items: ChartItem[] = [{ kind: 'line', x1: 0, y1: 0, x2: 1, y2: 1, stroke: stroke() }];
    const on = build(sceneOf(items, { backdrop: true }));
    const off = build(sceneOf(items, { backdrop: false }));
    expect(on.children).toHaveLength(1);
    expect(on.firstElementChild!.tagName).toBe('line');
    expect(on.querySelector('rect')).toBeNull();
    expect(on.querySelector('[class*="backdrop"]')).toBeNull();
    expect(on.innerHTML).toBe(off.innerHTML);
  });

  it('title を渡すと role="img"＋最初の子に <title>（textContent だけ）・aria-hidden なし。渡さなければ aria-hidden', () => {
    const named = build2(sceneOf([]), '<b>グラフ</b>"x');
    expect(named.getAttribute('role')).toBe('img');
    expect(named.hasAttribute('aria-hidden')).toBe(false);
    const t = named.firstElementChild!;
    expect(t.tagName).toBe('title');
    expect(t.textContent).toBe('<b>グラフ</b>"x');
    expect(t.children).toHaveLength(0);
    for (const a of named.attributes) expect(a.value).toMatch(SAFE_ATTR);
    const plain = build(sceneOf([]));
    expect(plain.getAttribute('aria-hidden')).toBe('true');
    expect(plain.querySelector('title')).toBeNull();
  });

  it('場面の backdrop はグラフの面が塗りなしの時に真（layoutChart と通しで）', () => {
    const none = layoutChart(barModel({ chartArea: { fill: { kind: 'none' } } }), 400, 300)!;
    const white = layoutChart(
      barModel({ chartArea: { fill: { kind: 'solid', color: { rgb: 0xffffff, alpha: 1 } } } }),
      400,
      300,
    )!;
    expect(none.backdrop).toBe(true);
    expect(white.backdrop).toBe(false);
    // 場面の真偽によらず、svg の中に下地は無い
    expect(build(none).querySelector('.xs-chart-backdrop')).toBeNull();
    expect(build(white).querySelector('.xs-chart-backdrop')).toBeNull();
  });
});

describe('TD-103-5 layoutChart の場面を通しで SVG にする', () => {
  it('縦棒：系列の棒の path・目盛の文字・軸の線が出る。全項目が要素になる', () => {
    const scene = layoutChart(barModel(), 400, 300)!;
    const svg = build(scene);
    expect(svg.children.length).toBe(scene.items.length);
    expect(svg.querySelectorAll('path').length).toBeGreaterThan(0);
    const labels = [...svg.querySelectorAll('text')].map((t) => t.textContent);
    expect(labels).toEqual(expect.arrayContaining(['A', 'B', 'C']));
  });
});

describe('TD-109-5 面の角丸（rect の rx・ry。ISSUE-109）', () => {
  const rectItem = (radius?: number): ChartItem => ({
    kind: 'rect',
    x: 0,
    y: 0,
    width: 300,
    height: 200,
    fill: { color: '#ffffff', opacity: 1 },
    stroke: null,
    ...(radius === undefined ? {} : { radius }),
  });

  it('radius が正の有限値なら rx・ry（数値）が付く。無い・0・負・有限でないなら付けない', () => {
    const e = build(sceneOf([rectItem(11)])).children[0]!;
    expect([e.getAttribute('rx'), e.getAttribute('ry')]).toEqual(['11', '11']);
    for (const r of [undefined, 0, -3, NaN, Infinity]) {
      const x = build(sceneOf([rectItem(r)])).children[0]!;
      expect(x.hasAttribute('rx')).toBe(false);
      expect(x.hasAttribute('ry')).toBe(false);
    }
    expect(build(sceneOf([rectItem(8.12345)])).children[0]!.getAttribute('rx')).toBe('8.123');
  });

  it('layoutChart の面の矩形は半径 11px で、svg の最初の rect が rx=11（roundedCorners=false は直角）', () => {
    const round = layoutChart(barModel(), 400, 300)!;
    const first = build(round).querySelector('rect')!;
    expect(first.getAttribute('rx')).toBe('11');
    expect(first.getAttribute('ry')).toBe('11');
    const square = layoutChart(barModel({ roundedCorners: false }), 400, 300)!;
    expect(build(square).querySelector('rect')!.hasAttribute('rx')).toBe(false);
  });
});
