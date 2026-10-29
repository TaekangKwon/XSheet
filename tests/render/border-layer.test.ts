// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import type { BorderEdge, BorderStyle, CellBorder } from '../../src/core/index';
import { createBorderLayer } from '../../src/render/border-layer';
import { buildBorderPaths } from '../../src/render/borders';
import type { GridRow } from '../../src/render/grid';

/**
 * 罫線層の DOM（ISSUE-98・ADR-0029 §決定 6・8）のテスト。jsdom は CSS を適用しないため、
 * 色は class（墨色）と CSSOM の style（明示色）で確かめ、見た目は目視（TV-43）で確かめる。
 */

const geo = {
  colWidth: () => 40,
  colLeft: (c: number) => Math.max(0, c - 1) * 40,
  rowTop: (r: number) => Math.max(0, r - 1) * 20,
};

const edge = (style: BorderStyle, rgb: number | null = null, alpha = 1): BorderEdge => ({
  style,
  color: rgb === null ? null : { rgb, alpha },
});
const cb = (p: Partial<CellBorder>): CellBorder => ({
  diagonalUp: false,
  diagonalDown: false,
  ...p,
});

/** 1 行・複数列（各列の指定）から結果を作る */
function build(cells: CellBorder[]) {
  const row: GridRow = { r: 1, cells: cells.map(() => undefined), borders: cells.map((_, i) => i) };
  return buildBorderPaths([row], cells, geo, { first: 1, last: cells.length });
}

const paths = (svg: SVGSVGElement): SVGPathElement[] => [...svg.querySelectorAll('path')];

describe('TD-98-9 罫線層（svg.xs-borders）', () => {
  it('svg は 1 枚。初期は空（幅 0）で、aria-hidden', () => {
    const layer = createBorderLayer();
    expect(layer.element.namespaceURI).toBe('http://www.w3.org/2000/svg');
    expect(layer.element.getAttribute('class')).toBe('xs-borders');
    expect(layer.element.getAttribute('aria-hidden')).toBe('true');
    expect(layer.element.childElementCount).toBe(0);
    expect(layer.element.getAttribute('width')).toBe('0');
  });

  it('位置と大きさは描いた線の外接矩形。座標は viewBox で A1 原点のまま保つ', () => {
    const layer = createBorderLayer();
    layer.render(build([cb({ right: edge('thin') }), cb({})]), 32, 24);
    const svg = layer.element;
    // 列 1 の右の線＝x 39〜40・y 0〜20（A1 原点）。行番号の列 32・見出し 24 だけずらして置く
    expect(svg.style.left).toBe('71px');
    expect(svg.style.top).toBe('24px');
    expect(svg.getAttribute('width')).toBe('1');
    expect(svg.getAttribute('height')).toBe('20');
    expect(svg.getAttribute('viewBox')).toBe('39 0 1 20');
  });

  it('墨色（自動・黒）は class is-ink だけで、style に色を入れない', () => {
    const layer = createBorderLayer();
    layer.render(build([cb({ bottom: edge('thin') }), cb({ bottom: edge('thin', 0) })]), 0, 0);
    const [p] = paths(layer.element);
    expect(paths(layer.element)).toHaveLength(1);
    expect(p!.getAttribute('class')).toBe('xs-border-path is-ink');
    expect(p!.getAttribute('style')).toBeNull();
    expect(p!.getAttribute('stroke')).toBeNull();
    expect(p!.getAttribute('fill')).toBe('none');
    expect(p!.getAttribute('stroke-width')).toBe('1');
    expect(p!.getAttribute('shape-rendering')).toBe('crispEdges');
  });

  it('明示色は #rrggbb を CSSOM の stroke に入れる（class は is-ink でない）。不透明度は stroke-opacity', () => {
    const layer = createBorderLayer();
    layer.render(
      build([
        cb({ bottom: edge('thin', 0xff0000) }),
        cb({ bottom: edge('dashed', 0x0000ff, 0.5) }),
      ]),
      0,
      0,
    );
    const [red, blue] = paths(layer.element);
    expect(red!.getAttribute('class')).toBe('xs-border-path');
    // jsdom は色を rgb() に正規化する（#ff0000 と同じ値）
    expect(red!.style.getPropertyValue('stroke')).toBe('rgb(255, 0, 0)');
    expect(red!.style.getPropertyValue('stroke-opacity')).toBe('');
    expect(blue!.style.getPropertyValue('stroke')).toBe('rgb(0, 0, 255)');
    expect(blue!.style.getPropertyValue('stroke-opacity')).toBe('0.5');
    expect(blue!.getAttribute('stroke-dasharray')).toBe('3 1');
  });

  it('path は線種×色ごとにまとまる（同じ線種・色の区間は 1 つの path に連なる）', () => {
    const layer = createBorderLayer();
    layer.render(
      build([
        cb({ bottom: edge('thin', 0xff0000) }),
        cb({ bottom: edge('thin') }),
        cb({ bottom: edge('thin', 0xff0000) }),
        cb({ bottom: edge('medium', 0xff0000) }),
      ]),
      0,
      0,
    );
    const ps = paths(layer.element);
    expect(ps).toHaveLength(3);
    expect(ps[0]!.getAttribute('d')).toBe('M0 19.5H40M80 19.5H120');
  });

  it('二重線：隙間の path（is-gap・幅 3）が先に来て、その上に 1px の線 2 本（同じ位置から 1px ずつ離れた d）', () => {
    const layer = createBorderLayer();
    layer.render(build([cb({ right: edge('double') }), cb({ right: edge('thin') }), cb({})]), 0, 0);
    const ps = paths(layer.element);
    // 隙間（double の中心線）→ double の 2 本 → thin
    expect(ps.map((p) => p.getAttribute('class'))).toEqual([
      'xs-border-path is-gap',
      'xs-border-path is-ink',
      'xs-border-path is-ink',
    ]);
    expect(ps[0]!.getAttribute('stroke-width')).toBe('3');
    expect(ps[0]!.getAttribute('d')).toBe('M39.5 0V20');
    expect(ps[0]!.getAttribute('style')).toBeNull();
    expect(ps[1]!.getAttribute('stroke-width')).toBe('1');
    expect(ps[1]!.getAttribute('d')).toBe('M38.5 0V20M40.5 0V20');
    expect(ps[2]!.getAttribute('d')).toBe('M79.5 0V20');
  });

  it('斜線の path は画素にそろえない（shape-rendering を付けない）', () => {
    const layer = createBorderLayer();
    layer.render(build([cb({ diagonal: edge('thin'), diagonalDown: true })]), 0, 0);
    const [p] = paths(layer.element);
    expect(p!.getAttribute('d')).toBe('M0 0L40 20');
    expect(p!.getAttribute('shape-rendering')).toBeNull();
  });

  it('作り直しで中身は入れ替わる。線が無い結果・clear で空になる（要素は残る）', () => {
    const layer = createBorderLayer();
    const el = layer.element;
    layer.render(build([cb({ right: edge('thin') })]), 0, 0);
    expect(paths(el)).toHaveLength(1);
    layer.render(build([cb({ right: edge('medium') }), cb({ right: edge('thin') })]), 0, 0);
    expect(paths(el)).toHaveLength(2);
    layer.render(build([cb({})]), 0, 0);
    expect(el.childElementCount).toBe(0);
    expect(el.getAttribute('width')).toBe('0');
    layer.render(build([cb({ right: edge('thin') })]), 0, 0);
    layer.clear();
    expect(el.childElementCount).toBe(0);
    expect(layer.element).toBe(el);
  });

  it('ファイル由来の文字列（色の文字列）は入らない：数値でない色は墨色に倒れ、style に何も入らない', () => {
    const layer = createBorderLayer();
    const bad = { style: 'thin', color: { rgb: Number.NaN, alpha: 1 } } as BorderEdge;
    layer.render(build([cb({ bottom: bad })]), 0, 0);
    const [p] = paths(layer.element);
    expect(p!.getAttribute('class')).toBe('xs-border-path is-ink');
    expect(p!.getAttribute('style')).toBeNull();
  });
});
