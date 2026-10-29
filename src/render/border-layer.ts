// 正: InterHub docs/plan/adr/0029-xlsx-cell-borders.md（§決定 1・6・8）
/**
 * `.xlsx` の罫線層（ISSUE-98）：`svg.xs-borders` 1 枚の中身を、{@link buildBorderPaths} の結果から作る。
 * - `createElementNS` だけで作る（`innerHTML` は使わない）。`d`・`stroke-width`・`stroke-dasharray` は数値から組み立てた文字列だけ。
 * - 墨色（自動・黒）は class `is-ink` だけ（色は CSS のテーマ変数 `--xs-border-ink`。テーマの切替で描き直さない）。
 *   ほかの明示色は整数から組み立てた `#rrggbb` を CSSOM（`style.setProperty`）で入れる（CSP のため style 属性の文字列は組み立てない）。
 * - 二重線は、セルの面の色の隙間（`is-gap`・幅 3）の上に 1px の線 2 本。隙間は全 path の線より先（下）に描く。
 * - 位置と大きさは描いた線の外接矩形だけ（A1 の左上を原点とする座標を viewBox で保つ）。線が無ければ中身を空にする。
 */
import type { BorderPath, BorderPaths } from './borders';

const SVG_NS = 'http://www.w3.org/2000/svg';

/** 罫線層 */
export interface BorderLayer {
  /** wrapper に置く svg（常に 1 枚。中身だけ入れ替わる） */
  readonly element: SVGSVGElement;
  /**
   * 結果で中身を作り直す。線が無ければ空にする。
   * @param result {@link buildBorderPaths} の結果（座標は A1 の左上が原点）
   * @param originX A1 の左上の x（行番号の列の幅）
   * @param originY A1 の左上の y（見出し行の高さ）
   */
  render(result: BorderPaths, originX: number, originY: number): void;
  /** 中身を空にする（シートの切替・後始末） */
  clear(): void;
}

/** SVG の path を作り、属性を入れる（値は数値・列挙値・数値だけの `d` のみ） */
function pathEl(attrs: Record<string, string | number>): SVGPathElement {
  const e = document.createElementNS(SVG_NS, 'path');
  e.setAttribute('fill', 'none');
  for (const [k, v] of Object.entries(attrs)) {
    e.setAttribute(k, String(v));
  }
  return e;
}

/** 線の path（二重線以外は `d`・幅・パターン、二重線は 1px の線 2 本の `lineD`）。色は墨色なら class、明示色なら CSSOM */
function linePath(p: BorderPath): SVGPathElement | null {
  const double = p.lineD !== undefined;
  const d = double ? p.lineD! : p.d;
  if (d === '') {
    return null;
  }
  const attrs: Record<string, string | number> = {
    class: p.ink ? 'xs-border-path is-ink' : 'xs-border-path',
    d,
    'stroke-width': double ? 1 : p.width,
  };
  if (!double && p.dasharray !== null) {
    attrs['stroke-dasharray'] = p.dasharray;
  }
  if (!p.diagonal) {
    attrs['shape-rendering'] = 'crispEdges';
  }
  const el = pathEl(attrs);
  if (!p.ink && p.color !== null) {
    el.style.setProperty('stroke', p.color);
    if (p.alpha < 1) {
      el.style.setProperty('stroke-opacity', String(p.alpha));
    }
  }
  return el;
}

/** 二重線の隙間の path（セルの面の色で、線 2 本の間とグリッド線を覆う） */
function gapPath(p: BorderPath): SVGPathElement {
  const attrs: Record<string, string | number> = {
    class: 'xs-border-path is-gap',
    d: p.d,
    'stroke-width': p.width,
  };
  if (!p.diagonal) {
    attrs['shape-rendering'] = 'crispEdges';
  }
  return pathEl(attrs);
}

/** 罫線層を作る（svg は常に 1 枚で、中身だけを `render` で入れ替える） */
export function createBorderLayer(): BorderLayer {
  const element = document.createElementNS(SVG_NS, 'svg');
  element.setAttribute('class', 'xs-borders');
  element.setAttribute('aria-hidden', 'true');
  element.setAttribute('focusable', 'false');

  const clear = (): void => {
    element.replaceChildren();
    element.removeAttribute('viewBox');
    element.setAttribute('width', '0');
    element.setAttribute('height', '0');
  };
  clear();

  return {
    element,
    clear,
    render(result, originX, originY) {
      const b = result.bounds;
      if (!b || result.paths.length === 0) {
        clear();
        return;
      }
      const x = Math.floor(b.left);
      const y = Math.floor(b.top);
      const w = Math.max(0, Math.ceil(b.right) - x);
      const h = Math.max(0, Math.ceil(b.bottom) - y);
      element.style.left = `${originX + x}px`;
      element.style.top = `${originY + y}px`;
      element.setAttribute('width', String(w));
      element.setAttribute('height', String(h));
      element.setAttribute('viewBox', `${x} ${y} ${w} ${h}`);
      const children: SVGPathElement[] = [];
      for (const p of result.paths) {
        if (p.lineD !== undefined && p.d !== '') {
          children.push(gapPath(p)); // 隙間は、ほかの線より先（下）に
        }
      }
      for (const p of result.paths) {
        const el = linePath(p);
        if (el) {
          children.push(el);
        }
      }
      element.replaceChildren(...children);
    },
  };
}
