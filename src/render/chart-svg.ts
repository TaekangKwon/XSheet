// 正: InterHub docs/plan/adr/0031-xlsx-charts.md（§決定 3・8・10）
/**
 * `.xlsx` のグラフの場面（{@link ChartScene}。ISSUE-102）を `<svg>` 1 枚にする（ISSUE-103）。
 * - 場面の項目を順に要素へ写すだけで、計算はしない。`rect`・`path`・`line`・`text`（回転・揃え・太字）。
 * - 属性に入れるのは数値・`#rrggbb`・不透明度・固定の語・数値と固定の命令だけの `d`。ファイル由来の文字列（系列名・項目・タイトル）は
 *   `<text>` の `textContent` にだけ入れる（`innerHTML`・属性には入れない）。`font-family` は書かず画面の文字を継承する。
 * - 場面の値が不正（有限でない数・`#rrggbb` でない色・`d` に想定外の文字）の項目は描かない。
 * - ダーク表示（ADR-0031 §決定 10・P5）：グラフは Excel の色のまま描く。場面の `backdrop`（グラフの面が塗りなし）が真のときの白い下地は、この `<svg>` の中には置かない
 *   （中に置くと、グラフより下の図形・画像まで隠れる）。画像レイヤー（`xlsx-image-layer.ts`）が、図形と同じ兄弟の `svg.xs-backdrop`（`z-index: -1`）として図の直前に置く。
 * - DOM は `doc`（`Document`）経由で作る（テストで差し替えられる）。`chrome.*` に触らない。
 */
import type {
  ChartAnchor,
  ChartBaseline,
  ChartColor,
  ChartItem,
  ChartScene,
  ChartStroke,
} from './chart-layout';
import { dashArray } from './shapes';

const SVG_NS = 'http://www.w3.org/2000/svg';

/** 文字の揃え → `text-anchor`（固定の語の表） */
const TEXT_ANCHOR: Record<ChartAnchor, string> = {
  start: 'start',
  middle: 'middle',
  end: 'end',
};

/**
 * 文字の縦位置 → `dominant-baseline`（固定の語の表）。場面の `middle` は行の中央に置く意味なので、em の箱の中央（`central`）に写す。
 */
const TEXT_BASELINE: Record<ChartBaseline, string> = {
  alphabetic: 'alphabetic',
  middle: 'central',
  hanging: 'hanging',
};

/** 線の端・結びの固定の語（場面の値を通す表。表に無い値は入れない） */
const LINE_CAP: Record<string, string> = { butt: 'butt', round: 'round' };
const LINE_JOIN: Record<string, string> = { miter: 'miter', round: 'round' };

/** `d` に許す文字（数値と固定の命令 `M L C A Z`・区切り・符号・小数点・指数） */
const PATH_DATA = /^[MLCAZ0-9\s.,+\-eE]*$/;

/** `#rrggbb` */
const HEX_COLOR = /^#[0-9a-fA-F]{6}$/;

/** 有限でない値を 0 にして小数 3 桁へ丸める（属性に入れる数値用） */
function num(n: number): number {
  return Number.isFinite(n) ? Math.round(n * 1000) / 1000 : 0;
}

/** 色が `#rrggbb` で、不透明度が有限か */
function validColor(c: ChartColor | null | undefined): c is ChartColor {
  return !!c && HEX_COLOR.test(c.color) && Number.isFinite(c.opacity);
}

/** 不透明度を 0〜1・小数 3 桁に収める */
function alpha(a: number): number {
  return Math.round(Math.min(1, Math.max(0, a)) * 1000) / 1000;
}

/** SVG の要素を作り、属性（数値・固定の語・検証済みの値だけ）を入れる */
function el<K extends keyof SVGElementTagNameMap>(
  doc: Document,
  tag: K,
  attrs: Record<string, string | number>,
): SVGElementTagNameMap[K] {
  const e = doc.createElementNS(SVG_NS, tag);
  for (const [k, v] of Object.entries(attrs)) {
    e.setAttribute(k, String(v));
  }
  return e;
}

/** 塗りの属性（`fill`・`fill-opacity`）。null・不正な色は `fill="none"` */
function fillAttrs(fill: ChartColor | null): Record<string, string | number> {
  if (!validColor(fill)) {
    return { fill: 'none' };
  }
  const a = alpha(fill.opacity);
  return a < 1 ? { fill: fill.color, 'fill-opacity': a } : { fill: fill.color };
}

/** 線の属性。null・不正な色は `stroke="none"` */
function strokeAttrs(stroke: ChartStroke | null): Record<string, string | number> {
  if (!stroke || !validColor(stroke)) {
    return { stroke: 'none' };
  }
  const w = num(stroke.width);
  const out: Record<string, string | number> = {
    stroke: stroke.color,
    'stroke-width': w,
    'stroke-linecap': LINE_CAP[stroke.cap] ?? 'butt',
    'stroke-linejoin': LINE_JOIN[stroke.join] ?? 'miter',
  };
  const a = alpha(stroke.opacity);
  if (a < 1) {
    out['stroke-opacity'] = a;
  }
  const dash = dashArray(stroke.dash, w).filter((n) => Number.isFinite(n) && n >= 0);
  if (dash.length > 0) {
    out['stroke-dasharray'] = dash.map(num).join(' ');
  }
  return out;
}

/** 項目 1 つを要素にする（不正なら null） */
function itemEl(doc: Document, item: ChartItem): SVGElement | null {
  switch (item.kind) {
    case 'rect':
      return el(doc, 'rect', {
        x: num(item.x),
        y: num(item.y),
        width: num(item.width),
        height: num(item.height),
        // 角丸（グラフの面だけ。数値の rx・ry。ISSUE-109）
        ...(item.radius !== undefined && Number.isFinite(item.radius) && item.radius > 0
          ? { rx: num(item.radius), ry: num(item.radius) }
          : {}),
        ...fillAttrs(item.fill),
        ...strokeAttrs(item.stroke),
      });
    case 'path':
      if (item.d === '' || !PATH_DATA.test(item.d)) {
        return null;
      }
      return el(doc, 'path', {
        d: item.d,
        ...fillAttrs(item.fill),
        ...strokeAttrs(item.stroke),
      });
    case 'line':
      return el(doc, 'line', {
        x1: num(item.x1),
        y1: num(item.y1),
        x2: num(item.x2),
        y2: num(item.y2),
        fill: 'none',
        ...strokeAttrs(item.stroke),
      });
    case 'text': {
      if (!validColor(item.color)) {
        return null;
      }
      const x = num(item.x);
      const y = num(item.y);
      const attrs: Record<string, string | number> = {
        x,
        y,
        'font-size': num(item.size),
        'text-anchor': TEXT_ANCHOR[item.anchor] ?? 'start',
        'dominant-baseline': TEXT_BASELINE[item.baseline] ?? 'alphabetic',
        ...fillAttrs(item.color),
      };
      if (item.bold) {
        attrs['font-weight'] = 'bold';
      }
      const rot = num(item.rotate);
      if (rot !== 0) {
        attrs.transform = `rotate(${rot} ${x} ${y})`;
      }
      const t = el(doc, 'text', attrs);
      // ファイル由来の文字列はここ（textContent）だけ
      t.textContent = item.text;
      return t;
    }
    default:
      return null;
  }
}

/**
 * グラフの場面を `<svg>` にする。大きさ（`width`・`height`・`viewBox`）は場面の px。クラス `xs-chart`（CSS が枠いっぱいに伸ばす）。
 * 下地（`scene.backdrop`）はこの `<svg>` には含めない（呼び出し側の画像レイヤーが図の層の最背面に置く）。
 *
 * @param scene `layoutChart` の結果（null は呼び出し側が破線枠にする）
 * @param doc 要素を作る文書
 * @param title 読み上げの名前（呼び出し側の i18n の文言。`role="img"` の `<title>` に `textContent` で入れる。省略すると `aria-hidden`）
 */
export function buildChartSvg(scene: ChartScene, doc: Document, title?: string): SVGSVGElement {
  const w = num(scene.width);
  const h = num(scene.height);
  const svg = el(doc, 'svg', {
    width: w,
    height: h,
    viewBox: `0 0 ${w} ${h}`,
    focusable: 'false',
    ...(title === undefined ? { 'aria-hidden': 'true' } : { role: 'img' }),
  });
  svg.setAttribute('class', 'xs-chart');
  if (title !== undefined) {
    // 属性ではなく <title> の textContent（描画はされない。ファイル由来の文字列は渡さない）
    const t = doc.createElementNS(SVG_NS, 'title');
    t.textContent = title;
    svg.append(t);
  }
  for (const item of scene.items) {
    const e = itemEl(doc, item);
    if (e) {
      svg.append(e);
    }
  }
  return svg;
}
