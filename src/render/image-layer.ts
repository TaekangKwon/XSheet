// 正: InterHub docs/plan/adr/0026-xlsx-image-overlay.md（§決定 6・8・9）
/**
 * `.xlsx` グリッドの上に重ねる画像レイヤー（ISSUE-86）。
 * - 構造：`position: absolute` のレイヤー要素（原点＝A1 の左上）の中に、図ごとの要素を `left/top/width/height`（px）で置く。
 *   表とは独立で、スクロールは枠ごと動くため位置の再計算は要らない。可視範囲（上下左右 1 画面分）と交わる図だけを DOM に出し入れする。
 * - 画像は `<img>`（`decoding="async"`・`alt` は説明か名前）。画像以外・描けない画像・壊れた画像は種類名入りの破線枠。
 *   種類名は文言（`messages`）を `textContent` で入れる（図の名前・パスは画面にも console にも出さない。名前は `alt` のみ）。
 * - 図形・テキストボックスは、形を評価できれば形どおりに描く（ISSUE-92）：`div.xs-figure`（外接矩形）＞ `<svg>`（path 1 つにつき `<path>` 1 つ。枠線・塗り・面の濃淡はクラスと CSS）
 *   ＋ `div.xs-shape-box`（回転前の矩形・回転・図形の矩形で切る）＞ `div.xs-shape-text`（文字領域）。枠も塗りも文字も無い図形は描かない（配置の時点で除く）。
 * - 形を評価できない図形（代替）で文字のあるものは種類名ではなく文字を描く（ISSUE-88。段落＝`div`・ラン＝`span`・改行＝`br`、すべて `textContent`）。
 *   枠線・塗りの有無はクラス、回転・縦書きは列挙値の対応表から `style` に入れる（ファイルの文字列は style に入れない）。
 * - 線・矢印（ISSUE-88）は 1 本につき `<svg>` 1 つ（外接矩形の位置）。`<marker>`・`innerHTML` は使わず、矢じりは `<path>`・`<ellipse>`。
 * - 色（ISSUE-96）：lib の整数 RGB・不透明度を検証して `#rrggbb`・数値にし、要素の style（CSSOM）で入れる（図形の fill・stroke、線・矢じり、文字の color、代替の文字枠の背景・枠線）。
 *   面の濃淡は lib の `applyFillMode`。色が無い（省略）ときは style を入れず、暫定のテーマ変数（CSS）のまま。
 * - ダーク表示の下地（ISSUE-96）：図形・線・代替の文字枠は、図の要素の直前に独立した `svg.xs-backdrop`（`z-index: -1`・片側 3px 広げる）を兄弟として置く。
 *   形の塗りの path と線の縁取り（太さ＋6px・実線・丸）だけを持ち、色は CSS のテーマ変数（ライトは透明）で決める＝style に色を入れず、テーマの切り替えで描き直さない。
 *   レイヤーが `isolation: isolate` で重ね順の文脈を作り、下地は全図の下・セルの上に描かれる。図と一緒に出し入れする。
 * - blob URL は {@link createImageStore}（シート 1 枚ぶん）が持つ。シートの差し替え・{@link ImageLayer.dispose} で revoke する。
 * 位置の計算は DOM 非依存の `drawings.ts`、blob の管理は `image-store.ts` に切り出している。
 */
import type {
  ImageResult,
  SheetDrawing,
  DrawingColor,
  DrawingText,
  TextAlign,
  TextAnchor,
  TextVert,
} from '../core/index';
import { applyFillMode } from '../core/internal';
import { EMU_PER_PX, pickVisible, type DrawingBox, type PxRect } from './drawings';
import {
  arrowPolygonData,
  BACKDROP_PAD,
  filledShown,
  preRotationRect,
  SHAPE_MITER_LIMIT,
  shapeHasText,
  textShown,
  xfrmExtPx,
  type LineRender,
  type ShapeDrawn,
  type ShapeShade,
} from './shapes';
import { createImageStore, type ImageEntry, type ImageStore } from './image-store';
import { formatMessage, MESSAGES_JA, type XSheetMessages } from './messages';

/** 画像レイヤーに渡すシートの図 */
export interface LayerSheet {
  /** 図の一覧（drawing XML の順） */
  drawings: readonly SheetDrawing[];
  /** 配置済みの図（`layoutDrawings` の結果。`index` は `drawings` の位置） */
  boxes: readonly DrawingBox[];
  /** 画像を読む（`XlsxBook.readImage` をシート番号込みで束ねたもの） */
  readImage: (path: string) => Promise<ImageResult>;
  /** 画像の合計の上限を超えたとき（バナー用）。1 シートにつき 1 回 */
  onOverBudget?: () => void;
}

/** 画像レイヤー */
export interface ImageLayer {
  /** スクロール枠の中の wrapper に取り付けるレイヤー要素 */
  readonly element: HTMLElement;
  /** 描くシートを差し替える（前のシートの blob URL はすべて revoke し、DOM の図も外す） */
  setSheet(sheet: LayerSheet | null): void;
  /** 原点（A1 の左上）の位置とレイヤーの大きさ（px。wrapper の左上基準） */
  place(left: number, top: number, width: number, height: number): void;
  /** 可視範囲（A1 基準・余白込み。`viewWindow`）と交わる図だけを DOM に出し、外れた図を外す */
  update(view: PxRect): void;
  /** すべての blob URL を revoke し、DOM の図を外す。以後 `setSheet`・`update` は何もしない */
  dispose(): void;
}

/** 図の種類名（破線枠の文字）の文言のキー。画像の `image` は「表示できない画像」の枠に使う */
const KIND_KEYS: Record<SheetDrawing['kind'], keyof XSheetMessages> = {
  image: 'imageBroken',
  chart: 'chart',
  smartart: 'smartart',
  shape: 'shape',
  textbox: 'textbox',
  group: 'group',
  // 線は `<svg>` で描き、枠にならない。型を満たすための値
  line: 'shape',
};

/**
 * 破線枠に入れる種類名。画像（`kind: 'image'`）は、描けない形式で形式名が分かれば「画像（EMF）」、それ以外は「画像（表示できない）」。
 *
 * @param entry 画像の読み込み失敗（画像以外・読み込み前は不要）
 * @param messages 文言の表（既定は日本語）
 */
export function frameLabel(
  drawing: SheetDrawing,
  entry?: ImageEntry,
  messages: Readonly<XSheetMessages> = MESSAGES_JA,
): string {
  if (
    drawing.kind === 'image' &&
    entry &&
    !entry.ok &&
    entry.reason === 'unsupported' &&
    entry.label
  ) {
    return formatMessage(messages.imageFormat, { format: entry.label });
  }
  return messages[KIND_KEYS[drawing.kind]];
}

/** 図 1 つの要素（画像・破線枠・文字の枠は `div`、線は `<svg>`） */
type FigureEl = HTMLElement | SVGSVGElement;

/** 図 1 つの DOM（位置・大きさは固定。中身は画像・破線枠・文字・線のいずれか） */
interface Figure {
  box: DrawingBox;
  drawing: SheetDrawing;
  el: FigureEl;
  /** ダーク表示の下地（図形・線・代替の文字枠だけ。図の要素の直前に置く。ISSUE-96） */
  backdrop?: SVGSVGElement;
}

/** 種類名入りの破線枠にする（中身を置き換える） */
function fillFrame(el: FigureEl, label: string): void {
  const frame = document.createElement('div');
  frame.className = 'xs-frame';
  frame.textContent = label;
  el.replaceChildren(frame);
}

const SVG_NS = 'http://www.w3.org/2000/svg';
/** CSS の px（1pt = 96/72 px） */
const PX_PER_PT = 96 / 72;
/** 文字の大きさの既定（pt。Excel の既定） */
const DEFAULT_PT = 11;
/** 内側の余白の既定（EMU。Excel の既定 l/r 0.1 インチ・t/b 0.05 インチ） */
const DEFAULT_INSETS = { l: 91440, t: 45720, r: 91440, b: 45720 };
/** 行の高さ（文字の大きさの倍率） */
const LINE_HEIGHT = 1.2;

/** 段落の揃え → `text-align`（論理値。縦書きでも合う） */
const TEXT_ALIGN: Record<TextAlign, string> = {
  l: 'start',
  ctr: 'center',
  r: 'end',
  just: 'justify',
  dist: 'justify',
};
/** 縦位置 → flex の `justify-content`（column の主軸＝ブロック方向なので縦書きでも合う） */
const TEXT_ANCHOR: Record<TextAnchor, string> = {
  t: 'flex-start',
  ctr: 'center',
  b: 'flex-end',
};
/** 文字の向き → `writing-mode`・`text-orientation`（表に無い値＝横書き） */
const TEXT_VERT: Partial<Record<TextVert, { mode: string; orientation?: string }>> = {
  eaVert: { mode: 'vertical-rl' },
  wordArtVertRtl: { mode: 'vertical-rl' },
  mongolianVert: { mode: 'vertical-lr' },
  vert: { mode: 'vertical-rl', orientation: 'sideways' },
  vert270: { mode: 'sideways-lr' },
  wordArtVert: { mode: 'vertical-rl', orientation: 'upright' },
};

/**
 * 文字を描く図形か（kind が shape・textbox で、表示する文字が 1 文字以上ある。改行だけ・空のランは含めない）。
 * 該当しない図形は種類名入りの破線枠になる。
 */
export function hasShapeText(drawing: SheetDrawing): boolean {
  return shapeHasText(drawing);
}

/** 有限でない値を 0 にして小数 2 桁へ丸める（style・属性に入れる数値用） */
function num(n: number): number {
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
}

/** 検証済みの色：`#rrggbb` と不透明度（0〜1・小数 3 桁）。DOM に入れるのはこの 2 つだけ */
interface Paint {
  hex: string;
  alpha: number;
}

/**
 * lib の色（整数 RGB・不透明度）を検証して `Paint` にする。RGB が 0〜0xFFFFFF の整数でない・不透明度が有限でないときは null（入れない）。
 * 不透明度は 0〜1 に収める。ファイルの文字列は通らない（数値だけから `#rrggbb` を組み立てる）。
 */
function paintOf(c: DrawingColor | undefined): Paint | null {
  if (
    !c ||
    !Number.isInteger(c.rgb) ||
    c.rgb < 0 ||
    c.rgb > 0xffffff ||
    !Number.isFinite(c.alpha)
  ) {
    return null;
  }
  return {
    hex: `#${c.rgb.toString(16).padStart(6, '0')}`,
    alpha: Math.round(Math.min(1, Math.max(0, c.alpha)) * 1000) / 1000,
  };
}

/** SVG 要素の style に塗り（`fill`）か線（`stroke`）の色を入れる。不透明度が 1 未満なら `-opacity` も入れる（要素の style は CSS のクラスの規則に勝つ） */
function setPaint(el: SVGElement, prop: 'fill' | 'stroke', p: Paint | null): void {
  if (!p) {
    return;
  }
  el.style.setProperty(prop, p.hex);
  if (p.alpha < 1) {
    el.style.setProperty(`${prop}-opacity`, String(p.alpha));
  }
}

/** HTML 用の色（`color`・背景・枠線）：不透明なら `#rrggbb`、不透明度があれば `#rrggbbaa` */
function cssColor(p: Paint): string {
  return p.alpha < 1
    ? `${p.hex}${Math.round(p.alpha * 255)
        .toString(16)
        .padStart(2, '0')}`
    : p.hex;
}

/** 文字の大きさ（pt）を px にする。有限でない・0 以下は既定、`fontScale` を掛ける */
function fontPx(pt: number | undefined, scale: number | undefined): number {
  const base = pt !== undefined && Number.isFinite(pt) && pt > 0 ? pt : DEFAULT_PT;
  const k = scale !== undefined && Number.isFinite(scale) && scale > 0 ? scale : 1;
  return num(base * PX_PER_PT * k);
}

/** 余白（EMU）を px にする。未指定・有限でない・負は既定値 */
function insetPx(emu: number | undefined, fallback: number): number {
  return num((emu !== undefined && Number.isFinite(emu) && emu >= 0 ? emu : fallback) / EMU_PER_PX);
}

/** 文字の枠（図形の中身）を作る。段落＝`div`・ラン＝`span`・改行＝`br`。本文はすべて `textContent` */
function fillText(el: HTMLElement, text: DrawingText, w: number, h: number): void {
  const scale = text.fontScale;
  const ins = text.insets ?? DEFAULT_INSETS;
  // 余白は各辺とも枠（回転前の矩形）の半分まで。巨大な値で要素が膨らまないようにする
  const cap = (px: number, size: number): number => num(Math.min(px, Math.max(0, size) / 2));
  el.style.paddingLeft = `${cap(insetPx(ins.l, DEFAULT_INSETS.l), w)}px`;
  el.style.paddingTop = `${cap(insetPx(ins.t, DEFAULT_INSETS.t), h)}px`;
  el.style.paddingRight = `${cap(insetPx(ins.r, DEFAULT_INSETS.r), w)}px`;
  el.style.paddingBottom = `${cap(insetPx(ins.b, DEFAULT_INSETS.b), h)}px`;
  el.style.justifyContent = TEXT_ANCHOR[text.anchor] ?? TEXT_ANCHOR.t;
  el.style.whiteSpace = text.noWrap ? 'pre' : 'pre-wrap';
  const vert = TEXT_VERT[text.vert];
  if (vert) {
    el.style.setProperty('writing-mode', vert.mode);
    if (vert.orientation) {
      el.style.setProperty('text-orientation', vert.orientation);
    }
  }
  const paras: HTMLElement[] = [];
  for (const p of text.paragraphs) {
    const para = document.createElement('div');
    para.className = 'xs-para';
    para.style.textAlign = TEXT_ALIGN[p.align] ?? TEXT_ALIGN.l;
    const shown = p.runs.filter((r) => !r.br && r.text.length > 0);
    if (shown.length === 0) {
      // 文字の無い段落だけ段落記号（endParaRPr）の大きさで行の高さを保つ（論理方向＝縦書きでも効く）
      const endPx = fontPx(p.endSizePt, scale);
      para.style.fontSize = `${endPx}px`;
      para.style.setProperty('min-block-size', `${num(endPx * LINE_HEIGHT)}px`);
    } else {
      // 文字のある行は最大のランの大きさで決まる
      para.style.fontSize = `${Math.max(...shown.map((r) => fontPx(r.sizePt, scale)))}px`;
    }
    if (p.bullet) {
      const firstRun = p.runs.find((r) => !r.br);
      const b = document.createElement('span');
      b.textContent = `${p.bullet} `;
      b.style.fontSize = `${fontPx(firstRun?.sizePt ?? p.endSizePt, scale)}px`;
      // 記号は先頭のランの色（Excel の既定は文字と同じ色）
      const bulletPaint = paintOf(firstRun?.color);
      if (bulletPaint) {
        b.style.color = cssColor(bulletPaint);
      }
      para.append(b);
    }
    for (const r of p.runs) {
      if (r.br) {
        para.append(document.createElement('br'));
        continue;
      }
      const span = document.createElement('span');
      span.textContent = r.text;
      span.style.fontSize = `${fontPx(r.sizePt, scale)}px`;
      // 文字の色（`a:noFill` は不透明度 0＝透明）。色が無ければ継承（暫定の `--xs-fg`）
      const runPaint = paintOf(r.color);
      if (runPaint) {
        span.style.color = cssColor(runPaint);
      }
      if (r.bold) {
        span.style.fontWeight = 'bold';
      }
      if (r.italic) {
        span.style.fontStyle = 'italic';
      }
      const deco = [r.underline ? 'underline' : '', r.strike ? 'line-through' : '']
        .filter(Boolean)
        .join(' ');
      if (deco) {
        span.style.textDecoration = deco;
      }
      para.append(span);
    }
    if (p.runs.length > 0 && p.runs[p.runs.length - 1]!.br) {
      // 末尾の改行は HTML では空行にならない（Excel は 1 行空く）ので 1 つ足す
      para.append(document.createElement('br'));
    }
    paras.push(para);
  }
  el.replaceChildren(...paras);
}

/** 時計回りの回転角（度。0〜360 未満）。`rot` は 1/60000 度で、無い・有限でない値は 0 */
function rotationDeg(rot: number | undefined): number {
  if (rot === undefined || !Number.isFinite(rot)) {
    return 0;
  }
  return num((((rot / 60000) % 360) + 360) % 360);
}

/** SVG の要素を作り、属性を入れる（値は数値・列挙値・数値だけの `d` のみ） */
function svgEl<K extends keyof SVGElementTagNameMap>(
  tag: K,
  attrs: Record<string, string | number>,
): SVGElementTagNameMap[K] {
  const e = document.createElementNS(SVG_NS, tag);
  for (const [k, v] of Object.entries(attrs)) {
    e.setAttribute(k, String(v));
  }
  return e;
}

/**
 * 線 1 本の `<svg>` を作る（外接矩形に置き、矢じりは `<path>`・`<ellipse>`。`<marker>` は使わない）。
 *
 * @param color 線の色（lib の `DrawingLine.color`）。矢じりの塗り・線も同じ。無い・不正なら style を入れず、属性の `currentColor`（暫定のテーマ変数）のまま
 */
function buildLine(line: LineRender, rect: PxRect, color?: DrawingColor): SVGSVGElement {
  const paint = paintOf(color);
  const svg = svgEl('svg', {
    width: rect.w,
    height: rect.h,
    viewBox: `0 0 ${num(rect.w)} ${num(rect.h)}`,
    'aria-hidden': 'true',
    focusable: 'false',
  });
  svg.setAttribute('class', 'xs-figure xs-line');
  svg.style.left = `${rect.x}px`;
  svg.style.top = `${rect.y}px`;
  svg.style.width = `${rect.w}px`;
  svg.style.height = `${rect.h}px`;
  const stroke = num(line.strokeWidth);
  const attrs: Record<string, string | number> = {
    d: line.d,
    fill: 'none',
    stroke: 'currentColor',
    'stroke-width': stroke,
    'stroke-linejoin': 'round',
    'stroke-linecap': 'butt',
  };
  const dash = line.dash.filter((n) => Number.isFinite(n) && n >= 0).map(num);
  if (dash.length > 0) {
    attrs['stroke-dasharray'] = dash.join(' ');
  }
  const main = svgEl('path', attrs);
  setPaint(main, 'stroke', paint);
  svg.append(main);
  for (const a of line.heads) {
    if (a.kind === 'polygon') {
      const head = svgEl('path', {
        d: arrowPolygonData(a),
        fill: a.filled ? 'currentColor' : 'none',
        stroke: a.filled ? 'none' : 'currentColor',
        'stroke-width': stroke,
        'stroke-linejoin': 'round',
        'stroke-linecap': 'round',
      });
      // 塗りの矢じりは fill、開いた V は stroke に線の色を入れる
      setPaint(head, a.filled ? 'fill' : 'stroke', paint);
      svg.append(head);
    } else {
      const oval = svgEl('ellipse', {
        cx: num(a.cx),
        cy: num(a.cy),
        rx: num(a.rx),
        ry: num(a.ry),
        transform: `rotate(${num(a.angle)} ${num(a.cx)} ${num(a.cy)})`,
        fill: 'currentColor',
        stroke: 'none',
      });
      setPaint(oval, 'fill', paint);
      svg.append(oval);
    }
  }
  return svg;
}

/**
 * ダーク表示の下地の `<svg>` を作る（ISSUE-96）。外接矩形（図の位置）から四方に `pad` だけ広げて置き、`viewBox` を `-pad` からにして中の座標は図と同じにする。
 * クラス `xs-backdrop`：`z-index: -1`・色・不透明度は CSS（テーマ変数）が決める。図の要素ではない（`xs-figure` を付けない）。
 */
function backdropSvg(rect: PxRect, pad: number): SVGSVGElement {
  const w = num(rect.w + 2 * pad);
  const h = num(rect.h + 2 * pad);
  const svg = svgEl('svg', {
    width: w,
    height: h,
    viewBox: `${-pad} ${-pad} ${w} ${h}`,
    'aria-hidden': 'true',
    focusable: 'false',
  });
  svg.setAttribute('class', 'xs-backdrop');
  svg.style.left = `${num(rect.x - pad)}px`;
  svg.style.top = `${num(rect.y - pad)}px`;
  svg.style.width = `${w}px`;
  svg.style.height = `${h}px`;
  return svg;
}

/** 下地の要素（`path`・`ellipse`・`rect`）。`fill` なら形を白く塗り、`strokeWidth` があれば線の縁取り（実線・角と端は丸）。色は付けない（クラスと CSS のテーマ変数） */
function backdropPart<K extends 'path' | 'ellipse' | 'rect'>(
  tag: K,
  attrs: Record<string, string | number>,
  fill: boolean,
  strokeWidth?: number,
): SVGElementTagNameMap[K] {
  const all = { ...attrs };
  const classes = ['xs-backdrop-part'];
  if (fill) {
    classes.push('is-fill');
  }
  if (strokeWidth !== undefined) {
    classes.push('is-stroke');
    all['stroke-width'] = strokeWidth;
    all['stroke-linejoin'] = 'round';
    all['stroke-linecap'] = 'round';
  }
  all.class = classes.join(' ');
  return svgEl(tag, all);
}

/** 線・矢印の下地：線の path に沿った白い縁取り（太さ＋片側 3px。実線）。矢じりは形を白で塗り、同じ縁取りを付ける（開いた V は縁取りだけ） */
function buildLineBackdrop(line: LineRender, rect: PxRect): SVGSVGElement {
  const svg = backdropSvg(rect, BACKDROP_PAD);
  const border = num(line.strokeWidth + 2 * BACKDROP_PAD);
  svg.append(backdropPart('path', { d: line.d }, false, border));
  for (const a of line.heads) {
    if (a.kind === 'polygon') {
      svg.append(backdropPart('path', { d: arrowPolygonData(a) }, a.filled, border));
    } else {
      svg.append(
        backdropPart(
          'ellipse',
          {
            cx: num(a.cx),
            cy: num(a.cy),
            rx: num(a.rx),
            ry: num(a.ry),
            transform: `rotate(${num(a.angle)} ${num(a.cx)} ${num(a.cy)})`,
          },
          true,
          border,
        ),
      );
    }
  }
  return svg;
}

/** 面の濃淡 → クラス（色が無いときだけ CSS が文字色・背景色の薄い膜にする暫定の色。色があれば style の色が勝つ） */
const SHADE_CLASS: Record<ShapeShade, string> = {
  darken: 'is-darken',
  darkenLess: 'is-darken-less',
  lighten: 'is-lighten',
  lightenLess: 'is-lighten-less',
};

/**
 * 図形の形の `<svg>` を作る（外接矩形いっぱい。path 1 つにつき `<path>` 1 つ）。
 * 色は Excel の色（`drawing.fillColor`・`outlineColor`）を style の fill・stroke に入れる。面の濃淡は塗りの色に lib の `applyFillMode` を掛けた色。
 * 色が無い（省略・不正）ときは style を入れず、クラスと CSS の暫定の色（テーマ変数）のまま。
 */
function buildShapeSvg(shape: ShapeDrawn, rect: PxRect, drawing: SheetDrawing): SVGSVGElement {
  // 濃淡の計算に渡す前に検証する（不正な値は計算に通さない）
  const fillColor = paintOf(drawing.fillColor) ? drawing.fillColor : undefined;
  const outlinePaint = paintOf(drawing.outlineColor);
  const svg = svgEl('svg', {
    width: rect.w,
    height: rect.h,
    viewBox: `0 0 ${num(rect.w)} ${num(rect.h)}`,
    'aria-hidden': 'true',
    focusable: 'false',
  });
  svg.setAttribute('class', 'xs-shape-svg');
  const stroke = num(shape.strokeWidth);
  const dash = shape.dash.filter((n) => Number.isFinite(n) && n >= 0).map(num);
  for (const f of shape.faces) {
    const attrs: Record<string, string | number> = { d: f.d };
    const classes = ['xs-shape-path'];
    if (f.fill) {
      classes.push('has-fill');
      if (f.shade) {
        classes.push(SHADE_CLASS[f.shade]);
      }
    }
    if (f.stroke) {
      classes.push('has-stroke');
      attrs['stroke-width'] = stroke;
      attrs['stroke-linejoin'] = 'miter';
      attrs['stroke-miterlimit'] = SHAPE_MITER_LIMIT;
      if (dash.length > 0) {
        attrs['stroke-dasharray'] = dash.join(' ');
      }
    }
    attrs.class = classes.join(' ');
    const path = svgEl('path', attrs);
    if (f.fill && fillColor) {
      setPaint(path, 'fill', paintOf(f.shade ? applyFillMode(fillColor, f.shade) : fillColor));
    }
    if (f.stroke) {
      setPaint(path, 'stroke', outlinePaint);
    }
    svg.append(path);
  }
  return svg;
}

/** 図形の下地の `<svg>`（ISSUE-96）：形の塗りの path は白で塗り、枠線は線の縁取り（太さ＋片側 3px・実線）。何も無ければ null */
function buildShapeBackdrop(shape: ShapeDrawn, rect: PxRect): SVGSVGElement | null {
  if (shape.backdrop.length === 0) {
    return null;
  }
  const svg = backdropSvg(rect, BACKDROP_PAD);
  const border = num(shape.strokeWidth + 2 * BACKDROP_PAD);
  for (const b of shape.backdrop) {
    svg.append(backdropPart('path', { d: b.d }, b.fill, b.stroke ? border : undefined));
  }
  return svg;
}

/**
 * 図が Excel の色を 1 つでも持つか（塗り・枠線・線・文字のランの色）。持たない図（色を省略した SheetDrawing）は暫定のテーマ変数
 * （`--xs-fg` など）で描かれ、白い下地に乗ると読めないため、下地を敷かない（ISSUE-96）。
 */
function hasAnyColor(drawing: SheetDrawing): boolean {
  return (
    !!drawing.fillColor ||
    !!drawing.outlineColor ||
    !!drawing.line?.color ||
    !!drawing.text?.paragraphs.some((p) => p.runs.some((r) => !!r.color))
  );
}

/** 代替の矩形の文字枠の下地（ISSUE-96）：回転前の矩形（`r`）の位置に置き、図の要素と同じく中心で回す。余白は付けない（縁取りの無い塗りだけ） */
function buildFrameBackdrop(r: PxRect, deg: number): SVGSVGElement {
  const svg = backdropSvg(r, 0);
  if (deg !== 0) {
    svg.style.transform = `rotate(${deg}deg)`;
  }
  svg.append(backdropPart('rect', { width: num(r.w), height: num(r.h) }, true));
  return svg;
}

/**
 * 形で描く図形の要素を作る（ADR-0028 §決定 5）：`div.xs-figure`（外接矩形）＞ `<svg>` ＋ `div.xs-shape-box`（回転前の矩形）＞ `div.xs-shape-text`（文字領域）。
 * 文字は図形の矩形（box）で切り、文字領域では切らない。回転は box ごと中心で掛け、反転は文字領域の位置（`textRect`）に反映済みで文字は鏡写しにしない。
 * 上下反転（`flipV`・`flipHV`）の図形は、Excel と同じく文字を 180° 回す（文字領域の中心まわり。領域の位置は鏡写し済み。左右反転は文字を回さない。ISSUE-93）。
 */
function buildShape(shape: ShapeDrawn, rect: PxRect, drawing: SheetDrawing): HTMLElement {
  const el = document.createElement('div');
  el.className = 'xs-figure';
  el.style.left = `${rect.x}px`;
  el.style.top = `${rect.y}px`;
  el.style.width = `${rect.w}px`;
  el.style.height = `${rect.h}px`;
  if (shape.faces.length > 0) {
    el.append(buildShapeSvg(shape, rect, drawing));
  }
  if (shapeHasText(drawing)) {
    const box = document.createElement('div');
    box.className = 'xs-shape-box';
    box.style.left = `${num(shape.box.x - rect.x)}px`;
    box.style.top = `${num(shape.box.y - rect.y)}px`;
    box.style.width = `${num(shape.box.w)}px`;
    box.style.height = `${num(shape.box.h)}px`;
    const deg = rotationDeg(drawing.rot);
    if (deg !== 0) {
      box.style.transform = `rotate(${deg}deg)`;
    }
    const tr = shape.textRect;
    const text = document.createElement('div');
    text.className = 'xs-shape-text';
    text.style.left = `${num(tr.x)}px`;
    text.style.top = `${num(tr.y)}px`;
    text.style.width = `${num(tr.w)}px`;
    text.style.height = `${num(tr.h)}px`;
    if (drawing.flipV === true) {
      text.style.transform = 'rotate(180deg)';
    }
    fillText(text, drawing.text!, tr.w, tr.h);
    box.append(text);
    el.append(box);
  }
  return el;
}

/**
 * 画像レイヤーを作る。
 *
 * @param deps テスト用に差し替える依存（blob URL の作成・解放）。通常は省略（`URL.createObjectURL`・`revokeObjectURL`）
 */
export function createImageLayer(
  deps: {
    createUrl?: (blob: Blob) => string;
    revokeUrl?: (url: string) => void;
    /** 破線枠の文言（既定は日本語） */
    messages?: Readonly<XSheetMessages>;
  } = {},
): ImageLayer {
  const messages = deps.messages ?? MESSAGES_JA;
  const element = document.createElement('div');
  element.className = 'xs-layer';

  let sheet: LayerSheet | null = null;
  let store: ImageStore | null = null;
  /** 今 DOM にある図（図の位置 → 要素）。重ね順は `order`（図の位置の昇順） */
  let mounted = new Map<number, Figure>();
  let order: number[] = [];
  let disposed = false;

  /** DOM の図をすべて外す */
  const unmountAll = (): void => {
    mounted = new Map();
    order = [];
    element.replaceChildren();
  };

  /** 画像の読み込み結果を図へ反映する（`<img>` か破線枠） */
  const applyEntry = (fig: Figure, entry: ImageEntry): void => {
    if (entry.ok) {
      const img = document.createElement('img');
      img.decoding = 'async';
      img.alt = fig.drawing.descr || fig.drawing.name;
      img.draggable = false;
      img.className = 'xs-img';
      // 壊れた画像（デコードできない）は破線枠へ差し替える。以後 onerror は呼ばれない
      img.onerror = () => fillFrame(fig.el, messages.imageBroken);
      img.src = entry.url;
      fig.el.replaceChildren(img);
    } else {
      fillFrame(fig.el, frameLabel(fig.drawing, entry, messages));
    }
  };

  /** 図 1 つの要素を作る。画像は読み込み済みなら反映し、未読なら読み込みを始める（読み込み中は空） */
  const mount = (box: DrawingBox): Figure => {
    const drawing = sheet!.drawings[box.index]!;
    if (drawing.kind === 'line' && box.line) {
      return {
        box,
        drawing,
        el: buildLine(box.line, box.rect, drawing.line?.color),
        backdrop: hasAnyColor(drawing) ? buildLineBackdrop(box.line, box.rect) : undefined,
      };
    }
    if (box.shape) {
      return {
        box,
        drawing,
        el: buildShape(box.shape, box.rect, drawing),
        backdrop: hasAnyColor(drawing)
          ? (buildShapeBackdrop(box.shape, box.rect) ?? undefined)
          : undefined,
      };
    }
    const el = document.createElement('div');
    el.className = 'xs-figure';
    const fig: Figure = { box, drawing, el };
    if (hasShapeText(drawing)) {
      // 文字のある図形：回転前の矩形に置いて回す（box.rect は可視判定用のアンカー矩形）。反転は掛けない。
      // グループの中の図の矩形は xfrm（回転前の矩形）から写したものなので、縦横を入れ替えない（ISSUE-89）
      const r = drawing.group
        ? { ...box.rect }
        : preRotationRect(box.rect, drawing.rot, xfrmExtPx(drawing));
      el.style.left = `${r.x}px`;
      el.style.top = `${r.y}px`;
      el.style.width = `${r.w}px`;
      el.style.height = `${r.h}px`;
      el.classList.add('xs-shape-text');
      el.classList.toggle('has-outline', drawing.hasOutline === true);
      el.classList.toggle('has-fill', drawing.hasFill === true);
      const deg = rotationDeg(drawing.rot);
      if (deg !== 0) {
        el.style.transform = `rotate(${deg}deg)`;
      }
      // 塗りは塗りの色を背景、枠線は枠線の色の破線に（色が無ければ暫定のテーマ変数のまま）
      const fillPaint = drawing.hasFill === true ? paintOf(drawing.fillColor) : null;
      if (fillPaint) {
        el.style.backgroundColor = cssColor(fillPaint);
      }
      const outlinePaint = drawing.hasOutline === true ? paintOf(drawing.outlineColor) : null;
      if (outlinePaint) {
        el.style.outlineColor = cssColor(outlinePaint);
      }
      fillText(el, drawing.text!, r.w, r.h);
      // 下地は、見える塗り（不透明度 0 超）か見える文字があるときだけ。色が無い図には敷かない
      if (hasAnyColor(drawing) && (filledShown(drawing) || textShown(drawing))) {
        fig.backdrop = buildFrameBackdrop(r, deg);
      }
      return fig;
    }
    el.style.left = `${box.rect.x}px`;
    el.style.top = `${box.rect.y}px`;
    el.style.width = `${box.rect.w}px`;
    el.style.height = `${box.rect.h}px`;
    if (drawing.kind !== 'image') {
      fillFrame(el, frameLabel(drawing, undefined, messages));
    } else if (drawing.mediaPath === null || !store) {
      fillFrame(el, messages.imageBroken);
    } else {
      const path = drawing.mediaPath;
      const known = store.peek(path);
      if (known) {
        applyEntry(fig, known);
      } else {
        const current = store;
        void current.load(path).then((entry) => {
          // 破棄・シートの切替で null になった／別のシートの図になった結果は捨てる
          if (!entry || disposed || store !== current) {
            return;
          }
          // この図がまだ DOM にある（外れていたら、次に出す時に peek で反映する）
          if (mounted.get(box.index) === fig) {
            applyEntry(fig, entry);
          }
        });
      }
    }
    return fig;
  };

  return {
    element,
    setSheet(next) {
      if (disposed) {
        return;
      }
      store?.dispose();
      store = null;
      unmountAll();
      sheet = next;
      if (next) {
        store = createImageStore({
          read: next.readImage,
          createUrl: deps.createUrl,
          revokeUrl: deps.revokeUrl,
          onOverBudget: next.onOverBudget,
        });
      }
    },
    place(left, top, width, height) {
      element.style.left = `${left}px`;
      element.style.top = `${top}px`;
      element.style.width = `${Math.max(0, width)}px`;
      element.style.height = `${Math.max(0, height)}px`;
    },
    update(view) {
      if (disposed || !sheet) {
        return;
      }
      const wanted = pickVisible(sheet.boxes, view);
      if (wanted.length === order.length && wanted.every((b, i) => b.index === order[i])) {
        return; // 出ている図が同じ（スクロールのたびの DOM 操作を避ける）
      }
      const next = new Map<number, Figure>();
      for (const box of wanted) {
        next.set(box.index, mounted.get(box.index) ?? mount(box));
      }
      mounted = next;
      order = wanted.map((b) => b.index);
      // 重ね順は drawing の順（wanted は元の順）。既存の要素は再利用する。下地（z-index: -1）は図の直前に置く（全図の下に描かれる）
      const nodes: Element[] = [];
      for (const b of wanted) {
        const fig = next.get(b.index)!;
        if (fig.backdrop) {
          nodes.push(fig.backdrop);
        }
        nodes.push(fig.el);
      }
      element.replaceChildren(...nodes);
    },
    dispose() {
      if (disposed) {
        return;
      }
      disposed = true;
      store?.dispose();
      store = null;
      sheet = null;
      unmountAll();
    },
  };
}
