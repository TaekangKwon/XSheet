// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createGeometry, type SheetDrawing } from '../../src/core/index';
import {
  emptyLayout,
  type DrawingGroup,
  type DrawingLine,
  type GroupRect,
} from '../../src/core/types';
import {
  groupChildRect,
  layoutDrawings,
  placeDrawing,
  viewWindow,
  EMU_PER_PX,
} from '../../src/render/drawings';
import { createImageLayer } from '../../src/render/image-layer';
import { renderLine } from '../../src/render/shapes';
import { MESSAGES_JA } from '../../src/render/messages';

/**
 * グループの中の図の位置と描画（ISSUE-89・ADR-0027 §決定 5）。
 * - `groupChildRect`：グループの子の座標系（`chOff`・`chExt`）から、最も外側のグループのアンカー（px）への線形な写像。
 * - `layoutDrawings`：写した矩形で配置。回転した子は縦横を入れ替えない（トップレベルの図のアンカーにだけ掛ける）。
 * - 画像レイヤー（jsdom）：子が文字の枠・SVG の線・`<img>` として出て、グループの破線枠は出ない。
 * 破線枠の文字は既定の文言（日本語の `MESSAGES_JA`）。
 */

const rect = (x: number, y: number, cx: number, cy: number): GroupRect => ({ x, y, cx, cy });
const NO_GROUP_OWN = null;

/** 最も外側のグループの座標系 1 段だけのグループ */
const oneLevel = (ch: GroupRect, xfrm: GroupRect): DrawingGroup => ({
  levels: [{ ch, own: NO_GROUP_OWN }],
  xfrm,
});

describe('TD-89-7 groupChildRect：座標系 → アンカーの矩形（px）への線形な写像', () => {
  const anchor = { x: 100, y: 50, w: 400, h: 300 };

  it('chOff が 0 でない・拡大縮小あり：x = 矩形.x + (子.off.x − chOff.x) × 矩形.w / chExt.cx（y・幅・高さも同じ）', () => {
    // 4000×3000 の座標系を 400×300 px に（0.1 倍）。chOff=(1000, 2000)
    const g = oneLevel(rect(1000, 2000, 4000, 3000), rect(2000, 3000, 1000, 750));
    expect(groupChildRect(anchor, g)).toEqual({ x: 200, y: 150, w: 100, h: 75 });
  });

  it('拡大（chExt がアンカーより小さい）：縦と横を別の倍率で写す', () => {
    // 幅 2 倍・高さ 3 倍
    const g = oneLevel(rect(0, 0, 200, 100), rect(10, 20, 30, 40));
    expect(groupChildRect(anchor, g)).toEqual({ x: 100 + 20, y: 50 + 60, w: 60, h: 120 });
  });

  it('入れ子 2 段：内側のグループ自身の off・ext を外側の座標系から写し、その矩形を次の親にする', () => {
    const g: DrawingGroup = {
      levels: [
        { ch: rect(0, 0, 1000, 1000), own: null },
        { ch: rect(0, 0, 200, 200), own: rect(500, 500, 400, 400) },
      ],
      xfrm: rect(100, 50, 50, 100),
    };
    // 外側：400×300 px に 1000×1000 → 倍率 (0.4, 0.3)。内側の矩形 = (100+200, 50+150, 160, 120)
    // 内側：160×120 px に 200×200 → 倍率 (0.8, 0.6)。図 = (300+80, 200+30, 40, 60)
    expect(groupChildRect(anchor, g)).toEqual({ x: 380, y: 230, w: 40, h: 60 });
  });

  it('入れ子 3 段でも外側から順に写す（グループ自身の位置・倍率が掛け合わさる）', () => {
    const g: DrawingGroup = {
      levels: [
        { ch: rect(0, 0, 100, 100), own: null },
        { ch: rect(0, 0, 100, 100), own: rect(0, 0, 50, 50) },
        { ch: rect(0, 0, 100, 100), own: rect(50, 50, 50, 50) },
      ],
      xfrm: rect(0, 0, 100, 100),
    };
    // 400×300 → 内側 1：(0,0,200,150) → 内側 2：(100,75,100,75)。図は内側 2 の全体
    expect(groupChildRect({ x: 0, y: 0, w: 400, h: 300 }, g)).toEqual({
      x: 100,
      y: 75,
      w: 100,
      h: 75,
    });
  });

  it('端を四捨五入し、幅・高さは端の差で出す（0.5px ずれた子の幅が 1px ずれない）', () => {
    const g = oneLevel(rect(0, 0, 3, 3), rect(1, 1, 1, 1));
    // 100×100 px を 3 等分：x 33.33〜66.67 → 33〜67、幅 34
    expect(groupChildRect({ x: 0, y: 0, w: 100, h: 100 }, g)).toEqual({
      x: 33,
      y: 33,
      w: 34,
      h: 34,
    });
  });

  it('負の位置の子・アンカーの外へはみ出す子は、矩形を切らず負の位置のまま持つ（大きさが縮まない。隠すのはレイヤー）', () => {
    const g = oneLevel(rect(0, 0, 100, 100), rect(-50, -20, 100, 100));
    expect(groupChildRect({ x: 0, y: 0, w: 100, h: 100 }, g)).toEqual({
      x: -50,
      y: -20,
      w: 100,
      h: 100,
    });
    // 完全に外でも大きさは保つ
    const out = oneLevel(rect(0, 0, 100, 100), rect(-300, 0, 100, 100));
    expect(groupChildRect({ x: 0, y: 0, w: 100, h: 100 }, out)).toEqual({
      x: -300,
      y: 0,
      w: 100,
      h: 100,
    });
  });

  it('座標系の大きさが 0 以下・自身の位置が無い入れ子・段が無い：null（有限でない値も null）', () => {
    expect(groupChildRect(anchor, oneLevel(rect(0, 0, 0, 100), rect(0, 0, 1, 1)))).toBeNull();
    expect(groupChildRect(anchor, oneLevel(rect(0, 0, 100, -1), rect(0, 0, 1, 1)))).toBeNull();
    expect(
      groupChildRect(anchor, {
        levels: [
          { ch: rect(0, 0, 100, 100), own: null },
          { ch: rect(0, 0, 100, 100), own: null },
        ],
        xfrm: rect(0, 0, 1, 1),
      }),
    ).toBeNull();
    expect(groupChildRect(anchor, { levels: [], xfrm: rect(0, 0, 1, 1) })).toBeNull();
    expect(
      groupChildRect(anchor, oneLevel(rect(0, 0, 100, 100), rect(Number.NaN, 0, 1, 1))),
    ).toBeNull();
  });

  it('アンカーの幅・高さが 0 以下（逆向きの twoCell）は 0 とみなす：子の大きさも 0 になる', () => {
    const g = oneLevel(rect(0, 0, 100, 100), rect(10, 10, 50, 50));
    expect(groupChildRect({ x: 40, y: 40, w: -5, h: 0 }, g)).toEqual({ x: 40, y: 40, w: 0, h: 0 });
  });

  it('巨大な座標でも有限の矩形に丸める（位置は ±33,000,000px、大きさは 33,000,000px まで。片方の次元だけ 0 の線でも）', () => {
    const MAX = 33_000_000;
    const g = oneLevel(rect(0, 0, 1, 1), rect(0, 0, 1e15, 1e15));
    const r = groupChildRect({ x: 0, y: 0, w: 1e6, h: 1e6 }, g)!;
    expect(r.x + r.w).toBeLessThanOrEqual(MAX);
    expect(r.y + r.h).toBeLessThanOrEqual(MAX);
    // 位置が桁違いに大きい・小さい子（水平の線＝高さ 0）
    const far = oneLevel(rect(0, 0, 1, 1), rect(1e15, -1e15, 1e15, 0));
    const f = groupChildRect({ x: 0, y: 0, w: 1e6, h: 1e6 }, far)!;
    expect(f.x).toBe(MAX);
    expect(f.y).toBe(-MAX);
    expect(f.w).toBeLessThanOrEqual(MAX);
    expect(f.h).toBe(0);
  });
});

// ---------------------------------------------------------------------------

const px = (n: number): number => n * EMU_PER_PX;
/** 既定の列幅 64px・行高 20px の geometry */
const geo = createGeometry(emptyLayout(), 1000, 100);

/** twoCell（A1 から 列 cols・行 rows ぶん）のアンカー */
const anchorOf = (cols: number, rows: number): SheetDrawing['anchor'] => ({
  type: 'twoCell',
  from: { col: 0, colOff: 0, row: 0, rowOff: 0 },
  to: { col: cols, colOff: 0, row: rows, rowOff: 0 },
});

const LINE: DrawingLine = {
  geom: 'straightConnector1',
  adj: [50000, 50000, 50000],
  widthEmu: 12700,
  dash: 'solid',
  head: null,
  tail: { type: 'triangle', w: 'med', len: 'med' },
};

/** グループの子（A1 起点・4 列×3 行＝256×60px のアンカーに、座標系＝0〜2560×600 の恒等な 0.1 倍） */
function child(over: Partial<SheetDrawing>, xfrm: GroupRect): SheetDrawing {
  return {
    anchor: anchorOf(4, 3),
    kind: 'shape',
    name: 'c',
    descr: '',
    mediaPath: null,
    group: oneLevel(rect(0, 0, 2560, 600), xfrm),
    ...over,
  };
}

describe('TD-89-7 layoutDrawings：グループの中の図は、写した矩形で配置する', () => {
  it('子の矩形は写像のとおり。トップレベルの図はアンカーのまま。重ね順（index）は一覧の順', () => {
    const top: SheetDrawing = {
      anchor: anchorOf(1, 1),
      kind: 'chart',
      name: 't',
      descr: '',
      mediaPath: null,
    };
    const boxes = layoutDrawings(geo, [top, child({}, rect(256, 100, 640, 200))]);
    expect(boxes).toEqual([
      { index: 0, rect: { x: 0, y: 0, w: 64, h: 20 } },
      { index: 1, rect: { x: 26, y: 10, w: 64, h: 20 } },
    ]);
  });

  it('大きさ 0 の子（xfrm の幅か高さが 0）は出さない。線は幅か高さの片方だけ 0 でも出す', () => {
    const boxes = layoutDrawings(geo, [
      child({ name: 'zero' }, rect(0, 0, 0, 100)),
      child({ kind: 'line', name: 'hline', line: LINE }, rect(0, 300, 2560, 0)),
      child({ kind: 'line', name: 'dot', line: LINE }, rect(0, 0, 0, 0)),
    ]);
    expect(boxes.map((b) => b.index)).toEqual([1]);
    expect(boxes[0]!.line).toBeDefined();
  });

  it('座標系が壊れたグループの子（chExt が 0）は出さない', () => {
    const bad = child({}, rect(0, 0, 100, 100));
    bad.group = oneLevel(rect(0, 0, 0, 600), rect(0, 0, 100, 100));
    expect(layoutDrawings(geo, [bad])).toEqual([]);
  });

  it('入れ子の中で展開できなかったグループ（kind が group・親の座標系の位置）も、写した矩形の破線枠になる', () => {
    const boxes = layoutDrawings(geo, [
      child({ kind: 'group', name: 'deep' }, rect(1280, 300, 640, 120)),
    ]);
    expect(boxes).toEqual([{ index: 0, rect: { x: 128, y: 30, w: 64, h: 12 } }]);
  });

  /** 1 px = 1 単位の座標系（256×60 px の絶対アンカー）に置いた線の子 */
  function lineChild(xfrm: GroupRect, rot: number): SheetDrawing {
    return {
      anchor: { type: 'absolute', x: 0, y: 0, cx: px(256), cy: px(60) },
      kind: 'line',
      name: 'l',
      descr: '',
      mediaPath: null,
      line: LINE,
      rot,
      flipH: false,
      flipV: false,
      group: oneLevel(rect(0, 0, 256, 60), xfrm),
    };
  }
  const BIG = { ...geo, totalWidth: 5000, totalHeight: 5000 };
  const SHEET = { totalWidth: 5000, totalHeight: 5000 };

  it('子の回転は縦横を入れ替えない：rot=90° の子の線（xfrm＝回転前の 100×20）は、回転後の外接（20×100）をアンカーに持つトップレベルの線と同じ外接矩形になる', () => {
    const rot = 5400000;
    // 中心 (100, 100)。トップレベルのアンカーは回転後の外接 (90, 50, 20, 100)
    const top = renderLine({ x: 90, y: 50, w: 20, h: 100 }, LINE, { rot }, SHEET);
    const boxes = layoutDrawings(BIG, [lineChild(rect(50, 90, 100, 20), rot)]);
    expect(boxes[0]!.rect).toEqual(top.bbox);
    expect(boxes[0]!.line!.d).toBe(top.d);
    // 入れ替えていたら（20×100 を xfrm とみなして回すと）横長になってしまう
    expect(boxes[0]!.rect.h).toBeGreaterThan(boxes[0]!.rect.w);
  });

  it('回転 45°（トップレベルなら縦横を入れ替える角度）でも、子は xfrm の矩形のまま回す：回転前の矩形をそのまま renderLine に渡した結果と等しい（w≠h）', () => {
    const rot = 2700000;
    const xfrm = rect(50, 90, 100, 20);
    // 入れ替えを掛けない描画（回転前の矩形をそのまま渡す）／入れ替えを掛けた描画
    const kept = renderLine({ x: 50, y: 90, w: 100, h: 20 }, LINE, { rot }, SHEET, true);
    const swapped = renderLine({ x: 50, y: 90, w: 100, h: 20 }, LINE, { rot }, SHEET, false);
    expect(swapped.d).not.toBe(kept.d); // 入力が入れ替えの有無を見分けられる
    const boxes = layoutDrawings(BIG, [lineChild(xfrm, rot)]);
    expect(boxes[0]!.rect).toEqual(kept.bbox);
    expect(boxes[0]!.line!.d).toBe(kept.d);
  });

  it('placeDrawing：グループの外の図はアンカーの矩形、中の図は写した矩形（大きさの判定はしない）', () => {
    const top: SheetDrawing = {
      anchor: anchorOf(2, 2),
      kind: 'image',
      name: 'p',
      descr: '',
      mediaPath: null,
    };
    expect(placeDrawing(geo, top)).toEqual({ x: 0, y: 0, w: 128, h: 40 });
    expect(placeDrawing(geo, child({}, rect(0, 0, 0, 0)))).toEqual({ x: 0, y: 0, w: 0, h: 0 });
  });
});

// ---------------------------------------------------------------------------
// 画像レイヤー（DOM）
// ---------------------------------------------------------------------------

describe('TD-89-9 画像レイヤー：グループの子が文字の枠・SVG の線・<img>・破線枠として出る（グループの破線枠は出ない）', () => {
  const ALL = viewWindow(0, 0, 3000, 3000, 0, 0);

  /** 文字を持つテキストボックスの子 */
  const textChild = (): SheetDrawing =>
    child(
      {
        kind: 'textbox',
        name: 'tb',
        text: {
          paragraphs: [
            {
              runs: [
                {
                  br: false,
                  text: '中の文字',
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
        },
        rot: 0,
        flipH: false,
        flipV: false,
        hasOutline: true,
        hasFill: true,
      },
      rect(256, 100, 640, 200),
    );

  function mountAll(drawings: SheetDrawing[]) {
    const layer = createImageLayer({ createUrl: () => 'blob:x', revokeUrl: () => undefined });
    document.body.append(layer.element);
    const boxes = layoutDrawings({ ...geo, totalWidth: 5000, totalHeight: 5000 }, drawings);
    layer.setSheet({
      drawings,
      boxes,
      readImage: async () => ({
        ok: true,
        format: 'png',
        mime: 'image/png',
        width: 40,
        height: 30,
        bytes: new Uint8Array([1]),
      }),
    });
    layer.update(ALL);
    return layer;
  }

  it('文字の枠（.xs-shape-text）・SVG の線・種類名の破線枠が、写した矩形の位置に出る。グループの破線枠（spreadsheet.drawing.group）は無い', () => {
    const layer = mountAll([
      textChild(),
      child({ kind: 'line', name: 'ln', line: LINE }, rect(0, 300, 2560, 0)),
      child({ kind: 'shape', name: 'plain' }, rect(1280, 300, 640, 120)),
    ]);
    const figs = [...layer.element.querySelectorAll<HTMLElement>('.xs-figure')];
    expect(figs).toHaveLength(3);
    // 文字の枠：xfrm (256,100,640,200) の 0.1 倍＝(26,10,64,20)。子は回転前の矩形そのもの
    expect(figs[0]!.classList.contains('xs-shape-text')).toBe(true);
    expect([
      figs[0]!.style.left,
      figs[0]!.style.top,
      figs[0]!.style.width,
      figs[0]!.style.height,
    ]).toEqual(['26px', '10px', '64px', '20px']);
    expect(figs[0]!.textContent).toBe('中の文字');
    // 線：SVG
    expect(figs[1]!.matches('svg.xs-line')).toBe(true);
    // 文字の無い図形：種類名の破線枠
    expect(figs[2]!.querySelector('.xs-frame')?.textContent).toBe(MESSAGES_JA.shape);
    expect(layer.element.textContent).not.toContain(MESSAGES_JA.group);
  });

  it('回転した文字入りの子は、縦横を入れ替えず xfrm の矩形に置いて回す（rot=90°・幅 100・高さ 20）', () => {
    const d = textChild();
    d.rot = 5400000;
    d.group = oneLevel(rect(0, 0, 256, 60), rect(50, 20, 100, 20));
    const layer = mountAll([d]);
    const fig = layer.element.querySelector<HTMLElement>('.xs-figure')!;
    expect([fig.style.left, fig.style.top, fig.style.width, fig.style.height]).toEqual([
      '50px',
      '20px',
      '100px',
      '20px',
    ]);
    expect(fig.style.transform).toBe('rotate(90deg)');
  });

  it('グループの中の画像は <img> で出る（グループの外と同じ readImage・blob URL の経路）', async () => {
    const layer = mountAll([
      child(
        { kind: 'image', name: 'pic', mediaPath: 'xl/media/image1.png' },
        rect(256, 100, 640, 200),
      ),
    ]);
    await new Promise((r) => setTimeout(r, 10));
    const img = layer.element.querySelector<HTMLImageElement>('.xs-figure img');
    expect(img).not.toBeNull();
    expect(img!.getAttribute('src')).toBe('blob:x');
  });

  it('グループの外の画像と同じ画像を共有しても、blob URL は 1 つ（ISSUE-85・86 の画像の保管庫をそのまま使う）', async () => {
    let n = 0;
    const layer = createImageLayer({
      createUrl: () => `blob:s/${++n}`,
      revokeUrl: () => undefined,
    });
    document.body.append(layer.element);
    const top: SheetDrawing = {
      anchor: anchorOf(1, 1),
      kind: 'image',
      name: 'top',
      descr: '',
      mediaPath: 'xl/media/image1.png',
    };
    const drawings = [
      top,
      child(
        { kind: 'image', name: 'in', mediaPath: 'xl/media/image1.png' },
        rect(256, 100, 640, 200),
      ),
    ];
    layer.setSheet({
      drawings,
      boxes: layoutDrawings({ ...geo, totalWidth: 5000, totalHeight: 5000 }, drawings),
      readImage: async () => ({
        ok: true,
        format: 'png',
        mime: 'image/png',
        width: 40,
        height: 30,
        bytes: new Uint8Array([1]),
      }),
    });
    layer.update(ALL);
    await new Promise((r) => setTimeout(r, 10));
    const imgs = [...layer.element.querySelectorAll('img')];
    expect(imgs.map((i) => i.getAttribute('src'))).toEqual(['blob:s/1', 'blob:s/1']);
    expect(n).toBe(1);
  });

  it('展開できなかったグループ（kind が group）は、親の座標系の位置に「グループ」の破線枠のまま出る', () => {
    const layer = mountAll([child({ kind: 'group', name: 'deep' }, rect(1280, 300, 640, 120))]);
    const fig = layer.element.querySelector<HTMLElement>('.xs-figure')!;
    expect(fig.querySelector('.xs-frame')?.textContent).toBe(MESSAGES_JA.group);
    expect([fig.style.left, fig.style.top, fig.style.width, fig.style.height]).toEqual([
      '128px',
      '30px',
      '64px',
      '12px',
    ]);
  });
});
