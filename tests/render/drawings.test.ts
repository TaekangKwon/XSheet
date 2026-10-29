import { describe, expect, it } from 'vitest';
import { createGeometry, type SheetDrawing } from '../../src/core/index';
import { emptyLayout, type DrawingLine, type SheetLayout } from '../../src/core/types';
import {
  anchorPointX,
  anchorRect,
  drawingExtent,
  layoutDrawings,
  pickVisible,
  rectsIntersect,
  viewWindow,
  EMU_PER_PX,
} from '../../src/render/drawings';

/** 指定した項目だけ替えた layout（既定は emptyLayout：MDW 7＝列 64px・行 20px） */
const layoutOf = (over: Partial<SheetLayout> = {}): SheetLayout => ({ ...emptyLayout(), ...over });

const px = (n: number): number => n * EMU_PER_PX;
const pt = (col: number, colOff: number, row: number, rowOff: number) => ({
  col,
  colOff,
  row,
  rowOff,
});

/** 位置の計算に使う geometry（既定の寸法 64×20） */
const geo = createGeometry(layoutOf(), 50, 20);

const draw = (anchor: SheetDrawing['anchor'], over: Partial<SheetDrawing> = {}): SheetDrawing => ({
  anchor,
  kind: 'image',
  name: 'p',
  descr: '',
  mediaPath: 'xl/media/image1.png',
  ...over,
});

describe('TD-86-1 anchorRect：アンカー → A1 の左上基準の px（geometry の番号は 1 始まり・アンカーは 0 始まり）', () => {
  it('twoCell：from の列・行の左上＋ずれ 〜 to の同様。番号の 0/1 始まりが 1 つずれない', () => {
    // C 列（col 2）左端＋4px・行 4（row 3）上端＋3px から、E 列（col 4）左端＋10px・行 6（row 5）上端＋5px まで
    const r = anchorRect(geo, {
      type: 'twoCell',
      from: pt(2, px(4), 3, px(3)),
      to: pt(4, px(10), 5, px(5)),
    });
    expect(r).toEqual({ x: 2 * 64 + 4, y: 3 * 20 + 3, w: 2 * 64 + 6, h: 2 * 20 + 2 });
  });

  it('ずれは EMU / 9525 を四捨五入した整数 px（Excel の rowOff 25400 は約 3px）', () => {
    const r = anchorRect(geo, {
      type: 'twoCell',
      from: pt(0, 38100, 0, 25400),
      to: pt(1, 0, 1, 0),
    });
    expect(r).toEqual({ x: 4, y: 3, w: 64 - 4, h: 20 - 3 });
  });

  it('ずれが列・行の大きさを超える時は、セルの端で止める（隣のセルへ飛び出さない）', () => {
    expect(anchorPointX(geo, pt(1, px(500), 0, 0))).toBe(2 * 64); // 列 B の右端＝列 C の左端
    const r = anchorRect(geo, {
      type: 'twoCell',
      from: pt(0, px(500), 0, px(500)),
      to: pt(2, 0, 2, 0),
    });
    // from は A1 の右下の角（64, 20）で止まる
    expect(r).toEqual({ x: 64, y: 20, w: 2 * 64 - 64, h: 2 * 20 - 20 });
  });

  it('負・非有限のずれは 0 として扱う（NaN・巨大な値で座標が壊れない）', () => {
    const r = anchorRect(geo, {
      type: 'twoCell',
      from: pt(0, -5, 0, Number.NaN),
      to: pt(1, Number.POSITIVE_INFINITY, 1, 0),
    });
    expect(r).toEqual({ x: 0, y: 0, w: 64 + 0, h: 20 }); // Infinity は 0（無効値は採らない）
  });

  it('oneCell：from を起点に cx×cy（EMU → px）', () => {
    const r = anchorRect(geo, {
      type: 'oneCell',
      from: pt(1, px(2), 2, 0),
      cx: px(40),
      cy: px(30),
    });
    expect(r).toEqual({ x: 66, y: 40, w: 40, h: 30 });
  });

  it('absolute：シートの左上からの x・y に cx×cy', () => {
    const r = anchorRect(geo, { type: 'absolute', x: px(100), y: px(50), cx: px(64), cy: px(48) });
    expect(r).toEqual({ x: 100, y: 50, w: 64, h: 48 });
  });

  it('大きさ 0 以下の図は出さない（0 の ext・逆向きの twoCell）', () => {
    expect(
      anchorRect(geo, { type: 'oneCell', from: pt(0, 0, 0, 0), cx: 0, cy: px(10) }),
    ).toBeNull();
    expect(anchorRect(geo, { type: 'absolute', x: 0, y: 0, cx: px(10), cy: 0 })).toBeNull();
    expect(
      anchorRect(geo, { type: 'twoCell', from: pt(3, 0, 3, 0), to: pt(1, 0, 1, 0) }),
    ).toBeNull();
    expect(
      anchorRect(geo, { type: 'twoCell', from: pt(1, 0, 1, 0), to: pt(1, 0, 1, 0) }),
    ).toBeNull();
  });

  it('非表示の列を跨ぐ twoCell：幅 0 の列は 0 として畳まれ、画像は表と同じ位置に並ぶ', () => {
    // 列 D（4）が非表示。C 列左端＋4px から E 列（5）の左端＋10px まで
    const g = createGeometry(
      layoutOf({ cols: [{ min: 4, max: 4, width: null, hidden: true }] }),
      20,
      10,
    );
    expect(g.colWidth(4)).toBe(0);
    const r = anchorRect(g, { type: 'twoCell', from: pt(2, px(4), 0, 0), to: pt(4, px(10), 1, 0) });
    expect(r).toEqual({ x: 2 * 64 + 4, y: 0, w: 64 - 4 + 0 + 10, h: 20 });
  });

  it('非表示の行・列だけにかかる図（幅または高さが 0）は null', () => {
    const g = createGeometry(
      layoutOf({
        cols: [{ min: 2, max: 2, width: null, hidden: true }],
        rows: [{ min: 3, max: 3, ht: null, custom: false, hidden: true }],
      }),
      20,
      10,
    );
    // 列 B だけ（col 1 の左端〜右端）
    expect(
      anchorRect(g, { type: 'twoCell', from: pt(1, 0, 0, 0), to: pt(1, px(30), 2, 0) }),
    ).toBeNull();
    // 行 3 だけ
    expect(
      anchorRect(g, { type: 'twoCell', from: pt(0, 0, 2, 0), to: pt(2, 0, 2, px(10)) }),
    ).toBeNull();
  });

  it('範囲外の列・行（既定の寸法で延びる）にも置ける。上限（Excel の最大）を超える番号は丸める', () => {
    const r = anchorRect(geo, {
      type: 'oneCell',
      from: pt(100, 0, 100, 0),
      cx: px(10),
      cy: px(10),
    });
    expect(r).toEqual({ x: 100 * 64, y: 100 * 20, w: 10, h: 10 });
    const huge = anchorRect(geo, {
      type: 'oneCell',
      from: pt(2 ** 40, 0, 2 ** 40, 0),
      cx: px(10),
      cy: px(10),
    });
    expect(huge).not.toBeNull();
    expect(Number.isFinite(huge!.x) && Number.isFinite(huge!.y)).toBe(true);
  });
});

describe('TD-86-2 layoutDrawings・drawingExtent：配置と表の広げ方', () => {
  const list: SheetDrawing[] = [
    draw({ type: 'oneCell', from: pt(0, 0, 0, 0), cx: px(10), cy: px(10) }), // 0
    draw({ type: 'oneCell', from: pt(0, 0, 0, 0), cx: 0, cy: 0 }), // 1：出さない
    draw(
      { type: 'twoCell', from: pt(2, 0, 4, 0), to: pt(5, px(1), 9, px(1)) },
      { kind: 'chart', mediaPath: null },
    ), // 2
  ];

  it('大きさ 0 の図を除き、index は元の位置のまま（重ね順は元の順）', () => {
    const boxes = layoutDrawings(geo, list);
    expect(boxes.map((b) => b.index)).toEqual([0, 2]);
    expect(boxes[1]!.rect).toEqual({ x: 128, y: 80, w: 3 * 64 + 1, h: 5 * 20 + 1 });
  });

  it('最下行・最右列：図に含まれる最後の px を含む行・列（端がちょうど境界のとき、次の行・列を足さない）', () => {
    const big = createGeometry(layoutOf(), 1048576, 16384);
    const boxes = layoutDrawings(big, list);
    // 図 2：右端 x=128+193=321 → 列 6（320〜384）、下端 y=80+101=181 → 行 10（180〜200）
    expect(drawingExtent(big, boxes)).toEqual({ rows: 10, cols: 6 });
    // 端が境界ちょうど（幅 64px → 列 1 だけ）
    const edge = layoutDrawings(big, [
      draw({ type: 'oneCell', from: pt(0, 0, 0, 0), cx: px(64), cy: px(20) }),
    ]);
    expect(drawingExtent(big, edge)).toEqual({ rows: 1, cols: 1 });
  });

  it('図が無ければ 0・0', () => {
    expect(drawingExtent(createGeometry(layoutOf(), 10, 10), [])).toEqual({ rows: 0, cols: 0 });
  });
});

describe('TD-86-3 viewWindow・rectsIntersect・pickVisible：出し入れの判定', () => {
  it('可視範囲に上下左右 1 画面分の余白を足し、A1 基準に直す（原点＝行番号の列の幅・見出し行の高さ）', () => {
    const w = viewWindow(100, 200, 800, 400, 32, 24);
    expect(w).toEqual({ x: 100 - 32 - 800, y: 200 - 24 - 400, w: 2400, h: 1200 });
  });

  it('枠の大きさが 0（DOM 未取り付け）なら何も交差しない', () => {
    const w = viewWindow(0, 0, 0, 0, 32, 24);
    expect(w.w).toBe(0);
    expect(rectsIntersect({ x: 0, y: 0, w: 10, h: 10 }, w)).toBe(false);
  });

  it('接するだけは交わらない・1px でも重なれば交わる', () => {
    const a = { x: 0, y: 0, w: 10, h: 10 };
    expect(rectsIntersect(a, { x: 10, y: 0, w: 5, h: 5 })).toBe(false);
    expect(rectsIntersect(a, { x: 9, y: 9, w: 5, h: 5 })).toBe(true);
    expect(rectsIntersect(a, { x: 0, y: 10, w: 5, h: 5 })).toBe(false);
  });

  it('pickVisible：交わる図だけを元の順で返す（画面の 1 画面分外までは出す）', () => {
    const boxes = [
      { index: 0, rect: { x: 0, y: 0, w: 50, h: 50 } },
      { index: 1, rect: { x: 0, y: 5000, w: 50, h: 50 } }, // 遠い
      { index: 2, rect: { x: 0, y: 700, w: 50, h: 50 } }, // 1 画面分（高さ 400）の外側余白に入る
      { index: 3, rect: { x: 9000, y: 0, w: 50, h: 50 } }, // 右に遠い
    ];
    const win = viewWindow(0, 0, 800, 400, 32, 24);
    expect(pickVisible(boxes, win).map((b) => b.index)).toEqual([0, 2]);
    // 縦にスクロールして 1 を可視範囲へ
    const scrolled = viewWindow(0, 4800, 800, 400, 32, 24);
    expect(pickVisible(boxes, scrolled).map((b) => b.index)).toEqual([1]);
  });
});

/** 線の見た目（既定は 1px の直線・矢じりなし） */
function plainLine(over: Partial<DrawingLine> = {}): DrawingLine {
  return {
    geom: 'line',
    adj: [50000, 50000, 50000],
    widthEmu: 9525,
    dash: 'solid',
    head: null,
    tail: null,
    ...over,
  };
}

describe('TD-88-9 layoutDrawings：線は幅か高さが 0 でも残す（両方 0 のときだけ捨てる）', () => {
  const line = (anchor: SheetDrawing['anchor'], over: Partial<SheetDrawing> = {}): SheetDrawing =>
    draw(anchor, { kind: 'line', mediaPath: null, line: plainLine(), ...over });

  it('高さ 0 の水平線・幅 0 の垂直線は残る。両方 0 の線は捨てる', () => {
    const boxes = layoutDrawings(geo, [
      line({ type: 'oneCell', from: pt(1, 0, 1, 0), cx: px(100), cy: 0 }), // 0：水平
      line({ type: 'oneCell', from: pt(1, 0, 1, 0), cx: 0, cy: px(50) }), // 1：垂直
      line({ type: 'oneCell', from: pt(1, 0, 1, 0), cx: 0, cy: 0 }), // 2：点
      line({ type: 'twoCell', from: pt(1, 0, 1, 0), to: pt(1, 0, 3, 0) }), // 3：幅 0 の twoCell
    ]);
    expect(boxes.map((b) => b.index)).toEqual([0, 1, 3]);
    // 外接矩形は太さ分の大きさを持つ（幅・高さとも 1px 以上）
    for (const b of boxes) {
      expect(b.rect.w).toBeGreaterThan(0);
      expect(b.rect.h).toBeGreaterThan(0);
    }
    expect(boxes[0]!.rect.w).toBeGreaterThanOrEqual(100);
    expect(boxes[1]!.rect.h).toBeGreaterThanOrEqual(50);
  });

  it('逆向きの twoCell の線は、負の側を 0 とみなす（もう一方が正なら残る）', () => {
    const boxes = layoutDrawings(geo, [
      line({ type: 'twoCell', from: pt(3, 0, 1, 0), to: pt(1, 0, 3, 0) }), // 幅が負・高さが正
      line({ type: 'twoCell', from: pt(3, 0, 3, 0), to: pt(1, 0, 1, 0) }), // 両方が負
    ]);
    expect(boxes.map((b) => b.index)).toEqual([0]);
  });

  it('線以外の大きさ 0（幅だけ・高さだけ・両方）は従来どおり捨てる。anchorRect の期待も変わらない', () => {
    const boxes = layoutDrawings(geo, [
      draw({ type: 'oneCell', from: pt(0, 0, 0, 0), cx: px(100), cy: 0 }),
      draw({ type: 'oneCell', from: pt(0, 0, 0, 0), cx: 0, cy: px(100) }, { kind: 'shape' }),
      draw({ type: 'oneCell', from: pt(0, 0, 0, 0), cx: 0, cy: 0 }, { kind: 'textbox' }),
    ]);
    expect(boxes).toEqual([]);
    expect(
      anchorRect(geo, { type: 'oneCell', from: pt(0, 0, 0, 0), cx: px(100), cy: 0 }),
    ).toBeNull();
  });

  it('線の DrawingBox は line（描画結果）を持ち、線以外は持たない。line が無い壊れた線は既定の直線で残る', () => {
    const boxes = layoutDrawings(geo, [
      line({ type: 'oneCell', from: pt(1, 0, 1, 0), cx: px(100), cy: px(30) }),
      draw({ type: 'oneCell', from: pt(0, 0, 0, 0), cx: px(10), cy: px(10) }),
      line({ type: 'oneCell', from: pt(1, 0, 1, 0), cx: px(100), cy: px(30) }, { line: undefined }),
    ]);
    expect(boxes.map((b) => b.index)).toEqual([0, 1, 2]);
    expect(boxes[0]!.line).toBeDefined();
    expect(boxes[0]!.rect).toBe(boxes[0]!.line!.bbox);
    expect(boxes[1]!.line).toBeUndefined();
    expect(boxes[1]!.rect).toEqual({ x: 0, y: 0, w: 10, h: 10 });
    expect(boxes[2]!.line!.d).toMatch(/^M[\d.-]+ [\d.-]+ L/);
  });
});

describe('TD-88-10 線の外接矩形で可視判定する（U 字線の張り出し）', () => {
  // 幅 10px・高さ 60px のアンカー（x 400〜410）に、右へ大きく張り出す U 字（bentConnector3 の adj1＝500%）
  const u = draw(
    { type: 'absolute', x: px(400), y: px(100), cx: px(10), cy: px(60) },
    {
      kind: 'line',
      mediaPath: null,
      line: plainLine({ geom: 'bentConnector3', adj: [500000, 50000, 50000], widthEmu: 12700 }),
    },
  );

  it('rect は外接矩形（アンカーの矩形より右へ広い）', () => {
    const [b] = layoutDrawings(geo, [u]);
    expect(b!.rect).toBe(b!.line!.bbox);
    expect(b!.rect.x).toBeLessThanOrEqual(400);
    expect(b!.rect.x + b!.rect.w).toBeGreaterThanOrEqual(450);
    expect(b!.rect.y).toBeLessThanOrEqual(100);
    expect(b!.rect.y + b!.rect.h).toBeGreaterThanOrEqual(160);
  });

  it('pickVisible：アンカーの矩形が窓の外でも、張り出しが窓に入れば可視。張り出しも外なら不可視', () => {
    const boxes = layoutDrawings(geo, [u]);
    expect(rectsIntersect({ x: 400, y: 100, w: 10, h: 60 }, { x: 430, y: 0, w: 100, h: 300 })).toBe(
      false,
    );
    expect(pickVisible(boxes, { x: 430, y: 0, w: 100, h: 300 }).map((b) => b.index)).toEqual([0]);
    expect(pickVisible(boxes, { x: 600, y: 0, w: 100, h: 300 })).toEqual([]);
  });
});

describe('TD-88-11 drawingExtent：線の外接矩形を含み、シートの範囲を超えない', () => {
  it('U 字線の張り出しの分だけアンカーの矩形より広い列まで表を広げる', () => {
    const big = createGeometry(layoutOf(), 1048576, 16384);
    const mk = (adj1: number): SheetDrawing =>
      draw(
        { type: 'absolute', x: px(400), y: px(100), cx: px(10), cy: px(60) },
        {
          kind: 'line',
          mediaPath: null,
          line: plainLine({ geom: 'bentConnector3', adj: [adj1, 50000, 50000], widthEmu: 12700 }),
        },
      );
    // アンカーの右端 x=410 → 列 7（384〜448）。張り出し（x=450 付近）は列 8（448〜512）
    expect(drawingExtent(big, layoutDrawings(big, [mk(50000)])).cols).toBe(7);
    expect(drawingExtent(big, layoutDrawings(big, [mk(500000)])).cols).toBe(8);
  });

  it('adj=±2147483647 の線でも、外接矩形・表の広がりが totalWidth・totalHeight を超えない', () => {
    // 小さなシート（20 列×50 行＝1280×1000px）
    const sheet = createGeometry(layoutOf(), 50, 20);
    for (const adj of [2147483647, -2147483647]) {
      const d = draw(
        { type: 'absolute', x: px(400), y: px(100), cx: px(100), cy: px(60) },
        {
          kind: 'line',
          mediaPath: null,
          line: plainLine({ geom: 'bentConnector5', adj: [adj, adj, adj], widthEmu: 12700 }),
        },
      );
      const boxes = layoutDrawings(sheet, [d]);
      for (const b of boxes) {
        expect(b.rect.x).toBeGreaterThanOrEqual(0);
        expect(b.rect.y).toBeGreaterThanOrEqual(0);
        expect(b.rect.x + b.rect.w).toBeLessThanOrEqual(sheet.totalWidth);
        expect(b.rect.y + b.rect.h).toBeLessThanOrEqual(sheet.totalHeight);
      }
      const ext = drawingExtent(sheet, boxes);
      expect(ext.cols).toBeLessThanOrEqual(20);
      expect(ext.rows).toBeLessThanOrEqual(50);
      // 範囲を使い切った図の端（シートの端の 1px 手前）を含む行・列
      for (const b of boxes) {
        expect(ext.cols).toBeGreaterThanOrEqual(sheet.colAt(b.rect.x + b.rect.w - 1));
        expect(ext.rows).toBeGreaterThanOrEqual(sheet.rowAt(b.rect.y + b.rect.h - 1));
      }
    }
  });
});
