/// <reference types="node" />
/**
 * 実際の Excel で作った基準ファイルによる、図形の形・回転と反転・枠線の読み取りの照合（ISSUE-91・ADR-0028 §決定 3・10・U6）。
 * `fixtures/excel/ref-geoms.xlsx` は Excel 16 の COM（`make-ref-geoms.ps1`・PowerShell 5.1・画面 125%）で作成・保存した実ファイル。
 * 1 シートに図形 225 個（AutoShapeType 1〜183 の preset 182 個・調整値違い 6・回転反転 6・文字入り 2・自由図形 11・枠線の太さ 5・線種 8・そのほか）が入る。
 * 期待値は 2 系統で、同じフォルダの `ref-geoms.expected.json`（`make-ref-geoms.extract.mjs` が作る）に図形名ごとに入っている：
 * - COM の値（`Rotation`・`HorizontalFlip`＝-1 が反転・`Adjustments`・`Line.Weight`＝pt・`Line.DashStyle`・`Line.Visible`・`Fill.Visible`・`NodeCount`・`AutoShapeType`）
 * - Excel が書いた drawing XML の値（`a:prstGeom`・`a:avLst`・`a:custGeom`・`a:ln`・`a:xfrm`）
 * AutoShapeType 138（`msoShapeNotPrimitive`）は `AddShape` で作れず（COM が「指定された値は境界を超えています」を返す）、一覧に無い。
 * 自由図形（`BuildFreeform`・`AddCurve`・`AddPolyline`）は COM 上では `Type` 5・`AutoShapeType` 138 を返す。
 */
import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { evaluateGeom, getPresetGeom } from '../../src/geom/evaluate';
import { openXlsx, type SheetDrawing, type SheetDrawings } from '../../src/core/index';

/** expected.json の図形 1 つ */
interface RefShape {
  com: {
    Type: number;
    Rotation: number;
    HorizontalFlip: number;
    VerticalFlip: number;
    AutoShapeType: number;
    AdjCount: number;
    Adjustments?: number[];
    NodeCount: number;
    Width: number;
    Height: number;
    Line: { Visible: number; Weight: number; DashStyle: number };
    FillVisible: number;
  };
  xml: {
    txBox?: string;
    xfrm: { ext: { cx: number; cy: number }; rot?: number; flipH?: string; flipV?: string };
    geom:
      | { type: 'preset'; prst: string; avLst: { name: string; fmla: string }[] }
      | {
          type: 'custom';
          rect?: Record<string, string>;
          paths: { attrs: Record<string, string> }[];
        };
    ln?: {
      w?: number;
      noFill: boolean;
      prstDash?: string;
      headEnd?: { type: string };
      tailEnd?: { type: string };
    };
    fill?: string;
    style?: { lnRef?: number; fillRef?: number };
  };
}

const dir = new URL('../../fixtures/excel/', import.meta.url);
const expected = JSON.parse(readFileSync(new URL('ref-geoms.expected.json', dir), 'utf8')) as {
  excelVersion: string;
  notes: string[];
  shapes: Record<string, RefShape>;
};
const refNames = Object.keys(expected.shapes);

let result: SheetDrawings;
let byName: Map<string, SheetDrawing>;

beforeAll(async () => {
  const book = await openXlsx(readFileSync(new URL('ref-geoms.xlsx', dir)));
  result = await book.openDrawings(1);
  byName = new Map(result.drawings.map((d) => [d.name, d]));
});

/** 名前で図を 1 つ取る（無ければテストを失敗させる） */
function get(name: string): SheetDrawing {
  const d = byName.get(name);
  if (!d) throw new Error(`図が無い: ${name}`);
  return d;
}

/** 調整値（`adjN` の `val N`）を `[名前, 値]` にしたもの */
const avPairs = (av: { name: string; fmla: string }[]): [string, number][] =>
  av.map((a) => [a.name, Number(a.fmla.replace(/^val\s+/, ''))]);

/** 自由図形の名前（COM の `Type` が 5） */
const CUSTOM_NAMES = refNames.filter((n) => expected.shapes[n]!.xml.geom.type === 'custom');

describe('TD-91-9 ref-geoms.xlsx：図形の一覧と種類', () => {
  it('図形 225 個がすべて出る（通知なし）。AutoShapeType 138 は作れず、preset は 182 個', () => {
    expect(expected.excelVersion).toBe('16.0');
    // preset138 は Excel 16 の AddShape が拒否する（msoShapeNotPrimitive。自由図形がこの値を返す）
    expect(expected.notes).toEqual(['preset138 failed: 指定された値は境界を超えています。']);
    expect(refNames).toHaveLength(225);
    expect(result.drawings).toHaveLength(225);
    expect(result.notices).toEqual([]);
    const presets = refNames.filter((n) => /^p\d+$/.test(n));
    expect(presets).toHaveLength(182);
    expect(presets).not.toContain('p138');
  });

  it('種類：COM の Type 17（テキストボックス）だけ textbox、ほかは shape。線として出る図は無い', () => {
    for (const n of refNames) {
      const ref = expected.shapes[n]!;
      expect(get(n).kind, n).toBe(ref.com.Type === 17 ? 'textbox' : 'shape');
      expect(ref.xml.txBox === '1', n).toBe(ref.com.Type === 17);
    }
    expect(CUSTOM_NAMES).toHaveLength(11);
    for (const n of CUSTOM_NAMES) expect(expected.shapes[n]!.com.Type, n).toBe(5);
  });

  it('文字のある図形は textOval・textDiamond だけが text を持つ（形・枠線の項目も持つ）', () => {
    const withText = result.drawings.filter((d) => d.text !== undefined).map((d) => d.name);
    expect(withText.sort()).toEqual(['textDiamond', 'textOval']);
    expect(get('textOval').geom).toMatchObject({ type: 'preset', prst: 'ellipse' });
    expect(get('textDiamond').geom).toMatchObject({ type: 'preset', prst: 'diamond' });
  });
});

describe('TD-91-10 ref-geoms.xlsx：回転・反転・枠線と塗りの有無（全 225 図形）', () => {
  it('rot・flipH・flipV が Excel が書いた xfrm と COM の値に一致する', () => {
    for (const n of refNames) {
      const ref = expected.shapes[n]!;
      const d = get(n);
      const xrot = (((ref.xml.xfrm.rot ?? 0) % 21_600_000) + 21_600_000) % 21_600_000;
      expect(d.rot, n).toBe(xrot);
      // COM の Rotation（度）と 1/60000 度の差は 0.001° 未満
      expect(Math.abs(d.rot! - ref.com.Rotation * 60000), n).toBeLessThan(60);
      expect(d.flipH, n).toBe(ref.xml.xfrm.flipH === '1');
      expect(d.flipV, n).toBe(ref.xml.xfrm.flipV === '1');
      expect(d.flipH, n).toBe(ref.com.HorizontalFlip !== 0);
      expect(d.flipV, n).toBe(ref.com.VerticalFlip !== 0);
    }
  });

  it('回転・反転を付けた 7 図形（rot30・flipH・flipV・rot90・flipHRot45・flipHVRot270・ffRot）の値', () => {
    expect(get('rot30')).toMatchObject({ rot: 1_800_000, flipH: false, flipV: false });
    expect(get('flipH')).toMatchObject({ rot: 0, flipH: true, flipV: false });
    expect(get('flipV')).toMatchObject({ rot: 0, flipH: false, flipV: true });
    expect(get('rot90')).toMatchObject({ rot: 5_400_000 });
    expect(get('flipHRot45')).toMatchObject({ rot: 2_700_000, flipH: true });
    expect(get('flipHVRot270')).toMatchObject({ rot: 16_200_000, flipH: true, flipV: true });
    expect(get('ffRot')).toMatchObject({ rot: 19_200_000, flipH: true, geom: { type: 'custom' } });
  });

  it('枠線あり・塗りありが COM の Line.Visible・Fill.Visible に一致する', () => {
    for (const n of refNames) {
      const ref = expected.shapes[n]!;
      const d = get(n);
      expect(d.hasOutline, n).toBe(ref.com.Line.Visible !== 0);
      expect(d.hasFill, n).toBe(ref.com.FillVisible !== 0);
    }
    expect(get('lineNone')).toMatchObject({ hasOutline: false, hasFill: false });
    expect(get('emptyDiamond')).toMatchObject({ hasOutline: false, hasFill: false });
    expect(get('defaultRect')).toMatchObject({ hasOutline: true, hasFill: true });
    expect(get('p1')).toMatchObject({ hasOutline: true, hasFill: false });
  });
});

describe('TD-91-11 ref-geoms.xlsx：枠線の太さと線種', () => {
  it('太さが a:ln w（無ければ lnRef idx）に一致し、a:ln w は COM の Line.Weight（pt）と 1 EMU 以内で一致する', () => {
    const LN_REF: Record<number, number> = { 1: 6350, 2: 12700, 3: 19050 };
    for (const n of refNames) {
      const ref = expected.shapes[n]!;
      const w = ref.xml.ln?.w ?? LN_REF[ref.xml.style?.lnRef ?? -1] ?? 0;
      expect(get(n).outline!.widthEmu, n).toBe(w);
      if (ref.xml.ln?.w !== undefined) {
        expect(Math.abs(ref.xml.ln.w - ref.com.Line.Weight * 12700), n).toBeLessThanOrEqual(1);
      }
    }
  });

  it('太さ違い（0.25・1・3・6・12pt）：3175・12700・38100・76200・152400 EMU', () => {
    expect([0, 1, 2, 3, 4].map((i) => get(`w_${i}`).outline!.widthEmu)).toEqual([
      3175, 12700, 38100, 76200, 152400,
    ]);
  });

  it('線種が a:prstDash に一致する（Excel の DashStyle 1〜8 → solid・dot・dot・dash・dashDot・sysDashDotDot・lgDash・lgDashDot）', () => {
    const dashes = [
      'solid',
      'squareDot',
      'roundDot',
      'dash',
      'dashDot',
      'dashDotDot',
      'longDash',
      'longDashDot',
    ];
    expect(dashes.map((s) => get(`d_${s}`).outline!.dash)).toEqual([
      'solid',
      'dot',
      'dot',
      'dash',
      'dashDot',
      'sysDashDotDot',
      'lgDash',
      'lgDashDot',
    ]);
    for (const n of refNames) {
      expect(get(n).outline!.dash, n).toBe(expected.shapes[n]!.xml.ln?.prstDash ?? 'solid');
    }
  });

  it('テキストボックスの枠線 tbBordered は a:ln w=38100・線種の指定なし（solid）。style の lnRef idx は 0', () => {
    expect(get('tbBordered')).toMatchObject({
      kind: 'textbox',
      hasOutline: true,
      outline: { widthEmu: 38100, dash: 'solid' },
    });
  });
});

describe('TD-91-12 ref-geoms.xlsx：preset の形と調整値（225 図形のうち自由図形を除く 214 個）', () => {
  const presetNames = refNames.filter((n) => expected.shapes[n]!.xml.geom.type === 'preset');

  it('prst と avLst の調整値が Excel が書いた XML に一致する', () => {
    expect(presetNames).toHaveLength(214);
    for (const n of presetNames) {
      const g = expected.shapes[n]!.xml.geom;
      if (g.type !== 'preset') throw new Error(n);
      expect(get(n).geom, n).toEqual({ type: 'preset', prst: g.prst, adj: avPairs(g.avLst) });
    }
  });

  it('AutoShapeType が同じなら prst も同じ。COM の調整値の数は avLst の数以上', () => {
    const seen = new Map<number, string>();
    for (const n of presetNames) {
      const ref = expected.shapes[n]!;
      const g = get(n).geom;
      if (g?.type !== 'preset') throw new Error(n);
      const prev = seen.get(ref.com.AutoShapeType);
      if (prev !== undefined) expect(g.prst, n).toBe(prev);
      else seen.set(ref.com.AutoShapeType, g.prst);
      expect(ref.com.AdjCount, n).toBeGreaterThanOrEqual(g.adj.length);
    }
  });

  it('調整値違いの 6 図形：avLst の val が COM の Adjustments（割合）の 100000 倍に一致する', () => {
    const cases: [string, string, [string, number][]][] = [
      [
        'adjRightArrow',
        'rightArrow',
        [
          ['adj1', 30000],
          ['adj2', 80000],
        ],
      ],
      ['adjRoundRectSmall', 'roundRect', [['adj', 10000]]],
      ['adjRoundRectMax', 'roundRect', [['adj', 50000]]],
      [
        'adjCallout',
        'wedgeRectCallout',
        [
          ['adj1', -40000],
          ['adj2', 90000],
        ],
      ],
      ['adjCan', 'can', [['adj', 50000]]],
      [
        'adjRightArrowNeg',
        'rightArrow',
        [
          ['adj1', 90000],
          ['adj2', 10000],
        ],
      ],
    ];
    for (const [n, prst, adj] of cases) {
      const g = get(n).geom;
      if (g?.type !== 'preset') throw new Error(n);
      expect(g.prst, n).toBe(prst);
      // 既定と同じ値を設定した adjRoundRectMax と adjCan も avLst に書かれる（既定は 16667・25000 でなく、50000 は既定でない値）
      expect(g.adj, n).toEqual(adj);
      const com = expected.shapes[n]!.com.Adjustments!;
      adj.forEach(([, v], i) => expect(Math.abs(v / 100000 - com[i]!), n).toBeLessThan(1e-4));
    }
  });

  it('調整しない図形は avLst が空（既定は形の表）。rot30・flipH も同じ rightArrow', () => {
    expect(get('p33').geom).toEqual({ type: 'preset', prst: 'rightArrow', adj: [] });
    expect(get('rot30').geom).toEqual({ type: 'preset', prst: 'rightArrow', adj: [] });
    expect(get('p9').geom).toEqual({ type: 'preset', prst: 'ellipse', adj: [] });
  });

  it('全 preset（214 個。upArrow を含む）を評価器に通すと、すべて失敗せず有限の経路を返す', () => {
    for (const n of presetNames) {
      const g = get(n).geom;
      if (g?.type !== 'preset') throw new Error(n);
      // upArrow は ECMA-376 の原本に定義が無く、getPresetGeom が downArrow の上下反転として返す
      const def = getPresetGeom(g.prst);
      expect(def, `${n}:${g.prst}`).toBeDefined();
      if (def === undefined) continue;
      const ext = expected.shapes[n]!.xml.xfrm.ext;
      const w = ext.cx / 9525;
      const h = ext.cy / 9525;
      const r = evaluateGeom(def, new Map(g.adj), w, h);
      expect(r.ok, n).toBe(true);
      if (!r.ok) continue;
      const big = 5 * Math.max(w, h);
      for (const p of r.paths) {
        for (const sp of p.subpaths) {
          expect(Number.isFinite(sp.start.x) && Number.isFinite(sp.start.y), n).toBe(true);
          for (const s of sp.segs) {
            const v = s.t === 'L' ? [s.x, s.y] : [s.x1, s.y1, s.x2, s.y2, s.x, s.y];
            for (const c of v) expect(Math.abs(c), n).toBeLessThan(big);
          }
        }
      }
    }
  });

  it('Excel が書く upArrow（p35）：矢先が上（頂点が y=0・x が中央）', () => {
    const g = get('p35').geom;
    if (g?.type !== 'preset') throw new Error('p35');
    expect(g.prst).toBe('upArrow');
    const ext = expected.shapes.p35!.xml.xfrm.ext;
    const w = ext.cx / 9525;
    const h = ext.cy / 9525;
    const r = evaluateGeom(getPresetGeom('upArrow')!, new Map(g.adj), w, h);
    if (!r.ok) throw new Error(r.reason);
    const sp = r.paths[0]!.subpaths[0]!;
    const all = [sp.start, ...sp.segs.map((s) => ({ x: s.x, y: s.y }))];
    expect(Math.min(...all.map((p) => p.y))).toBeCloseTo(0, 9);
    const tip = all.find((p) => Math.abs(p.y) < 1e-9)!;
    expect(tip.x).toBeCloseTo(w / 2, 9);
    expect(Math.max(...all.map((p) => p.y))).toBeCloseTo(h, 9);
  });
});

describe('TD-91-13 ref-geoms.xlsx：自由図形（U6）', () => {
  /** 自由図形の定義（custom でなければテストを失敗させる） */
  const custom = (n: string): Extract<NonNullable<SheetDrawing['geom']>, { type: 'custom' }> => {
    const g = get(n).geom;
    if (g?.type !== 'custom') throw new Error(`自由図形でない: ${n}`);
    return g;
  };

  it('11 図形すべてが custom で、path は 1 つ。w・h は xfrm の ext に一致し、点はすべて整数（U6）', () => {
    expect(CUSTOM_NAMES.sort()).toEqual(
      [
        'ffTriangle',
        'ffCurve',
        'ffMixed',
        'curve4',
        'curve7',
        'polyOpen',
        'polyClosed',
        'polyArrow',
        'polyFilled',
        'curveArrow',
        'ffRot',
      ].sort(),
    );
    for (const n of CUSTOM_NAMES) {
      const def = custom(n).def;
      const ext = expected.shapes[n]!.xml.xfrm.ext;
      expect(def.paths, n).toHaveLength(1);
      const p = def.paths[0]!;
      expect(p.w, n).toBe(ext.cx);
      expect(p.h, n).toBe(ext.cy);
      // Excel は fill・stroke を書かない（既定＝塗る・線を引く）
      expect(p.fill, n).toBeUndefined();
      expect(p.stroke, n).toBeUndefined();
      for (const cmd of p.c) {
        for (const a of cmd.slice(1))
          expect(Number.isInteger(a) || typeof a === 'string', n).toBe(true);
      }
    }
  });

  it('命令は moveTo・lnTo・cubicBezTo・close だけ。節点の数（M＋L＋3×C）が COM の NodeCount に一致する', () => {
    for (const n of CUSTOM_NAMES) {
      const p = custom(n).def.paths[0]!;
      let nodes = 0;
      for (const cmd of p.c) {
        expect(['M', 'L', 'C', 'Z'], n).toContain(cmd[0]);
        nodes += cmd[0] === 'M' || cmd[0] === 'L' ? 1 : cmd[0] === 'C' ? 3 : 0;
      }
      expect(nodes, n).toBe(expected.shapes[n]!.com.NodeCount);
    }
  });

  it('評価器が失敗せず、経路は 1 つの部分経路で、終点が矩形の中に入る（閉じるかは close の有無）', () => {
    for (const n of CUSTOM_NAMES) {
      const def = custom(n).def;
      const ext = expected.shapes[n]!.xml.xfrm.ext;
      const w = ext.cx / 9525;
      const h = ext.cy / 9525;
      const r = evaluateGeom(def, undefined, w, h);
      expect(r.ok, n).toBe(true);
      if (!r.ok) continue;
      expect(r.paths, n).toHaveLength(1);
      const sub = r.paths[0]!.subpaths;
      expect(sub, n).toHaveLength(1);
      expect(sub[0]!.closed, n).toBe(def.paths[0]!.c.some((c) => c[0] === 'Z'));
      const ends = [sub[0]!.start, ...sub[0]!.segs.map((s) => ({ x: s.x, y: s.y }))];
      for (const e of ends) {
        expect(e.x, n).toBeGreaterThanOrEqual(-0.01);
        expect(e.x, n).toBeLessThanOrEqual(w + 0.01);
        expect(e.y, n).toBeGreaterThanOrEqual(-0.01);
        expect(e.y, n).toBeLessThanOrEqual(h + 0.01);
      }
      // 制御点は矩形の外に出てよいが、極端には出ない（ffMixed の曲線は矩形の上へ 1.7px はみ出す）
      for (const s of sub[0]!.segs) {
        if (s.t !== 'C') continue;
        for (const c of [s.x1, s.y1, s.x2, s.y2])
          expect(Math.abs(c), n).toBeLessThan(2 * Math.max(w, h));
      }
      // 文字領域：rect が l・t・r・b（0,0,0,0）の幅なしなら矩形全体に戻る
      expect(r.textRect.r - r.textRect.l, n).toBeGreaterThan(0);
    }
  });

  it('rect：自動のものは 0・0・0・0（幅なし）、ffMixed（曲線の節点が自動）は l・t・r・b と guide 10 個', () => {
    for (const n of CUSTOM_NAMES) {
      const def = custom(n).def;
      if (n === 'ffMixed') {
        expect(def.rect).toEqual(['l', 't', 'r', 'b']);
        expect(def.gd).toHaveLength(10);
        expect(def.gd[0]).toEqual(['csX0', '*/', 0, 'w', 857250]);
      } else {
        expect(def.rect, n).toEqual([0, 0, 0, 0]);
        expect(def.gd, n).toHaveLength(0);
      }
      expect(def.av, n).toEqual([]);
    }
  });

  it('枠線・回転は preset と同じに持つ。自由図形の矢じり（U6）は a:ln の headEnd・tailEnd に付くが、図形は読まない（描かない）', () => {
    expect(get('polyArrow').outline).toEqual({ widthEmu: 12700, dash: 'solid' });
    // Excel の xdr:sp の自由図形（開いた線）にも矢じりが付く：polyArrow は始点 diamond・終点 triangle、curveArrow は終点 triangle
    const e = (n: string) => expected.shapes[n]!.xml.ln;
    expect([e('polyArrow')?.headEnd?.type, e('polyArrow')?.tailEnd?.type]).toEqual([
      'diamond',
      'triangle',
    ]);
    expect([e('curveArrow')?.headEnd?.type, e('curveArrow')?.tailEnd?.type]).toEqual([
      undefined,
      'triangle',
    ]);
    // 矢じりの無い自由図形は headEnd・tailEnd を書かない
    for (const n of CUSTOM_NAMES.filter((x) => x !== 'polyArrow' && x !== 'curveArrow')) {
      expect(e(n)?.headEnd, n).toBeUndefined();
      expect(e(n)?.tailEnd, n).toBeUndefined();
    }
    for (const n of CUSTOM_NAMES) expect(get(n).line, n).toBeUndefined();
  });
});
