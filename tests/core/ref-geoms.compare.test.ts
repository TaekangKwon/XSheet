/// <reference types="node" />
/**
 * Excel との図形の形の照合の結果を固定するテスト（ISSUE-93・ADR-0028 §決定 10 単位 d・U3・U4）。
 * 照合そのもの（Excel 16 の COM で画像を書き出し、InterHub の描画を Playwright で写して輪郭の距離を測る）は vitest の常時実行に入れない
 * （`fixtures/excel/make-ref-geom-pngs.ps1` → `compare-ref-geoms.mjs`。手順は各ファイルの冒頭）。ここでは、その要約の JSON
 * （`fixtures/excel/ref-geoms.compare.json`・`ref-flips.compare.json`。PNG の山はリポジトリに入れない）を読み、次を確かめる：
 * - U3：全 preset（線の 10 種を除く 182 個）・調整値違い・回転反転・自由図形・既定の図形の輪郭の距離（95 パーセンタイル・100% 換算）が 1.5px 以内。
 *   超える形は `KNOWN_OVER` に原因つきで載せたものだけ（載せていないものが超えたら失敗。載せたのに超えていなければ失敗＝古い記載を残さない）
 * - U4：左右・上下に反転した文字入りの図形で、文字の向き（180° 回っていない）と位置（枠線の外接矩形の左上からの重心の差 5px 以内）が Excel に合う。合わない組は `KNOWN_TEXT`
 * JSON が無い（Excel を実行していない）間は、結果のテストは skip。判定の関数（`checkCompare`・`checkFlips`）は合成の JSON で常に検査する。
 */
import { existsSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

/** ref-geoms.compare.json の図形 1 つ */
interface GeomEntry {
  group: string;
  prst: string;
  p95: number;
  ok: boolean;
  excelPxPerPt?: number;
}
interface GeomCompare {
  version: number;
  excelVersion: string | null;
  arcMode: string;
  thresholdPx: number;
  scale: number;
  summary: {
    total: number;
    ok: number;
    ng: number;
    byGroup: Record<string, { total: number; ok: number }>;
    ngNames: string[];
  };
  notes: string[];
  shapes: Record<string, GeomEntry>;
}
/** ref-flips.compare.json の図形 1 つ */
interface FlipEntry {
  kind: string;
  variant: string;
  outline: { p95: number; ok: boolean };
  text: { orientation: string; dCentroid: { x: number; y: number } | null };
  textOk: boolean;
}
interface FlipCompare {
  version: number;
  thresholdPx: number;
  textPosTolPx: number;
  notes: string[];
  shapes: Record<string, FlipEntry>;
}

/** 許容（輪郭の距離。100% 換算の px） */
const TOLERANCE_PX = 1.5;
/** 文字の位置の許容（100% 換算の px） */
const TEXT_POS_TOL_PX = 5;

/**
 * 許容を超えた形の既知の差（図形名 → 原因）。ISSUE-93 の実装メモ・TV-41 の既知の差異と同じ内容を載せる。
 * 実行して許容を超えた形があれば、原因（変換・評価・arcTo・Excel 側の描画の癖）を切り分けてここに書く（直せるものは直して載せない）。
 */
const KNOWN_OVER: Record<string, string> = {
  polyOpen:
    'Excel の画像側の限界（InterHub の描画は合っている）：開いた折れ線は Excel の画像が図形の矩形ちょうどで切られ、端の線が半分欠け、右端の列に背景（マゼンタ）が混ざって縦の偽の線ができる（輪郭の距離・幅の差 +40px はこの偽の線による）。さらに斜線が淡く、濃さの閾値で途切れる。頂点の位置は一致（谷 1/3・山 2/3・端＝幅の 1/3 刻み）。閉じた図形・曲線の自由図形は合う',
  polyArrow:
    '自由図形：InterHub は始点・終点の印（Excel は四角と矢じり）を描かず、中心線の長さが違う。自由図形の線端の印は未対応（ISSUE-88 の既知の差）',
  p6: '画素の量子化：平均距離 0.58px。斜辺の端 2.5% の画素が最大 10px 離れる（Excel の斜辺の端の画素）。形は一致',
  flipHVRot270:
    '回転図形は twoCellAnchor のまま比べる（行の高さの丸めで InterHub の図が約 4px 大きい）。形は一致（平均距離 0.44px）',
  rot90:
    '回転図形は twoCellAnchor のまま比べる（行の高さの丸めで InterHub の図が約 4px 大きい）。形は一致（平均距離 0.88px）',
  ffMixed: '画素の量子化：p95 が 2 画素（1.6〜1.8px）。平均距離は 0.3〜0.8px で、形は一致',
  p33: '画素の量子化：p95 が 2 画素（1.6〜1.8px）。平均距離は 0.3〜0.8px で、形は一致',
  p96: '画素の量子化：p95 が 2 画素（1.6〜1.8px）。平均距離は 0.3〜0.8px で、形は一致',
  polyClosed: '画素の量子化：p95 が 2 画素（1.6〜1.8px）。平均距離は 0.3〜0.8px で、形は一致',
  p49: '画素の量子化：p95 が 2 画素（1.6〜1.8px）。平均距離は 0.3〜0.8px で、形は一致',
  p111: '画素の量子化：p95 が 2 画素（1.6〜1.8px）。平均距離は 0.3〜0.8px で、形は一致',
  p124: '画素の量子化：p95 が 2 画素（1.6〜1.8px）。平均距離は 0.3〜0.8px で、形は一致',
  p140: '画素の量子化：p95 が 2 画素（1.6〜1.8px）。平均距離は 0.3〜0.8px で、形は一致',
  p169: '画素の量子化：p95 が 2 画素（1.6〜1.8px）。平均距離は 0.3〜0.8px で、形は一致',
  adjRightArrowNeg: '画素の量子化：p95 が 2 画素（1.6〜1.8px）。平均距離は 0.3〜0.8px で、形は一致',
  ffRot: '画素の量子化：p95 が 2 画素（1.6〜1.8px）。平均距離は 0.3〜0.8px で、形は一致',
};
/** 文字の向き・位置が合わない反転の組の既知の差（図形名 → 原因） */
const KNOWN_TEXT: Record<string, string> = {
  'u4_ellipse_flipVRot30:outline':
    '回転図形は twoCellAnchor のまま比べる（行の高さの丸めで InterHub の図が数 px 大きい）。形は一致',
  'u4_diamond_flipVRot30:outline':
    '回転図形は twoCellAnchor のまま比べる（行の高さの丸めで InterHub の図が数 px 大きい）。形は一致',
  'u4_triangle_flipVRot30:outline':
    '回転図形は twoCellAnchor のまま比べる（行の高さの丸めで InterHub の図が数 px 大きい）。形は一致',
  'u4_rightArrow_flipVRot30:outline':
    '回転図形は twoCellAnchor のまま比べる（行の高さの丸めで InterHub の図が数 px 大きい）。形は一致',
  u4_homePlate_none:
    '測定が不安定：homePlate は文字が枠線に接するため、枠線を除くと文字が欠け、反転なしでも rot180 と判定される。向き・位置の結論は他の 4 種から出す',
  u4_homePlate_flipH:
    '測定が不安定：homePlate は文字が枠線に接するため、枠線を除くと文字が欠け、反転なしでも rot180 と判定される。向き・位置の結論は他の 4 種から出す',
  'u4_homePlate_flipVRot30:outline':
    '回転図形は twoCellAnchor のまま比べる（行の高さの丸めで InterHub の図が数 px 大きい）。形は一致',
};

/** 輪郭の照合の結果の検査。問題の一覧（空なら通る） */
function checkCompare(c: GeomCompare, known: Record<string, string>): string[] {
  const bad: string[] = [];
  const names = Object.keys(c.shapes);
  if (c.version !== 1) bad.push(`version ${c.version}`);
  if (c.thresholdPx !== TOLERANCE_PX) bad.push(`許容 ${c.thresholdPx}`);
  if (c.summary.total !== names.length) bad.push(`total ${c.summary.total} ≠ 図形 ${names.length}`);
  const over = names.filter((n) => !(c.shapes[n]!.p95 <= TOLERANCE_PX));
  if (c.summary.ng !== over.length) bad.push(`ng ${c.summary.ng} ≠ 超えた図形 ${over.length}`);
  for (const n of names) {
    const e = c.shapes[n]!;
    if (e.ok !== e.p95 <= TOLERANCE_PX) bad.push(`${n}: ok と p95 が食い違う`);
    if (e.p95 > TOLERANCE_PX && known[n] === undefined)
      bad.push(`${n}（${e.prst}）が許容超え p95=${e.p95}`);
  }
  for (const n of Object.keys(known)) {
    const e = c.shapes[n];
    if (e === undefined) bad.push(`既知の差 ${n} が結果に無い`);
    else if (e.p95 <= TOLERANCE_PX) bad.push(`既知の差 ${n} は許容内（記載を消す）`);
  }
  return bad;
}

/** 反転した文字の照合の結果の検査 */
function checkFlips(c: FlipCompare, known: Record<string, string>): string[] {
  const bad: string[] = [];
  if (c.version !== 1) bad.push(`version ${c.version}`);
  for (const [n, e] of Object.entries(c.shapes)) {
    const pos =
      e.text.dCentroid === null ? Infinity : Math.hypot(e.text.dCentroid.x, e.text.dCentroid.y);
    const good = e.text.orientation === 'upright' && pos <= TEXT_POS_TOL_PX;
    if (e.textOk !== good) bad.push(`${n}: textOk と向き・位置が食い違う`);
    if (!e.outline.ok && known[`${n}:outline`] === undefined)
      bad.push(`${n}: 枠線の輪郭が許容超え p95=${e.outline.p95}`);
    if (!good && known[n] === undefined)
      bad.push(`${n}: 文字の向き ${e.text.orientation}・位置の差 ${pos}`);
  }
  for (const n of Object.keys(known)) {
    const base = n.replace(/:outline$/, '');
    const e = c.shapes[base];
    if (e === undefined) bad.push(`既知の差 ${n} が結果に無い`);
    else if (n.endsWith(':outline') ? e.outline.ok : e.textOk)
      bad.push(`既知の差 ${n} は合っている（記載を消す）`);
  }
  return bad;
}

const dir = new URL('../../fixtures/excel/', import.meta.url);
const geomsUrl = new URL('ref-geoms.compare.json', dir);
const flipsUrl = new URL('ref-flips.compare.json', dir);

describe('TD-93-8 検査の関数（合成の結果で）', () => {
  const entry = (p95: number, prst = 'rect'): GeomEntry => ({
    group: 'preset',
    prst,
    p95,
    ok: p95 <= TOLERANCE_PX,
  });
  const make = (shapes: Record<string, GeomEntry>): GeomCompare => ({
    version: 1,
    excelVersion: '16.0',
    arcMode: 'visual',
    thresholdPx: TOLERANCE_PX,
    scale: 1.25,
    summary: {
      total: Object.keys(shapes).length,
      ok: Object.values(shapes).filter((e) => e.ok).length,
      ng: Object.values(shapes).filter((e) => !e.ok).length,
      byGroup: {},
      ngNames: [],
    },
    notes: [],
    shapes,
  });

  it('全部が許容内なら問題なし。超えた形は、既知に載せていなければ失敗', () => {
    expect(checkCompare(make({ a: entry(0.4), b: entry(1.5) }), {})).toEqual([]);
    expect(checkCompare(make({ a: entry(0.4), b: entry(2.2, 'cloud') }), {})).toEqual([
      'b（cloud）が許容超え p95=2.2',
    ]);
  });

  it('既知の差に載せた形は超えていてよい。載せたのに許容内・結果に無いのは失敗（古い記載を残さない）', () => {
    expect(checkCompare(make({ a: entry(3) }), { a: '原因' })).toEqual([]);
    expect(checkCompare(make({ a: entry(1) }), { a: '原因' })).toEqual([
      '既知の差 a は許容内（記載を消す）',
    ]);
    expect(checkCompare(make({ a: entry(1) }), { z: '原因' })).toEqual(['既知の差 z が結果に無い']);
  });

  it('集計が図形と食い違うものを見つける', () => {
    const c = make({ a: entry(0.4), b: entry(2) });
    c.summary.ng = 0;
    expect(checkCompare(c, { b: 'x' })).toEqual(['ng 0 ≠ 超えた図形 1']);
  });

  it('反転した文字：180° 回っている・位置が 5px を超えて違う組は、既知に載せていなければ失敗', () => {
    const flip = (
      orientation: string,
      d: { x: number; y: number } | null,
      outlineOk = true,
    ): FlipEntry => ({
      kind: 'ellipse',
      variant: 'flipV',
      outline: { p95: outlineOk ? 0.3 : 4, ok: outlineOk },
      text: { orientation, dCentroid: d },
      textOk: orientation === 'upright' && d !== null && Math.hypot(d.x, d.y) <= TEXT_POS_TOL_PX,
    });
    const make2 = (shapes: Record<string, FlipEntry>): FlipCompare => ({
      version: 1,
      thresholdPx: TOLERANCE_PX,
      textPosTolPx: TEXT_POS_TOL_PX,
      notes: [],
      shapes,
    });
    expect(checkFlips(make2({ a: flip('upright', { x: 1, y: 1 }) }), {})).toEqual([]);
    expect(checkFlips(make2({ a: flip('rot180', { x: 1, y: 1 }) }), {})).toEqual([
      'a: 文字の向き rot180・位置の差 1.4142135623730951',
    ]);
    expect(checkFlips(make2({ a: flip('upright', { x: 4, y: 4 }) }), {})[0]).toContain('位置の差');
    expect(checkFlips(make2({ a: flip('rot180', { x: 1, y: 1 }) }), { a: '原因' })).toEqual([]);
    expect(checkFlips(make2({ a: flip('upright', { x: 0, y: 0 }) }), { a: '原因' })).toEqual([
      '既知の差 a は合っている（記載を消す）',
    ]);
    expect(checkFlips(make2({ a: flip('unknown', null, false) }), {}).length).toBe(2);
    expect(
      checkFlips(make2({ a: flip('upright', { x: 0, y: 0 }, false) }), { 'a:outline': '原因' }),
    ).toEqual([]);
  });
});

describe.skipIf(!existsSync(geomsUrl))(
  'TD-93-9 ref-geoms.compare.json：全図形の輪郭が Excel と 1.5px 以内（U3）',
  () => {
    const c = (
      existsSync(geomsUrl) ? JSON.parse(readFileSync(geomsUrl, 'utf8')) : {}
    ) as GeomCompare;

    it('Excel 16・既定の arcTo の換算（視角）・画像の欠けなし。比べた図形は 207 個（preset 182・調整値違い 6・回転反転 6・自由図形 11・既定 2）', () => {
      expect(c.excelVersion).toBe('16.0');
      expect(c.arcMode).toBe('visual');
      expect(c.notes.filter((n) => n.startsWith('Excel の画像が無い'))).toEqual([]);
      expect(c.summary.total).toBe(207);
      expect(c.summary.byGroup.preset?.total).toBe(182);
      expect(c.summary.byGroup.adjusted?.total).toBe(6);
      expect(c.summary.byGroup.xform?.total).toBe(6);
      expect(c.summary.byGroup.custom?.total).toBe(11);
      expect(c.summary.byGroup.default?.total).toBe(2);
    });

    it('Excel の画像の倍率が 1pt = 1.667 画素（画面 125%）に近い（preset の画像の幅 ÷ COM の幅の中央値。吹き出し・弧は枠の外に出るため外れ値になる）', () => {
      const v = Object.values(c.shapes)
        .map((e) => e.excelPxPerPt)
        .filter((x): x is number => x !== undefined)
        .sort((a, b) => a - b);
      expect(v.length).toBeGreaterThan(100);
      const median = v[Math.floor(v.length / 2)]!;
      expect(median).toBeGreaterThan(1.6);
      expect(median).toBeLessThan(1.8);
    });

    it('輪郭の距離（95 パーセンタイル）が 1.5px 以内。超える形は KNOWN_OVER に原因つきで載せたものだけ', () => {
      expect(checkCompare(c, KNOWN_OVER)).toEqual([]);
    });

    it('U3：楕円の弧の形（cloud・pie・chord・arc・blockArc・moon・smileyFace）が一致する', () => {
      const arcs = new Set(['cloud', 'pie', 'chord', 'arc', 'blockArc', 'moon', 'smileyFace']);
      const found = new Set<string>();
      for (const [n, e] of Object.entries(c.shapes)) {
        if (e.group !== 'preset' || !arcs.has(e.prst)) continue;
        found.add(e.prst);
        if (KNOWN_OVER[n] === undefined)
          expect(e.p95, `${n}:${e.prst}`).toBeLessThanOrEqual(TOLERANCE_PX);
      }
      expect([...found].sort()).toEqual([...arcs].sort());
    });
  },
);

describe.skipIf(!existsSync(flipsUrl))(
  'TD-93-10 ref-flips.compare.json：反転した文字入りの図形の枠線・文字が Excel に合う（U4）',
  () => {
    const c = (
      existsSync(flipsUrl) ? JSON.parse(readFileSync(flipsUrl, 'utf8')) : {}
    ) as FlipCompare;

    it('文字入りの 25 図形（5 種 × 反転なし・左右・上下・左右上下・上下＋30° 回転）が全部ある。画像の欠けなし', () => {
      expect(Object.keys(c.shapes)).toHaveLength(25);
      expect(c.notes.filter((n) => n.startsWith('画像が足りない'))).toEqual([]);
    });

    it('枠線の輪郭は 1.5px 以内、文字は Excel と同じ向きで、位置が 5px 以内。合わない組は KNOWN_TEXT に載せたものだけ', () => {
      expect(checkFlips(c, KNOWN_TEXT)).toEqual([]);
    });
  },
);
