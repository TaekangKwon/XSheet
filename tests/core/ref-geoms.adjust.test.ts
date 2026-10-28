/// <reference types="node" />
/**
 * 実際の Excel（COM の `Shape.Adjustments`）と形の表（`xlsx-preset-geom.json`）の調整値の突き合わせ（ISSUE-93・ADR-0028 §決定 2・U3 の補助）。
 * 入力は ISSUE-91 の `fixtures/excel/ref-geoms.expected.json`（図形 `p1`〜`p183`＝AutoShapeType ごとの既定の図形。調整値は触っていない）。
 * COM の `Adjustments` は「長さ・比率の調整値は値 ÷ 100000」「角度の調整値は度（値 ÷ 60000）」で返る（`pie` の 0・270° は 0・-90）ため、
 * 表の `av`（既定の調整値）をその 2 通りに直して、COM の既定値と比べる。調整値の数（`AdjCount`）は `av` の数と一致するはず。
 */
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { getPresetGeom } from '../../src/geom/evaluate';

interface RefShape {
  com: { AutoShapeType: number; AdjCount: number; Adjustments?: number[] };
  xml: { geom: { type: string; prst?: string } };
}

const dir = new URL('../../fixtures/excel/', import.meta.url);
const expected = JSON.parse(readFileSync(new URL('ref-geoms.expected.json', dir), 'utf8')) as {
  shapes: Record<string, RefShape>;
};
/** AutoShapeType ごとの既定の図形（`p<番号>`）のうち、preset のもの */
const presetShapes = Object.entries(expected.shapes).filter(
  ([n, s]) => /^p\d+$/.test(n) && s.xml.geom.type === 'preset',
);

/** 角度（度）を (-180, 180] にそろえる */
function wrapDeg(a: number): number {
  const m = ((a % 360) + 360) % 360;
  return m > 180 ? m - 360 : m;
}

/** 調整値 1 つ（表の値）が COM の値と一致するか。長さ・比率なら ÷100000、角度なら ÷60000（度。360° の周期を同一視） */
function sameAdjust(av: number, com: number): boolean {
  const asRatio = Math.abs(av / 100000 - com) <= 2e-5 + 1e-9;
  const asAngle = Math.abs(wrapDeg(av / 60000 - com)) <= 2e-3;
  return asRatio || asAngle;
}

describe('TD-93-1 COM の調整値の数と既定値が、形の表の av と一致する（ref-geoms.xlsx の preset 182 図形）', () => {
  it('図形は 182 個（AutoShapeType 1〜183 のうち 138 は作れず、upArrow を含む）', () => {
    expect(presetShapes).toHaveLength(182);
  });

  it('調整値の数（AdjCount）が表の av の数に一致する（表に無い upArrow は downArrow の上下反転として引く）', () => {
    const bad: string[] = [];
    for (const [n, s] of presetShapes) {
      const def = getPresetGeom(s.xml.geom.prst!);
      if (def === undefined) {
        bad.push(`${n}:${s.xml.geom.prst}（表に無い）`);
        continue;
      }
      const count = def.av?.length ?? 0;
      if (count !== s.com.AdjCount)
        bad.push(`${n}:${s.xml.geom.prst} 表 ${count} / COM ${s.com.AdjCount}`);
    }
    expect(bad).toEqual([]);
  });

  it('既定値（COM の Adjustments）が av の値に一致する（比率＝値÷100000、角度＝値÷60000 度）', () => {
    const bad: string[] = [];
    let checked = 0;
    for (const [n, s] of presetShapes) {
      const av = getPresetGeom(s.xml.geom.prst!)?.av ?? [];
      const com = s.com.Adjustments ?? [];
      av.forEach(([name, v], i) => {
        const c = com[i];
        if (c === undefined) return;
        checked++;
        if (!sameAdjust(v, c)) bad.push(`${n}:${s.xml.geom.prst}.${name} 表 ${v} / COM ${c}`);
      });
    }
    expect(bad).toEqual([]);
    // 調整値のある形の調整値すべてを突き合わせた（空振りで通らない）
    expect(checked).toBeGreaterThan(150);
  });

  it('調整値のある形のうち、角度の調整値が COM では度で返る形（pie・arc・blockArc・chord・circularArrow など）を確かめた', () => {
    // arc の既定は adj1 = 270°（16200000）・adj2 = 0。COM は -90・0
    expect(expected.shapes.p25!.com.Adjustments).toEqual([-90, 0]);
    expect(sameAdjust(16200000, -90)).toBe(true);
    // blockArc の既定は 180°・0・0.25
    expect(expected.shapes.p20!.com.Adjustments).toEqual([180, 0, 0.25]);
    // 比率と角度を取り違えない：16200000 は比率として 162 で、COM の -90 とは合わない
    expect(Math.abs(16200000 / 100000 - -90)).toBeGreaterThan(1);
    // 値の違いは検出する
    expect(sameAdjust(16200000, 0)).toBe(false);
    expect(sameAdjust(25000, 0.3)).toBe(false);
  });
});
