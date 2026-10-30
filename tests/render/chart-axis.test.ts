import { describe, expect, it } from 'vitest';
import type { ChartAxis } from '../../src/core/types';
import {
  AXIS_CONSTANTS,
  categoryLabelIndices,
  categoryLabelSkip,
  categoryPositions,
  categorySpacing,
  computeValueScale,
  crossValue,
  dataExtent,
  niceUnit,
  tickCountLimit,
  tickFormatCode,
  tickLabel,
  valueFraction,
} from '../../src/render/chart-axis';

/** 数値軸の最小の `ChartAxis` */
function valAxis(over: Partial<ChartAxis> = {}): ChartAxis {
  return { kind: 'val', deleted: false, orientation: 'minMax', text: {}, ...over };
}

/** 大きいプロット（N=10）で範囲を決める */
function scaleOf(min: number, max: number, axis?: ChartAxis, length = 400, labelH = 12) {
  return computeValueScale({ axis, extent: { min, max }, lengthPx: length, labelHeightPx: labelH });
}

describe('dataExtent（データの範囲）', () => {
  it('TD-102-1: 集合は全点の最小・最大。null・非有限は数えない', () => {
    const e = dataExtent('clustered', [{ values: [3, null, 10] }, { values: [-2, Number.NaN, 5] }]);
    expect(e).toEqual({ min: -2, max: 10 });
  });

  it('TD-102-1: 全部 null なら null', () => {
    expect(dataExtent('clustered', [{ values: [null, null] }])).toBeNull();
    expect(dataExtent('standard', [])).toBeNull();
  });

  it('TD-102-1: 積み上げは項目ごとの正の和の最大と負の和の最小', () => {
    const e = dataExtent('stacked', [
      { values: [1, 5, -2] },
      { values: [4, 1, -3] },
      { values: [2, -1, 1] },
    ]);
    expect(e).toEqual({ min: -5, max: 7 });
  });

  it('TD-102-1: 正だけの積み上げの最小は 0', () => {
    expect(dataExtent('stacked', [{ values: [1, 2] }, { values: [3, 4] }])).toEqual({
      min: 0,
      max: 6,
    });
  });

  it('TD-102-1: percentStacked は 0〜1（点があれば）', () => {
    expect(dataExtent('percentStacked', [{ values: [1, 2] }])).toEqual({ min: 0, max: 1 });
    expect(dataExtent('percentStacked', [{ values: [null] }])).toBeNull();
  });
});

describe('niceUnit・tickCountLimit（刻みと目盛の数の上限）', () => {
  it('TD-102-2: {1,2,5}×10ⁿ の最小で、span ÷ u が上限以内', () => {
    expect(niceUnit(100, 10)).toBe(10);
    expect(niceUnit(91.35, 10)).toBe(10);
    expect(niceUnit(100, 5)).toBe(20);
    expect(niceUnit(100, 4)).toBe(50);
    expect(niceUnit(1, 10)).toBe(0.1);
    expect(niceUnit(0.35, 10)).toBe(0.05);
    expect(niceUnit(12_345_678, 10)).toBe(2_000_000);
  });

  it('TD-102-2: span が 0・負・非有限でも有限の正の値を返す', () => {
    for (const s of [0, -5, Number.NaN, Infinity]) {
      const u = niceUnit(s, 10);
      expect(Number.isFinite(u)).toBe(true);
      expect(u).toBeGreaterThan(0);
    }
  });

  it('TD-102-2: 目盛の数の上限は 長さ ÷ (文字の高さ × 1.0) を 2〜10 に丸める（ISSUE-104：Excel 16 の照合）', () => {
    expect(tickCountLimit(400, 12)).toBe(10);
    expect(tickCountLimit(60, 12)).toBe(5);
    expect(tickCountLimit(30, 12)).toBe(2);
    expect(tickCountLimit(Number.NaN, 12)).toBe(2);
  });
});

describe('computeValueScale（数値軸の範囲）', () => {
  it('TD-102-3: 0〜87 の正のデータは 0〜100・刻み 10', () => {
    const s = scaleOf(0, 87);
    expect([s.min, s.max, s.majorUnit]).toEqual([0, 100, 10]);
    expect(s.majorTicks).toEqual([0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]);
    expect(s.minorUnit).toBe(2);
    expect(s.minorTicks).toHaveLength(40); // 主と重ならない 2 おき
    expect(s.reversed).toBe(false);
  });

  it('TD-102-3: dmin が dmax×5/6 未満なら最小は 0（例 10〜87）', () => {
    const s = scaleOf(10, 87);
    expect(s.min).toBe(0);
  });

  it('TD-102-3: dmin が dmax×5/6 以上なら 0 にしない（例 80〜90 → 下端に余りを取って切り下げ）', () => {
    const s = scaleOf(80, 90);
    expect(s.min).toBeGreaterThan(0);
    expect(s.min).toBeLessThanOrEqual(80);
    expect(s.max).toBeGreaterThanOrEqual(90);
    // ちょうど境界（dmin = dmax×5/6）も 0 にしない
    expect(scaleOf(75, 90).min).toBeGreaterThan(0);
    expect(scaleOf(74, 90).min).toBe(0);
  });

  it('TD-102-3: 負を含むと両側に余りを取り、刻みの倍数に切り下げ・切り上げる', () => {
    const s = scaleOf(-30, 50);
    expect(s.min).toBeLessThanOrEqual(-30);
    expect(s.max).toBeGreaterThanOrEqual(50);
    expect(Math.abs(s.min % s.majorUnit)).toBe(0);
    expect(s.max % s.majorUnit).toBe(0);
  });

  it('TD-102-3: 全部が負なら最大を 0 にする（対称）', () => {
    const s = scaleOf(-90, -10);
    expect(s.max).toBe(0);
    expect(s.min).toBeLessThanOrEqual(-90);
    const t = scaleOf(-90, -80);
    expect(t.max).toBeLessThan(0);
  });

  it('TD-102-3: 小数のデータ（0〜0.8）', () => {
    const s = scaleOf(0, 0.8);
    expect(s.min).toBe(0);
    expect(s.max).toBeCloseTo(0.9, 10);
    expect(s.majorUnit).toBe(0.1);
    expect(s.majorTicks[1]).toBe(0.1); // 誤差を丸めている
    expect(s.majorTicks[3]).toBe(0.3);
  });

  it('TD-102-3: 大きな値（0〜12,345,678）', () => {
    const s = scaleOf(0, 12_345_678);
    expect(s.majorUnit).toBe(2_000_000);
    expect(s.max).toBe(14_000_000);
  });

  it('TD-102-3: 全部 null（extent が null）なら 0〜1', () => {
    const s = computeValueScale({ extent: null, lengthPx: 400, labelHeightPx: 12 });
    expect([s.min, s.max]).toEqual([0, 1]);
    expect(s.majorTicks[0]).toBe(0);
    expect(s.majorTicks.at(-1)).toBe(1);
  });

  it('TD-102-3: 全部同じ値でも範囲が作れる（5 → 0〜…・0 → 0〜1）', () => {
    const s = scaleOf(5, 5);
    expect(s.min).toBe(0);
    expect(s.max).toBeGreaterThanOrEqual(5);
    const z = scaleOf(0, 0);
    expect([z.min, z.max]).toEqual([0, 1]);
    const n = scaleOf(-5, -5);
    expect(n.max).toBe(0);
    expect(n.min).toBeLessThanOrEqual(-5);
  });

  it('TD-102-3: 積み上げの正負の和から範囲を決める', () => {
    const e = dataExtent('stacked', [{ values: [10, 20] }, { values: [-5, 30] }]);
    const s = computeValueScale({ extent: e, lengthPx: 400, labelHeightPx: 12 });
    expect(s.min).toBeLessThanOrEqual(-5);
    expect(s.max).toBeGreaterThanOrEqual(30);
  });

  it('TD-102-4: percentStacked は 0〜1・刻み 0.1・書式 0%', () => {
    const s = computeValueScale({
      extent: { min: 0, max: 1 },
      percentStacked: true,
      lengthPx: 400,
      labelHeightPx: 12,
    });
    expect([s.min, s.max, s.majorUnit, s.percent]).toEqual([0, 1, 0.1, true]);
    expect(s.majorTicks).toEqual([0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1]);
    expect(tickFormatCode(undefined, 'General', true)).toBe('0%');
    expect(tickLabel('0%', 0.3)).toBe('30%');
  });

  it('TD-102-5: 明示の min・max・majorUnit が優先される', () => {
    const s = scaleOf(0, 87, valAxis({ min: 10, max: 60, majorUnit: 25 }));
    expect([s.min, s.max, s.majorUnit]).toEqual([10, 60, 25]);
    expect(s.majorTicks).toEqual([10, 35, 60]);
  });

  it('TD-102-5: 一部だけ明示（min だけ）なら max は自動', () => {
    const s = scaleOf(20, 87, valAxis({ min: 20 }));
    expect(s.min).toBe(20);
    expect(s.max).toBeGreaterThanOrEqual(87);
    const t = scaleOf(20, 87, valAxis({ max: 200 }));
    expect(t.max).toBe(200);
    expect(t.min).toBe(0);
  });

  it('TD-102-5: 明示の max が min 以下なら解ける形に倒す（例外にしない）', () => {
    const s = scaleOf(0, 10, valAxis({ min: 50, max: 40 }));
    expect(s.max).toBeGreaterThan(s.min);
  });

  it('TD-102-6: プロットの長さで刻みが変わる（小さいグラフは粗い）', () => {
    const big = scaleOf(0, 87, undefined, 400, 12);
    const small = scaleOf(0, 87, undefined, 100, 12);
    expect(small.majorUnit).toBeGreaterThan(big.majorUnit);
    expect(small.majorTicks.length).toBeLessThan(big.majorTicks.length);
  });

  it('TD-102-7: maxMin の反転は valueFraction に効き、範囲は変わらない', () => {
    const s = scaleOf(0, 100, valAxis({ orientation: 'maxMin' }));
    expect(s.reversed).toBe(true);
    expect(valueFraction(s, 0)).toBe(1);
    expect(valueFraction(s, s.max)).toBe(0);
    const n = scaleOf(0, 100);
    expect(valueFraction(n, 0)).toBe(0);
    expect(valueFraction(n, n.max / 2)).toBeCloseTo(0.5, 10);
  });

  it('TD-102-7: valueFraction は範囲外を端に収め、非有限は 0 にする', () => {
    const s = scaleOf(0, 100);
    expect(valueFraction(s, -50)).toBe(0);
    expect(valueFraction(s, 1e9)).toBe(1);
    expect(valueFraction(s, Number.NaN)).toBe(0);
  });

  it('TD-102-8: 交点：autoZero は 0（範囲外は近い端）・max・crossesAt', () => {
    const s = scaleOf(-30, 50);
    expect(crossValue(s)).toBe(0);
    expect(crossValue(s, { ...valAxis(), kind: 'cat', crosses: 'autoZero' })).toBe(0);
    expect(crossValue(s, { ...valAxis(), kind: 'cat', crosses: 'max' })).toBe(s.max);
    expect(crossValue(s, { ...valAxis(), kind: 'cat', crosses: 'min' })).toBe(s.min);
    expect(crossValue(s, { ...valAxis(), kind: 'cat', crossesAt: 20 })).toBe(20);
    expect(crossValue(s, { ...valAxis(), kind: 'cat', crossesAt: 9999 })).toBe(s.max);
    const pos = scaleOf(80, 90);
    expect(crossValue(pos)).toBe(pos.min); // 0 が範囲外
    const neg = scaleOf(-90, -80);
    expect(crossValue(neg)).toBe(neg.max);
  });

  it('TD-102-9: 目盛の文字の書式（0%・#,##0・sourceLinked・無指定）', () => {
    expect(
      tickFormatCode(valAxis({ numFmt: { code: '#,##0', sourceLinked: false } }), '0.0', false),
    ).toBe('#,##0');
    expect(tickLabel('#,##0', 1234567)).toBe('1,234,567');
    expect(
      tickFormatCode(valAxis({ numFmt: { code: '0%', sourceLinked: false } }), undefined, false),
    ).toBe('0%');
    expect(
      tickFormatCode(valAxis({ numFmt: { code: '#,##0', sourceLinked: true } }), '0.00', false),
    ).toBe('0.00');
    expect(
      tickFormatCode(valAxis({ numFmt: { code: '#,##0', sourceLinked: true } }), undefined, false),
    ).toBe('General');
    expect(tickFormatCode(undefined, '0.0', false)).toBe('0.0');
    // percentStacked でも明示の書式が優先
    expect(
      tickFormatCode(valAxis({ numFmt: { code: '0.0%', sourceLinked: false } }), undefined, true),
    ).toBe('0.0%');
    expect(tickLabel('General', -0)).toBe('0');
  });
});

describe('目盛の数の上限（無限・巨大なループを防ぐ）', () => {
  it('TD-102-10: 明示の刻みが小さすぎても主の目盛は maxMajorTicks 以内', () => {
    const s = scaleOf(0, 1e6, valAxis({ majorUnit: 1e-9 }));
    expect(s.majorTicks.length).toBeLessThanOrEqual(AXIS_CONSTANTS.maxMajorTicks + 1);
    expect(s.majorUnit).toBeGreaterThan(1e-9);
  });

  it('TD-102-10: 補助の目盛が maxMinorTicks を超える明示の刻みは捨てる', () => {
    const s = scaleOf(0, 100, valAxis({ minorUnit: 1e-6 }));
    expect(s.minorUnit).toBeNull();
    expect(s.minorTicks).toEqual([]);
  });

  it('TD-102-10: 自動でも極端な値（1e300・1e-300・巨大な幅）で有限の目盛が出る', () => {
    for (const [a, b] of [
      [0, 1e300],
      [0, 1e-300],
      [-1e300, 1e300],
      [0, 1e15],
    ] as const) {
      const s = scaleOf(a, b);
      expect(s.majorTicks.length).toBeGreaterThan(1);
      expect(s.majorTicks.length).toBeLessThanOrEqual(AXIS_CONSTANTS.maxMajorTicks + 1);
      expect(s.minorTicks.length).toBeLessThanOrEqual(AXIS_CONSTANTS.maxMinorTicks + 1);
      expect(s.majorTicks.every((t) => Number.isFinite(t))).toBe(true);
    }
  });

  it('TD-102-10: 明示の max が巨大（1e300）でも目盛は上限以内', () => {
    const s = scaleOf(0, 1, valAxis({ max: 1e300, majorUnit: 1 }));
    expect(s.majorTicks.length).toBeLessThanOrEqual(AXIS_CONSTANTS.maxMajorTicks + 1);
  });
});

describe('項目軸（帯・midCat・間引き）', () => {
  it('TD-102-11: 帯の等分（中心は (i+0.5)×帯の幅）', () => {
    expect(categoryPositions(4, 400)).toEqual([50, 150, 250, 350]);
    expect(categorySpacing(4, 400)).toBe(100);
    expect(categoryPositions(0, 400)).toEqual([]);
  });

  it('TD-102-11: midCat は目盛の上（端から端まで (n−1) 等分・1 項目は中央）', () => {
    expect(categoryPositions(5, 400, { midCat: true })).toEqual([0, 100, 200, 300, 400]);
    expect(categorySpacing(5, 400, true)).toBe(100);
    expect(categoryPositions(1, 400, { midCat: true })).toEqual([200]);
  });

  it('TD-102-12: reversed（maxMin）は始点と終点を入れ替える（横棒は先頭が下＝縦軸を始点が下の向きにする）', () => {
    expect(categoryPositions(4, 400, { reversed: true })).toEqual([350, 250, 150, 50]);
    // 横棒の縦の項目軸：始点（下）からの距離。先頭の項目が小さい距離＝下、maxMin なら上
    const fromBottom = categoryPositions(3, 300);
    expect(fromBottom[0]).toBeLessThan(fromBottom[2]!);
  });

  it('TD-102-13: tickLblSkip の明示はそのまま従う（1 以上・項目数まで）', () => {
    expect(categoryLabelSkip(10, 100, 500, 3)).toBe(3);
    expect(categoryLabelSkip(10, 100, 20, 1)).toBe(1);
    expect(categoryLabelSkip(10, 100, 20, 99)).toBe(10);
    expect(categoryLabelSkip(10, 100, 500, 0)).toBeGreaterThan(1); // 0 は無効＝自動へ
  });

  it('TD-102-13: 推定の幅が帯に入らない時は 2・3…個おきに間引く', () => {
    expect(categoryLabelSkip(10, 100, 50)).toBe(1);
    expect(categoryLabelSkip(10, 50, 60)).toBe(2); // 60+4 > 50
    expect(categoryLabelSkip(10, 20, 60)).toBe(4); // ceil(64/20)
    expect(categoryLabelSkip(3, 1, 500)).toBe(3); // 項目数を超えない
    expect(categoryLabelSkip(10, 0, 5)).toBe(10);
  });

  it('TD-102-13: 文字を出す項目は 0, skip, 2·skip…', () => {
    expect(categoryLabelIndices(7, 3)).toEqual([0, 3, 6]);
    expect(categoryLabelIndices(3, 1)).toEqual([0, 1, 2]);
    expect(categoryLabelIndices(0, 2)).toEqual([]);
  });
});

describe('レビュー指摘の修正（負を含む 100% 積み上げ・横向きの軸の目盛の数）', () => {
  it('TD-102-14: percentStacked は最大 1 固定・最小は負の値が 1 つでもあれば -1（無ければ 0。ISSUE-104：Excel 16 の照合 C2）', () => {
    const neg = dataExtent('percentStacked', [{ values: [-30, 10] }, { values: [70, 10] }]);
    expect(neg).toEqual({ min: -1, max: 1 });
    expect(dataExtent('percentStacked', [{ values: [1, 2] }])).toEqual({ min: 0, max: 1 });
    // 全部が負なら -1〜0（ISSUE-105：Excel 16 の照合 C2）
    expect(dataExtent('percentStacked', [{ values: [-30, -10] }, { values: [-70, -10] }])).toEqual({
      min: -1,
      max: 0,
    });
  });

  it('TD-105-5: 全部が負の 100% 積み上げの軸は -1〜0（刻みは他の軸と同じく自動）', () => {
    const s = computeValueScale({
      extent: { min: -1, max: 0 },
      percentStacked: true,
      lengthPx: 400,
      labelHeightPx: 12,
    });
    expect([s.min, s.max]).toEqual([-1, 0]);
    expect(s.majorUnit).toBe(0.1);
    expect(s.majorTicks[0]).toBe(-1);
    expect(s.majorTicks.at(-1)).toBe(0);
  });

  it('TD-102-14: 負を含む 100% 積み上げの自動の最小は -1（刻みの倍数に切り下げ。範囲 2 を 10 本以内の刻み 0.2）', () => {
    const at = (min: number) =>
      computeValueScale({
        extent: { min, max: 1 },
        percentStacked: true,
        lengthPx: 400,
        labelHeightPx: 12,
      });
    expect(at(-1).min).toBe(-1);
    expect(at(-1).max).toBe(1);
    expect(at(-1).majorUnit).toBe(0.2);
    expect(at(-1).majorTicks.slice(0, 3)).toEqual([-1, -0.8, -0.6]);
    // 明示の min が優先
    const ex = computeValueScale({
      axis: valAxis({ min: -1 }),
      extent: { min: -0.3, max: 1 },
      percentStacked: true,
      lengthPx: 400,
      labelHeightPx: 12,
    });
    expect(ex.min).toBe(-1);
  });

  it('TD-102-15: 横向きの軸は最長の目盛の文字の幅（最小は文字の高さ × 3.25）で目盛の数の上限を決める', () => {
    expect(tickCountLimit(300, 12)).toBe(10); // 縦：高さ 12 × 1.0 ごと（上限 10）
    expect(tickCountLimit(300, 12, 66)).toBe(4); // 横：文字の幅 66 ごと
    expect(tickCountLimit(300, 12, 10)).toBe(7); // 横：文字が狭くても 12 × 3.25 ごと
    // 横：文字の幅に隙間（2）を足した間隔。幅 59 だけなら 5 目盛・隙間を足すと 4（隣の文字との隙間を 0 にしない）
    expect(tickCountLimit(300, 12, 59)).toBe(4);
    expect(tickCountLimit(30, 12, 66)).toBe(2);
    const wide = computeValueScale({
      extent: { min: 0, max: 87 },
      lengthPx: 300,
      labelHeightPx: 12,
      labelWidthPx: 66,
    });
    const tall = computeValueScale({
      extent: { min: 0, max: 87 },
      lengthPx: 300,
      labelHeightPx: 12,
    });
    expect(wide.majorTicks.length).toBeLessThan(tall.majorTicks.length);
  });
});
