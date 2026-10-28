/// <reference types="node" />
/**
 * 実際の Excel で作った基準ファイルによる、図形・線・文字の色の照合（ISSUE-95・ADR-0028 §決定 6・U7・U8）。
 * `fixtures/excel/ref-colors.xlsx` は Excel 16 の COM（`make-ref-colors.ps1`・PowerShell 5.1・画面 125%）で作成・保存した実ファイル（図形 128 個）。
 * 期待値は同じフォルダの `ref-colors.expected.json`（`make-ref-colors.ps1` が書く `com.json` そのもの）に図形名ごとに入っている：
 * - COM の色の値（`Fill.ForeColor.RGB`・`Line.ForeColor.RGB`・`Fill.Transparency`・`GradientStops`）。Excel 自身が解いた色
 * - Excel が描いた画素（`samples`＝塗りの中央・面ごとの点、`stroke`＝太い線の筋。CopyPicture → グラフ → PNG の書き出し。`make-ref-colors.extract.mjs` が PNG から `stroke` を足す）
 * 貼られた絵は白い背景の矩形を含むので、透過の混色の基準は背景のマゼンタ（`bg`）でなく白（alpha 0.5 の赤が #FF7F7F）。
 * 解いた色は各成分 ±3 以内（模様だけは画素が点描なので ±8）で比べる。
 */
import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { applyFillMode } from '../../src/core/color';
import { openXlsx, type SheetDrawing } from '../../src/core/index';
import type { DrawingColor } from '../../src/core/types';

interface Com {
  fill: {
    Visible: number;
    Type: number;
    Transparency: number;
    ForeRGB: string;
    GradientStops?: { Position: number; RGB: string }[];
  };
  line: { Visible: number; Weight: number; Transparency: number; ForeRGB: string };
}
interface RefShape extends Com {
  name: string;
  group: string;
  input: Record<string, unknown>;
  bg: number[];
  box: number[];
  samples?: { fx: number; fy: number; rgb: string }[];
  stroke?: { axis: string; rgb: string };
}

const dir = new URL('../../fixtures/excel/', import.meta.url);
const expected = JSON.parse(readFileSync(new URL('ref-colors.expected.json', dir), 'utf8')) as {
  excelVersion: string;
  bg: number[];
  notes: string[];
  shapes: RefShape[];
};

/** 'RRGGBB' → [r, g, b] */
const rgbOf = (hex: string): [number, number, number] => [
  parseInt(hex.slice(0, 2), 16),
  parseInt(hex.slice(2, 4), 16),
  parseInt(hex.slice(4, 6), 16),
];
const unpack = (n: number): [number, number, number] => [(n >> 16) & 255, (n >> 8) & 255, n & 255];
/** 各成分の差の最大 */
const diff = (a: readonly number[], b: readonly number[]): number =>
  Math.max(...a.map((v, i) => Math.abs(v - b[i]!)));
/** 貼られた絵の背景（白）。透過の色はこの上に重なった見た目になる */
const WHITE = [255, 255, 255] as const;
/** 色を白の上に重ねた見た目 */
const over = (c: DrawingColor): number[] =>
  unpack(c.rgb).map((v, i) => v * c.alpha + WHITE[i]! * (1 - c.alpha));

let byName: Map<string, SheetDrawing>;
const refs = expected.shapes;
/** 比べた差の最大（失敗したときの手がかり） */
const worst = new Map<string, number>();
const track = (label: string, d: number): number => {
  worst.set(label, Math.max(worst.get(label) ?? 0, d));
  return d;
};

beforeAll(async () => {
  const bytes = new Uint8Array(readFileSync(new URL('ref-colors.xlsx', dir)));
  const { drawings, notices } = await (await openXlsx(bytes)).openDrawings(1);
  expect(notices).toEqual([]);
  byName = new Map(drawings.map((d) => [d.name, d]));
});

describe('TD-95-20 ref-colors.xlsx：全図形が読める', () => {
  it('図形 128 個が名前で引け、塗り・枠線・線の色がある（有無は Excel の COM と同じ）', () => {
    expect(expected.excelVersion).toBe('16.0');
    expect(refs).toHaveLength(128);
    for (const r of refs) {
      const d = byName.get(r.name);
      expect(d, r.name).toBeDefined();
      if (d!.kind === 'line') {
        expect(d!.line!.color, r.name).toBeDefined();
        continue;
      }
      expect(d!.hasFill === true, `${r.name} hasFill`).toBe(r.fill.Visible === -1);
      expect(d!.fillColor !== undefined, `${r.name} fillColor`).toBe(d!.hasFill);
      expect(d!.outlineColor !== undefined, `${r.name} outlineColor`).toBe(d!.hasOutline);
    }
  });
});

describe('TD-95-21 U7：解いた色は Excel の COM の色（Fill/Line.ForeColor.RGB）と各成分 ±3 以内', () => {
  /** 単色の塗り（COM の Fill.Type が 1＝msoFillSolid）でテーマ色・明るさ・RGB・既定の図形 */
  const solids = refs.filter((r) => r.fill.Visible === -1 && r.fill.Type === 1);

  it('塗り：テーマ色 12 種×明るさ 6 段階（72）・srgb・面の濃淡の図形・既定の図形・透過（alpha＝1−Transparency）', () => {
    expect(solids.length).toBeGreaterThanOrEqual(100);
    for (const r of solids) {
      const c = byName.get(r.name)!.fillColor!;
      const d = track(`fill:${r.group}`, diff(unpack(c.rgb), rgbOf(r.fill.ForeRGB)));
      expect(
        d,
        `${r.name} 塗り ${c.rgb.toString(16)} vs COM ${r.fill.ForeRGB}`,
      ).toBeLessThanOrEqual(3);
      expect(Math.abs(c.alpha - (1 - r.fill.Transparency)), `${r.name} alpha`).toBeLessThanOrEqual(
        0.01,
      );
    }
  });

  it('枠線・線：既定の図形（accent1＋shade）・テキストボックス（lt1＋shade）・太い線（RGB・テーマ色＋明るさ）・線', () => {
    const lines = refs.filter(
      (r) => r.line.Visible === -1 && ['default', 'outline'].includes(r.group),
    );
    expect(lines.length).toBeGreaterThanOrEqual(10);
    for (const r of lines) {
      const d0 = byName.get(r.name)!;
      const c = (d0.kind === 'line' ? d0.line!.color : d0.outlineColor)!;
      const d = track(`line:${r.group}`, diff(unpack(c.rgb), rgbOf(r.line.ForeRGB)));
      expect(d, `${r.name} 線 ${c.rgb.toString(16)} vs COM ${r.line.ForeRGB}`).toBeLessThanOrEqual(
        3,
      );
      expect(Math.abs(c.alpha - (1 - r.line.Transparency)), `${r.name} alpha`).toBeLessThanOrEqual(
        0.01,
      );
    }
  });

  it('グラデーション：COM の stop の数と位置（Excel が書いた順は乱れていても 0〜1 の範囲）', () => {
    // 色の照合は画素（次の describe）で行う。ここは stop が読めていること（COM の stop は 2 つ・3 つ）
    expect(
      refs.filter((x) => x.group === 'gradient').map((r) => r.fill.GradientStops!.length),
    ).toEqual([2, 2, 2, 3]);
    for (const r of refs.filter((x) => x.group === 'gradient'))
      expect(byName.get(r.name)!.fillColor, r.name).toBeDefined();
  });
});

describe('TD-95-22 Excel が描いた画素との照合（画素が記録された図形だけ）', () => {
  const withPixels = refs.filter((r) => (r.samples?.length ?? 0) > 0 || r.stroke !== undefined);
  /**
   * 許容。模様は ±8：Excel は 8x8 のタイルを画面 125% で描くので、被覆率が名前の割合からずれる（pct25 は 27.5%）。
   * グラデーションは ±6：勾配が 1 行で約 5 進むので、中央の 1px の取り方で 3 ずれる（赤→青の中央の画素は #BC00B8、2 stop は線形 RGB の混色 #BC00BC）。
   */
  const tol = (r: RefShape): number => (r.group === 'pattern' ? 8 : r.group === 'gradient' ? 6 : 3);
  /**
   * 模様のうち照合するもの。pct5・pct25・pct50 は被覆率が名前の割合に近い。pct75（画素 #272727＝被覆 85%）・pct90（#0A0A0A＝96%）・horz（#E3E3E3＝11%）・dkHorz（#848484）は
   * 名前の割合・「その他は 50%」の近似から外れる（1 色の代表色にする近似の限界。ADR-0028 §決定 6.4）ので、差を記録するだけで期待にしない。
   */
  const PATTERNS_CHECKED = new Set([
    'patt_pct5',
    'patt_pct25',
    'patt_pct50',
    'patt_pct25_red_blue',
    'patt_smCheck',
  ]);

  it.runIf(withPixels.length > 0)(
    '塗り・太い線・文字：解いた色（透過は背景との混色）が画素と ±3（模様は ±8）',
    () => {
      for (const r of withPixels.filter((x) => x.group !== 'face' && x.group !== 'blip')) {
        const d0 = byName.get(r.name)!;
        if (r.group === 'text') {
          const runs = d0.text!.paragraphs[0]!.runs;
          r.samples?.forEach((s, i) => {
            const c = (runs[i] ?? runs[0]!).color!;
            // 文字の色の点（図形の塗りの隅の点＝2 つ目以降のうち塗りがあるものは対象外）
            if (i < runs.length || i === 0)
              expect(
                track('pixel:text', diff(over(c), rgbOf(s.rgb))),
                `${r.name} #${i}`,
              ).toBeLessThanOrEqual(3);
          });
          continue;
        }
        // 塗り：最初の点（図形の中央）
        const fillPx = r.samples?.[0]?.rgb;
        if (fillPx !== undefined && d0.fillColor !== undefined && r.fill.Visible === -1) {
          if (r.group === 'pattern' && !PATTERNS_CHECKED.has(r.name)) {
            track(`pixel:pattern(not asserted)`, diff(over(d0.fillColor), rgbOf(fillPx)));
            continue;
          }
          expect(
            track(`pixel:${r.group}`, diff(over(d0.fillColor), rgbOf(fillPx))),
            `${r.name} 塗り`,
          ).toBeLessThanOrEqual(tol(r));
        }
        // 枠線・線：太い線（3pt 以上）の最初の筋。細い線は縁がにじむので対象外
        const lineColor = d0.kind === 'line' ? d0.line!.color : d0.outlineColor;
        if (r.stroke !== undefined && lineColor !== undefined) {
          expect(
            track('pixel:line', diff(over(lineColor), rgbOf(r.stroke.rgb))),
            `${r.name} 線`,
          ).toBeLessThanOrEqual(3);
        }
      }
    },
  );

  /**
   * 面ごとの path の `fill`（`xlsx-preset-geom.json`）：U8 の照合に使う。
   * Excel の画素から決めた係数（accent1 #4472C4 の面）：lighten #8FAADC・lightenLess #698ED0・darkenLess #375C9E・darken #294476
   * （FF0000 では #FF6666・#FF3232・#CD0000・#990000。808080 では #B3B3B3・#999999・#676767・#4D4D4D）。
   * 成分の比が darken・lighten は 0.6、darkenLess・lightenLess は 0.802〜0.809（→ 0.805）で、いずれも sRGB の値に直接掛けた結果。
   * 額縁（foldedCorner）の折り返しは、サンプル点が三角形の縁にかかって混色になる（#3F6AB7・#A0B2D2）ので期待に使わない。
   */
  const FACE_MODES: Record<string, Record<string, string>> = {
    can: { top: 'lighten', body: 'norm' },
    cube: { front: 'norm', top: 'lightenLess', side: 'darkenLess' },
    foldedCorner: { body: 'norm' },
    bevel: {
      center: 'norm',
      top: 'lightenLess',
      left: 'lighten',
      right: 'darken',
      bottom: 'darkenLess',
    },
  };

  it.runIf(withPixels.some((r) => r.group === 'face'))(
    'U8：面の濃淡（darken・lighten）が Excel の面の画素と ±3',
    () => {
      for (const r of withPixels.filter((x) => x.group === 'face')) {
        const names = r.input.sampleNames as string[];
        const fill = byName.get(r.name)!.fillColor!;
        r.samples!.forEach((s, i) => {
          const mode = FACE_MODES[r.input.shape as string]![names[i]!];
          if (mode === undefined) return;
          const want = applyFillMode(fill, mode);
          expect(
            track('face', diff(unpack(want.rgb), rgbOf(s.rgb))),
            `${r.name} ${names[i]} (${mode})`,
          ).toBeLessThanOrEqual(3);
        });
      }
    },
  );
});
