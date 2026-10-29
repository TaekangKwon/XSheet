import { describe, expect, it } from 'vitest';
import type { BorderEdge, BorderStyle, CellBorder } from '../../src/core/index';
import type { GridCell, GridRow } from '../../src/render/grid';
import {
  BORDER_STYLES,
  MAX_BORDER_SEGMENTS,
  buildBorderPaths,
  pickBorderEdge,
  type BorderPath,
} from '../../src/render/borders';

/**
 * 罫線の純関数（ISSUE-98・ADR-0029 §決定 3〜5・7）の単体テスト。
 * 寸法は固定の擬似 geometry（既定の列幅 40px・行高 20px。非表示の列は幅 0）にして、座標を手計算で確かめる。
 */

const ink = (style: BorderStyle): BorderEdge => ({ style, color: null });
const thin = ink('thin');

/** 擬似 geometry：列 c の幅は `widths[c]`（無ければ 40）、行高は 20 */
function geoOf(widths: Record<number, number> = {}) {
  const w = (c: number): number => (c < 1 ? 0 : (widths[c] ?? 40));
  return {
    colWidth: w,
    colLeft(c: number): number {
      let x = 0;
      for (let k = 1; k < c; k++) x += w(k);
      return x;
    },
    rowTop: (r: number): number => Math.max(0, r - 1) * 20,
  };
}

/** 行 2 が非表示（高さ 0）の geometry：行 3 の上端は y＝20 */
function foldedGeo() {
  return {
    ...geoOf(),
    rowTop: (r: number): number => (r <= 2 ? Math.max(0, r - 1) * 20 : (r - 2) * 20),
  };
}

type Spec = CellBorder | undefined;

/** 罫線の指定（列ごと）の 2 次元配列から、GridRow（行番号は 1 始まり）と罫線の表を作る */
function gridOf(
  specs: Spec[][],
  cellsOf: (GridCell | null | undefined)[][] = [],
  firstRow = 1,
): { rows: GridRow[]; table: CellBorder[] } {
  const table: CellBorder[] = [];
  const rows = specs.map((line, i) => {
    const borders = line.map((b) => {
      if (!b) return undefined;
      table.push(b);
      return table.length - 1;
    });
    const cells = cellsOf[i] ?? line.map(() => undefined);
    return {
      r: firstRow + i,
      cells,
      borders: borders.some((v) => v !== undefined) ? borders : undefined,
    };
  });
  return { rows: rows.map((row) => (row.borders ? row : { r: row.r, cells: row.cells })), table };
}

const b4 = (e: BorderEdge): CellBorder => ({
  left: e,
  right: e,
  top: e,
  bottom: e,
  diagonalUp: false,
  diagonalDown: false,
});
const sides = (p: Partial<CellBorder>): CellBorder => ({
  diagonalUp: false,
  diagonalDown: false,
  ...p,
});

const merged = (colSpan: number, rowSpan: number): GridCell => ({
  text: '',
  colSpan,
  rowSpan,
  merged: true,
});

/** path をまとめて見やすくする（線種・色・d） */
const dOf = (paths: BorderPath[]): string[] =>
  paths.map((p) => `${p.style}:${p.ink ? 'ink' : p.color}:${p.d}`);

describe('TD-98-2 線種の表と位置', () => {
  it('thin の縦の辺はグリッド線の画素（列 1 の右＝x 39）に重なる（中心 39.5・幅 1）', () => {
    const { rows, table } = gridOf([[sides({ right: thin })]]);
    const r = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 3 });
    expect(r.paths).toHaveLength(1);
    expect(r.paths[0]).toMatchObject({
      style: 'thin',
      ink: true,
      color: null,
      width: 1,
      dasharray: null,
      d: 'M39.5 0V20',
    });
    expect(r.bounds).toEqual({ left: 39, top: 0, right: 40, bottom: 20 });
  });

  it('thin の横の辺は行 1 の下（y＝rowTop(2)−1＝19）の画素に重なる', () => {
    const { rows, table } = gridOf([[sides({ bottom: thin })]]);
    const r = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 1 });
    expect(r.paths[0]!.d).toBe('M0 19.5H40');
  });

  it('medium はその画素と手前の 1 画素（中心 g・幅 2）、thick は中心の画素を挟む 3 画素（中心 g+0.5・幅 3）', () => {
    const m = buildBorderPaths(...args({ right: ink('medium') }));
    expect(m.paths[0]).toMatchObject({ width: 2, d: 'M39 0V20' });
    const t = buildBorderPaths(...args({ right: ink('thick') }));
    expect(t.paths[0]).toMatchObject({ width: 3, d: 'M39.5 0V20' });
    expect(t.bounds).toEqual({ left: 38, top: 0, right: 41, bottom: 20 });
  });

  it('二重線は中心線（隙間用・幅 3）と、1px ずつ離した 2 本の d（lineD）を持つ', () => {
    const r = buildBorderPaths(...args({ right: ink('double') }));
    expect(r.paths[0]).toMatchObject({
      style: 'double',
      width: 3,
      d: 'M39.5 0V20',
      lineD: 'M38.5 0V20M40.5 0V20',
    });
    const h = buildBorderPaths(...args({ bottom: ink('double') }));
    expect(h.paths[0]!.lineD).toBe('M0 18.5H40M0 20.5H40');
  });

  it('A 列の左の辺・1 行目の上の辺は、線の幅だけ内側へ寄せる（端を 0 にそろえる）', () => {
    const l1 = buildBorderPaths(...args({ left: thin }));
    expect(l1.paths[0]!.d).toBe('M0.5 0V20');
    const l2 = buildBorderPaths(...args({ left: ink('medium') }));
    expect(l2.paths[0]!.d).toBe('M1 0V20');
    const l3 = buildBorderPaths(...args({ left: ink('thick') }));
    expect(l3.paths[0]!.d).toBe('M1.5 0V20');
    const t1 = buildBorderPaths(...args({ top: ink('medium') }));
    expect(t1.paths[0]!.d).toBe('M0 1H40');
  });

  it('13 線種の幅と stroke-dasharray が表どおり', () => {
    const expected: Record<BorderStyle, [number, string | null]> = {
      thin: [1, null],
      hair: [1, '1 1'],
      dotted: [1, '1 2'],
      dashed: [1, '3 1'],
      dashDot: [1, '8 3 3 3'],
      dashDotDot: [1, '8 3 3 3 3 3'],
      medium: [2, null],
      mediumDashed: [2, '9 3'],
      mediumDashDot: [2, '9 3 3 3'],
      mediumDashDotDot: [2, '9 3 3 3 3 3'],
      slantDashDot: [2, '11 1 5 1'],
      thick: [3, null],
      double: [3, null],
    };
    expect(Object.keys(BORDER_STYLES).sort()).toEqual(Object.keys(expected).sort());
    for (const [style, [width, dash]] of Object.entries(expected)) {
      const r = buildBorderPaths(...args({ right: ink(style as BorderStyle) }));
      expect(r.paths[0], style).toMatchObject({ style, width, dasharray: dash });
    }
  });
});

/** 1 行 1 列・範囲 1〜1 の geo 付き引数 */
function args(spec: Partial<CellBorder>): Parameters<typeof buildBorderPaths> {
  const { rows, table } = gridOf([[sides(spec)]]);
  return [rows, table, geoOf(), { first: 1, last: 1 }];
}

describe('TD-98-3 辺の優先', () => {
  const e = (style: BorderStyle, rgb: number | null = null): BorderEdge => ({
    style,
    color: rgb === null ? null : { rgb, alpha: 1 },
  });

  it('片方だけの指定はそれを採る（どちら側でも）', () => {
    expect(pickBorderEdge(thin, undefined)).toBe(thin);
    expect(pickBorderEdge(undefined, thin)).toBe(thin);
    expect(pickBorderEdge(undefined, undefined)).toBeUndefined();
  });

  it('太さの順位が大きい方（thick ＞ medium ＞ thin ＞ hair）', () => {
    expect(pickBorderEdge(e('thin'), e('medium'))).toEqual(e('medium'));
    expect(pickBorderEdge(e('thick'), e('medium'))).toEqual(e('thick'));
    expect(pickBorderEdge(e('hair'), e('thin'))).toEqual(e('thin'));
    expect(pickBorderEdge(e('mediumDashed'), e('thin'))).toEqual(e('mediumDashed'));
  });

  it('同じ順位なら線種の分類：double ＞ 実線 ＞ 破線系 ＞ dotted', () => {
    expect(pickBorderEdge(e('thick', 1), e('double', 2))).toEqual(e('double', 2));
    expect(pickBorderEdge(e('dashed', 1), e('thin', 2))).toEqual(e('thin', 2));
    expect(pickBorderEdge(e('dotted', 1), e('dashDot', 2))).toEqual(e('dashDot', 2));
    expect(pickBorderEdge(e('dotted', 1), e('dashed', 2))).toEqual(e('dashed', 2));
    expect(pickBorderEdge(e('medium', 1), e('mediumDashDot', 2))).toEqual(e('medium', 1));
  });

  it('それでも同じなら左（上）のセルの指定 a', () => {
    const a = e('thin', 0xff0000);
    const b = e('thin', 0x0000ff);
    expect(pickBorderEdge(a, b)).toBe(a);
    expect(pickBorderEdge(e('dashed', 1), e('dashDot', 2))).toEqual(e('dashed', 1));
  });

  it('縦の辺：A の右 thin と B の左 medium は medium を描く。横の辺：上の下と下の上でも同じ規則', () => {
    const { rows, table } = gridOf([
      [sides({ right: thin }), sides({ left: ink('medium') })],
      [sides({ bottom: ink('dashed') }), sides({})],
      [sides({ top: ink('thin') }), undefined],
    ]);
    const r = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 2 });
    expect(dOf(r.paths)).toEqual(['medium:ink:M39 0V20', 'thin:ink:M0 39.5H40']);
  });
});

describe('TD-98-4 区間の連結と path のまとめ', () => {
  it('同じ境目の上で同じ線が続く区間は 1 本、違う線で切れる', () => {
    const { rows, table } = gridOf([
      [sides({ right: thin })],
      [sides({ right: thin })],
      [sides({ right: ink('dashed') })],
      [sides({ right: thin })],
    ]);
    const r = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 1 });
    expect(dOf(r.paths)).toEqual(['thin:ink:M39.5 0V40M39.5 60V80', 'dashed:ink:M39.5 40V60']);
    expect(r.segments).toBe(3);
  });

  it('横の辺も同じ：続く列を 1 本にして、線種×色ごとの path に連ねる', () => {
    const { rows, table } = gridOf([
      [sides({ bottom: thin }), sides({ bottom: thin }), sides({ bottom: thin })],
    ]);
    const r = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 3 });
    expect(dOf(r.paths)).toEqual(['thin:ink:M0 19.5H120']);
  });

  it('色：自動・不透明な #000000 は墨色、ほかは #rrggbb。同じ線種でも色が違えば別の path', () => {
    const red: BorderEdge = { style: 'thin', color: { rgb: 0xff0000, alpha: 1 } };
    const black: BorderEdge = { style: 'thin', color: { rgb: 0, alpha: 1 } };
    const faint: BorderEdge = { style: 'thin', color: { rgb: 0, alpha: 0.5 } };
    const { rows, table } = gridOf([
      [
        sides({ bottom: thin }),
        sides({ bottom: red }),
        sides({ bottom: black }),
        sides({ bottom: faint }),
      ],
    ]);
    const r = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 4 });
    expect(r.paths.map((p) => [p.ink, p.color, p.alpha])).toEqual([
      [true, null, 1],
      [false, '#ff0000', 1],
      [false, '#000000', 0.5],
    ]);
    // 自動（列 1）と黒（列 3）は同じ墨色の同じ線。間に別の線（列 2）があるので 2 区間
    expect(r.paths[0]!.d).toBe('M0 19.5H40M80 19.5H120');
  });

  it('罫線が無い・表が空なら path も bounds も無い', () => {
    const { rows } = gridOf([[undefined, undefined]]);
    expect(buildBorderPaths(rows, [], geoOf(), { first: 1, last: 2 })).toMatchObject({
      paths: [],
      bounds: null,
    });
    const r2 = buildBorderPaths([{ r: 1, cells: [] }], [b4(thin)], geoOf(), { first: 1, last: 2 });
    expect(r2.paths).toEqual([]);
  });
});

describe('TD-98-5 結合セル', () => {
  /** A1:B2 の結合。4 つの構成セルすべてに 4 辺の thin（内側にも指定がある） */
  it('内側の辺（構成セルの間）は描かず、外周だけを 1 本ずつ描く', () => {
    const c = b4(thin);
    const { rows, table } = gridOf(
      [
        [c, c],
        [c, c],
      ],
      [
        [merged(2, 2), null],
        [null, null],
      ],
    );
    const r = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 2 });
    // 縦：左端・右端（中の境目 x=39 は結合の内側）／横：上端・下端（中の境目 y=19 は内側）
    expect(dOf(r.paths)).toEqual(['thin:ink:M0.5 0V40M79.5 0V40M0 0.5H80M0 39.5H80']);
  });

  it('外周の辺は面する構成セルの指定から取る（左上のセルの指定を外周全体に広げない）', () => {
    // A1:A2 の結合。下の辺は A2（構成セル）の bottom だけが持つ
    const { rows, table } = gridOf(
      [[sides({})], [sides({ bottom: thin })]],
      [[merged(1, 2)], [null]],
    );
    expect(dOf(buildBorderPaths(rows, table, geoOf(), { first: 1, last: 1 }).paths)).toEqual([
      'thin:ink:M0 39.5H40',
    ]);
    // 左上のセルだけが bottom を持つ場合、それは結合の内側の辺なので描かない
    const g2 = gridOf([[sides({ bottom: thin })], [sides({})]], [[merged(1, 2)], [null]]);
    expect(buildBorderPaths(g2.rows, g2.table, geoOf(), { first: 1, last: 1 }).paths).toEqual([]);
  });

  it('外周の外側のセルとの優先は通常の規則。縮んで 1×1 になった結合（merged）は結合として扱わない', () => {
    const { rows, table } = gridOf(
      [[sides({ right: thin }), sides({ left: ink('medium') })]],
      [[merged(1, 1), undefined]],
    );
    expect(dOf(buildBorderPaths(rows, table, geoOf(), { first: 1, last: 2 }).paths)).toEqual([
      'medium:ink:M39 0V20',
    ]);
  });

  it('範囲の左外に起点がある結合でも、内側の辺を描かない（行は全列を持つ）', () => {
    // 列 2〜3 が結合（起点は列 2）。範囲は 3〜3：左端の縦の辺（列 2 の右・列 3 の左）は結合の内側
    const c = b4(thin);
    const { rows, table } = gridOf([[undefined, c, c]], [[undefined, merged(2, 1), null]]);
    const r = buildBorderPaths(rows, table, geoOf(), { first: 3, last: 3 });
    expect(dOf(r.paths)).toEqual(['thin:ink:M119.5 0V20M80 0.5H120M80 19.5H120']);
  });

  it('縦に結合した行の畳み（rowSpan は見える行の数）：見える行の添字で内側を判定する', () => {
    // 行 1・3 が見える（行 2 は非表示で rows に無い）。A1:A3 の結合は rowSpan 2 に縮んで来る
    const c = b4(thin);
    const { rows, table } = gridOf([[c], [c]], [[merged(1, 2)], [null]]);
    rows[1] = { ...rows[1]!, r: 3 };
    const r = buildBorderPaths(rows, table, foldedGeo(), { first: 1, last: 1 });
    // 行 1 の下・行 3 の上の境目（内側）は描かない。外周：左・右（高さは 2 行ぶん 40）・上（行 1 の上）・下（行 3 の下）
    expect(r.paths[0]!.d).toBe('M0.5 0V40M39.5 0V40M0 0.5H40M0 39.5H40');
  });
});

describe('TD-98-6 斜線', () => {
  const diagCell = (up: boolean, down: boolean, e: BorderEdge = thin): CellBorder => ({
    diagonal: e,
    diagonalUp: up,
    diagonalDown: down,
  });

  it('diagonalDown は左上→右下、diagonalUp は左下→右上（セルの対角）', () => {
    const down = gridOf([[diagCell(false, true)]]);
    expect(
      dOf(buildBorderPaths(down.rows, down.table, geoOf(), { first: 1, last: 1 }).paths),
    ).toEqual(['thin:ink:M0 0L40 20']);
    const up = gridOf([[diagCell(true, false)]]);
    expect(dOf(buildBorderPaths(up.rows, up.table, geoOf(), { first: 1, last: 1 }).paths)).toEqual([
      'thin:ink:M0 20L40 0',
    ]);
    const both = gridOf([[diagCell(true, true)]]);
    const r = buildBorderPaths(both.rows, both.table, geoOf(), { first: 1, last: 1 });
    expect(r.paths).toHaveLength(1);
    expect(r.paths[0]).toMatchObject({ diagonal: true, d: 'M0 0L40 20M0 20L40 0' });
  });

  it('斜線は縦横の線と別の path（画素にそろえない）', () => {
    const g = gridOf([[{ ...sides({ right: thin }), ...diagCell(false, true) }]]);
    const r = buildBorderPaths(g.rows, g.table, geoOf(), { first: 1, last: 1 });
    expect(r.paths.map((p) => p.diagonal)).toEqual([false, true]);
  });

  it('結合セルでは左上のセルの指定で、結合範囲の対角に引く（構成セルの斜線は引かない）', () => {
    const { rows, table } = gridOf(
      [
        [diagCell(false, true), diagCell(false, true)],
        [diagCell(false, true), diagCell(false, true)],
      ],
      [
        [merged(2, 2), null],
        [null, null],
      ],
    );
    const r = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 2 });
    expect(dOf(r.paths)).toEqual(['thin:ink:M0 0L80 40']);
  });

  it('二重線の斜線は法線方向に 1px ずつ離した 2 本', () => {
    const g = gridOf([[diagCell(false, true, ink('double'))]]);
    const r = buildBorderPaths(g.rows, g.table, geoOf(), { first: 1, last: 1 });
    // 対角 (0,0)→(40,20)：法線 (-20, 40)/44.721
    expect(r.paths[0]!.d).toBe('M0 0L40 20');
    expect(r.paths[0]!.lineD).toBe('M-0.447 0.894L39.553 20.894M0.447 -0.894L40.447 19.106');
  });

  it('斜線の指定が無い（diagonalUp・Down とも false）なら描かない', () => {
    const g = gridOf([[{ diagonal: thin, diagonalUp: false, diagonalDown: false }]]);
    expect(buildBorderPaths(g.rows, g.table, geoOf(), { first: 1, last: 1 }).paths).toEqual([]);
  });
});

describe('TD-98-7 範囲の端・非表示', () => {
  it('左端の縦の辺は左隣の列（範囲の外）の右の辺も見る', () => {
    const { rows, table } = gridOf([[sides({ right: thin }), undefined]]);
    const r = buildBorderPaths(rows, table, geoOf(), { first: 2, last: 2 });
    expect(dOf(r.paths)).toEqual(['thin:ink:M39.5 0V20']);
  });

  it('右端の縦の辺は右隣の列（範囲の外）の左の辺も見る', () => {
    const { rows, table } = gridOf([[undefined, sides({ left: thin })]]);
    const r = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 1 });
    expect(dOf(r.paths)).toEqual(['thin:ink:M39.5 0V20']);
  });

  it('上端の横の辺は下のセルの上の辺だけで決める。下端は上のセルの下の辺だけ', () => {
    const { rows, table } = gridOf([[sides({ top: ink('medium'), bottom: thin })]]);
    const r = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 1 });
    expect(dOf(r.paths)).toEqual(['medium:ink:M0 1H40', 'thin:ink:M0 19.5H40']);
  });

  it('幅 0 の列の辺は描かず、非表示の列をはさむ辺は前後の見える列の指定で決める', () => {
    // 列 2 が非表示。列 1 の右 thin → 見える次の列 3 の左との境目（x＝colLeft(3)−1＝39）に 1 本
    const { rows, table } = gridOf([
      [sides({ right: thin }), sides({ right: thin, left: thin }), sides({})],
    ]);
    const r = buildBorderPaths(rows, table, geoOf({ 2: 0 }), { first: 1, last: 3 });
    expect(dOf(r.paths)).toEqual(['thin:ink:M39.5 0V20']);
  });

  it('非表示の行をはさむ横の辺は、上下の描いた行（rows は見える行だけ）の指定で決める', () => {
    // 行 1 と行 3 が隣り合う（行 2 は非表示）。行 1 の下 thin と行 3 の上 medium → medium
    const { rows, table } = gridOf([[sides({ bottom: thin })], [sides({ top: ink('medium') })]]);
    rows[1] = { ...rows[1]!, r: 3 };
    const r = buildBorderPaths(rows, table, foldedGeo(), { first: 1, last: 1 });
    // 行 3 の上＝y 20 の境目（行 2 は高さ 0）。medium は画素 19 と手前の 1 画素（中心 19）
    expect(dOf(r.paths)).toEqual(['medium:ink:M0 19H40']);
  });
});

describe('TD-98-8 区間の上限', () => {
  it('MAX_BORDER_SEGMENTS は 50,000', () => {
    expect(MAX_BORDER_SEGMENTS).toBe(50_000);
  });

  it('区間の数が上限を超えたら以後を描かない（truncated）', () => {
    // 行ごとに thin・dashed を交互に右の辺へ：区間は 10 本
    const specs = Array.from({ length: 10 }, (_, i) => [
      sides({ right: i % 2 ? ink('dashed') : thin }),
    ]);
    const { rows, table } = gridOf(specs);
    const full = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 1 });
    expect(full).toMatchObject({ segments: 10, truncated: false });
    const cut = buildBorderPaths(rows, table, geoOf(), { first: 1, last: 1 }, 3);
    expect(cut).toMatchObject({ segments: 3, truncated: true });
    const drawn = cut.paths.reduce((n, p) => n + (p.d.match(/M/g)?.length ?? 0), 0);
    expect(drawn).toBe(3);
  });
});
