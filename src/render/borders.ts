// 正: InterHub docs/plan/adr/0029-xlsx-cell-borders.md（§決定 3〜5・7）
/**
 * `.xlsx` のセルの罫線を、SVG の path にするための純関数（DOM 非依存・ISSUE-98）。
 * - 描いた行（`GridRow.borders`）と罫線の表（`XlsxSheet.borders`）から、すべての横の辺・縦の辺（と斜線）について
 *   「隣り合う 2 つの指定」から 1 つの線を決め（太さの順位 → 線種 → 左・上のセル）、結合セルの内側の辺は描かず、
 *   同じ境目の上で同じ線が続く区間を 1 本にまとめて、線種×色ごとの path の `d` にする。
 * - 線種の幅・パターン・太さの順位は {@link BORDER_STYLES}（定数 1 か所）。Excel の画像との照合（E1〜E5）で直すのはここだけ。
 * - 座標は A1 の左上を原点とする px（geometry の `colLeft`・`rowTop` と同じ座標）。1px の線はグリッド線の画素に重ねる。
 */
import type { BorderEdge, BorderStyle, CellBorder, Geometry } from '../core/index';
import type { GridRow } from './grid';

/** 1 回の描画で作る区間（線 1 本・斜線 1 本）の数の上限。超えた分は描かない（ADR-0029 §決定 7） */
export const MAX_BORDER_SEGMENTS = 50_000;

/** 結合セルの範囲を辺の判定用に印す作業量の上限（重なった巨大な結合を並べた細工でも固まらないため） */
const MERGE_MARK_BUDGET = 2_000_000;

/** 線種の見た目の分類（同じ太さの順位の中の優先：double ＞ 実線 ＞ 破線系 ＞ dotted ＞ hair。ADR-0029 §決定 4） */
const KIND_HAIR = 0;
const KIND_DOTTED = 1;
const KIND_DASHED = 2;
const KIND_SOLID = 3;
const KIND_DOUBLE = 4;

/** 線種 1 つの描き方 */
export interface BorderStyleSpec {
  /** 線の幅（px）。二重線は隙間を含む全体の幅（1px の線 2 本＋ 1px の隙間） */
  width: number;
  /** `stroke-dasharray` の数値（px）。実線は null */
  dash: readonly number[] | null;
  /** 太さの順位（大きいほど強い。辺の優先の 1 番目の基準） */
  weight: number;
  /** 線種の分類（同じ順位の中の優先。大きいほど強い） */
  kind: number;
  /** 二重線（隙間＋ 1px の線 2 本で描く） */
  double?: boolean;
}

/**
 * 13 線種の幅・パターン・優先の表（ADR-0029 §決定 3。暫定値で、E2 の照合で直す定数 1 か所）。
 * パターンの数値は px。`slantDashDot` の斜めの切り口は再現しない。
 */
export const BORDER_STYLES: Readonly<Record<BorderStyle, BorderStyleSpec>> = {
  thin: { width: 1, dash: null, weight: 2, kind: KIND_SOLID },
  hair: { width: 1, dash: [1, 1], weight: 1, kind: KIND_HAIR },
  dotted: { width: 1, dash: [1, 2], weight: 2, kind: KIND_DOTTED },
  dashed: { width: 1, dash: [3, 1], weight: 2, kind: KIND_DASHED },
  dashDot: { width: 1, dash: [8, 3, 3, 3], weight: 2, kind: KIND_DASHED },
  dashDotDot: { width: 1, dash: [8, 3, 3, 3, 3, 3], weight: 2, kind: KIND_DASHED },
  medium: { width: 2, dash: null, weight: 3, kind: KIND_SOLID },
  mediumDashed: { width: 2, dash: [9, 3], weight: 3, kind: KIND_DASHED },
  mediumDashDot: { width: 2, dash: [9, 3, 3, 3], weight: 3, kind: KIND_DASHED },
  mediumDashDotDot: { width: 2, dash: [9, 3, 3, 3, 3, 3], weight: 3, kind: KIND_DASHED },
  slantDashDot: { width: 2, dash: [11, 1, 5, 1], weight: 3, kind: KIND_DASHED },
  thick: { width: 3, dash: null, weight: 4, kind: KIND_SOLID },
  double: { width: 3, dash: null, weight: 4, kind: KIND_DOUBLE, double: true },
};

/** 線種の表の 1 件（未知の線種・Object.prototype の名前は undefined） */
function specOf(style: BorderStyle): BorderStyleSpec | undefined {
  return Object.prototype.hasOwnProperty.call(BORDER_STYLES, style)
    ? BORDER_STYLES[style]
    : undefined;
}

/**
 * 同じ辺の 2 つの指定から、描く 1 つを選ぶ（ADR-0029 §決定 4）。
 * 片方だけならそれ。両方あれば、太さの順位が大きい方 → 同じなら線種の分類（double ＞ 実線 ＞ 破線系 ＞ dotted ＞ hair）
 * → それでも同じなら左（上）のセルの指定 `a`。線種が表に無い指定は無いものとして扱う。
 *
 * @param a 左（上）のセルの、この辺（右・下）の指定
 * @param b 右（下）のセルの、この辺（左・上）の指定
 */
export function pickBorderEdge(
  a: BorderEdge | undefined,
  b: BorderEdge | undefined,
): BorderEdge | undefined {
  const sa = a ? specOf(a.style) : undefined;
  const sb = b ? specOf(b.style) : undefined;
  if (!sa) return sb ? b : undefined;
  if (!sb) return a;
  if (sa.weight !== sb.weight) return sa.weight > sb.weight ? a : b;
  if (sa.kind !== sb.kind) return sa.kind > sb.kind ? a : b;
  return a;
}

/** 1 本の path にまとめた線（線種×色×斜線かどうかごと） */
export interface BorderPath {
  /** 線種 */
  style: BorderStyle;
  /** 墨色（自動・黒）。色はテーマ変数で決める（`color` は null） */
  ink: boolean;
  /** 明示色（`#rrggbb`）。墨色なら null */
  color: string | null;
  /** 明示色の不透明度（0〜1。1 なら不透明） */
  alpha: number;
  /** `stroke-width`（px）。二重線は隙間の幅（3） */
  width: number;
  /** `stroke-dasharray`。実線・二重線は null */
  dasharray: string | null;
  /** 斜線の path か（縦横の線は画素にそろえて描くが、斜線はそろえない） */
  diagonal: boolean;
  /** path の `d`。二重線は中心線（隙間の描画に使う） */
  d: string;
  /** 二重線だけ：1px の線 2 本の `d`（中心線から 1px ずつ離した 2 本） */
  lineD?: string;
}

/** 描いた線の外接矩形（A1 の左上を 0 とする px。線の幅を含む） */
export interface BorderBounds {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** {@link buildBorderPaths} の結果 */
export interface BorderPaths {
  paths: BorderPath[];
  /** 線が 1 本も無ければ null */
  bounds: BorderBounds | null;
  /** 作った区間の数 */
  segments: number;
  /** 区間の上限で打ち切ったか */
  truncated: boolean;
}

/** 辺の判定で使う、解決済みの線（同じ線種・色なら同じ参照にする） */
interface Line {
  key: string;
  style: BorderStyle;
  spec: BorderStyleSpec;
  ink: boolean;
  color: string | null;
  alpha: number;
}

/** 数値を小数 3 桁に丸める（path の文字列を短く・安定させる） */
const num = (v: number): number => Math.round(v * 1000) / 1000;

/**
 * 辺の色が墨色（自動・黒）か、明示色（`#rrggbb`＋不透明度）かを決める。
 * null（自動）・不正な値・不透明な `#000000` は墨色。文字列は整数から組み立てる（ファイルの文字列は通らない）。
 */
function colorOf(edge: BorderEdge): { ink: boolean; color: string | null; alpha: number } {
  const c = edge.color;
  if (
    !c ||
    !Number.isInteger(c.rgb) ||
    c.rgb < 0 ||
    c.rgb > 0xffffff ||
    !Number.isFinite(c.alpha)
  ) {
    return { ink: true, color: null, alpha: 1 };
  }
  const alpha = Math.round(Math.min(1, Math.max(0, c.alpha)) * 1000) / 1000;
  if (c.rgb === 0 && alpha >= 1) {
    return { ink: true, color: null, alpha: 1 };
  }
  return { ink: false, color: `#${c.rgb.toString(16).padStart(6, '0')}`, alpha };
}

/** 隣り合う線が「同じ線」か（線種と色が同じ）を参照の比較で済ませるための、線の解決器 */
function createLineResolver(): (edge: BorderEdge | undefined) => Line | undefined {
  const byEdge = new Map<BorderEdge, Line | null>();
  const byKey = new Map<string, Line>();
  return (edge) => {
    if (!edge) return undefined;
    const cached = byEdge.get(edge);
    if (cached !== undefined) return cached ?? undefined;
    const spec = specOf(edge.style);
    let line: Line | null = null;
    if (spec) {
      const { ink, color, alpha } = colorOf(edge);
      const key = `${edge.style}|${ink ? 'ink' : `${color}@${alpha}`}`;
      line = byKey.get(key) ?? null;
      if (!line) {
        line = { key, style: edge.style, spec, ink, color, alpha };
        byKey.set(key, line);
      }
    }
    byEdge.set(edge, line);
    return line ?? undefined;
  };
}

/**
 * 線 1 本の「グリッド線の画素の位置 g」から、stroke の中心の座標を出す。
 * 1px はその画素の中心（g + 0.5）、2px はその画素と手前（左・上）の 1 画素（中心は g）、3px はその画素を中心に 1 画素ずつ（g + 0.5）。
 * 手前の端が 0 より前に出る時（A 列の左・1 行目の上）は、線の幅だけ内側へ寄せて、端を 0 にそろえる（sticky の見出しの下に隠れない）。
 */
function centerOf(g: number, width: number): number {
  const c = width === 2 ? g : g + 0.5;
  return c - width / 2 < 0 ? width / 2 : c;
}

/** 結合セルの範囲（rows の添字と列番号。両端を含む） */
interface MergeRect {
  i1: number;
  i2: number;
  c1: number;
  c2: number;
  id: number;
}

/**
 * 描いた範囲の罫線を、線種×色ごとの path にする。
 *
 * - 範囲の左端の縦の辺は左隣の列、右端の縦の辺は右隣の列のセル（範囲外でも `rows` は全列を持つ）の指定も見る。
 *   上端の横の辺は下のセルの上の辺だけ、下端の横の辺は上のセルの下の辺だけで決める。
 * - 結合セルは `GridRow` の左上のセルの `colSpan`・`rowSpan`（行の畳みで縮んだ見える行の数）から範囲を復元し、
 *   内側の辺は描かない。外周は面する構成セルの辺から取る。斜線は左上のセルの指定で結合範囲の対角に引く。
 * - 幅 0 の列・行の辺は描かない（`rows` は見える行だけ）。非表示の列・行をはさむ辺は、前後の見える列・行の指定で決める。
 * - 同じ境目の上で同じ線が続く区間は 1 本（`M x1 y H x2`・`M x y1 V y2`）にまとめる。区間の数が `maxSegments` を超えたら
 *   以後は描かない（`truncated`）。
 *
 * @param rows 描く行（見える行だけ・行番号の昇順。全列を持つ）
 * @param table 罫線の表（`XlsxSheet.borders`）
 * @param geo 寸法（`colLeft`・`colWidth`・`rowTop`）
 * @param range 描く列の範囲（1 始まり・両端を含む）
 * @param maxSegments 区間の数の上限（既定 {@link MAX_BORDER_SEGMENTS}）
 */
export function buildBorderPaths(
  rows: readonly GridRow[],
  table: readonly CellBorder[],
  geo: Pick<Geometry, 'colLeft' | 'colWidth' | 'rowTop'>,
  range: { first: number; last: number },
  maxSegments: number = MAX_BORDER_SEGMENTS,
): BorderPaths {
  const out: BorderPaths = { paths: [], bounds: null, segments: 0, truncated: false };
  const { first, last } = range;
  const R = rows.length;
  if (R === 0 || table.length === 0 || last < first || first < 1) return out;
  if (!rows.some((row) => row.borders !== undefined)) return out;

  const resolve = createLineResolver();
  const acc = new Map<string, { line: Line; diagonal: boolean; d: string[]; lineD: string[] }>();
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;

  const cellBorder = (i: number, c: number): CellBorder | undefined => {
    if (c < 1) return undefined;
    const idx = rows[i]!.borders?.[c - 1];
    return idx === undefined ? undefined : table[idx];
  };

  /** 区間 1 つを積む。上限なら false（以後を描かない） */
  const room = (): boolean => {
    if (out.segments >= maxSegments) {
      out.truncated = true;
      return false;
    }
    out.segments++;
    return true;
  };
  const bucket = (line: Line, diagonal: boolean) => {
    const key = `${diagonal ? 'd' : 'o'}|${line.key}`;
    let b = acc.get(key);
    if (!b) {
      b = { line, diagonal, d: [], lineD: [] };
      acc.set(key, b);
    }
    return b;
  };
  /** 外接矩形を広げる（padX・padY は線の幅の半分。線の向きに沿う側には足さない） */
  const grow = (
    x0: number,
    y0: number,
    x1: number,
    y1: number,
    padX: number,
    padY: number,
  ): void => {
    left = Math.min(left, x0 - padX, x1 - padX);
    right = Math.max(right, x0 + padX, x1 + padX);
    top = Math.min(top, y0 - padY, y1 - padY);
    bottom = Math.max(bottom, y0 + padY, y1 + padY);
  };

  /** 横線：y の中心を g（グリッド線の画素）から出し、x0〜x1 を 1 区間にする */
  const emitH = (line: Line, g: number, x0: number, x1: number): void => {
    if (!room()) return;
    const w = line.spec.width;
    const cy = centerOf(g, w);
    const b = bucket(line, false);
    b.d.push(`M${num(x0)} ${num(cy)}H${num(x1)}`);
    if (line.spec.double) {
      b.lineD.push(`M${num(x0)} ${num(cy - 1)}H${num(x1)}M${num(x0)} ${num(cy + 1)}H${num(x1)}`);
    }
    grow(x0, cy, x1, cy, 0, w / 2);
  };
  /** 縦線：x の中心を g から出し、y0〜y1 を 1 区間にする */
  const emitV = (line: Line, g: number, y0: number, y1: number): void => {
    if (!room()) return;
    const w = line.spec.width;
    const cx = centerOf(g, w);
    const b = bucket(line, false);
    b.d.push(`M${num(cx)} ${num(y0)}V${num(y1)}`);
    if (line.spec.double) {
      b.lineD.push(`M${num(cx - 1)} ${num(y0)}V${num(y1)}M${num(cx + 1)} ${num(y0)}V${num(y1)}`);
    }
    grow(cx, y0, cx, y1, w / 2, 0);
  };
  /** 斜線：(ax, ay) → (bx, by)。二重線は法線方向に 1px ずつ離した 2 本 */
  const emitDiag = (line: Line, ax: number, ay: number, bx: number, by: number): void => {
    if (!room()) return;
    const b = bucket(line, true);
    b.d.push(`M${num(ax)} ${num(ay)}L${num(bx)} ${num(by)}`);
    if (line.spec.double) {
      const len = Math.hypot(bx - ax, by - ay) || 1;
      const nx = -(by - ay) / len;
      const ny = (bx - ax) / len;
      b.lineD.push(
        `M${num(ax + nx)} ${num(ay + ny)}L${num(bx + nx)} ${num(by + ny)}` +
          `M${num(ax - nx)} ${num(ay - ny)}L${num(bx - nx)} ${num(by - ny)}`,
      );
    }
    grow(ax, ay, bx, by, line.spec.width / 2, line.spec.width / 2);
  };

  // --- 結合セルの範囲（内側の辺を除くため）。窓は列 first-1〜last+1 ---
  const W = last - first + 3;
  const col0 = first - 1; // 窓の添字 0 の列
  const ids = new Uint32Array(R * W);
  const merges: MergeRect[] = [];
  let budget = MERGE_MARK_BUDGET;
  for (let i = 0; i < R; i++) {
    const cells = rows[i]!.cells;
    const upto = Math.min(cells.length, last);
    for (let ci = 0; ci < upto; ci++) {
      const cell = cells[ci];
      if (!cell || (cell.colSpan <= 1 && cell.rowSpan <= 1)) continue;
      const c1 = ci + 1;
      const c2 = c1 + Math.max(1, cell.colSpan) - 1;
      if (c2 < col0) continue;
      const i2 = Math.min(R - 1, i + Math.max(1, cell.rowSpan) - 1);
      const m: MergeRect = { i1: i, i2, c1, c2, id: merges.length + 1 };
      merges.push(m);
      const a = Math.max(c1, col0) - col0;
      const z = Math.min(c2, last + 1) - col0;
      for (let k = i; k <= i2 && budget > 0; k++) {
        for (let x = a; x <= z && budget > 0; x++, budget--) ids[k * W + x] = m.id;
      }
    }
  }
  const idAt = (i: number, c: number): number => {
    const x = c - col0;
    return x < 0 || x >= W ? 0 : ids[i * W + x]!;
  };

  // --- 見える列（幅 > 0）と、範囲の外の最も近い見える列 ---
  const vis: number[] = [];
  for (let c = first; c <= last; c++) if (geo.colWidth(c) > 0) vis.push(c);
  let widest = 0;
  for (const row of rows) widest = Math.max(widest, row.cells.length, row.borders?.length ?? 0);
  let prevVis = 0;
  for (let c = first - 1; c >= 1; c--) {
    if (geo.colWidth(c) > 0) {
      prevVis = c;
      break;
    }
  }
  let nextVis = 0;
  for (let c = last + 1; c <= widest; c++) {
    if (geo.colWidth(c) > 0) {
      nextVis = c;
      break;
    }
  }

  /** 連続する番号 0..n-1 に、線 `at(k)` を割り当て、同じ線の続く区間ごとに `emit(start, end, line)` を呼ぶ */
  const runs = (
    n: number,
    at: (k: number) => Line | undefined,
    emit: (start: number, end: number, line: Line) => void,
  ): void => {
    let cur: Line | undefined;
    let start = 0;
    for (let k = 0; k < n; k++) {
      const line = at(k);
      if (line !== cur) {
        if (cur) emit(start, k - 1, cur);
        cur = line;
        start = k;
        if (out.truncated) return;
      }
    }
    if (cur && !out.truncated) emit(start, n - 1, cur);
  };

  // --- 縦の辺（列の境目ごと。左のセルの右・右のセルの左） ---
  for (let t = 0; t <= vis.length && !out.truncated; t++) {
    const lc = t === 0 ? prevVis : vis[t - 1]!;
    const rc = t < vis.length ? vis[t]! : nextVis;
    if (vis.length === 0) break;
    const gx = (t < vis.length ? geo.colLeft(rc) : geo.colLeft(lc + 1)) - 1;
    runs(
      R,
      (i) => {
        if (lc > 0 && rc > 0) {
          const id = idAt(i, lc);
          if (id !== 0 && id === idAt(i, rc)) return undefined; // 結合の内側
        }
        const a = lc > 0 ? cellBorder(i, lc)?.right : undefined;
        const b = rc > 0 ? cellBorder(i, rc)?.left : undefined;
        return resolve(pickBorderEdge(a, b));
      },
      (s, e, line) => emitV(line, gx, geo.rowTop(rows[s]!.r), geo.rowTop(rows[e]!.r + 1)),
    );
  }

  // --- 横の辺（行の境目ごと。上のセルの下・下のセルの上。端は片側だけ） ---
  for (let j = 0; j <= R && !out.truncated && vis.length > 0; j++) {
    const gy = (j < R ? geo.rowTop(rows[j]!.r) : geo.rowTop(rows[R - 1]!.r + 1)) - 1;
    runs(
      vis.length,
      (t) => {
        const c = vis[t]!;
        if (j > 0 && j < R) {
          const id = idAt(j - 1, c);
          if (id !== 0 && id === idAt(j, c)) return undefined; // 結合の内側
        }
        const a = j > 0 ? cellBorder(j - 1, c)?.bottom : undefined;
        const b = j < R ? cellBorder(j, c)?.top : undefined;
        return resolve(pickBorderEdge(a, b));
      },
      (s, e, line) => emitH(line, gy, geo.colLeft(vis[s]!), geo.colLeft(vis[e]! + 1)),
    );
  }

  // --- 斜線：結合セルは左上のセルの指定で結合範囲の対角、ほかは各セルの対角 ---
  const diag = (
    cb: CellBorder | undefined,
    i1: number,
    i2: number,
    c1: number,
    c2: number,
  ): void => {
    if (!cb?.diagonal || (!cb.diagonalUp && !cb.diagonalDown)) return;
    const line = resolve(cb.diagonal);
    if (!line) return;
    const x0 = geo.colLeft(c1);
    const x1 = geo.colLeft(c2 + 1);
    const y0 = geo.rowTop(rows[i1]!.r);
    const y1 = geo.rowTop(rows[i2]!.r + 1);
    if (x1 <= x0 || y1 <= y0) return;
    if (cb.diagonalDown) emitDiag(line, x0, y0, x1, y1);
    if (cb.diagonalUp) emitDiag(line, x0, y1, x1, y0);
  };
  for (const m of merges) {
    if (out.truncated) break;
    if (m.c1 > last || m.c2 < first) continue;
    diag(cellBorder(m.i1, m.c1), m.i1, m.i2, m.c1, m.c2);
  }
  for (let i = 0; i < R && !out.truncated; i++) {
    const row = rows[i]!;
    if (row.borders === undefined) continue;
    for (const c of vis) {
      const cell = row.cells[c - 1];
      if (cell === null || (cell && (cell.colSpan > 1 || cell.rowSpan > 1))) continue;
      diag(cellBorder(i, c), i, i, c, c);
    }
  }

  for (const b of acc.values()) {
    const { line } = b;
    const path: BorderPath = {
      style: line.style,
      ink: line.ink,
      color: line.color,
      alpha: line.alpha,
      width: line.spec.width,
      dasharray: line.spec.dash ? line.spec.dash.join(' ') : null,
      diagonal: b.diagonal,
      d: b.d.join(''),
    };
    if (line.spec.double) path.lineD = b.lineD.join('');
    out.paths.push(path);
  }
  if (out.paths.length > 0) out.bounds = { left, top, right, bottom };
  return out;
}
