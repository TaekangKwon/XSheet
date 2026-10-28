// 正: InterHub docs/plan/adr/0028-xlsx-preset-shape-geometry.md（§決定 1・2・5・7）
/**
 * `.xlsx` の図形の「形」の定義（DrawingML の guide の式と path 命令）を評価する純ロジック（DOM 非依存・InterHub ISSUE-90）。
 * 流れ：形の定義（表の preset か `custGeom`）＋調整値＋回転前の矩形の幅・高さ（px）
 * → 組み込み変数 → `avLst`（既定を調整値で上書き）→ `gdLst` を上から 1 回ずつ → path 命令を経路（直線・3 次ベジェ）へ。
 * 式は自前の解釈器で評価し、`eval`・`new Function` は使わない（docs/01_requirements.md N1）。
 * 名前（guide・調整値・`prst`）は `Map` のキーにだけ使い、普通のオブジェクトのキーにしない。
 * 反転・回転・シート範囲への丸め・`d` の組み立ては描画側（ISSUE-92・`xlsx-shapes.ts` の写像）が行う。ここは回転前の矩形の中の経路まで。
 * 形の表は生成物 `preset-geom.json`（変換器 `tools/geom/build-preset-geom.mjs`・出典はリポジトリ直下の NOTICE）。
 * コメント中の ISSUE-N・ADR-N は InterHub の番号。
 */
import { GUIDE_ARITY, MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES, type PathSeg, type Pt } from './defs';
import presetTable from './preset-geom.json';

// 1 図形の guide・path 命令の数の上限は `defs.ts`（ADR-0028 §決定 7。読み取り側と同じ値）。評価器の利用側・テストの import 先は変えない
export { MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES };

/** 1px あたりの EMU（式は EMU に換算した大きさで評価し、結果を px に戻す） */
const EMU_PER_PX = 9525;
/**
 * guide の値を有限に保つための丸め（±10^30）。弧矢印の「直線と円の交点」の式は EMU の 4 乗（約 10^20）の中間値を作る。
 * ここを 10^15 にすると判別式 `r²·d² − D²` の両辺が同じ上限に張り付いて 0 になり、交点が壊れる（ISSUE-93）
 */
const VALUE_LIMIT = 1e30;
/** path の座標・半径・角度と文字領域に使う値の丸め（±10^15）。描画側へ巨大な値を渡さない */
const COORD_LIMIT = 1e15;
/** 角度の単位（1/60000 度）→ ラジアン */
const UNIT_TO_RAD = Math.PI / 180 / 60000;
/** ラジアン → 角度の単位（1/60000 度） */
const RAD_TO_UNIT = (180 / Math.PI) * 60000;
/** 弧の振り角の上限（360°。1/60000 度）。これを超える分は同じ円周をなぞるだけ */
const MAX_SWEEP = 360 * 60000;

/** path の塗りの濃淡の種類（`none`＝塗らない） */
export type GeomFill = 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';

/** 式・命令の引数：整数、または組み込み変数・guide の名前 */
export type GeomArg = number | string;

/** guide 1 つ：`[名前, 演算子, 引数…]`（例 `['x1', '+-', 'w', 'adj1', 0]`） */
export type GeomGuide = readonly [name: string, op: string, ...args: GeomArg[]];

/** path 命令 1 つ：`['M', x, y]`・`['L', x, y]`・`['A', wR, hR, stAng, swAng]`・`['Q', x1, y1, x, y]`・`['C', x1, y1, x2, y2, x, y]`・`['Z']` */
export type GeomCmd = readonly [op: string, ...args: GeomArg[]];

/** path 1 つの定義 */
export interface GeomPathDef {
  /** 座標系の幅・高さ（無い・0 なら形の座標系＝EMU のまま） */
  w?: number;
  h?: number;
  /** 塗りの濃淡の種類（無ければ `norm`） */
  fill?: GeomFill;
  /** 線を引くか（無ければ true） */
  stroke?: boolean;
  /** 命令の列 */
  c: readonly GeomCmd[];
}

/** 形の定義（表の preset も `custGeom` もこの形に揃える） */
export interface GeomDef {
  /** 調整値の既定（`avLst`）：`[名前, 値]` */
  av?: ReadonlyArray<readonly [string, number]>;
  /** guide の列（`gdLst`）。名前は「それより前」に定義したものだけ引ける */
  gd?: readonly GeomGuide[];
  /** 文字領域（l・t・r・b の引数。無ければ矩形全体） */
  rect?: readonly GeomArg[];
  /** path の列 */
  paths: readonly GeomPathDef[];
  /**
   * 評価した経路・文字領域を矩形の中で上下反転する（y → h − y）。原本に定義の無い `upArrow` を `downArrow` から作るために使う（`getPresetGeom`）。
   * ファイルの `custGeom` は持たない
   */
  flipY?: boolean;
}

/** 評価の失敗の理由（代替の描き方は ISSUE-92）。`name`＝未知の名前・前方参照、`op`＝未知の演算子、`args`＝引数の数の誤り、`cmd`＝未知の path 命令、`limit`＝上限超え */
export type GeomFailure = 'name' | 'op' | 'args' | 'cmd' | 'limit';

/** 部分経路：始点・区間の列・閉じるか（区間は `xlsx-shapes.ts` の直線 `L`・3 次ベジェ `C` と同じ表し方） */
export interface GeomSubpath {
  start: Pt;
  segs: PathSeg[];
  closed: boolean;
}

/** path 1 つの経路（px・回転前の矩形の左上が原点・y は下向き） */
export interface GeomPath {
  fill: GeomFill;
  stroke: boolean;
  subpaths: GeomSubpath[];
}

/** 文字領域（px・回転前の矩形の左上が原点） */
export interface GeomTextRect {
  l: number;
  t: number;
  r: number;
  b: number;
}

/** 評価の結果：成功なら経路と文字領域、失敗なら理由（例外にはしない） */
export type GeomResult =
  { ok: true; paths: GeomPath[]; textRect: GeomTextRect } | { ok: false; reason: GeomFailure };

/** guide の評価の結果（値は EMU） */
export type GuidesResult =
  { ok: true; values: Map<string, number> } | { ok: false; reason: GeomFailure };

/** 演算子ごとの引数の数（読み取り側の字句の検査と同じ表。`defs.ts`） */
const ARITY: ReadonlyMap<string, number> = GUIDE_ARITY;

/** path 命令ごとの引数の数 */
const CMD_ARITY: ReadonlyMap<string, number> = new Map([
  ['M', 2],
  ['L', 2],
  ['A', 4],
  ['Q', 4],
  ['C', 6],
  ['Z', 0],
]);

const FILLS: ReadonlySet<string> = new Set<GeomFill>([
  'none',
  'norm',
  'lighten',
  'lightenLess',
  'darken',
  'darkenLess',
]);

/** 有限でなければ 0、±10^30 に丸める（-0 も 0 にする） */
function fin(v: number): number {
  if (!Number.isFinite(v) || v === 0) return 0;
  return v > VALUE_LIMIT ? VALUE_LIMIT : v < -VALUE_LIMIT ? -VALUE_LIMIT : v;
}

/** 0 除算は 0 */
function div(a: number, b: number): number {
  return b === 0 ? 0 : a / b;
}

/** 演算子 1 つを評価する（演算子と引数の数は呼び出し側が検査済み） */
function applyOp(op: string, a: readonly number[]): number {
  const [x, y, z] = a;
  switch (op) {
    case '*/':
      return div(x * y, z);
    case '+-':
      return x + y - z;
    case '+/':
      return div(x + y, z);
    case '?:':
      return x > 0 ? y : z;
    case 'abs':
      return Math.abs(x);
    case 'sqrt':
      return x < 0 ? 0 : Math.sqrt(x);
    case 'val':
      return x;
    case 'max':
      return Math.max(x, y);
    case 'min':
      return Math.min(x, y);
    case 'mod':
      return Math.sqrt(x * x + y * y + z * z);
    case 'pin':
      return y < x ? x : y > z ? z : y;
    case 'sat2':
      return x * Math.sin(Math.atan2(z, y));
    case 'cat2':
      return x * Math.cos(Math.atan2(z, y));
    case 'sin':
      return x * Math.sin(y * UNIT_TO_RAD);
    case 'cos':
      return x * Math.cos(y * UNIT_TO_RAD);
    case 'tan':
      return x * Math.tan(y * UNIT_TO_RAD);
    case 'at2':
      return Math.atan2(y, x) * RAD_TO_UNIT;
    default:
      return 0;
  }
}

/** 組み込み変数（約 40 個）の表。W・H は EMU に換算した幅・高さ */
function builtins(W: number, H: number): Map<string, number> {
  const ss = Math.min(W, H);
  return new Map<string, number>([
    ['w', W],
    ['h', H],
    ['l', 0],
    ['t', 0],
    ['r', W],
    ['b', H],
    ['hc', W / 2],
    ['vc', H / 2],
    ['ss', ss],
    ['ls', Math.max(W, H)],
    ['wd2', W / 2],
    ['wd3', W / 3],
    ['wd4', W / 4],
    ['wd5', W / 5],
    ['wd6', W / 6],
    ['wd8', W / 8],
    ['wd10', W / 10],
    ['wd12', W / 12],
    ['wd32', W / 32],
    ['hd2', H / 2],
    ['hd3', H / 3],
    ['hd4', H / 4],
    ['hd5', H / 5],
    ['hd6', H / 6],
    ['hd8', H / 8],
    ['ssd2', ss / 2],
    ['ssd4', ss / 4],
    ['ssd6', ss / 6],
    ['ssd8', ss / 8],
    ['ssd16', ss / 16],
    ['ssd32', ss / 32],
    ['cd2', 180 * 60000],
    ['cd4', 90 * 60000],
    ['cd8', 45 * 60000],
    ['3cd4', 270 * 60000],
    ['3cd8', 135 * 60000],
    ['5cd8', 225 * 60000],
    ['7cd8', 315 * 60000],
  ]);
}

/** 引数を数値にする：数値はそのまま（有限に丸める）、文字列は名前として引く。引けなければ null */
function resolve(a: GeomArg, ctx: ReadonlyMap<string, number>): number | null {
  if (typeof a === 'number') return fin(a);
  if (typeof a === 'string') return ctx.get(a) ?? null;
  return null;
}

/** {@link resolve} の結果を path の座標の範囲（±10^15）に丸める（path 命令・文字領域用） */
function resolveCoord(a: GeomArg, ctx: ReadonlyMap<string, number>): number | null {
  const v = resolve(a, ctx);
  return v === null ? null : v > COORD_LIMIT ? COORD_LIMIT : v < -COORD_LIMIT ? -COORD_LIMIT : v;
}

/** 寸法（px）を有限・非負にする */
function dim(v: number): number {
  return Number.isFinite(v) && v > 0 ? v : 0;
}

/**
 * guide を評価して名前 → 値（EMU）の表にする。
 * 順序は 組み込み変数 → `av`（`adjust` の同じ名前の値で上書き。`av` に無い名前は無視）→ `gd` を上から 1 回ずつ。
 * 前方参照・未知の名前・未知の演算子・引数の数の誤り・上限超えは例外にせず失敗として返す。
 *
 * @param def 形の定義
 * @param adjust ファイルの調整値（名前 → 値）。無ければ既定のまま
 * @param w 回転前の矩形の幅（px）
 * @param h 回転前の矩形の高さ（px）
 */
export function evaluateGuides(
  def: GeomDef,
  adjust: ReadonlyMap<string, number> | undefined,
  w: number,
  h: number,
): GuidesResult {
  const av = def.av ?? [];
  const gd = def.gd ?? [];
  if (av.length + gd.length > MAX_GEOM_GUIDES) return { ok: false, reason: 'limit' };
  const ctx = builtins(dim(w) * EMU_PER_PX, dim(h) * EMU_PER_PX);
  for (const [name, dflt] of av) {
    // 渡された調整値が有限でなければ（NaN・Infinity）既定値に戻す
    const over = adjust?.get(name);
    ctx.set(name, fin(typeof over === 'number' && Number.isFinite(over) ? over : dflt));
  }
  for (const g of gd) {
    const [name, op, ...args] = g;
    const arity = ARITY.get(op);
    if (arity === undefined) return { ok: false, reason: 'op' };
    if (args.length !== arity) return { ok: false, reason: 'args' };
    const nums: number[] = [];
    for (const a of args) {
      const v = resolve(a, ctx);
      if (v === null) return { ok: false, reason: 'name' };
      nums.push(v);
    }
    ctx.set(name, fin(applyOp(op, nums)));
  }
  return { ok: true, values: ctx };
}

/** 視角 a（中心から見た方向・ラジアン）を、楕円（半径 rx・ry）の媒介変数の角度へ換算する（巻き数を保つ） */
function parametricAngle(a: number, rx: number, ry: number): number {
  const wrapped = Math.atan2(Math.sin(a), Math.cos(a));
  const turns = Math.round((a - wrapped) / (2 * Math.PI));
  return Math.atan2(rx * Math.sin(a), ry * Math.cos(a)) + 2 * Math.PI * turns;
}

/** path 1 つを経路にする（座標は path の座標系で計算し、最後に矩形の大きさへ線形に写す） */
function buildPath(
  pd: GeomPathDef,
  ctx: ReadonlyMap<string, number>,
  w: number,
  h: number,
): GeomPath | GeomFailure {
  const sx =
    pd.w !== undefined && Number.isFinite(pd.w) && pd.w > 0 ? dim(w) / pd.w : 1 / EMU_PER_PX;
  const sy =
    pd.h !== undefined && Number.isFinite(pd.h) && pd.h > 0 ? dim(h) / pd.h : 1 / EMU_PER_PX;
  const subpaths: GeomSubpath[] = [];
  /** 今の点と、今の部分経路の始点（path の座標系） */
  let cx = 0;
  let cy = 0;
  let ssx = 0;
  let ssy = 0;
  const st: { cur: GeomSubpath | null } = { cur: null };

  /** 開いている部分経路を確定する（区間が無ければ捨てる） */
  const flush = (): void => {
    if (st.cur && st.cur.segs.length > 0) subpaths.push(st.cur);
    st.cur = null;
  };
  /** 部分経路が無ければ今の点から始める */
  const open = (): GeomSubpath => {
    if (!st.cur) {
      ssx = cx;
      ssy = cy;
      st.cur = { start: { x: cx * sx, y: cy * sy }, segs: [], closed: false };
    }
    return st.cur;
  };
  const cubic = (x1: number, y1: number, x2: number, y2: number, x: number, y: number): void => {
    open().segs.push({
      t: 'C',
      x1: x1 * sx,
      y1: y1 * sy,
      x2: x2 * sx,
      y2: y2 * sy,
      x: x * sx,
      y: y * sy,
    });
    cx = x;
    cy = y;
  };

  for (const cmd of pd.c) {
    const [op, ...args] = cmd;
    const arity = CMD_ARITY.get(op);
    if (arity === undefined) return 'cmd';
    if (args.length !== arity) return 'args';
    const v: number[] = [];
    for (const a of args) {
      const n = resolveCoord(a, ctx);
      if (n === null) return 'name';
      v.push(n);
    }
    switch (op) {
      case 'M':
        flush();
        cx = ssx = v[0];
        cy = ssy = v[1];
        break;
      case 'L':
        open().segs.push({ t: 'L', x: v[0] * sx, y: v[1] * sy });
        cx = v[0];
        cy = v[1];
        break;
      case 'Q': {
        // 2 次ベジェ → 3 次ベジェ（誤差なし）：制御点を端点から 2/3 だけ制御点へ寄せる
        const [qx, qy, x, y] = v;
        cubic(
          cx + ((qx - cx) * 2) / 3,
          cy + ((qy - cy) * 2) / 3,
          x + ((qx - x) * 2) / 3,
          y + ((qy - y) * 2) / 3,
          x,
          y,
        );
        break;
      }
      case 'C':
        cubic(v[0], v[1], v[2], v[3], v[4], v[5]);
        break;
      case 'A': {
        const [rx, ry, stAng, swAng] = v;
        const sweep = Math.max(-MAX_SWEEP, Math.min(MAX_SWEEP, swAng));
        if (!(rx > 0 && ry > 0) || sweep === 0) break;
        // 角度は視角（中心から見た方向）。媒介変数の角度へ換算して 90° 以下に分け、3 次ベジェで近似する
        const a1 = stAng * UNIT_TO_RAD;
        const t1 = parametricAngle(a1, rx, ry);
        const total = parametricAngle(a1 + sweep * UNIT_TO_RAD, rx, ry) - t1;
        if (!Number.isFinite(total) || Math.abs(total) < 1e-12) break;
        const n = Math.max(1, Math.ceil(Math.abs(total) / (Math.PI / 2) - 1e-9));
        const delta = total / n;
        const k = (4 / 3) * Math.tan(delta / 4);
        const ox = cx - rx * Math.cos(t1);
        const oy = cy - ry * Math.sin(t1);
        for (let i = 0; i < n; i++) {
          const ta = t1 + i * delta;
          const tb = ta + delta;
          const [ca, sa, cb, sb] = [Math.cos(ta), Math.sin(ta), Math.cos(tb), Math.sin(tb)];
          cubic(
            ox + rx * (ca - k * sa),
            oy + ry * (sa + k * ca),
            ox + rx * (cb + k * sb),
            oy + ry * (sb - k * cb),
            ox + rx * cb,
            oy + ry * sb,
          );
        }
        break;
      }
      case 'Z':
        if (st.cur) {
          st.cur.closed = true;
          flush();
        }
        // 閉じた後の今の点は、その部分経路の始点（続く lnTo は始点から新しい部分経路を始める）
        cx = ssx;
        cy = ssy;
        break;
    }
  }
  flush();
  return {
    fill: pd.fill !== undefined && FILLS.has(pd.fill) ? pd.fill : 'norm',
    stroke: pd.stroke !== false,
    subpaths,
  };
}

/**
 * 形の定義を評価して、回転前の矩形の中の経路と文字領域を返す（座標は px・矩形の左上が原点）。
 * 失敗（未知の名前・演算子・命令、引数の数の誤り、上限超え）は例外にせず `{ ok: false, reason }` で返す。
 * 経路は直線と 3 次ベジェだけ（`quadBezTo` は誤差なしで変換・`arcTo` は 90° 以下に分けて近似。誤差は半径の 0.03% 未満）。
 * 文字領域は `rect` の l・t・r・b を評価した値。`rect` が無い・幅か高さが 0 以下なら矩形全体。
 * 反転した `rect`（r ≦ l など）は、調整値の極端な値で式が裏返ったものとみなして矩形全体へ戻す
 * （幅 0 以下の領域には文字が並ばず、図形ごと文字が消えるより、図形の矩形に並べる方が読める）。
 *
 * @param def 形の定義（表の preset か `custGeom`）
 * @param adjust ファイルの調整値（名前 → 値）
 * @param w 回転前の矩形の幅（px）
 * @param h 回転前の矩形の高さ（px）
 */
export function evaluateGeom(
  def: GeomDef,
  adjust: ReadonlyMap<string, number> | undefined,
  w: number,
  h: number,
): GeomResult {
  let commands = 0;
  for (const p of def.paths) commands += p.c.length;
  if (commands > MAX_GEOM_COMMANDS) return { ok: false, reason: 'limit' };
  const g = evaluateGuides(def, adjust, w, h);
  if (!g.ok) return g;
  const paths: GeomPath[] = [];
  for (const pd of def.paths) {
    const p = buildPath(pd, g.values, w, h);
    if (typeof p === 'string') return { ok: false, reason: p };
    paths.push(p);
  }
  let textRect: GeomTextRect = { l: 0, t: 0, r: dim(w), b: dim(h) };
  if (def.rect && def.rect.length === 4) {
    const r = def.rect.map((a) => resolveCoord(a, g.values));
    if (r.some((n) => n === null)) return { ok: false, reason: 'name' };
    const [l, t, rr, b] = (r as number[]).map((n) => n / EMU_PER_PX);
    if (rr - l > 0 && b - t > 0) textRect = { l, t, r: rr, b };
  }
  if (def.flipY === true) {
    // 点の写像だけで足りる（`arcTo` は 3 次ベジェに変換済み）。文字領域は t' = h − b・b' = h − t
    const H = dim(h);
    const fy = (y: number): number => H - y;
    for (const p of paths) {
      for (const sp of p.subpaths) {
        sp.start = { x: sp.start.x, y: fy(sp.start.y) };
        sp.segs = sp.segs.map((s) =>
          s.t === 'L'
            ? { t: 'L', x: s.x, y: fy(s.y) }
            : { t: 'C', x1: s.x1, y1: fy(s.y1), x2: s.x2, y2: fy(s.y2), x: s.x, y: fy(s.y) },
        );
      }
    }
    textRect = { l: textRect.l, t: fy(textRect.b), r: textRect.r, b: fy(textRect.t) };
  }
  return { ok: true, paths, textRect };
}

/** 形の表（`prst` → 定義）。生成物から初回に 1 度だけ `Map` に積む */
let presetMap: Map<string, GeomDef> | null = null;

function presets(): Map<string, GeomDef> {
  if (!presetMap) {
    const shapes = (presetTable as unknown as { shapes: Record<string, GeomDef> }).shapes;
    presetMap = new Map(Object.entries(shapes));
  }
  return presetMap;
}

/** 上下反転の別名：原本（ECMA-376 の `presetShapeDefinitions.xml`）に定義の無い preset を、ある preset の上下反転として扱う（`upArrow`＝`downArrow` の上下反転。矢先が上） */
const FLIPPED_ALIASES: ReadonlyMap<string, string> = new Map([['upArrow', 'downArrow']]);

/** 別名の定義（初回に 1 度だけ作る）。元の定義を共有し、`flipY` だけ足す */
const aliasCache = new Map<string, GeomDef>();

/** preset 名（`prst`）から形の定義を引く。表に無い（線の 10 種・未知の名前）なら undefined。`upArrow` は原本に無いので `downArrow` の上下反転（`flipY`）で返す */
export function getPresetGeom(prst: string): GeomDef | undefined {
  const own = presets().get(prst);
  if (own !== undefined) return own;
  const base = FLIPPED_ALIASES.get(prst);
  if (base === undefined) return undefined;
  let def = aliasCache.get(prst);
  if (def === undefined) {
    const src = presets().get(base);
    if (src === undefined) return undefined;
    def = { ...src, flipY: true };
    aliasCache.set(prst, def);
  }
  return def;
}

/** 表にある preset 名の一覧（原本の出現順。別名の `upArrow` は含めない） */
export function presetNames(): string[] {
  return [...presets().keys()];
}
