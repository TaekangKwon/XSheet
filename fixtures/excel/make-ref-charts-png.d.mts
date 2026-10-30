// make-ref-charts-png.mjs の型宣言（テストから import するため。本体は依存なしの Node 用の純ロジック）

/** COM の座標の原点のずれ（pt） */
export const COM_ORIGIN_PT: number;
/** 0xRRGGBB → 'RRGGBB' */
export function toHex(p: number): string;

/** png-index（make-ref-charts2.ps1 の png 段）の 1 件 */
export interface PngIndexEntry {
  Name: string;
  Source?: string;
  File?: string;
  ChartStyle?: number | null;
  WidthPt: number;
  HeightPt: number;
  PlotInside?: { Left?: number; Top?: number; Width?: number; Height?: number };
  Horizontal?: boolean;
  HasAxes?: boolean;
  ValueAxis?: { Min?: number; Max?: number; Major?: number; Minor?: number } | null;
  ExcludeRGB?: string[];
  /** ISSUE-110：真ならグラフの面の左上の角を読む（corner） */
  Corner?: boolean;
  /** ISSUE-110：真なら数値軸の目盛の印を読む（valTicks） */
  Ticks?: boolean;
  /** ISSUE-110：COM の CategoryAxis の位置（pt）。あれば項目軸の文字の塊を読む（labels） */
  LabelBand?: { Left?: number; Top?: number; Width?: number; Height?: number } | null;
}

/** グラフの面の左上の角（scanCorner） */
export interface PngCorner {
  rounded: boolean;
  transparent: boolean;
  outside: string | null;
  radiusPx: number;
  runs: number[];
}

/** 項目軸の文字の塊 1 つ（px・角度は度で 0＝水平・正＝右上がり） */
export interface PngLabelCluster {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
  w: number;
  h: number;
  n: number;
  angle: number;
}

/** 項目軸の文字の塊の読み取り（scanLabels） */
export interface PngLabels {
  region: { x0: number; y0: number; x1: number; y1: number };
  count: number;
  clusters: PngLabelCluster[];
  medianW: number | null;
  medianH: number | null;
  medianAngle: number | null;
}

/** 数値軸の目盛の印の片側（scanTicks） */
export interface PngTickSide {
  major: number;
  minor: number;
  other: number;
  majorLen: number;
  minorLen: number;
  list: { pos: number; len: number; kind: 'major' | 'minor' | 'other' }[];
}

/** 数値軸の目盛の印（外側・内側） */
export interface PngValTicks {
  cap: number;
  out: PngTickSide;
  in: PngTickSide;
}

/** 読んだ線 1 本 */
export interface PngLine {
  dir: 'h' | 'v';
  /** 画素の位置（行・列。線の行を濃さで重み付けした平均） */
  pos: number;
  /** プロットの内側に対する位置（値の方向の線は上端・左端が 0、値の軸の線は縦棒なら左端 0） */
  rel: number;
  role: 'majorGrid' | 'minorGrid' | 'catAxisLine' | 'valAxisLine' | 'catGrid' | 'other';
  rows: { p: number; color: string; count: number }[];
  /** 背景から最も離れた行の色（RRGGBB） */
  peak: string;
  /** （背景 − 画素）の行ごとの合計（R, G, B） */
  ink: number[];
}

/** 1 枚の PNG の読み取り結果 */
export interface PngAnalysis {
  width: number;
  height: number;
  scale: number;
  scaleY: number;
  bg: string;
  hasAxes: boolean;
  plotPx?: { x0: number; y0: number; x1: number; y1: number };
  lines: PngLine[];
  /** ISSUE-110：index に Corner があるグラフ */
  corner?: PngCorner;
  /** ISSUE-110：index に LabelBand があるグラフ */
  labels?: PngLabels;
  /** ISSUE-110：index に Ticks があるグラフ */
  valTicks?: PngValTicks;
  summary: Partial<
    Record<PngLine['role'], { count: number; peak: string; ink: number[]; rows: number }>
  >;
}

/** ref-charts-png.json の 1 件 */
export interface PngChartResult extends PngAnalysis {
  source?: string;
  chartStyle: number | null;
}

/** 1 枚の PNG から線を読む */
export function analyzeChartPng(
  img: { width: number; height: number; data: Uint8Array },
  entry: PngIndexEntry,
  opts?: { d?: number; minFraction?: number; marginPt?: number },
): PngAnalysis;

/** グラフの面の左上の角を読む（ISSUE-110） */
export function scanCorner(img: { width: number; height: number; data: Uint8Array }): PngCorner;

/** 範囲の中の文字の塊を読む（ISSUE-110）。bg＝背景の色（0xRRGGBB）・g＝同じ塊とみなす画素の距離 */
export function scanLabels(
  img: { width: number; height: number; data: Uint8Array },
  bg: number,
  region: { x0: number; y0: number; x1: number; y1: number },
  g: number,
): PngLabels;

/** 数値軸の目盛の印を読む（ISSUE-110） */
export function scanTicks(
  img: { width: number; height: number; data: Uint8Array },
  bg: number,
  geo: {
    vert: boolean;
    x0: number;
    y0: number;
    x1: number;
    y1: number;
    scale: number;
    majorPos: number[];
    minorPos: number[];
    tol: number;
  },
): PngValTicks;

/**
 * <out dir>/png の png-index*.json と PNG を読み、<out dir>/ref-charts-png.json を書く。書いた内容を返す。
 * opts.out＝出力ファイル名（既定 ref-charts-png.json）・opts.index＝png-index の名前の正規表現（既定 ^png-index.*\.json$）
 */
export function run(
  dir: string,
  opts?: { out?: string; index?: string },
): {
  note: string;
  excelVersion: string | null;
  notes: string[];
  charts: Record<string, PngChartResult>;
};
