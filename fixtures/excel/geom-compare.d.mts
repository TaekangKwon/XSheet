// geom-compare.mjs の型宣言（テストから import するため。本体は依存なしの Node 用の純ロジック）

/** 読んだ画像（RGBA 8bit・width×height×4） */
export interface RgbaImage {
  width: number;
  height: number;
  data: Uint8Array;
}

/** インクのマスク（1＝インク・行優先） */
export interface Mask {
  w: number;
  h: number;
  data: Uint8Array;
}

/** 矩形（x1・y1 を含む） */
export interface Box {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

/** 輪郭の距離（95 パーセンタイル）の許容（100% 換算の px） */
export const TOLERANCE_PX: number;
/** Excel（画面 125%）と XSheet の写し（deviceScaleFactor 1.25）の、100% の 1px あたりの画素数 */
export const SCALE_125: number;

/** フィルタ済みの走査線から PNG を組み立てる（テスト・検査用） */
export function encodePngRaw(o: {
  width: number;
  height: number;
  colorType: number;
  bitDepth: number;
  raw: Uint8Array;
  plte?: Uint8Array;
  trns?: Uint8Array;
}): Uint8Array;
/** RGBA（8bit）を PNG にする */
export function encodePng(width: number, height: number, rgba: Uint8Array): Uint8Array;
/** PNG を RGBA（8bit）に読む */
export function decodePng(buf: Uint8Array): RgbaImage;

/** Excel の画像の、背景（マゼンタ）でない部分の外接矩形。無ければ null */
export function pictureRect(img: RgbaImage, bg?: number[], tol?: number): Box | null;
/** 画像（の矩形）からインクのマスクを取る */
export function inkMask(
  img: RgbaImage,
  rect?: Box | null,
  opts?: { fraction?: number; floor?: number },
): Mask;
/** マスクのインクの外接矩形。無ければ null */
export function maskBox(mask: Mask): Box | null;
/** マスクを矩形で切り出す */
export function cropMask(mask: Mask, box: Box): Mask;
/** マスクをインクの外接矩形で切り出す（インクが無ければ空の 0×0） */
export function cropToInk(mask: Mask): Mask;
/** マスクを最近傍で w×h に拡縮する */
export function resizeMask(mask: Mask, w: number, h: number): Mask;
/** 輪郭の画素を `[x0, y0, x1, y1, …]` で返す */
export function contourPoints(mask: Mask): number[];
/** マスクを Chebyshev 距離 r だけ膨らませる */
export function dilate(mask: Mask, r: number): Mask;

/** 格子の全画素について、点までのユークリッド距離を返す（点が無ければ Infinity） */
export function distanceField(w: number, h: number, pts: number[]): Float64Array;
/** 昇順の配列の p 分位（最近順位法） */
export function percentile(sorted: number[], p: number): number;
/** 2 つの点の集合の双方向の最近点距離（昇順） */
export function pointDistances(w: number, h: number, ptsA: number[], ptsB: number[]): number[];

/** 輪郭の比較の結果（距離・寸法の差は 100% 換算の px） */
export interface ContourCompare {
  p95: number;
  max: number;
  mean: number;
  nA: number;
  nB: number;
  dw: number;
  dh: number;
  /** 空の側（'interhub'＝XSheet の側。名前は InterHub から写したまま） */
  empty: null | 'both' | 'excel' | 'interhub';
}
/** 輪郭の比較（a＝Excel・b＝XSheet。どちらもインクの外接矩形で切り出したマスク） */
export function compareContours(a: Mask, b: Mask, scale?: number): ContourCompare;
/** 細線化（Zhang-Suen。枠線を 1 画素幅の中心線にする） */
export function thin(mask: Mask): Mask;
/** 中心線どうしの比較（a＝Excel・b＝XSheet のインクのマスク。normalize なら b を a の寸法へ拡縮、searchPx > 0 なら ±searchPx 画素の位置合わせをして測る） */
export function compareSkeletons(
  a: Mask,
  b: Mask,
  scale?: number,
  normalize?: boolean,
  searchPx?: number,
): ContourCompare & { shift: { x: number; y: number } };
/** 重ね合わせ画像（a のみ＝赤・b のみ＝青・両方＝黒） */
export function overlayImage(a: Mask, b: Mask): RgbaImage;

/** 文字のインクの重心・外接矩形（100% 換算の px。枠線の外接矩形の左上が原点） */
export interface TextStats {
  cx: number;
  cy: number;
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}
/** 反転した図形の文字の比較の結果 */
export interface TextCompare {
  orientation: 'upright' | 'rot180' | 'unknown';
  reason?: string;
  p95Upright?: number;
  p95Rot180?: number;
  excel?: TextStats | null;
  interhub?: TextStats | null;
  dCentroid?: { x: number; y: number };
  nA?: number;
  nB?: number;
}
/** 反転した図形の文字の比較（e＝Excel・i＝XSheet。文字入りの図形と枠線だけの双子のマスク） */
export function compareText(o: {
  eText: Mask;
  eOutline: Mask;
  iText: Mask;
  iOutline: Mask;
  scale?: number;
  sepPx?: number;
}): TextCompare;

/** ref-geoms.xlsx の図形名 → 比較の群 */
export function groupOfShape(
  name: string,
): 'preset' | 'adjusted' | 'xform' | 'custom' | 'default' | 'skip';
/** 比較の結果の集計 */
export function summarize(
  entries: { name: string; group: string; p95: number }[],
  tolerance?: number,
): {
  total: number;
  ok: number;
  ng: number;
  okRatio: number;
  byGroup: Record<string, { total: number; ok: number }>;
  ngNames: string[];
};
