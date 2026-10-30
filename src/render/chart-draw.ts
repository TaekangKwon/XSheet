// XSheet: グラフの描画（遅れて読む部品の入口。ADR-0003 §1）
/**
 * グラフの配置と SVG（`chart-layout.ts`・`chart-axis.ts`・`chart-thin.ts`・`chart-svg.ts`）をまとめた入口。
 * `image-layer.ts` がグラフを初めて描く時だけ動的 import する（グラフの無いブックでは読まない）。
 * ライブラリのビルドでは別のチャンク、単独版では `tools/standalone/parts.mjs` の一覧に従う。
 */
export { layoutChart } from './chart-layout';
export { buildChartSvg } from './chart-svg';
