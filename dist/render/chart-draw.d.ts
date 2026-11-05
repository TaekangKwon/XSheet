/**
 * グラフの配置と SVG（`chart-layout.ts`・`chart-axis.ts`・`chart-thin.ts`・`chart-svg.ts`）をまとめた入口。
 * `image-layer.ts` がグラフを初めて描く時だけ動的 import する（グラフの無いブックでは読まない）。
 * ライブラリのビルドでは別のチャンク、単独版では `tools/standalone/parts.mjs` の一覧に従う。
 */
export { layoutChart } from './chart-layout';
export { buildChartSvg } from './chart-svg';
