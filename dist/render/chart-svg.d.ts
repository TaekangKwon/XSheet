/**
 * `.xlsx` のグラフの場面（{@link ChartScene}。ISSUE-102）を `<svg>` 1 枚にする（ISSUE-103）。
 * - 場面の項目を順に要素へ写すだけで、計算はしない。`rect`・`path`・`line`・`text`（回転・揃え・太字）。
 * - 属性に入れるのは数値・`#rrggbb`・不透明度・固定の語・数値と固定の命令だけの `d`。ファイル由来の文字列（系列名・項目・タイトル）は
 *   `<text>` の `textContent` にだけ入れる（`innerHTML`・属性には入れない）。`font-family` は書かず画面の文字を継承する。
 * - 場面の値が不正（有限でない数・`#rrggbb` でない色・`d` に想定外の文字）の項目は描かない。
 * - ダーク表示（ADR-0031 §決定 10・P5）：グラフは Excel の色のまま描く。場面の `backdrop`（グラフの面が塗りなし）が真のときの白い下地は、この `<svg>` の中には置かない
 *   （中に置くと、グラフより下の図形・画像まで隠れる）。画像レイヤー（`xlsx-image-layer.ts`）が、図形と同じ兄弟の `svg.xs-backdrop`（`z-index: -1`）として図の直前に置く。
 * - DOM は `doc`（`Document`）経由で作る（テストで差し替えられる）。`chrome.*` に触らない。
 */
import type { ChartScene } from './chart-layout';
/**
 * グラフの場面を `<svg>` にする。大きさ（`width`・`height`・`viewBox`）は場面の px。クラス `xs-chart`（CSS が枠いっぱいに伸ばす）。
 * 下地（`scene.backdrop`）はこの `<svg>` には含めない（呼び出し側の画像レイヤーが図の層の最背面に置く）。
 *
 * @param scene `layoutChart` の結果（null は呼び出し側が破線枠にする）
 * @param doc 要素を作る文書
 * @param title 読み上げの名前（呼び出し側の i18n の文言。`role="img"` の `<title>` に `textContent` で入れる。省略すると `aria-hidden`）
 */
export declare function buildChartSvg(scene: ChartScene, doc: Document, title?: string): SVGSVGElement;
