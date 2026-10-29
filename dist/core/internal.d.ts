/**
 * core の半公開の入口（`xsheet/internal`）。render が使う core の部品のうち、公開の入口に無いものを出す。
 * 版の互換を約束しない（中身は render の都合で変わる）。src/render は core をこの入口と公開の入口からだけ引く。
 */
/** 図形の塗りの明暗（`a:path fill` の lighten・darken など）を色に掛ける（図形の SVG で使う） */
export { applyFillMode } from './color';
/** 既定の最大桁幅と、寸法の指定が無いシートの layout（描画の部品の試験で使う） */
export { DEFAULT_MDW, emptyLayout } from './types';
