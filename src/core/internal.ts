// 正: docs/02_architecture.md §2
/**
 * core の半公開の入口（`xsheet/internal`）。render が使う core の部品のうち、公開の入口に無いものを出す。
 * 版の互換を約束しない（中身は render の都合で変わる）。src/render は core をこの入口と公開の入口からだけ引く。
 */
/** 図形の塗りの明暗（`a:path fill` の lighten・darken など）を色に掛ける（図形の SVG で使う） */
export { applyFillMode } from './color';
/** 既定の最大桁幅と、寸法の指定が無いシートの layout（描画の部品の試験で使う） */
export { DEFAULT_MDW, emptyLayout } from './types';
/** 色の変換（lumMod など）を色に掛ける（グラフの既定の色で使う。ISSUE-103） */
export { applyTransforms, type Rgba } from './color';
/** 数値の表示形式でひとつの値を書式にする（グラフの軸・データラベルで使う。ISSUE-102） */
export { formatNumberByCode } from './format';
/** 既定のテーマ（Office 2013〜）とテーマの型（グラフの既定の系列色で使う。ISSUE-103） */
export { DEFAULT_THEME, type Theme, type ThemeSlot } from './theme';
