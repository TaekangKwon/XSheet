// render のテストの前準備（vitest の setupFiles）。
// 形の表（geom）は render から動的 import する（M2-3・`shapes.ts` の `loadGeom`）。部品のテストは形を評価できる前提なので、
// 各テストファイルの前に読み込んでおく（読み込まない状態は `render-sheet.geom.test.ts` が `vi.resetModules` で作る）。
import { loadGeom } from '../../src/render/shapes';

await loadGeom();
