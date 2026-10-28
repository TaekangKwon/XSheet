// 正: docs/02_architecture.md §1
/**
 * geom の入口（`xsheet/geom`）。図形の形の表と式の評価。他の層に依存しない。
 * 形の表（`preset-geom.json`・約 171 KB）を読むので、render からは動的 import する（M2-3）。
 * core は表を読まない `defs.ts` だけを引く。
 */
export { GUIDE_ARITY, MAX_GEOM_COMMANDS, MAX_GEOM_GUIDES, type PathSeg, type Pt } from './defs';
export {
  evaluateGeom,
  evaluateGuides,
  getPresetGeom,
  presetNames,
  type GeomArg,
  type GeomCmd,
  type GeomDef,
  type GeomFailure,
  type GeomFill,
  type GeomGuide,
  type GeomPath,
  type GeomPathDef,
  type GeomResult,
  type GeomSubpath,
  type GeomTextRect,
  type GuidesResult,
} from './evaluate';
