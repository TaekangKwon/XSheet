// 単独版の部品とファイルの一覧（M6-3）。ビルド（build.mjs）・検査（check.mjs）・大きさ（tools/size.mjs）が共有する。

/**
 * 遅れて読む部品の一覧。部品を足すときはここに 1 行足す（ADR-0003 §1）。
 * `module` は render などが動的 import する入口のファイル。ESM では xsheet-<name>.mjs になり、IIFE では中に入る。
 */
export const LAZY_PARTS = [{ name: 'geom', module: 'src/geom/index.ts' }];

/** 単独版の出力ファイル（dist/ の根からの相対） */
export const STANDALONE_FILES = [
  'xsheet.min.js',
  'xsheet.mjs',
  ...LAZY_PARTS.map((p) => `xsheet-${p.name}.mjs`),
  'xsheet.css',
  'xsheet.d.ts',
  'xsheet.d.mts',
  'xsheet.licenses.txt',
];
