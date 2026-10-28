# 作業日誌

1 件ごとに 5 行以内を末尾に追記する（新しいものが下）。

## 2026-10-28 M0（雛形）
- 立ち上げ資料を InterHub `docs/guides/xlsx-lib-project/` から `docs/` へ写し、名前を XSheet（npm 名 `xsheet`）に直した。
- `CLAUDE.md`・ADR-0001（D1〜D8 と開発用の道具の版）を作った。
- `package.json`・tsconfig（全体と core）・vitest＋jsdom・ESLint＋Prettier を置き、jszip 3.10.2・ssf 0.11.2 を完全一致で固定した。

## 2026-10-28 M1-1（geom を写す）
- InterHub 943e4f0 から評価器・形の表・変換器・XML・NOTICE を写した（バイト一致を確認）。`npm run geom:build` を作った。
- 評価器が core・render から引いていた演算子の表・上限・経路の型を `src/geom/defs.ts` へ移し、geom を他の層に依存しない層にした。
- 再生成した JSON は 176 形とも元と同じ（`_source`・`_notice` のパスだけ違う）。テスト 60 件が通る（LINE_PRESETS の照合は M1-2 まで todo）。
