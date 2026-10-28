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

## 2026-10-28 M1-2（core を写す）
- InterHub 943e4f0 から core 24・`src/testing/` 2・テスト 32・fixtures 23（ルートの `fixtures/excel/`。データは blob と一致）を写した。
- `drawing-geom.ts` の `GUIDE_ARITY`・`limits.ts` の `MAX_GEOM_*` を `src/geom/defs.ts` から引く形にし、LINE_PRESETS の照合のテストを戻した。
- jszip の型が `@types/node` を引き込み `tsconfig.core.json` で Node を止められないため、ESLint で core・geom の Node API を止めた（ADR-0001 追記）。
- テスト 810 件（36 ファイル）・typecheck・lint・format:check が通る。

## 2026-10-28 M1-3（公開の範囲）
- 公開の入口に `types.ts` の公開の型を全部と `Rgba` を出し、半公開の `src/core/internal.ts`（`xsheet/internal`：`applyFillMode`・`emptyLayout`・`DEFAULT_MDW`）を作った。
- `src/render/**` は core を index・internal からだけ引く（ESLint）。値・型の公開の範囲をテストで固定した。
- core の入口を Vite で束ね、素の Node（コード生成の禁止付き）で Excel のファイルを開いて元のソースと同じ結果になることを確かめるテストを足した。
- テスト 817 件（38 ファイル）・typecheck・lint・format:check が通る。

## 2026-10-28 M2-1（描画の部品を写す）
- D6 を `xs-` に決め（ユーザー）、ADR-0001・README・02〜04 の `xv-` を直した。
- InterHub 943e4f0 の `xlsx-*.ts` 8 本を `src/render/`（名前から `xlsx-` を外す）、テスト 17 本を `tests/render/` へ写し、`ih-` を `xs-` 系に置き換えた（0 件）。
- 文言の `t()` は仮置きの `src/render/i18n.ts`（キーを返す）から引く（M2-2 で消す）。perf の 1 本は M2-3、CSS の 3 本は M3-1 で写す。
- テスト 1,497 件（55 ファイル）・typecheck・lint・format:check が通る。

## 2026-10-28 M2-2（文言を messages に）
- `src/render/messages.ts`：`XSheetMessages`（23 キー。InterHub の 22 キー＋空のシートの `data.noData`）・凍結した `MESSAGES_JA`（既定）・`MESSAGES_EN`・`resolveMessages`・`formatMessage`。
- `image-layer.ts` の `t()` を、`createImageLayer({ messages })`・`frameLabel(…, messages)` に替え、仮置きの `i18n.ts` を消した（`t(` は 0 件）。
- テストの期待値をキーの文字列から日本語の既定値へ直し、文言の表・差し込み・上書きのテストを足した。型の名前は `XlsxViewMessages` をやめて `XSheetMessages` にした。
- テスト 1,510 件（55 ファイル）・typecheck・lint・format:check が通る。
