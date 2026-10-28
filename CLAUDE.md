# CLAUDE.md

Claude Code 向けのプロジェクト指示書（XSheet）。

## TL;DR
- XSheet は .xlsx を Excel どおりに表示する JavaScript ライブラリ（npm のパッケージ名は `xsheet`）。core（読み取り・DOM 非依存）／geom（図形の形）／render（DOM の表示）の 3 層。
- 元の実装は InterHub（`C:\GitHub\InterHub` コミット `943e4f0`）。**InterHub は読むだけで変更しない**。
- 文書・コメント・コミットメッセージは日本語。識別子は英語。ユーザーとの会話の言語は各自の個人設定に従う。

## 運用ルール
- 計画と決定は `docs/`（README・01〜05）と `docs/adr/` にある。設計の判断は ADR に残す。
- 1 指示＝1 Issue（`docs/04_roadmap.md` の M0-1 などの単位）。終わったら `docs/.history/progress.md` に 5 行以内を追記し、`docs/.history/handoff.md`（Git 管理外）を「更新日 / ブランチ / 現在のエラー / 次にやること」で上書きする。
- 大きなファイルは必要な範囲だけ読む。広い探索はサブエージェントに任せる。
- コミット・push・PR・タグは、ユーザーの指示があるときだけ行う。`main` で作業するときは先にブランチを切る。
- 破壊的な操作（削除・上書き・force push）は事前に確認する。

## 守ること
- 外部通信をしない。`eval`・`new Function` を使わない。
- core（`src/core/`・`src/geom/`）は DOM・Node 固有の API を使わない。
- 上限（`XlsxLimits`）を外さない。例外のメッセージにセル値・シート名を含めない。
- 依存を増やさない（今は jszip・ssf だけ）。増やすときは ADR に理由を書き、2 週間以上前の版を完全一致で固定する。開発用の依存も完全一致で固定する。
- 配布物に `NOTICE`（Ecma の著作権表示）を必ず含める。
- テストデータは自作の合成（`src/testing/make-xlsx.ts`）か、Excel で自作したファイルだけ。

## コマンド

| コマンド | 用途 |
| --- | --- |
| `npm run test` | vitest（1 回だけ実行） |
| `npm run typecheck` | tsc（core と全体の 2 つの設定） |
| `npm run lint` / `npm run format:check` | ESLint ／ Prettier の確認（`--write` で一括整形しない） |
| `npm run build` | Vite のライブラリモード＋d.ts（M4 で作る） |
| `npm run geom:build` | 図形の形の表を `tools/geom/presetShapeDefinitions.xml` から作り直す（M1-1 で作る） |

## 技術
- TypeScript strict・ESM・`verbatimModuleSyntax`。Vite（ライブラリモード）・vitest＋jsdom。
- TypeScript は 6.0 に固定（typescript-eslint 8.71 が `<6.1` までのため。理由は ADR-0001）。
- 対応環境：core は ES2022 のブラウザ・Worker・Node 20 以上。render は主要ブラウザの最新 2 版。開発には Node 22.12 以上が要る（vitest 5 の条件）。
- リポジトリ：https://github.com/TaekangKwon/XSheet（非公開）・既定ブランチ `main`。GitHub CLI などの場所は各自の個人設定に従う。
