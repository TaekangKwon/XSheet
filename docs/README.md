# XSheet — .xlsx を Excel どおりに表示する JavaScript ライブラリ：立ち上げ資料

| 項目 | 値 |
| --- | --- |
| 種別 | 新しいプロジェクトの立ち上げ資料（このフォルダ一式を新リポジトリの `docs/` へ写して使う） |
| 元の実装 | InterHub（作者の非公開リポジトリ）main `943e4f0` の `.xlsx` ビューア。xrag-Viewer（作者の非公開リポジトリ）の自前リーダー（xrag-Viewer 側の ADR-0005）が起点 |
| 最初の利用者 | xrag-Viewer・InterHub |
| 更新日 | 2026-10-30 |
| 担当 | kwon |

## TL;DR
- InterHub で作った .xlsx の読み取りと描画（列幅・行高・罫線・セル書式・画像・図形・色）を、どのプロジェクトからでも `import` して使える 1 つのライブラリにする。
- SheetJS が「多形式の読み書き」の汎用ライブラリなのに対し、これは **.xlsx 読み取り専用・Excel と同じ見た目の表示** に絞る。100 万セルでも開け、核は DOM 非依存で Worker でも動く。
- ほぼ完成したコード（テストを除き約 1.25 万行・テスト約 2,000 件）を写して整えるのが中心。新規に書くのは公開 API の入口・インスタンス化・ビルドと配布。

---

## 1. この資料の読み方

| 順 | 文書 | 内容 | 読む人 |
| --- | --- | --- | --- |
| 1 | 本書 | 何を作るか・決めること・着手の手順 | 全員 |
| 2 | [01_requirements.md](01_requirements.md) | 目的・範囲（やる／やらない）・SheetJS との違い・非機能要件 | 全員 |
| 3 | [02_architecture.md](02_architecture.md) | 構成・公開 API・データモデル・CSS 変数・文言 | 実装する人 |
| 4 | [03_source-map.md](03_source-map.md) | InterHub のどのファイルをどこへ写し、何を直すか | 実装する人 |
| 5 | [04_roadmap.md](04_roadmap.md) | マイルストーン M0〜M7 と Issue の切り方・完了条件 | 計画する人 |
| 6 | [05_quality.md](05_quality.md) | テスト・Excel との照合・性能・既知の制限 | 実装・確認する人 |
| — | [guide/](guide/README.md) | 使い方の手引き：単独版・React と Vue・Worker・CSP・テーマ（M6） | 使う人 |
| — | [../CLAUDE.md](../CLAUDE.md) | Claude Code 向けのプロジェクト指示書（立ち上げ資料の雛形から作った） | 全員 |
| — | [adr/](adr/) | 設計の判断の記録（ADR-0001＝ライブラリ化の決定と D1〜D8・0002＝ビルドと配り方・0003＝誰でも使える配り方と InterHub の取り込み・0004＝公開・0005＝GitHub Pages のデモ・0006＝npm への公開の名前と入口） | 全員 |

## 2. 先に決めること

M0 で決める項目は [ADR-0001](adr/0001-library-decision.md)、D5・D7 は [ADR-0002](adr/0002-build-and-distribution.md) で決めた。表は立ち上げ時の推奨のまま残す。

| # | 決めること | 推奨 | 決める時期 |
| --- | --- | --- | --- |
| D1 | 名前 | **決定**：`XSheet`（npm のパッケージ名は `xsheet`。2026-10-28 に npm で未使用を確認） | M0 |
| D2 | リポジトリ | **決定**：GitHub `TaekangKwon/XSheet`。公開は [ADR-0004](adr/0004-public-release.md) | M0 |
| D3 | ライセンス | **決定**（[ADR-0004](adr/0004-public-release.md)）：MIT。preset 定義の Ecma の著作権表示（`NOTICE`）を同梱。依存の jszip は MIT/GPLv3 の二重ライセンス（MIT を選ぶ）、ssf は Apache-2.0 | M0（公開は M4 以降） |
| D4 | パッケージの形 | npm パッケージ 1 つ＋サブパス（`xsheet`・`xsheet/render`・`xsheet/style.css`）。monorepo にはしない | M0 |
| D5 | 配布 | **決定**（[ADR-0002](adr/0002-build-and-distribution.md)）：GitHub のタグに `dist/` を含めて git 依存で配る（使う側が `--ignore-scripts` でも動く）。公開後も続け、npm は別に決める（[ADR-0004](adr/0004-public-release.md)） | M4 |
| D6 | クラス名・CSS 変数の接頭辞 | `xs-`・`--xs-*`（XSheet に合わせる。旧仮称 xlsx-view 由来の `xv-` はやめた。2026-10-28 にユーザーが決定） | M2 |
| D7 | 既定の文言 | **決定**（[ADR-0002](adr/0002-build-and-distribution.md)）：日本語と英語を同梱し、使う側が差し替えられる | M3 |
| D8 | 対応環境 | 核：ES2022 のブラウザ・Web Worker・Node 20 以上。描画：Chrome/Edge/Firefox/Safari の最新 2 版 | M0 |

## 3. 着手の手順

1. リポジトリを作る（D2）。このフォルダ一式を新リポジトリの `docs/` へ写す。
2. 雛形（InterHub `docs/guides/xlsx-lib-project/CLAUDE.template.md`）を `CLAUDE.md` として置き、`<…>` を埋める。
3. [04_roadmap.md](04_roadmap.md) の M0 の Issue から始める。1 指示＝1 Issue で進める。
4. Claude のセッションへの最初の指示の例：

   > docs/README.md と docs/04_roadmap.md を読み、M0 の最初の Issue を始めて。InterHub（コミット 943e4f0）は読むだけで変更しない。

## 4. 関連する資料（InterHub 側）

元の設計資料（InterHub の ADR-0025〜0031・作業指示書 ISSUE-82〜111）は作者の非公開リポジトリにある。公開されていないので、写したファイルの対応と直した点は [03_source-map.md](03_source-map.md)、判断の要点は各 ADR に写してある。
