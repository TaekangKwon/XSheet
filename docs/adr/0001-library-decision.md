# 0001. InterHub の .xlsx ビューアを XSheet というライブラリにする

> 版: 1.0 ・ 更新日: 2026-10-28 ・ 担当: kwon
> 関連文書: [../README.md](../README.md)（D1〜D8） / [../01_requirements.md](../01_requirements.md) / [../04_roadmap.md](../04_roadmap.md)（M0） / 元（別リポジトリ・パスのみ）: InterHub `docs/plan/adr/0025`〜`0030`・`docs/guides/xlsx-lib-project/`

## TL;DR
- InterHub（main `943e4f0`）の .xlsx の読み取りと描画を写し、どのプロジェクトからも `import` できるライブラリ XSheet（npm 名 `xsheet`）にする。
- M0 で決める D1・D2・D3・D4・D8 を決めた。D5（M4）・D7（M3）は推奨のまま、その時期に決める。D6 は M2-1 の前に `xs-` に決めた。
- 開発用の道具の版を決めた。TypeScript は typescript-eslint の対応範囲に合わせて 6.0.3 にした。

---

## ステータス

accepted（2026-10-28 起票・名前はユーザーが決定）

## コンテキスト

- InterHub と xrag-Viewer が、それぞれ自前の .xlsx リーダーと表示を持っている。同じ直しを 2 か所で続けないため、1 つのライブラリにまとめる（[01_requirements.md](../01_requirements.md) §1）。
- 立ち上げ資料は InterHub の `docs/guides/xlsx-lib-project/` で、仮称は `xlsx-view` だった。本リポジトリの `docs/` はその写しで、名前だけを XSheet に直した。

## 決定

### 1. D1〜D8

| # | 項目 | 決定 | 状態 |
| --- | --- | --- | --- |
| D1 | 名前 | `XSheet`。npm のパッケージ名は `xsheet`（2026-10-28 に npm で未使用を確認） | 決定 |
| D2 | リポジトリ | `C:\GitHub\XSheet`・GitHub `TaekangKwon/XSheet`（非公開）・既定ブランチ `main` | 決定 |
| D3 | ライセンス | 非公開の間は書かない（`package.json` は `private: true`・`license: UNLICENSED`）。公開を決めるときに MIT か Apache-2.0 を選ぶ。Ecma の著作権表示（`NOTICE`）は M1-1 で置き、配布物に必ず含める | 決定（公開時に再決定） |
| D4 | パッケージの形 | npm パッケージ 1 つ＋サブパス（`xsheet`・`xsheet/render`・`xsheet/style.css`、render が動的 import する `xsheet/geom`、半公開の `xsheet/internal`（M1-3 で追加。版の互換を約束しない））。monorepo にしない。`exports` は M0-1 で枠だけ置き、中身は M1〜M4 で埋める | 決定 |
| D5 | 配布 | 推奨のまま（非公開の間はタグに `dist/` を含めて git 依存）。[0002](0002-build-and-distribution.md) | 決定（M4） |
| D6 | 接頭辞 | `xs-`・`--xs-*`・`data-xs-*`。推奨の `xv-` は旧仮称 xlsx-view 由来なので、名前に合わせた（2026-10-28・ユーザーが決定。M2-1 で `ih-` を置き換えた） | 決定 |
| D7 | 既定の文言 | 推奨のまま（ja・en を同梱し差し替え可）。[0002](0002-build-and-distribution.md) | 決定（M3） |
| D8 | 対応環境 | core：ES2022 のブラウザ・Web Worker・Node 20 以上。render：Chrome/Edge/Firefox/Safari の最新 2 版 | 決定 |

### 2. 開発用の道具の版

2 週間以上前（2026-10-14 以前）に出た版を、完全一致で固定する。

| 依存 | 版 | 理由 |
| --- | --- | --- |
| jszip | 3.10.2 | InterHub と同じ（M0-3） |
| ssf | 0.11.2 | InterHub・xrag-Viewer と同じ（M0-3） |
| typescript | 6.0.3 | typescript-eslint 8.71.2 の対応が `>=4.8.4 <6.1.0`。7.0（xrag-Viewer の版）は ESLint の型の解析が使えない |
| vitest | 5.0.4 | xrag-Viewer と同じ系列。Node 22.12 以上が要る |
| jsdom | 30.1.2 | xrag-Viewer と同じ |
| vite | 8.3.4 | vitest の peer。M4 のライブラリモードでも使う |
| eslint・@eslint/js・typescript-eslint | 10.12.0・10.0.1・8.71.2 | 平らな設定（flat config）。InterHub の設定に倣う |
| prettier | 3.9.10 | InterHub の設定（`printWidth: 100`・`singleQuote`・`trailingComma: all`）に倣う |
| @types/node | 24.19.2 | テストが `node:fs` で fixtures を読むため。core の型には入れない |

- core の DOM 非依存は、`tsconfig.core.json`（`lib` を ES2022 だけにし `types` を空にする）で型検査して確かめる（N8）。
- 追記（M1-2）：jszip の型（`index.d.ts`）が `/// <reference types="node" />` で `@types/node` を引き込むため、`tsconfig.core.json` では `process`・`Buffer` を止められない（DOM は止まる）。Node 固有の API は ESLint（`src/core/**`・`src/geom/**` に `no-restricted-globals` と、`node:*`・組み込みモジュールへの `no-restricted-imports`）で止める。`TextDecoder` はブラウザ・Worker・Node に共通の API なので使ってよい（型は `@types/node` から来る）。

## 却下した案

| 案 | 却下の理由 |
| --- | --- |
| TypeScript 7.0 にそろえる | typescript-eslint が対応していない。ESLint を外すか、型の解析なしで動かすことになる |
| 名前を仮称の `xlsx-view` のまま | ユーザーが XSheet に決めた |

## 結果

- 立ち上げ資料の `xlsx-view` は `XSheet`（文中）・`xsheet`（パッケージ名・サブパス）に置き換えた。接頭辞は D6 のとおり M2-1 で `xs-` に決め、文書の `xv-` も `xs-` に直した。
- TypeScript を 7 系へ上げるのは、typescript-eslint が対応してから別の ADR で決める。
