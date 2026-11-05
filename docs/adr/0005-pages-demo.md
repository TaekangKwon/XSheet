# 0005. GitHub Pages のデモ

> 版: 1.0 ・ 更新日: 2026-11-05 ・ 担当: kwon
> 関連文書: [0003](0003-universal-distribution.md)（例のページ・単独版） / [0004](0004-public-release.md)（公開。§4 の配り方は変えない）

## TL;DR
- 例のページ（`examples/index.html`）を、そのまま GitHub Pages のデモにする（https://taekangkwon.github.io/XSheet/）。入れずにブラウザで試せる。
- `npm run pages:build` が Vite で例のページをビルドし、見本のサンプルを写して `pages/`（Git の対象外）に置く。base は相対で、サイトの場所に依らない。
- `main` への push で GitHub Actions（`.github/workflows/pages.yml`）がビルドして出す。CI でもビルドだけ確かめる。
- ページは外へ通信しない。CSP は例のページのまま（`connect-src 'self'`。サンプルは同じ出所から読む）。選んだファイルはブラウザの中で読むだけで、どこへも送らない。

---

## ステータス

accepted（2026-11-05。ユーザーが公開後の候補を「v0.4.0 → Pages のデモ → npm の準備」の順に進めるよう指示した）。

## コンテキスト

- 公開した（ADR-0004）が、試すには git 依存で入れるか、Release の zip を落として `viewer.html` を開く必要がある。README の画面は静止画だけ。
- 例のページは、ファイルの選択・ドロップ・見本・テーマ・言語・図の有無・タブの位置を切り替えられ、CSP も厳しい形で付いている。デモにそのまま使える。
- 例のページはこれまで開発サーバー（`npm run examples`）でだけ動かしていた。サンプルの読み込みは根からの絶対の場所（`/fixtures/excel/…`）で、`https://<owner>.github.io/XSheet/` のような下の場所では動かない。

## 決定

### 1. 何を出すか

- 例のページ（`examples/index.html`）だけ。単独版の例（`standalone.html`・`esm.html`）はビルド済みの `dist/` を読む作りで、Release の zip の `viewer.html` と同じなので出さない。
- 見本のサンプルは `examples/main.ts` の `SAMPLES` に載せたもの（`fixtures/excel` の自作のファイル。合わせて約 1.9 MB）だけを写す。照合用のほかのファイルは出さない。
- 根の `index.html` は、見本（`showcase.xlsx`）を開いた例のページへ移るだけのページ（スクリプトなし・`default-src 'none'`）。

### 2. 作り方

- `tools/pages/build.mjs`（`npm run pages:build`）：Vite の `build` を `examples/vite.config.ts` に `base: './'` と出力先 `pages/` を足して呼ぶ。nonce（`html.cspNonce`）はビルドの `<script>`・`<link>` にも付く。
- 例のページのサンプルの読み込みを、ページからの相対（`../fixtures/excel/<名前>`）にした。開発サーバーの `/examples/` でも、Pages の `<サイト>/examples/` でも同じ場所を指す。
- 出力 `pages/` は Git に入れない（`/pages/`）。

### 3. 出し方

- `.github/workflows/pages.yml`：`main` への push と手動（`workflow_dispatch`）で、ビルド → `actions/upload-pages-artifact` → `actions/deploy-pages`。書き込みの権限（`pages: write`・`id-token: write`）は出す側のジョブだけに付ける。
- デモは `main` の今の姿を出す（版のタグではない）。版ごとの物は Release の zip とタグで配る（ADR-0004 §4 のまま）。
- CI（`ci.yml`）に `npm run pages:build` を足し、PR の時点でビルドが壊れていないことを確かめる。
- リポジトリの Pages の設定は「GitHub Actions から出す」にする（最初の 1 回だけ）。

## 却下した案

| 案 | 却下の理由 |
| --- | --- |
| `gh-pages` ブランチにビルドの結果をコミットする | 生成物が履歴に積もる。Actions から出せば要らない |
| 版のタグ（Release）のときだけ出す | 出すまでの間が空き、README の説明と画面がずれる。版ごとの物は zip がある |
| 単独版（`xsheet.min.js`）で別のデモのページを作る | 例のページと 2 つを保つことになる。例のページのほうが切り替えが多い |
| 照合用の `fixtures/excel` を全部出す | 大きい（グラフの照合用だけで約 1.7 MB）。見本の一覧に載せたものだけで足りる |

## 結果

- 良い点：入れずに試せる。README から 1 回のクリックで動く画面へ行ける。例のページが本番のビルドでも動くことを CI で確かめられる。
- 注意：例のページは `main` と一緒に公開される。例のページの文言・サンプルを変えるときは、外から見えることを前提にする。
- 注意：例のページは日本語の画面のまま（ライブラリの文言は切り替えられる）。
