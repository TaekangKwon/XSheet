# 0002. ビルドの形と配り方（D5・D7）

> 版: 1.0 ・ 更新日: 2026-10-29 ・ 担当: kwon
> 関連文書: [0001](0001-library-decision.md)（D1〜D8） / [../02_architecture.md](../02_architecture.md) §8 / [../04_roadmap.md](../04_roadmap.md)（M4）

## TL;DR
- `npm run build` は Vite のライブラリモードで ESM だけを `dist/` に出す。入口ごとに別のファイル、形の表（geom）は動的 import の別のチャンク。圧縮しない。型は `tsc` が出す。
- 配り方（D5）は推奨のとおり：非公開の間は GitHub のタグに `dist/` を含め、使う側は git 依存で入れる。`main` には `dist/` を入れない。
- 既定の文言（D7）は推奨のとおり：ja・en を同梱し、`messages` で差し替えられる（M2-2 で実装済み）。

---

## ステータス

accepted（2026-10-29。D5・D7 は 0001 の推奨のまま。タグを打つのはユーザーの指示があるとき）

## コンテキスト

- 0001 で D5（配布）は M4、D7（既定の文言）は M3 で決めることにしていた。
- 使う側（InterHub・xrag-Viewer）はどちらも Vite と TypeScript の `moduleResolution: "bundler"` を使う。
- 形の表（`preset-geom.json`）は大きい（約 171 KB）ので、図形を描くときだけ読む（N6）。render は geom を動的 import している（M2-3）。
- 非公開のうちは npm に出さない（D3）。使う側は `npm install` の前後でビルドを走らせたくない（`--ignore-scripts` でも動いてほしい）。

## 決定

### 1. ビルドの形（M4-1）

| 項目 | 決定 | 理由 |
| --- | --- | --- |
| 形式 | ESM だけ（`type: module`）。CommonJS は出さない | 使う側はどちらも ESM のバンドラ。core は Node 20 以上でも ESM で読める |
| 入口 | `dist/core/index.js`・`dist/core/internal.js`・`dist/geom/index.js`・`dist/render/index.js`（`package.json` の `exports` と同じ形）。入口どうしが共有する部品は `dist/chunks/` | render が core を重ねて持たない。サブパスごとに使う側のバンドラが要る分だけ取れる |
| 形の表 | render からは動的 import のまま（`import("../geom/index.js")`）。使う側のビルドでも別のチャンクになる | N6。図形の無いシートでは読まない |
| 依存 | jszip・ssf は束ねない（`dependencies` のまま） | 使う側で重複を除ける・版を合わせられる |
| 圧縮 | しない（ソースマップも出さない） | 使う側のビルドが圧縮する。dist を読んで調べやすい |
| 型 | `tsc -p tsconfig.build.json`（`emitDeclarationOnly`）。`src/testing` は配らない | Vite は型を出さない。追加の依存（d.ts を束ねる道具）を入れない |
| CSS・著作権表示 | `src/style.css` と `NOTICE` を `dist/` に写す（vite.config.ts の小さなプラグイン） | `xsheet/style.css` の実体。配布物に NOTICE を必ず含める（D3） |

- 型の import は拡張子なし（`./types`）のため、使う側の TypeScript は `moduleResolution: "bundler"` を前提にする（`node16`・`nodenext` では型を辿れない）。必要になったら d.ts に `.js` を付ける後処理を足す。
- 別のプロジェクトから使えることは `npm run check:consumer` で確かめる（一時フォルダに使う側を作り、`npm pack` で入るファイルの写しで `exports` を通して型検査と Vite のビルドをする。形の表が別のチャンクであること・CSS が入ることも見る）。

### 2. 配り方（D5）

- 非公開の間は、GitHub のタグ（`v0.1.0` など）に `dist/` を含める。使う側は `"xsheet": "github:TaekangKwon/XSheet#v0.1.0"` のように git 依存で入れる（`prepare` などのスクリプトを持たないので `--ignore-scripts` でも動く）。
- `main` には `dist/` を入れない（`.gitignore` のまま）。タグは、`main` の該当コミットから切ったリリース用のコミット（`npm run build` の後に `git add -f dist` しただけのもの）に打つ。
- タグを打つ・push するのはユーザーの指示があるときだけ（CLAUDE.md）。npm への公開は D3（ライセンス）を決めた後。

### 3. 既定の文言（D7）

- ja（`MESSAGES_JA`・既定）と en（`MESSAGES_EN`）を同梱し、`messages` に一部のキーだけを渡して差し替えられる（M2-2・[02](../02_architecture.md) §7）。

## 結果

- 良い点：使う側は git 依存を足すだけで使える。サブパスごとに要る分だけ入る。形の表は図形のあるシートでだけ読まれる。
- 注意：タグのコミットにだけ `dist/` があるため、タグを打つ前に `npm run build`・`npm run check:consumer`・`npm run size` を通す（M4-3 の手順）。
- 注意：`moduleResolution: "bundler"` 以外の使う側は型を辿れない（上の 1 の注）。
