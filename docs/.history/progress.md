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

## 2026-10-28 M2-3（renderSheet を切り出す）
- `spreadsheet.ts` の .xlsx の部分を、表の組み立て `grid-table.ts`・表示のインスタンス `sheet-view.ts`・公開の `renderSheet`（`render-sheet.ts`）に分けた。状態はインスタンスに持ち、window のリスナをやめて枠の ResizeObserver にした。
- geom は動的 import（`shapes.ts` の `loadGeom`）。図形のあるシートだけ形の表を読み、読めなければ図形を代替で描いて `libLoadFailed` を通知する（render から geom の値の import は ESLint で禁止）。
- 描画のテスト 5 本を `renderSheet` を呼ぶ形に書き直し、2 つの表示を同時に作って壊す後始末のテスト・geom の読み込みのテストを足した。タブの関係のテストは M2-4 で写す。
- M2-1 で残っていた `--xs-xlsx-*`（置き換えの順の誤り）を `--xs-*` に直した。テスト 1,586 件（63 ファイル）・typecheck・lint・format:check が通る（全体を並列で回すと perf の時間のテストがまれに落ちる）。

## 2026-10-28 M2-4（renderWorkbook）
- 公開の `renderWorkbook(container, data, { limits, … })` を作った。タブ列（`sheet-tabs.ts`）の後ろに 1 冊 1 つの表示を置き、タブの切り替えで書式の `<style>` を作り直さない。開けないファイルは種類ごとのバナー、シート 0 枚は `noSheets`。
- `destroy` はタブのリスナ・ResizeObserver・rAF → 表示（blob URL）→ `book.close()` の順に放す。タブのクリックは `addEventListener` にした。
- タブの関係のテスト（M2-3 で残したもの）を写し、切り替えて壊す後始末のテストを足した。perf のテストは別の project で 1 本ずつ最後に回す（全体の並列でまれに落ちたため）。
- テスト 1,610 件（65 ファイル）・typecheck・lint・format:check が通る。

## 2026-10-29 M3-1（style.css）
- `src/style.css` を作った（InterHub `global.css` の .xlsx の部分と共用の表・帯・タブの規則を `xs-` で）。変数 14 個の明るい既定値を `:where(:root, :host)` に置き、ダークは使う側が上書きする。
- 表示の根は `container` の高さいっぱいの縦の flex、`.xs-body` がスクロール枠。見出し・行番号は sticky（z 2・1・角 3）。
- 02 §6 の変数の表を実際の規則に合わせて直した（グリッド線は `--ih-border-soft`。`--xs-border`・`--xs-accent`・`--xs-warn*` を足した）。
- CSS の静的テスト 3 本を写し、クラス・変数・文書の突き合わせのテストを足した。テスト 1,632 件（69 ファイル）・typecheck・lint・format:check が通る。

## 2026-10-29 M3-2（examples）
- `examples/`（index.html・main.ts・example.css・vite.config.ts）を作った。`npm run examples` で http://localhost:5173/examples/ を開く。ファイルの選択・ドロップ・サンプル（fixtures/excel）・テーマ（OS・明るい・ダーク）・文言（ja/en）・図の有無を切り替えられる。
- ダークは `example.css` が変数を上書きする（`prefers-color-scheme` とテーマの属性の 2 通り）。値は 02 §6 の「ダークの例」の列と同じ。CSP の meta は `script-src 'self'`・`img-src blob:`（style-src は M3-3 で nonce にする）。
- `tests/render/examples.test.ts`（7 件）：ダークの値と文書・CSP・入口・サンプルの有無を突き合わせる。内蔵ブラウザで明るい・ダークの表示と CSP の違反が無いことを見た（実機の目視はユーザー）。テスト 1,639 件。

## 2026-10-29 M3-3（nonce・CSP）
- examples の CSP を `style-src 'nonce-xsheet-example'` だけにした。`<link>` に nonce を付け、main.ts はその nonce プロパティを renderWorkbook に渡す。
- 内蔵ブラウザで、nonce ありで書式の塗りが効き、nonce なしでは書式の `<style>` が止まる（sheet が null）ことを確かめた。画像（blob:）も出る。
- 02 §3 に「CSP」の節（3 つの指定と理由）を足し、01 N7 の基準を更新した。`tests/render/csp.test.ts`（style 属性・HTML の文字列・<style> の数・data: を render のソースで見る・nonce の受け渡し）。テスト 1,647 件。

## 2026-10-29 M4 の前の整理（照合のハーネス・線形時間のテスト）
- `fixtures/excel/compare-ref-geoms.mjs` を XSheet の `renderWorkbook` へ向け直した（esbuild → Vite の build()・クラス名は xs-・`npm run geom:compare`）。2 図形で撮影まで確かめた（Excel との距離の照合は Excel が要るので未実施）。
- `linear-time.test.ts` を `linear-time.perf.test.ts` にして perf の project（ほかの後に 1 ファイルずつ）へ移し、上限を超えたときだけ 3 回まで測り直す。テスト 1,647 件。

## 2026-10-29 M4-1（ビルド）
- `vite.config.ts`（ライブラリモード・ESM・入口 4 つ・jszip/ssf は束ねない・圧縮しない・style.css と NOTICE を写す）と `tsconfig.build.json`（d.ts だけ）。`npm run build`。
- `tools/consumer-check.mjs`（`npm run check:consumer`）：一時フォルダの使う側から exports を通して型検査と Vite のビルドが通る。形の表は使う側でも別のチャンク。
- ADR-0002 で D5（タグに dist を含めて git 依存）・D7 を決定として記録。02 §8 にビルドと配布物を書いた。

## 2026-10-29 M4-2（大きさ・ライセンス・NOTICE）
- `tools/size.mjs`（`npm run size`）：入口ごとの大きさ（dist・gzip・圧縮＋gzip）を測り、上限（実測＋約 1 割）を超えたら失敗。core 27,781・render 47,240・geom 32,519 B（圧縮＋gzip）。
- `check:consumer` に `npm pack --dry-run` の中身（exports の先・NOTICE・dist/NOTICE、src を含まない）の確認を足した。
- 02 §8 に大きさと依存のライセンス（15 パッケージ）の表。`tests/distribution.test.ts`（7 件）が lockfile・exports・vite.config・上限と突き合わせる。テスト 1,654 件。

## 2026-10-29 M4-3（README・CHANGELOG。タグは指示待ち）
- 利用者向けの `README.md`（入れ方・使い方 3 通り・指定・サブパス・CSS 変数とダークの例・CSP・上限・対応環境・開発）と `CHANGELOG.md`（0.1.0・未リリース）を書いた。`version` を 0.1.0 にした。
- 02 §8 に版を出す手順（build・check:consumer・size → release ブランチに dist を足す → 注釈付きのタグ）。タグ `v0.1.0` を打つのはユーザーの指示があるとき。
- テスト 1,654 件・typecheck・lint・format:check・build・check:consumer・size が通る。

## 2026-10-29 レビューの指摘の直し（M3・M4）
- check:consumer はリポジトリへの junction をやめ、npm pack で入るファイルの写し＋Vite の alias にした（一時フォルダを消してもリポジトリに届かない）。照合のハーネスのサーバーは 127.0.0.1 だけで待ち受け、site の外を配らない。
- examples：読み込みの依頼に番号を付け、遅れて終わった古い読み込み・エラーを捨てる。開発サーバーにも nonce（html.cspNonce）。size.mjs は副作用の import も辿り、辿り方の自己確認を足した。
- csp.test.ts：実際に描いて style 属性を付けないことを確かめる（2 件）・toDataURL などを禁止に追加。テスト 1,656 件。

## 2026-10-29 タグ v0.1.0
- PR #4 をマージ（main a6d9990）。main で test 1,656 件・typecheck・lint・format:check・build・check:consumer・size が通ることを確かめた。
- CHANGELOG に日付（45dff90）。release/v0.1.0 に `git add -f dist` だけのコミット（fe5a269）を作り、注釈付きのタグ v0.1.0 を打って push（タグだけ。release ブランチは push しない）。
- タグの中身を git archive で取り出し、npm pack で exports の先・NOTICE・dist/NOTICE が入り src が入らないことを確かめた。日付のコミットは chore/changelog-v0.1.0 から main へ入れる。

## 2026-10-30 ADR-0003（誰でも使える配り方・InterHub の取り込み。提案）
- 使いにくさ P1〜P8 と SheetJS・PDF.js の配り方を整理し、単独版（`xsheet.min.js`・`xsheet.mjs`）と 1 行の入口 `view()`（`src/facade/`）を提案した。バンドラ向けの出力は今のまま。
- InterHub の 943e4f0 以降（グラフ・22 コミット）を読むだけで確認し、写しの規約と同期の道具（`tools/sync/interhub.json`・`sync:status`・`sync:patch`）を提案した。グラフは M6 の後に取り込む。
- 04 に M6（誰でも使える配り方・v0.2.0）を足し、旧 M6（拡張）を M7 にした（01・02・05・docs/README の参照も）。
- ユーザーが Q1〜Q4 を提案のとおりに決め、ADR-0003 を accepted（1.0）にした（社内向けから・`xsheet` の入口を全部に・`view` の既定の言語は auto・`MESSAGES_KO` を同梱）。

## 2026-10-30 M6-1 同期の道具
- `tools/sync/interhub.json`（対応表・規則・手作業・除くもの・変換・取り込み済み 943e4f0）と `tools/sync/sync.mjs`（`npm run sync:status`・`sync:patch`）。git は読むだけのサブコマンドに限り、InterHub に書かない・作業ツリーを読まない。
- 実行の結果：943e4f0..b16f45d で対応表に触れたコミット 14 件・54 ファイル（対応表 8・規則 43・手作業 3・対応なし 0）。`git apply --check` は image-layer.ts・core/index.ts の 2 つだけ当たらない。
- 03 §6 に写しの規約と使い方。`tests/sync.test.ts`（31 件・一時の git の写しで InterHub に触れないことを確かめる）。

## 2026-10-30 M6-2〜M6-5 誰でも使える配り方（v0.2.0）
- 入口の層 `src/facade/`：`view`・`read`・`Viewer`、`lang`（auto）・`theme`（ダークの値を style.css の `data-xs-theme` へ）・`css` の自動挿入・`height` の既定・同じ container の差し替え。`MESSAGES_KO`。入口 `xsheet` を全部に、`xsheet/core` を足した。
- 単独版 `tools/standalone/build.mjs`：`xsheet.min.js`（gzip 110,818）・`xsheet.mjs`（80,520）＋`xsheet-geom.mjs`（30,588）・`xsheet.css`・`xsheet.d.ts`・`xsheet.d.mts`・`xsheet.licenses.txt`。jszip は lib から束ね、setimmediate を `new Function` の無い版に替えた。`npm run check:standalone`・上限を size に。
- README の「3 行で始める」・`docs/guide/`（6 本）・02 の更新、`examples/standalone.html`・`esm.html`（http で目視 OK。file:// はプレビューで確かめられず未確認）。
- `npm run release:zip`（release/xsheet-0.2.0-standalone.zip）・CHANGELOG 0.2.0（日付は未定）・版を 0.2.0 に。独立のレビューの指摘 7 件を直した。テスト 1,722 件・typecheck・lint・format・build・size・check:consumer・check:standalone が通る。
