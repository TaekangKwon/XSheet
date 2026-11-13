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

## 2026-10-30 タグ v0.2.0・GitHub Release
- PR #7 をマージ（main b9ae956）。main で test 1,722 件・typecheck・lint・format:check・build・check:consumer・check:standalone・size が通ることを確かめた。
- CHANGELOG に日付（8752d96）。release/v0.2.0 に `git add -f dist` だけのコミット（5e12803）を作り、注釈付きのタグ v0.2.0 を打って push（タグだけ）。git archive でタグの中に exports の先・NOTICE・単独版があることを確かめた。
- GitHub Release「XSheet 0.2.0」に xsheet-0.2.0-standalone.zip（275,196 B）を添付。zip の viewer.html を file:// で開く目視はユーザーに任せた。

## 2026-10-30 M7 グラフの取り込み（InterHub f035f68）
- `sync:patch`（943e4f0..f035f68・ISSUE-101〜111）を当て、`image-layer.ts`・`core/index.ts` の当たらない hunk と手作業の 3 ファイル（`readChart` の配線・`.xs-chart`・描画のテスト）を写した。公開の型に `Chart*`、`core/internal` にグラフが使う部品を足した。`baseCommit` を f035f68 に。
- グラフの描画は遅れて読む部品（`render/chart-draw.ts`・B では別のチャンク・単独版は `xsheet-chart.mjs`）。`size` に遅れて読むチャンクの上限、`check:standalone` にグラフの描画を足した。上限は core 36,000・render 59,000・chart-draw 17,500・min.js 144,000・mjs 94,000・chart.mjs 28,000。
- examples のサンプルに `ref-charts.xlsx`・`ref-charts2.xlsx` を足し、開発サーバーで明るい表示・ダークのグラフを目視（プレビュー）。zip の `file://` の目視はユーザーが OK（2026-10-30）。
- テスト 2,161 件・typecheck・lint・format:check・build・size・check:consumer・check:standalone が通る。

## 2026-10-30 タグ v0.3.0・GitHub Release
- PR #9（M7 のグラフ）を main に（5b42d04）。CHANGELOG の [Unreleased] を 0.3.0 に、package.json・README・guide・04 の版を合わせた（bd81712）。
- release/v0.3.0 に `git add -f dist` だけのコミット（3844df6）を作り、注釈付きのタグ v0.3.0 を打って push（タグだけ）。git archive でタグの中に exports の先 11 個・NOTICE・単独版（xsheet-chart.mjs を含む）があることを確かめた。
- GitHub Release「XSheet 0.3.0」に xsheet-0.3.0-standalone.zip（328,477 B・xsheet-chart.mjs を含む 10 ファイル）を添付。タグからの入れ方と zip の viewer.html の目視はユーザーに任せた。

## 2026-10-30 公開の準備（ADR-0004）
- ADR-0004：MIT・履歴を残す・英語の README＋ja／ko・配り方は変えず npm は後で。公開の切り替えは、InterHub 由来のコードの権利の確認とユーザーの指示が条件。
- `LICENSE`（MIT）、`package.json`（`private` を外す・`license`・`repository` など）、`files` と単独版の zip に `LICENSE`。README を英語にし、`README.ja.md`・`README.ko.md`、`SECURITY.md`、CI（`.github/workflows/ci.yml`）を足した。
- 作者の PC のパス（文書・`tools/sync/interhub.json` の `defaultDir` → `../InterHub`・`compare-ref-geoms.mjs`）と、非公開の InterHub の中のパスに頼る手順・無い資料の参照を外した。ESLint の対象から `release/` を外した。
- テスト 2,161 件・typecheck・lint・format:check・build・size・check:consumer・check:standalone が通る。

## 2026-10-30 公開前の見直しと README の訳の手直し
- 公開前の見直し：追跡中のファイルに秘密の値・作者の PC のパスは無い。残るのはコミットのメールアドレス（41 件）、fixtures の .xlsx の作成者名、非公開の InterHub・xrag-Viewer の名前と中のパス（03・ADR-0003 などの記録）。PR #11 の前に外した `.interhub/` を含むコミット 1dc9467 は GitHub 上にまだ残るが、中身はこのリポジトリのファイル一覧だけ。
- `README.ko.md`・`README.ja.md` を、InterHub-MCP の訳に合わせて丁寧な文体（합니다체・です／ます）の自然な文に書き直した（見出し・表・コード・リンクの数は英語版と同じ）。「なぜ作ったか」の 1 段落目はユーザーが書いた韓国語の内容（既存のビューアは大きな文書で遅く、文字以外を描けない）に 3 つともそろえた。
- docs/README の「ADR-0005」が xrag-Viewer 側の ADR だと分かるように直した。

## 2026-11-04 README に会社の紹介を追加
- `README.md`（About us）・`README.ja.md`（開発元）・`README.ko.md`（만든 곳）の「謝辞」と「ライセンス」の間に、株式会社システム・エムズの紹介を 1 段落ずつ入れた。中身は会社のホームページとブログから（所在地・設立年・事業・モットー「create rich time」・ブログへのリンク）。従業員数などすぐ古くなる数字は入れていない。
- `LICENSE`・`package.json` の作者が個人（TaekangKwon）なのに合わせ、紹介文の 1 文目は「システム・エムズの Kwon がメンテナンスしている」とした（会社が作った、とは書かない）。

## 2026-11-04 README の会社紹介の見た目を整える
- 3 つの README の会社紹介を、1 段落から「2 行の紹介（1 行目の末尾は `\` の改行）＋アイコン付きのリンク（🌐 ホームページ・✍️ ブログ）」の形にした。中身（Kwon がメンテナンス・所在地・設立年・事業・モットー）は前のまま。採用ページへのリンクはユーザーの指示で入れていない。

## 2026-11-04 README に実際の画面を入れる
- Excel の COM で見本のブック `fixtures/excel/showcase.xlsx` を作る `make-showcase.ps1` を足した（結合したタイトル・塗りと罫線の表・数値の書式・赤の負の値・縦棒グラフ・吹き出しの図形・2 枚目のシート）。XSheet がまだ描かないセルの配置は使わない（日本語版の Excel は COM で書式の `[Red]` を受け付けないため、負の値は文字の色で赤にした）。
- 例のページに見本を足し、`?sample=<ファイル名>` で最初から開けるようにした。Edge のヘッドレスで 2 倍の解像度で撮り、上の操作の帯を切り落として `docs/images/screenshot.png` にした。
- 3 つの README の言語の切り替えの下に画像と説明の 1 行を入れた。手順は `docs/05_quality.md` の表に書いた。

## 2026-11-04 リポジトリを公開する（ADR-0004 §6）
- ユーザーが InterHub 由来のコードの権利の確認を済ませ、公開を指示した。直前に `git grep` でパス・メールアドレス・秘密の値が無いことを確かめ（ADR-0004 の本文の例だけ）、main の CI が通っていることを確かめた。
- `gh repo edit --visibility public` で公開した。続けて、非公開の脆弱性報告を有効にし、説明（英語）と topics 8 個を付けた。
- main を保護した：PR を通す（承認 0 件）・CI の `check` が必須・管理者にも適用・force push と削除は不可。

## 2026-11-05 シートタブの位置を選ぶ（#17）
- `renderWorkbook`・`view` に `tabPosition`（`'top'`・`'bottom'`。既定 `'top'`）を足した。`'bottom'` ではタブ列を DOM でもシートの表示の後ろへ動かし、根に `xs-tabs-bottom` を付けて余白を上へ回す（style.css）。写したファイル `render-workbook.ts` の変更なので `XSheet:` の印を付けた。
- テスト（DOM の順・既定値・知らない値・入口の層からの受け渡し・CSS）、README 3 言語・02 §3・CHANGELOG・例のページ（タブの上・下の切り替え）を足した。例のページで上・下・ダークを目で確かめた。
## 2026-11-05 v0.4.0・GitHub Pages のデモ（ADR-0005）
- PR #19（版を 0.4.0・CHANGELOG の日付）をマージ（3efbc15）。main で build・check:consumer・check:standalone・size が通ることを確かめ、release/v0.4.0 に dist だけのコミット（3026b59）を作って注釈付きのタグ v0.4.0 を push（タグだけ）。タグの中に exports の先 11 個・NOTICE・LICENSE・単独版があることを確かめた。GitHub Release「XSheet 0.4.0」に xsheet-0.4.0-standalone.zip（329,629 B・11 ファイル）を添付。
- 例のページを Pages のデモにした：`npm run pages:build`（相対の base・見本のサンプル 10 個を写す）・`pages.yml`（main への push で出す）・CI に pages:build。サンプルの読み込みをページからの相対にした。`vite preview --base /XSheet/` で表・グラフ・図形と CSP の違反が無いことを確かめた。
- PR #20 をマージ（e616adc）。リポジトリの Pages の設定を「GitHub Actions から出す」にし（`gh api -X POST …/pages -f build_type=workflow`）、最初の配備が通った。https://taekangkwon.github.io/XSheet/ から見本の例のページへ移り、表・グラフ・図形（path 335 個）が出てコンソールのエラーが無いことを確かめた。

## 2026-11-05 README の zip の手順・ADR-0006（npm の名前と入口）
- README（英・日・韓）のインストールの節に、Releases からスタンドアロン版の zip を落として使う手順（最新のリリース → Assets → 展開して viewer.html／xsheet.min.js → 再配布ではライセンスのファイルも）を足し、クイックスタートと表からリンクした（PR #21）。
- npm の名前を確かめた：`xsheet` は無いが `x-sheet`（0.0.0）があり、記号だけが違う名前は断られるおそれが高い。ユーザーが推奨案で決め、ADR-0006 にした（`@taekangkwon/xsheet`・`xsheet/internal` を exports から外す・Release の公開で Actions から信頼された公開と provenance）。付け替えと公開は §4 の別の作業（PR #22）。

## 2026-11-12 表とシートタブを Excel に近い見た目に（Issue #23）
- Moobius の依頼の検討（docs/research・Git 管理外）の中でユーザーが見た目の変更を指示。Issue #23 を立て、ブランチ feat/excel-like-look で CSS だけ直した。
- 明るい既定値 3 つ（`--xs-cell-bg` 白・`--xs-grid`・`--xs-border`）を既定値のブロックで直し（`XSheet:` の印）、見出し・タブ列の上書きは `style.css` 末尾の `XSheet:` ブロック 1 つにまとめた。ダークの値・DOM・公開 API は変えない。
- 02 §3・§6、CHANGELOG（Unreleased「変更」）、`format-css-static.test.ts` の既定値を直した。examples で明るい・ダーク、タブの上・下、◀ ▶ が出る幅を目視した。

## 2026-11-12 ADR-0007（Moobius の依頼 R1〜R7 の公開 API）
- PR #24（Issue #23 の見た目）はユーザーがマージした（86a69a9）。
- Moobius の依頼（検索・ハイライト・移動・判定・拡大縮小・曜日と和暦・見出しを隠す）の検討（docs/research・Git 管理外）を ADR-0007 にした。ユーザーの決定：拡大縮小は寸法に倍率を掛ける案・見出しには掛けない・組み込みの書式 27〜36・50〜58 の和暦は読み取りの指定 `builtinDates` で有効にする。
- 検索は非同期・中断・件数の上限の形にした（依頼の案と違うので Moobius の同意待ち）。版は曜日を 0.4.1、残りを 0.5.0、npm は 0.6.0 以降。docs/README・04 の M7 の表を更新。
- 検索の形（§1）・受け入れから外す項目（§8）・版（§9）もユーザーが推奨どおりに決め、ADR-0007 を全部決定にした。作業を GitHub の Issue #25〜#34（M7-1〜M7-10）に分け、マイルストーン v0.4.1・v0.5.0 を作った。04 に「M7 Moobius の依頼」の表を足した。

## 2026-11-12 M7-1 曜日の書式（#25・ブランチ feat/weekday-format）
- `src/core/format-ja.ts`（書式の字句の分解・通し番号からの曜日・`withJaTokens`）を足し、`format.ts` は ssf を包む `// XSheet:` の口 1 つ。`aaa` だけの書式は ssf では日付にならないので `is_date` も包んだ。
- 置き換えで日付の字句が消えると ssf が `[$-411]` を「$」と出すため、表示の無い言語の指定は外す。日付の範囲の外は、使う区画に曜日があればセルを空にする。
- 正解は Excel 16（ja-JP）の COM で作った（`make-ref-weekday.ps1`・`ref-weekday(-1904).xlsx`・88 書式・1026 セル）。大文字の `AAA` も曜日、a 5 個以上は長い形、ja-JP の COM は `\` を円記号として読む（逃がしは `!`）。照合で除くのは ssf の受け持ち 171 セル（埋め・負の値の「-」・小文字の am/pm・1904 の上限の先）。
- 参考に見た SekisanHistoryTest の docs/_archive の Excel に曜日の書式は無かった（和暦 `[$-411]gge"年"m"月"d"日"` は .xls にあり、M7-8 で使える）。
- 02 §4・05 §2・CHANGELOG を更新。examples に「曜日の書式」を足し目視した。

## 2026-11-12 M7-3 v0.4.1 を出した（#27）
- PR #37（版 0.4.1・CHANGELOG・README 3 つ・guide）をマージし、main からビルドした `dist` だけのコミット（d80f306・`release/v0.4.1`）に注釈付きのタグ `v0.4.1` を打って push した（ブランチは push しない）。
- GitHub Release「XSheet 0.4.1」に `xsheet-0.4.1-standalone.zip`（11 ファイル・333,439 B）を添付。zip の `xsheet.mjs` と `viewer.html` で曜日の表示を確かめた（`viewer.html` は開発サーバー経由。`file://` は内蔵ブラウザで動かせなかった）。
- PR のマージは自動承認に一度止められ、ユーザーの承認の後にマージした。Moobius への知らせの下書きをユーザーに渡した。

## 2026-11-12 M7-2 0.5.0 の型の先出し（#26・ブランチ docs/v0.5.0-api-types）
- ADR-0007 §1〜§7 の型を `docs/guide/xsheet-0.5.0-preview.d.ts` にまとめた。単独版 0.4.1 の `xsheet.d.ts` の中の名前（`xsheet/_/…`）に重ねる形で、並べて置くと大域の `XSheet` と `'xsheet'` に 0.5.0 の追加が出る。
- 説明は `docs/guide/xsheet-0.5.0-preview.md`（置き方・依頼の案との違い・使い方の例・型で決めた既定・未決）。ADR に無かった `WorkbookOptions.builtinDates` を足し、ADR-0007 §6 にも書いた。
- 0.4.1 の `xsheet.d.ts` と並べて tsc（strict・skipLibCheck なし）が通ることと、使う側の例（誤りは `@ts-expect-error`）・文書の例が通ることを確かめた。先出しを外すと 18 件の型エラーになり、重なっていることも確かめた。

## 2026-11-12 M7-4 検索 findText（#28・ブランチ feat/find-text）
- `src/core/find-text.ts` を足し、`xsheet/core`・`xsheet`・単独版の `XSheet.findText` に出した（公開の値は 12 個）。照合は `CellView.text`。大文字と小文字は 1 文字ずつ小文字にして比べ（`İ` で長さが変わる・語末の `ς` は `σ` にそろえる）、改行を除くときと長さが変わるときだけ元の位置への対応表を作る。
- `getRows` は範囲の上で切れた結合の左上を範囲の最初の行に置くので、2 回目からは 1 行上から読んでその行は数えない（区切りをまたぐ結合を 2 回数えない）。
- 実機（Chrome・60 万セル）で、1 回の連続は約 9ms・全体で約 1.4 秒（一致なし）。初めは「10ms を超えたら休む」で 20〜25ms 握っていたので、次の区切りで超えそうなら先に休み、1 区切りが 10ms を超えたら行数を半分にするようにした。core は DOM の型を持たないので `setTimeout`・`DOMException` は大域から引き、`signal` は `{ aborted }` の形で受ける。
- テスト 17 件（ルール・位置の対応・結合・非表示・上限・中断・遅い区切り・5000×30）。02 §2・README 3 つ・CHANGELOG を更新。

## 2026-11-12 M7-5 ハイライト setHighlights（#29・ブランチ feat/highlights）
- `src/render/highlight.ts`（一致の表・セルの文字を分けて `span.xs-hits` 1 つで包む）を足し、写したファイルは `// XSheet:` の口だけ：`grid-table.ts`（引数 `hits` と呼び出し 1 か所）・`sheet-view.ts`（状態・`setHighlights`・シートを替えたら消す）。`SheetView`・`WorkbookView` に `setHighlights`。
- `###` のセルは入れ物ごと光らせ、塗りをまたぐ自動の色のはみ出しは片に `is-ink`（`--xs-fg`）。CSS 変数 `--xs-hit-bg`・`--xs-hit-current-bg` を既定値・ダーク 2 つ・examples・02 §6 に足した（変数は 16 個）。ダークの値（`rgba(255, 255, 0, 0.3)`・`#b45309`）は ADR に無く、ここで決めた。
- テスト 10 件（DOM の形・空白・重なり・### ・描き直しで残る・寸法が変わらない・はみ出し・結合・is-ink・WorkbookView）。実機で、色・スクロールで描き直した後も残ること・包んでも文字の幅が同じ（空白が消えない）ことを確かめた。ダーク表示の見た目はスクリーンショットに出ず、確かめていない。
- 全体のテストの負荷で findText の 5000×30 のテストが 5 秒で切れたので、そのブロックだけ 60 秒にした。

## 2026-11-12 M7-6 移動・位置の判定（#30・ブランチ feat/navigate）
- `src/render/navigate.ts` を足した：スクロール先（縦は上から 1/3・横は見えないときだけ最小の移動、左へは範囲が収まればセルの左端から）と、画面の位置 → セル（寸法の `rowAt`・`colAt`。結合・はみ出しは描いた行から引く）。写したファイルは `sheet-view.ts` に口だけ（描いた行を覚える・`createNavigation` に今の状態を渡す）。`highlight.ts` は今のヒットを替えられる形にし、描いた表の印だけを替える `markCurrent` を足した。
- はみ出す文字の中のヒットは、同じ書式のクラスを付けた画面の外の表（`.xs-measure`）で文字の幅を測って横の位置を出す。測れない（jsdom）・文字列でないときはセルの範囲。結合の高さは対象の行から 100 行を読んで出す。
- 実機（Chrome・2000 行 × 80 列）：遠い行のヒットが表示枠の上から 1/3、45 列のはみ出しの中のヒットが枠の右端にちょうど来る。1 回 25〜45ms（描き直しを含む）。`cellAt` が結合の左上・はみ出し元を返す。隠れた画面では rAF が止まるので、`cellAt` は先に今の位置を描くようにした。
- テスト 19 件。CHANGELOG・02 §3 を更新。単独版 `xsheet.mjs` は gzip 90,430 B（上限 94,000）で、M7-7 で上限に近づく。

## 2026-11-12 M7-7 拡大縮小 setZoom（#31・ブランチ feat/zoom）
- `src/render/zoom.ts` を足した：寸法を倍率で包む（境目を整数に丸める・`rowAt`・`colAt` は丸めた境目で合わせる）、指定の確かめ、fit、左上のセルを残す計算。`sheet-view.ts` には口だけ（寸法の包み・枠を埋める範囲・MDW・図の層の `scale()` と判定の範囲・fit の合わせ直し・`setZoom`）。文字は CSS 変数 `--xs-zoom`（JS が表示の根に入れる。テーマの 16 個には数えない）。見出しの文字は 13px のまま、行番号は縮小のときだけ小さくする。
- 実機（Chrome）：50%・75%・100%・150%・200% で罫線の層の線と td の境目が一致、図の位置と大きさがちょうど倍率どおり、倍率を替えても左上のセル（1001 行・11 列）が残る。3000 行 × 80 列で倍率の切り替えは 200〜400ms、50% の描き直しは td が 2.4 倍で約 240ms（100% で約 115ms。隠れた画面での値）。縮小のときは先にスクロールして 1 回で描く。
- `render/index` が上限 59,000 を 489 B 超えたので、理由（ADR-0007 の 3 機能）を 02 §8 に書いて上限を 65,000 にした。`xsheet.mjs` は 91,601 / 94,000。
- テスト 21 件（`zoom.test.ts`）。`style-css.test.ts` は `--xs-zoom` をテーマの変数から外して見るようにした。CHANGELOG・02 §3・§6・§8 を更新。

## 2026-11-12 M7-9 見出しを隠す headers・SheetView.showSheet（#33・ブランチ feat/headers）
- `RenderOptions.headers: false` で、列見出しの行を作らず行番号の列を幅 0 で畳む（`grid-table.ts` の口：引数 `headers`・列の畳み・`thead` を外す）。`sheet-view.ts` は見出しの寸法を定数からインスタンスの値（`headX`・`headY`）にし、表・罫線の層・図の層・判定・fit が同じ原点を使う。`renderWorkbook` も渡す。
- `renderSheet` の戻り値に `showSheet`（表示中のシートなら何もしない・ブックに無い番号は RangeError）。倍率は保ち、ハイライトと今のヒットは消す（中の表示が前から持つ動き）。
- 実機（Chrome・showcase.xlsx）：A1 が表示枠の左上 (0, 0)、行番号の列の幅 0、罫線の層の線と td の境目が一致、`cellAt` が A1 を返す。
- テスト 7 件。CHANGELOG・02 §3 を更新。README 3 つは M7-10 でまとめて直す。

## 2026-11-13 PR をまとめて出し、M7-10 の文書と例（#34・ブランチ docs/v0.5.0-readme）
- M7-5〜M7-9 を push して PR #40〜#43 を出し、CI の成功を見て #38〜#43 を順に main へマージした（main は 468e413）。
- README 3 つに「Search, highlights and zoom (v0.5.0)」の節と `zoom`・`headers` の行を足し、CSS 変数の数を 14 → 16 に直した。
- 例のページに検索（◀ ▶・クリックで今のヒット）・倍率・見出しの有無・クリックしたセルの番地を足した。`WorkbookView` に今のシートを返すメンバーが無いので、例はタブの `is-active` から引く。実機（showcase.xlsx）で全部の操作を確かめた。
- 先出しの型と実装の差は `signal` の型（広い・互換）と `builtinDates`（M7-8 で足す）だけ。guide の §6 に書いた。版の上げ・タグ・Release はまだ。

## 2026-11-13 今のシート activeSheet・onSheetChange（#34・ADR-0007 §7 版 1.1）
- 例のページで、タブを押された後に検索するシートを外から知る方法が無いと分かり、ユーザーの決定で ADR-0007 §7 に足した（DOM のイベント・タブの `is-active` を公開の印にする案は却下）。
- `SheetView.activeSheet`（選んでいるシート。押した時点で替わる・開けなかったブックは null）と `WorkbookOptions.onSheetChange`（替えたシートを描き終えた後。最初の表示・同じシート・追い越し・destroy の後は呼ばない。例外は setTimeout で投げ直す）。写したファイルは触っていない。
- テスト 9 件（`active-sheet.test.ts`。全体 2271）。例はタブの DOM を読むのをやめ、`activeSheet`・`onSheetChange` を使う。実機でタブを替えると結果が消え、次の検索が新しいシートを探すことを確かめた。
- README 3 つ・02 §3・CHANGELOG・先出しの型（0.4.1 の `xsheet.d.ts` と並べて tsc が通る）も直した。`xsheet.mjs` は gzip 91,805 / 94,000。

## 2026-11-13 M7-8 和暦の書式・builtinDates（#32・ブランチ feat/japanese-era）
- ユーザーの依頼のファイル（SekisanHistoryTest の compare_features.xlsx・03_表示形式）で `[$-ja-JP]ggge"年"m"月"d"日"` が `ggg2026年11月11日` と出るのを確かめ、ユーザーが M7-8 を全部 0.5.0 に入れると決めた。
- 正解は `make-ref-era.ps1`（XML から組んだ合成の xlsx を Excel 16 ja-JP の COM で開いて Range.Text を読む。保存しない）。組み込み 27〜36・50〜58 の書式、元年は `[$-ja-JP-x-gannen]` で 2 文字以上の元号の後だけ、`[$-ja]`・`[$-0411]` も日本語、などを ADR-0007 §6（版 1.2）に書いた。
- `format-ja.ts` に g・e・r の字句・区画の和暦の扱い・`eraOfSerial`・`JA_BUILTIN_DATES`。ssf が `m.d` で例外を投げるので和暦の区画では「.」を引用符で包む。`@` の区画は数値に使わない。`builtinDates` は openXlsx・read・view・renderWorkbook に通した（`format.ts` は引数 1 つの口）。
- テスト：ref-era 7 件（2952 セル中 261 を理由付きで除外）・format-era 11 件・builtin-dates 1 件（全体 2290）。依頼のファイルの 12 行目が `令和8年11月11日` になり、ほかの行は変わらない。例のページに「和暦の書式」と「組み込みの和暦」を足した。
- core/index が gzip 36,071 で上限 36,000 を超えたので、理由を 02 §8 に書いて 40,000 にした。
