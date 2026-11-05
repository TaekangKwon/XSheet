# 0006. npm への公開（名前と入口）

> 版: 1.0 ・ 更新日: 2026-11-05 ・ 担当: kwon
> 関連文書: [0002](0002-build-and-distribution.md)（D5・タグの `dist/`） / [0003](0003-universal-distribution.md)（入口 `xsheet`・`xsheet/core`） / [0004](0004-public-release.md)（§4「npm は名前と `xsheet/internal` を決めてから」）

## TL;DR
- npm の名前は **`@taekangkwon/xsheet`**（作者のスコープ・公開）。素の `xsheet` は空いているが、既にある `x-sheet` と記号だけが違うため、npm の名前の規則で断られるおそれが高い。
- 半公開の入口 `xsheet/internal` は **`exports` から外す**。npm には版の互換を約束する入口だけを載せる。
- npm へは GitHub Actions から、トークンを置かない信頼された公開（OIDC）と provenance で出す。出すのは版の Release を作ったとき。
- この ADR は決定だけ。名前の付け替え・`exports` の変更・公開のワークフローは別の作業（下の §4）で、npm に出すのはユーザーの指示があるとき。

---

## ステータス

accepted（2026-11-05。ユーザーが推奨案「`@taekangkwon/xsheet`・`internal` を外す」で決めた）。npm への公開そのものは §4 の準備がそろい、ユーザーが指示したとき。

## コンテキスト

- ADR-0004 §4 は、npm への公開を「名前 `xsheet` が空いているかの確認」と「`xsheet/internal` の扱い」を決めてから、とした。
- 名前の確認（2026-11-05・`npm view`）：
  - `xsheet`・`xsheets`・`x_sheet`・`x.sheet`・`xsheetjs`・`xsheet-viewer` は無い。
  - **`x-sheet` がある**（0.0.0・2024-01-29 更新）。npm は、既にある名前と記号（`-`・`_`・`.`）だけが違う新しい名前を「似すぎている」として断る。`xsheet` はこれに当たる。
  - `@taekangkwon/xsheet` も無い。スコープ付きの名前は、スコープの持ち主しか出せないので、似た名前の規則に掛からない。
- `xsheet/internal`（`src/core/internal.ts`）は、render が使う core の部品（`applyFillMode`・`applyTransforms`・`formatNumberByCode`・`DEFAULT_THEME` など）を出す入口で、版の互換を約束していない（README の入口の表）。`dist/` の render は core を内部のチャンクとして相対で読むため、`exports` に無くても動く。
- npm に載せた入口は、使う人が頼る。互換を約束しない入口を載せると、0.x の小さな版の上げでも使う側が壊れる。

## 決定

### 1. 名前

- `package.json` の `name` を `@taekangkwon/xsheet` にし、`publishConfig.access` を `public` にする。
- 入口は `@taekangkwon/xsheet`・`@taekangkwon/xsheet/core`・`/geom`・`/render`・`/style.css`（今の `xsheet…` の付け替え）。
- 変えないもの：リポジトリ名・ライブラリの呼び名（XSheet）・単独版のファイル名（`xsheet.min.js` など）と大域 `XSheet`・CSS の接頭辞 `xs-`。
- git 依存（`npm install github:TaekangKwon/XSheet#v<版>`）で入れた場合も、`node_modules` の中の名前は新しい名前になる。名前を変えた版から、README の `import` の例を新しい名前にする。

### 2. 入口

- `exports` から `./internal` を外す。`src/core/internal.ts` は残し、render と試験が使う（`src/render/**` が core を引く入口の制限は今のまま）。
- 外すのは互換を崩す変更なので、名前を変える版と同じ版（0.x の小さな版の上げ）にまとめ、CHANGELOG の「互換の注意」に書く。

### 3. 出し方

- `.github/workflows/publish.yml`：Release を公開したとき（`release: published`）に、タグを取り出し、CI と同じ検査のあと `npm publish --provenance --access public`。npm の「信頼された公開（Trusted Publisher）」で、このリポジトリのこのワークフローだけが出せるようにする。npm のトークンはリポジトリに置かない。
- タグの `dist/` で出す（ADR-0002 の配り方と同じ物を npm に載せる）。`npm pack --dry-run` の中身は `check:consumer` が確かめている。
- 版を出す手順（02 §8）に「Release を公開すると npm に出る」を足す。

### 4. この後の作業（それぞれ別の Issue）

1. （ユーザー）npm のアカウント `taekangkwon`（2 段階認証あり）を用意し、パッケージ `@taekangkwon/xsheet` の Trusted Publisher にこのリポジトリと `publish.yml` を登録する。
2. 名前の付け替えと `./internal` を外す：`package.json`・README（英・日・韓）・`docs/guide/`・examples・`tools/consumer-check.mjs`・`tests/core/public-api.test.ts`。
3. 公開のワークフロー `publish.yml` と、02 §8 の手順の更新。
4. 版を出す（0.5.0 の予定）。npm に出るのは、ユーザーの指示で Release を公開したとき。

## 却下した案

| 案 | 却下の理由 |
| --- | --- |
| 素の `xsheet` で出してみる | `x-sheet` と似すぎているとして断られるおそれが高い。断られてから名前を変えると、README・例の付け替えを 2 回することになる |
| スコープの無い別の名前（`xlsx-view` など） | ライブラリの呼び名（XSheet）と名前がずれる。空いていても、ほかの似た名前に掛かるかを毎回確かめることになる |
| `xsheet/internal` を残す（README に「互換を約束しない」と書いたまま） | npm に載ると頼る人が出て、render の都合の変更が使う側を壊す |
| 手元の PC から `npm publish`（トークン） | トークンの管理が要り、provenance（どのコミット・ワークフローから作ったか）を付けられない |

## 結果

- 良い点：`npm install @taekangkwon/xsheet` で入れられる。provenance で、npm の物がこのリポジトリのタグから作られたことを確かめられる。互換を約束する入口だけが残る。
- 注意：名前が変わるので、`xsheet` で `import` している使う側は付け替えが要る（CHANGELOG と README に書く）。
- 注意：npm に出した版は消せない（72 時間を過ぎると取り下げにも制限がある）。出す前に、タグからの入れ方と単独版を確かめる。
