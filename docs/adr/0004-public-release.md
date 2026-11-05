# 0004. XSheet を公開する（D2・D3・D5 の見直し）

> 版: 1.0 ・ 更新日: 2026-10-30 ・ 担当: kwon
> 関連文書: [0001](0001-library-decision.md)（D2・D3） / [0002](0002-build-and-distribution.md)（D5） / [0003](0003-universal-distribution.md)（誰でも使える配り方） / 参考: [InterHub-MCP](https://github.com/TaekangKwon/InterHub-MCP)（同じ作者の公開リポジトリ）

## TL;DR
- GitHub のリポジトリ `TaekangKwon/XSheet` を公開する。ライセンスは **MIT**（D3）。Git の履歴はそのまま残す。
- 公開の入口の README は英語（`README.md`）にし、日本語（`README.ja.md`）・韓国語（`README.ko.md`）の訳を置く。構成は InterHub-MCP に倣う。`docs/` は日本語のまま。
- 配り方（D5）は変えない：タグに `dist/` を含めた git 依存と、GitHub Release の単独版の zip。公開すれば権限なしで入れられる。npm への公開は別に決める。
- 公開の切り替えの前に、作者が InterHub 由来のコードを MIT で公開してよいことを確かめる（下の §6）。

---

## ステータス

accepted（2026-10-30。ユーザーが推奨案「MIT・履歴を残す・英語の README＋ja／ko・まず GitHub の公開だけ」で決めた）。リポジトリを公開に切り替えるのは §6 の条件がそろい、ユーザーが指示したとき。

## コンテキスト

- ADR-0003 は「まず社内（リポジトリの権限がある人）向けに作り、社外への公開は別に決める」とした。単独版・`view` の 1 行・3 言語の文言・CSP の手引きがそろい、社外の人がそのまま使える形になった。
- 非公開のままでは、git 依存にもタグの zip にもリポジトリの権限が要る（ADR-0003 の P2 が残っている）。
- 同じ作者の InterHub-MCP は MIT で公開済みで、英語の README＋日本語・韓国語の訳・CI のバッジという形をとっている。
- 公開の前に見つけた、外に出さないほうがよいもの：作者の PC のパス（`C:\GitHub\InterHub`・`D:\tools\…`）、非公開の InterHub の中のファイルのパスに頼った手順、存在しない資料への参照、「非公開」の前提の文。

## 決定

### 1. ライセンス（D3）

- MIT。根に `LICENSE` を置き、`package.json` を `"license": "MIT"` にして `private` を外す。`files` に `LICENSE` を足し、単独版の zip にも入れる。
- 形の表の元（ECMA-376 の `presetShapeDefinitions.xml`）の著作権表示 `NOTICE` は今までどおり配布物に含める。
- 依存との関係：jszip は MIT と GPL-3.0-or-later の選択（MIT を選ぶ）、ssf は Apache-2.0。どちらも MIT のライブラリから使える。単独版が中に持つ依存のライセンスは `xsheet.licenses.txt` に入る（M6-3 のまま）。

### 2. 履歴

- Git の履歴は書き換えない。コミットの作者名・メールアドレス・メッセージも公開される。

### 3. README と文書の言語

- `README.md` は英語。`README.ja.md`・`README.ko.md` はその訳で、3 つとも 1 行目の下に言語の切り替えを置き、同じ中身に保つ。
- 構成は InterHub-MCP に倣う：一言の説明とバッジ → 作った理由 → 3 行で始める → 入れ方 → API → 表示できるもの・制限 → セキュリティの考え方 → 対応環境 → 作り（3 層） → 開発 → 文書の一覧 → 謝辞 → ライセンス。
- `SECURITY.md`（英語）を置く。信頼できない .xlsx を開くライブラリなので、上限（`XlsxLimits`）と報告の窓口を書く。
- `docs/`・コメント・コミットメッセージは日本語のまま（CLAUDE.md）。

### 4. 配り方（D5）と npm

- 変えない。タグに `dist/` を含め（ADR-0002）、GitHub Release に単独版の zip を添付する（ADR-0003 §6）。公開すれば `npm install github:TaekangKwon/XSheet#v<版>` と zip が権限なしで使える。
- npm への公開は、名前 `xsheet` が空いているかの確認と、半公開の入口 `xsheet/internal` の扱い（版の互換を約束しない）を決めてから別に決める。→ [ADR-0006](0006-npm-publish.md) で決めた（2026-11-05）。

### 5. 外に出さないもの

- 作者の PC のパスを文書・道具から外す。InterHub は「作者の非公開リポジトリ」と書く。
- InterHub の取り込みの道具（`tools/sync/`）は作者だけが使える。InterHub の場所は環境変数 `INTERHUB_DIR`、無ければ隣のフォルダ `../InterHub`。
- InterHub・xrag-Viewer の名前は、作った理由（同じ .xlsx の表示を 2 か所で直していた）として残す。非公開のリポジトリの中のファイルのパスに頼る手順は書かない。
- CI（GitHub Actions）で test・typecheck・lint・format:check・build・size・check:consumer・check:standalone を回す。

### 6. 公開に切り替える条件

1. 作者が、InterHub 由来のコード（ADR-0001・03 の対応表のとおり写したもの）を MIT で公開してよいことを確かめる。InterHub は非公開で、ADR-0003 は社内向けとしていた。
2. この ADR の準備（README・LICENSE・CI・外に出さないものの整理）が `main` に入り、CI が通る。
3. 切り替えの直前に、パス・メールアドレス・秘密の値が残っていないことを `git grep` でもう一度確かめる。
4. 切り替え（`gh repo edit --visibility public`）はユーザーの指示があるときだけ行う。公開した後は `main` の保護（PR を通す）を付け、`SECURITY.md` の窓口にする非公開の脆弱性報告（Private vulnerability reporting。非公開のリポジトリでは使えない）を有効にする。

## 却下した案

| 案 | 却下の理由 |
| --- | --- |
| Apache-2.0 | 特許の条項はこのライブラリでは要らない。InterHub-MCP と同じ MIT にそろえる |
| 履歴を新しく始める（1 コミットにまとめる） | 決定と作業の記録（コミット・PR #1〜#10）が失われる。履歴を調べ、残るのは作者の PC のパス（文書・道具の古い版）とコミットのメールアドレスだけで、トークン・鍵などの秘密の値は無いことを確かめた（2026-10-30） |
| 公開と同時に npm に出す | 名前の確認と `xsheet/internal` の扱いが決まっていない。InterHub-MCP も GitHub の公開を先にした |
| README を日本語のまま | 使う人の大半が読めない。InterHub-MCP と同じく英語を入口にし、訳を並べる |

## 結果

- 良い点：権限なしで git 依存・zip が使える。ライセンスがはっきりし、社外のアプリに組み込める。
- 注意：README が 3 つになる。変えるときは 3 つとも直す（CLAUDE.md）。
- 注意：公開した後は Issue・PR が外から来る。外から来た PR のテストデータは、自作の合成か Excel で自作したファイルだけを受け付ける（CLAUDE.md の「守ること」）。
