# CSP（Content Security Policy）

| 項目 | 値 |
| --- | --- |
| 文書 | XSheet 使い方の手引き |
| 更新日 | 2026-10-30 |
| 関連 | [README.md](README.md)・[../02_architecture.md](../02_architecture.md) §3「CSP」・[../01_requirements.md](../01_requirements.md) N7 |

## TL;DR
- 要るのは 3 つ：`script-src 'self'`・`style-src` の nonce（か `'unsafe-inline'`）・`img-src blob:`。単独版でも同じ。
- `view` に `nonce` を渡すと、自動で入れる CSS の `<style>` と書式の `<style>` の両方に付く。
- 自動の `<style>` を避けたいなら `css: 'none'` にして `xsheet.css`（`style.css`）を `<link>` で読む。書式の `<style>` には nonce が要るまま。

---

## 1. 要る指定

| 指定 | 何のため |
| --- | --- |
| `script-src 'self'` | ライブラリ（`xsheet.min.js`・`xsheet.mjs`・バンドラの出力）と、図形を描くときに読む形の表（`xsheet-geom.mjs`・geom のチャンク）・グラフを描くときに読むグラフの描画（`xsheet-chart.mjs`・chart-draw のチャンク）。別の出所に置くならその出所。`eval`・`new Function` は使わないので `'unsafe-eval'` は要らない |
| `style-src 'nonce-<値>'`（または `'unsafe-inline'`） | ① 書式の `<style>`（表示の根に 1 枚。ブックの塗り・文字の装飾）② `view` が自動で入れる表示の CSS の `<style data-xsheet-style>`（`css: 'auto'` のとき、文書か ShadowRoot に 1 回だけ） |
| `img-src blob:` | 図の画像は blob URL で出す（`data:` の URL は作らない） |

- 要らないもの：`font-src`（フォントを読まない）・`connect-src`（通信しない。使う側が `fetch(url)` を渡すなら、その URL の分は使う側の都合で要る）。
- 位置・大きさ・色は CSSOM（`el.style.left = …`）で入れる。`style` 属性の文字列・`innerHTML` は使わないので、これらは `style-src` に止められない。`view` が container の高さの代わりに付ける 480px も CSSOM。

## 2. `view` と nonce

```html
<meta
  http-equiv="Content-Security-Policy"
  content="default-src 'none'; script-src 'self'; style-src 'nonce-r4nd0m'; img-src blob:"
/>
<script src="xsheet.min.js"></script>
<script src="app.js"></script>
```

```js
// app.js（classic のスクリプトなので、await は async 関数の中で）
document.querySelector('#file').addEventListener('change', async (e) => {
  await XSheet.view('#viewer', e.target.files[0], { nonce: 'r4nd0m' });
});
```

- nonce は応答ごとにサーバーが作る推測できない値にする（固定の値は例のページだけ）。
- ページの nonce は `<script>`・`<link>` の `nonce` プロパティで読める（`document.currentScript.nonce` など）。属性はブラウザが隠すことがあるので `getAttribute('nonce')` では読まない。
- `view` を呼ぶスクリプトをページの中に書くなら、その `<script>` にも nonce を付け、`script-src` に `'nonce-…'` を足す（[examples/standalone.html](../../examples/standalone.html) がこの形）。外のファイルにすれば `'self'` だけで済む。
- nonce を渡し忘れると、ブラウザが `<style>` を止める。表の枠・スクロール・塗りが崩れたら、まずここを見る。

## 3. `<link>` で CSS を読む形（`css: 'none'`）

自動で入れる `<style>` を CSP で許したくないとき。

```html
<meta
  http-equiv="Content-Security-Policy"
  content="default-src 'none'; script-src 'self'; style-src 'self' 'nonce-r4nd0m'; img-src blob:"
/>
<link rel="stylesheet" href="xsheet.css" />
```

```js
XSheet.view('#viewer', file, { css: 'none', nonce: 'r4nd0m' });
```

- `xsheet.css`（単独版）と `xsheet/style.css`（バンドラ）は同じもの。バンドラなら `import 'xsheet/style.css'` で使う側の CSS に入る。
- `css: 'auto'` のままでも、`style.css` が読み込み済み（`.xs-workbook` の規則が効く）なら自動の `<style>` は入らない。`'none'` は「決して入れない」を明示する指定。
- 書式の `<style>` はブックごとに中身が変わるので `<link>` にできない。nonce（か `'unsafe-inline'`）が要るまま。

## 4. Shadow DOM

container が Shadow DOM の中にあるとき、`view` は表示の CSS をその ShadowRoot に入れる（文書の head には入れない）。nonce の扱いは同じ。`css: 'none'` にするなら、ShadowRoot の中に `<link rel="stylesheet" href="xsheet.css">` を置く。

## 5. 確かめ方

- `examples/index.html`（`npm run examples`）は `style-src` を nonce だけにした CSP で動く（`tests/render/examples.test.ts` が meta を見る）。
- `examples/standalone.html` は `default-src 'none'` から始めて、nonce と `img-src blob:` だけを足した CSP で、`file://` でも動く。
