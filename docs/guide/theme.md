# テーマ（ダーク表示）と CSS 変数

| 項目 | 値 |
| --- | --- |
| 文書 | XSheet 使い方の手引き |
| 更新日 | 2026-10-30 |
| 関連 | [README.md](README.md)・[../02_architecture.md](../02_architecture.md) §6（変数の表） |

## TL;DR
- ダークは `view` の `theme: 'dark'`、OS に合わせるなら `'auto'`。値は `style.css` が持つ（変数を写す必要はない）。
- `renderWorkbook` などでは、表示の根か祖先（`<html>` など）に `data-xs-theme="dark"`（`"auto"`）を付ける。
- 色を替えるときは `--xs-*` を上書きする。明るい表示はどこで上書きしても勝つ。ダークの値は、テーマの属性の付いた要素より内側で上書きする。

---

## 1. テーマの選び方

| 使い方 | 書き方 | 付く場所 |
| --- | --- | --- |
| `view` | `view(el, file, { theme: 'dark' })`・`{ theme: 'auto' }` | 表示の根 `div.xs-viewer` に `data-xs-theme` |
| どの使い方でも | `<html data-xs-theme="dark">`（`"auto"`） | 書いた要素。内側の XSheet の表示が全部ダークになる |

- `theme: 'light'`（既定）は何も付けない。そのため、祖先に `data-xs-theme="dark"` があればダークになる（アプリ全体のテーマに従わせるなら、`theme` を渡さず `<html>` に付ける）。
- `auto` は `prefers-color-scheme: dark` のときだけダークの値になる。
- アプリが自前のテーマの属性（`data-theme` など）を持っているなら、切り替えのときに `document.documentElement.dataset.xsTheme = 'dark'` のように合わせて付ける。
- ダークでは `color-scheme: dark` も付く（スクロールバーなどがダークになる）。
- XSheet が色を付けるのは自分の要素だけ。ページの地の色は使う側が決める。

## 2. ダークでも変わらないもの

Excel の色（塗り・文字の色・図形の色）は変数ではなく、ファイルの色のまま描く（InterHub ADR-0030 の規則）。

| もの | ダークでの描き方 |
| --- | --- |
| 塗りのあるセル | Excel の色のまま |
| 白の塗り | 「紙」として `--xs-cell-bg` で描く |
| 自動・黒の罫線 | `--xs-border-ink`（墨色）で描く |
| 画像・図形 | 白い下地（`--xs-backdrop`）の上に描く |

## 3. 色を替える

変数は 14 個。役割と明るい既定値・ダークの値は [02](../02_architecture.md) §6。

### 明るい表示

既定値は詳細度 0（`:where(:root, :host)`）なので、どこで上書きしても勝つ。

```css
:root {
  --xs-accent: #0f766e;
  --xs-grid: #dde3ea;
}
```

### ダーク表示

ダークの値は `data-xs-theme` の付いた要素に直接付く。そのため `:root` に書いた値は、その要素より内側には届かない。テーマの属性の付いた要素より**内側**（表示の根 `.xs-workbook`・`.xs-sheet`）で上書きする。

```css
[data-xs-theme='dark'] :is(.xs-workbook, .xs-sheet) {
  --xs-accent: #f472b6;
}
@media (prefers-color-scheme: dark) {
  [data-xs-theme='auto'] :is(.xs-workbook, .xs-sheet) {
    --xs-accent: #f472b6;
  }
}
```

- 属性の付いた要素そのものに同じ詳細度（`[data-xs-theme='dark'] { … }`）で書くと、後から入る CSS が勝つ。`view` が自動で入れる `<style>` は head の最後に足されるので、使う側の CSS が負けることがある。内側に書けば順序に左右されない。
- `.xs-workbook` は `view`・`renderWorkbook` の表示の根、`.xs-sheet` は `renderSheet` の表示の根（`renderWorkbook` の中にもある）。

### テーマを使わず全部を自分で決める

`data-xs-theme` を付けず、14 個の変数を自前のテーマの属性で上書きしてもよい（v0.1.0 までの形。値の例は 02 §6 の「ダークの例」の列）。

```css
:root[data-theme='dark'] {
  --xs-fg: #e6e6e6;
  --xs-bg: #1e1e1e;
  /* …14 個 */
}
```

- `view` は、表示の CSS（`xsheet/style.css`・`xsheet.css`）が読み込み済みかを規則が効くかで見る（変数では見ない）。`:root` に `--xs-*` を書いただけなら、CSS は自動で入る。

## 4. Shadow DOM

- `view` は container が Shadow DOM の中なら、表示の CSS をその ShadowRoot に入れる。既定値は `:host` に付く。
- 文書の `<html data-xs-theme="dark">` は Shadow DOM の中の表示に届かない（中の既定値が host で定義し直すため）。`view` の `theme` を使うか、ShadowRoot の中の要素（container など）に `data-xs-theme` を付ける。
- 明るい表示の変数は、文書の CSS から host 要素に書けば中まで効く（`my-viewer { --xs-accent: #0f766e; }`）。
- ダークの値を替えるときは、ShadowRoot の中の CSS に 3 の形（テーマの属性より内側で上書き）で書く。
