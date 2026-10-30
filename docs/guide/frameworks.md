# React・Vue での置き方

| 項目 | 値 |
| --- | --- |
| 文書 | XSheet 使い方の手引き |
| 更新日 | 2026-10-30 |
| 関連 | [README.md](README.md)・[../../README.md](../../README.md)「`view`」 |

## TL;DR
- 空の `div` を 1 つ用意し、mount したら `view`、unmount（とファイルが替わったとき）に `destroy()`。
- `view` は非同期。待っている間に unmount・ファイルの替えが起きたら、解決したときにその `Viewer` を `destroy()` する（React の StrictMode でも同じ形で足りる）。
- XSheet が中に足す DOM（`div.xs-viewer`）は XSheet のもの。その `div` の中にフレームワークの子を描かない。

---

## 1. 共通の約束

| こと | 内容 |
| --- | --- |
| 置き場所 | 中身の無い要素を 1 つ渡す。`view` はその子に `div.xs-viewer` を足すだけで、要素自身の属性・クラスは変えない |
| 高さ | 要素に高さを付ける（`height`・flex の子など）。高さが 0 だと 480px で出し、`console.warn` を 1 回出す |
| 片付け | `destroy()` で表示・ブック・リスナ・blob URL を放す。何度呼んでもよい |
| 例外 | 開けないファイルは例外にならない（帯に文言が出る）。例外になるのは、要素が見つからない・受けない形の source・失敗した `Response` のとき（`TypeError`） |
| SSR | `xsheet` は読み込んだ時点では DOM に触れないので、サーバー側で import してもよい。`view` はブラウザで（React は `useEffect`、Vue は `onMounted` の中で）呼ぶ |
| 言語 | `lang` の既定 `auto` は `<html lang>` を先に見る。アプリの言語の設定があるなら `lang` を渡す |

## 2. React

```tsx
import { useEffect, useRef } from 'react';
import { view, type Viewer } from 'xsheet';

export function XlsxViewer({ file }: { file: Blob | null }) {
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const el = ref.current;
    if (el === null || file === null) return;
    let viewer: Viewer | undefined;
    let cancelled = false;
    view(el, file, { theme: 'auto' }).then(
      (v) => {
        // 待つ間に unmount した・file が替わった：出来た表示をすぐ捨てる
        if (cancelled) v.destroy();
        else viewer = v;
      },
      (e: unknown) => {
        if (!cancelled) console.error(e);
      },
    );
    return () => {
      cancelled = true;
      viewer?.destroy();
    };
  }, [file]);

  return <div ref={ref} style={{ height: 600 }} />;
}
```

- **StrictMode**：開発のときは effect が「実行 → 片付け → 実行」と 2 回走る。1 回目の `view` は片付けの後に解決するので、`cancelled` を見て `destroy()` される。残るのは 2 回目の表示だけになる。
- `file` が替わったときも同じ流れ（前の effect の片付け → 新しい `view`）。古い読み込みが後から解決しても、表示は重ならない。
- `theme`・`lang` などの指定を props から渡すなら、それも依存の配列に入れる（替わったら出し直す）。
- シートを外から替えるなら、`viewer` を `useRef` に入れて `viewer.showSheet(名前か番号)` を呼ぶ。シートの一覧は `viewer.sheets`。

## 3. Vue 3

```vue
<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref, watch } from 'vue';
import { view, type Viewer } from 'xsheet';

const props = defineProps<{ file: Blob | null }>();
const el = ref<HTMLDivElement>();
// Viewer は ref() に入れない（深い reactive の proxy で包まない）
let viewer: Viewer | undefined;
let seq = 0;

async function show(file: Blob | null): Promise<void> {
  const mine = ++seq;
  viewer?.destroy();
  viewer = undefined;
  if (el.value === undefined || file === null) return;
  try {
    const v = await view(el.value, file, { theme: 'auto' });
    // 待つ間に次の file が来た・unmount した：出来た表示をすぐ捨てる
    if (mine !== seq) v.destroy();
    else viewer = v;
  } catch (e) {
    if (mine === seq) console.error(e);
  }
}

onMounted(() => show(props.file));
watch(() => props.file, show);
onBeforeUnmount(() => {
  seq++;
  viewer?.destroy();
  viewer = undefined;
});
</script>

<template>
  <div ref="el" style="height: 600px"></div>
</template>
```

- 連番（`seq`）で「いちばん新しい呼び出しだけを残す」。unmount でも連番を進めるので、読み込み中に外れた表示は解決したときに捨てられる。
- `Viewer` を template から使う（シートの一覧を出すなど）なら、`shallowRef` に入れる。

## 4. `view` を使わない形

CSS を自分で読み込み、文言の表も自分で渡したいときは `renderWorkbook`（`xsheet/render`）を同じ形で使える。違いは次のとおり。

| | `view` | `renderWorkbook` |
| --- | --- | --- |
| 受ける source | File・Blob・Response など（Promise も） | `ArrayBuffer`・`Uint8Array` だけ |
| CSS | 自動で入れる | `xsheet/style.css` を import する |
| 高さが 0 の container | 480px にして警告 | 0px のまま（何も見えない） |
| 既定の言語 | `auto` | 日本語 |
| 名前でシートを替える | できる | 番号だけ |
