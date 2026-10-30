// Vite の `?inline`：CSS の文面を文字列として読む（facade/style.ts の CSS の自動挿入）
declare module '*.css?inline' {
  const css: string;
  export default css;
}
