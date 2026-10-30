// 単独版（tools/standalone/build.mjs）で jszip の `require("setimmediate")` の代わりに束ねる（M6-3）。
// npm の setimmediate は、関数でない引数を `new Function` で文字列から関数にする（N1 に反する）。
// jszip は関数しか渡さないので、その分岐の無い版を置く。大域の setImmediate が無いときだけ足す（元の polyfill と同じ）。
// MessageChannel で次のタスクに回す（setTimeout(0) は入れ子で 4ms ずつ遅れ、大きなファイルの展開が遅くなるため）。
const g = globalThis;
if (typeof g.setImmediate !== 'function') {
  if (typeof g.MessageChannel === 'function') {
    const queue = [];
    const channel = new g.MessageChannel();
    // 待ちの間に処理系（Node・Deno）の終了を止めない
    channel.port1.unref?.();
    channel.port2.unref?.();
    channel.port1.onmessage = () => {
      const task = queue.shift();
      if (task) task();
    };
    g.setImmediate = (fn, ...args) => {
      queue.push(() => fn(...args));
      channel.port2.postMessage(0);
    };
  } else {
    g.setImmediate = (fn, ...args) => g.setTimeout(fn, 0, ...args);
  }
}
