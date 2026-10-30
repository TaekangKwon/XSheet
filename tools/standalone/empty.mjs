// 単独版（tools/standalone/build.mjs）で、jszip が Node のストリームを調べる `require("stream")` の代わりに束ねる空のモジュール（M6-3）。
// jszip は Readable が無ければ Node のストリームを使わない（support.nodestream が false）。
export {};
