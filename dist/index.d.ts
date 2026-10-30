/**
 * パッケージの入口（`xsheet`）。入口の層（`view`・`read`）＋core（読み取り）＋render（表示）の公開物を全部出す（M6-2）。
 * 読み込んだ時点では DOM に触れない（Node で import してよい）。DOM・Node 固有の API を使わないことを約束する入口は `xsheet/core`。
 * 使わない部分はバンドラが落とす（package.json の sideEffects は CSS だけ）。
 */
export * from './core/index';
export * from './render/index';
export * from './facade/index';
