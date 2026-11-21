// 별도 작업 공간(Web Worker)에서 Pyodide로 model.py를 실행한다.
const BASE = 'https://cdn.jsdelivr.net/pyodide/v0.27.7/full/';
const ALLOWED = ['explore', 'overview', 'simulate', 'linkage', 'price', 'compare', 'profile'];
let py = null;
let chain = Promise.resolve();

async function ensurePython() {
  if (py) return;
  postMessage({ type: 'status', text: 'Python 실행 환경(Pyodide) 준비 중… 처음 한 번만 20~40초 걸립니다.' });
  importScripts(BASE + 'pyodide.js');
  py = await loadPyodide({ indexURL: BASE });
  await py.loadPackage(['numpy', 'micropip']);
  await py.runPythonAsync("import micropip\nawait micropip.install('openpyxl==3.1.5')");
  const response = await fetch('./model.py');
  if (!response.ok) throw Error('model.py를 불러올 수 없습니다.');
  await py.runPythonAsync(await response.text());
}

async function handle({ id, type, payload }) {
  try {
    await ensurePython();
    let result;
    if (type === 'load') {
      postMessage({ type: 'status', text: payload.label + ' Excel 읽기와 계수 검증 중…' });
      py.globals.set('raw_input', new Uint8Array(payload.buffer));
      py.globals.set('level_input', payload.level);
      await py.runPythonAsync('raw_bytes = raw_input.to_py()');
      result = py.runPython('load_level(level_input, raw_bytes)');
      py.runPython('del raw_input, raw_bytes');
    } else if (type === 'call') {
      if (!ALLOWED.includes(payload.fn)) throw Error('허용되지 않은 함수: ' + payload.fn);
      py.globals.set('fn_name', payload.fn);
      py.globals.set('args_json', JSON.stringify(payload.args));
      result = py.runPython('import json as _j\nglobals()[fn_name](*_j.loads(args_json))');
    } else {
      throw Error('알 수 없는 요청: ' + type);
    }
    postMessage({ id, ok: true, result: JSON.parse(result) });
  } catch (error) {
    let msg = String(error && error.message || error);
    const last = msg.trim().split('\n').filter(Boolean).pop();
    postMessage({ id, ok: false, error: last || msg });
  }
}

onmessage = e => { chain = chain.then(() => handle(e.data)); };
