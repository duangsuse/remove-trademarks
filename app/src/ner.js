// 离线 NER：多个 Worker 各持一个 onnxruntime 会话并行推理（WASM 线程池 + 可选 1 个 WebGPU 线程），内置 bert-base-chinese 词表
import NerWorker from './ner.worker.js?worker&inline';
import { toTrad } from './s2t.js';
import vocabText from './vocab.txt?raw';
import LABELS from './labels.json';

const vocab = new Map();
vocabText.split('\n').forEach((w, i) => vocab.set(w, i));
const UNK = vocab.get('[UNK]'), CLS = vocab.get('[CLS]'), SEP = vocab.get('[SEP]');

const isCJK = (c) => (c >= 0x4e00 && c <= 0x9fff) || (c >= 0x3400 && c <= 0x4dbf) || (c >= 0x20000 && c <= 0x2a6df) ||
  (c >= 0x2a700 && c <= 0x2ebef) || (c >= 0xf900 && c <= 0xfaff) || (c >= 0x2f800 && c <= 0x2fa1f);
const isPunct = (ch, c) => (c >= 33 && c <= 47) || (c >= 58 && c <= 64) || (c >= 91 && c <= 96) || (c >= 123 && c <= 126) || /\p{P}/u.test(ch);
const isSpace = (ch) => /[\s\u0000�]/.test(ch) || /\p{Cc}/u.test(ch);

// 返回 [{id, s, e}]（s/e 为原文 UTF-16 偏移）
export function tokenize(text) {
  const words = [];
  let cur = null, i = 0;
  const flush = () => { if (cur) { words.push(cur); cur = null; } };
  for (const ch of text) {
    const c = ch.codePointAt(0), n = ch.length;
    if (isSpace(ch)) flush();
    else if (isCJK(c) || isPunct(ch, c)) { flush(); words.push({ w: ch, s: i, e: i + n }); }
    else { if (!cur) cur = { w: '', s: i, e: i }; cur.w += ch; cur.e = i + n; }
    i += n;
  }
  flush();
  const out = [];
  for (const { w, s } of words) {
    const chars = [...w];
    if (chars.length > 100) { out.push({ id: UNK, s, e: s + w.length }); continue; }
    const pieces = []; let st = 0, bad = false;
    while (st < chars.length) {
      let en = chars.length, hit = null;
      while (st < en) {
        const sub = (st > 0 ? '##' : '') + chars.slice(st, en).join('');
        if (vocab.has(sub)) { hit = { id: vocab.get(sub), len: chars.slice(st, en).join('').length }; break; }
        en--;
      }
      if (!hit) { bad = true; break; }
      pieces.push(hit); st = en;
    }
    if (bad) { out.push({ id: UNK, s, e: s + w.length }); continue; }
    let p = s;
    for (const pc of pieces) { out.push({ id: pc.id, s: p, e: p + pc.len }); p += pc.len; }
  }
  return out;
}

// ---------------- 线程池 ----------------
// pool 元素：{ ep, w, state: 'init' | 'idle' | 'busy' | 'dead', job, timer }
let pool = [];
let cur = null; // 正在进行的 recognize：{ pump, fail }
export let backend = '';

export async function loadWasmBinary() {
  const el = document.getElementById('ort-wasm-gz');
  if (!el) return null; // 开发模式：Worker 回退到默认路径
  const bin = Uint8Array.from(atob(el.textContent.trim()), (c) => c.charCodeAt(0));
  const ds = new Response(new Blob([bin]).stream().pipeThrough(new DecompressionStream('gzip')));
  return new Uint8Array(await ds.arrayBuffer());
}

async function hasWebGPU() {
  try { return !!(navigator.gpu && (await navigator.gpu.requestAdapter())); } catch { return false; }
}

// 每个 WASM 线程实测约占 0.5~0.6GB 内存（模型权重 + ORT 堆，WASM 堆只增不减），所以默认最多 4 个：
// 取 CPU 核数的一半、内存(GB)的一半的较小者。想更快可在地址后加 ?threads=6（6 个线程全文约 30 秒，约多占 1GB）
export function workerCount() {
  const q = parseInt(new URLSearchParams(location.search).get('threads'), 10);
  if (q > 0) return Math.min(12, q);
  const hc = navigator.hardwareConcurrency || 4, mem = navigator.deviceMemory || 4;
  return Math.max(1, Math.min(4, hc >> 1, mem >> 1));
}

function describe() {
  const live = pool.filter((p) => p.state !== 'dead' && p.state !== 'init');
  const w = live.filter((p) => p.ep === 'wasm').length, g = live.filter((p) => p.ep === 'webgpu').length;
  backend = [w ? `WASM×${w}` : '', g ? 'WebGPU' : ''].filter(Boolean).join(' + ');
}

function kill(p, why) {
  if (p.state === 'dead') return;
  clearTimeout(p.timer);
  p.state = 'dead'; p.why = why;
  try { p.w.terminate(); } catch { /* 忽略 */ }
  describe();
}

function spawn(ep, model, wasm) {
  const p = { ep, state: 'init', w: new NerWorker(), job: null, timer: 0, fails: 0 };
  p.ready = new Promise((res, rej) => {
    p.w.onmessage = ({ data: m }) => {
      if (m.t === 'ready') { p.state = 'idle'; describe(); res(); cur?.pump(); }
      else if (m.t === 'error') { console.warn(ep, m.msg); kill(p, m.msg); rej(new Error(m.msg)); }
      else if (m.t === 'done' || m.t === 'fail') {
        clearTimeout(p.timer);
        const job = p.job; p.job = null;
        if (p.state === 'busy') p.state = 'idle';
        if (job && cur) m.t === 'done' ? job.done(m) : job.fail(p, m.msg);
        cur?.pump();
      }
    };
    p.w.onerror = (e) => { const msg = e.message || '线程崩溃'; const job = p.job; p.job = null; kill(p, msg); rej(new Error(msg)); if (job && cur) job.fail(p, msg, true); cur?.pump(); };
  });
  p.ready.catch(() => {});
  p.w.postMessage({ t: 'init', ep, model, wasm });
  return p;
}

let model = null; // 已选定的模型文件（Blob）；线程在第一次识别时才拉起
export const ready = () => !!model;

export function dispose() { for (const p of pool) kill(p, 'dispose'); pool = []; backend = ''; }

// 只登记模型文件，不启动线程（载入页面 / 模型时不占内存）。参数可以是 File / Blob / ArrayBuffer
export async function loadModel(data) {
  dispose();
  model = data instanceof Blob ? data : new Blob([data]); // Blob 在线程间按引用传递，不会拷贝 100MB
  modelHash = ''; cache.clear();
  backend = '待启动';
  return backend;
}

// 需要时才拉起线程：返回第一个线程就绪时的后端描述，其余线程在后台陆续加入；全部失效后再次调用会重新拉起
let starting = null;
function ensurePool(onStatus = () => {}) {
  if (pool.some((p) => p.state !== 'dead')) return Promise.resolve();
  return (starting ||= (async () => {
    try {
      onStatus('解压运行时…');
      const wasm = await loadWasmBinary();
      const eps = [...((await hasWebGPU()) ? ['webgpu'] : []), ...Array(workerCount()).fill('wasm')];
      onStatus(`启动 ${eps.length} 个推理线程…`);
      pool = eps.map((ep) => spawn(ep, model, wasm));
      try { await Promise.any(pool.map((p) => p.ready)); }
      catch (e) { const msg = pool.map((p) => `${p.ep}: ${p.why}`).join('；'); dispose(); throw new Error(msg || e.message); }
      // 第一个就绪后再多等其余线程几秒，让后端描述尽量完整
      await Promise.race([Promise.allSettled(pool.map((p) => p.ready)), new Promise((r) => setTimeout(r, 5000))]);
      describe();
    } finally { starting = null; }
  })());
}

// ---------------- 结果缓存：内容寻址 ----------------
// 键 = SHA-256( 模型文件内容的哈希 ‖ 缓存版本 ‖ 是否简转繁 ‖ 行文本 )，与文件名 / 大小 / 行号无关；值 = [[s, e, type, score], ...]
// 换了模型文件（内容不同）旧缓存自动失效；同一文本在任何文档里都复用。
const CACHE_VER = 2;
const cache = new Map();
let pending = null, modelHash = '', hashing = null;
const hex = (buf, n = 16) => [...new Uint8Array(buf).slice(0, n)].map((x) => x.toString(16).padStart(2, '0')).join('');
const sha = async (data) => crypto.subtle.digest('SHA-256', typeof data === 'string' ? new TextEncoder().encode(data) : data);
async function modelId() { // 模型内容哈希，每个模型只算一次（100MB 约 0.3 秒）
  if (modelHash) return modelHash;
  return (hashing ||= sha(await model.arrayBuffer()).then((h) => (modelHash = hex(h, 32)))).finally(() => { hashing = null; });
}
const lineKey = async (mh, pre, t) => hex(await sha(`${mh}\0${CACHE_VER}\0${pre}\0${t}`));
export const cacheExport = () => ({ v: CACHE_VER, model: modelHash, items: [...cache].slice(-40000) });
export function cacheImport(o) { if (o?.v === CACHE_VER && o.model && Array.isArray(o.items)) pending = o; } // 模型哈希对得上才会并入

// ---------------- 识别 ----------------
const hasCjk = /[㐀-鿿]/;
const WIN = 440, PART = 400;

function decode(win, lab, prob, out) {
  let span = null;
  const close = () => {
    if (span) out.push([span.line, span.s, span.e, span.type, span.sum / span.n]);
    span = null;
  };
  for (let j = 0; j < win.meta.length; j++) {
    const m = win.meta[j];
    if (!m) { close(); continue; }
    const l = LABELS[lab[j]];
    if (l === 'O') { close(); continue; }
    const bio = l[0], type = l.slice(2);
    if (bio === 'B' || bio === 'S' || !span || span.type !== type || span.line !== m.line) { close(); span = { type, line: m.line, s: m.t.s, e: m.t.e, sum: 0, n: 0 }; }
    span.e = m.t.e; span.sum += prob[j]; span.n++;
    if (bio === 'E' || bio === 'S') close();
  }
  close();
}

// 对多行文本做 NER，返回 [{text, type, score, line, s, e}]。
// 做了：跳过无汉字的行、相同行只算一次、命中缓存的行不再推理、简转繁（模型由繁体语料训练，偏移不变）、
//       ~400 token 打包、多线程并行；线程失败（如 WebGPU 设备丢失）时该窗口自动改由其他线程重做。
export async function recognize(lines, onProgress = () => {}, { trad = true, signal, onStatus } = {}) {
  const pre = trad ? 't' : 's';
  const first = new Map(); // 行文本 -> 首次出现的行号
  lines.forEach((t, i) => { if (hasCjk.test(t) && !first.has(t)) first.set(t, i); });
  const out = [];
  const emit = (line, t, ents) => { for (const [s, e, type, score] of ents) out.push({ text: t.slice(s, e), type, score, line, s, e }); };
  const mh = first.size ? await modelId() : '';
  if (pending?.model === mh) { for (const [k, v] of pending.items) cache.set(k, v); pending = null; }
  const items = [];
  for (const [t, i] of first) {
    const key = await lineKey(mh, pre, t), hit = cache.get(key);
    if (hit) emit(i, t, hit);
    else { const tok = tokenize(trad ? toTrad(t) : t); if (tok.length) items.push({ i, t, key, tok }); else cache.set(key, []); }
  }
  const windows = [];
  let w = null;
  for (const it of items) {
    for (let k = 0; k < it.tok.length; k += PART) {
      const part = it.tok.slice(k, k + PART);
      if (!w || w.n + part.length + 1 > WIN) { w = { n: 0, segs: [] }; windows.push(w); }
      w.segs.push({ line: it.i, tok: part }); w.n += part.length + 1;
    }
  }
  for (const win of windows) {
    const ids = [CLS], meta = [null];
    for (const sg of win.segs) { for (const t of sg.tok) { ids.push(t.id); meta.push({ line: sg.line, t }); } ids.push(SEP); meta.push(null); }
    win.ids = Int32Array.from(ids); win.meta = meta; win.ents = [];
  }
  const total = windows.reduce((a, x) => a + x.ids.length, 0), t0 = performance.now();
  if (!total) { onProgress(1, { tokens: 0, total: 0 }); return out; }
  if (!ready()) throw new Error('模型未载入');
  await ensurePool(onStatus);

  let doneTok = 0, finished = 0, next = 0, lastErr = '';
  const retry = [], tries = new Array(windows.length).fill(0);
  await new Promise((resolve, reject) => {
    const abort = () => { cur = null; reject(new DOMException('已取消', 'AbortError')); };
    const send = (p, wi) => {
      const win = windows[wi];
      p.state = 'busy';
      p.job = {
        done: (m) => {
          decode(win, m.lab, m.prob, win.ents); doneTok += win.ids.length; finished++;
          const sec = (performance.now() - t0) / 1000;
          onProgress(doneTok / total, { tokens: doneTok, total, tokPerSec: doneTok / sec, etaSec: (total - doneTok) / (doneTok / sec) });
        },
        fail: (pp, msg, dead) => { // 该窗口交给其他线程重做；WebGPU 一旦出错就不再使用
          lastErr = msg;
          if (++tries[wi] > 3) { cur = null; reject(new Error(`窗口 ${wi} 多次失败：${msg}`)); return; }
          retry.push(wi);
          if (!dead && (pp.ep === 'webgpu' || ++pp.fails > 1)) kill(pp, msg);
        },
      };
      p.timer = setTimeout(() => { const job = p.job; p.job = null; kill(p, '推理超时'); job?.fail(p, '推理超时', true); cur?.pump(); }, 90000);
      const ids = win.ids.slice(); // 窗口本身要留着以便失败重做，传出去的是副本
      p.w.postMessage({ t: 'run', id: wi, ids }, [ids.buffer]);
    };
    cur = {
      pump: () => {
        if (signal?.aborted) return abort();
        for (const p of pool) {
          if (p.state !== 'idle') continue;
          const wi = retry.length ? retry.shift() : next < windows.length ? next++ : -1;
          if (wi < 0) break;
          send(p, wi);
        }
        if (finished === windows.length) { cur = null; resolve(); }
        else if (!pool.some((p) => p.state !== 'dead')) { cur = null; reject(new Error(lastErr || '所有推理线程均已失败')); }
      },
    };
    signal?.addEventListener('abort', () => cur?.pump());
    cur.pump();
  });
  // 汇总：按行写缓存并输出
  const byLine = new Map();
  for (const win of windows) for (const e of win.ents) { if (!byLine.has(e[0])) byLine.set(e[0], []); byLine.get(e[0]).push(e); }
  for (const it of items) {
    const ents = (byLine.get(it.i) || []).map(([, s, e, type, score]) => [s, e, type, score]);
    cache.set(it.key, ents);
    emit(it.i, it.t, ents);
  }
  return out;
}
