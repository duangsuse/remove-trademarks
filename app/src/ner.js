// 离线 NER：onnxruntime-web (WebGPU 优先，失败回退 WASM) + 内置 bert-base-chinese 词表
import * as ort from 'onnxruntime-web/webgpu';
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

export const env = ort.env;
let session = null;
export let backend = '';

export async function loadWasmBinary() {
  const el = document.getElementById('ort-wasm-gz');
  if (!el) return null; // 开发模式：回退到默认路径
  const bin = Uint8Array.from(atob(el.textContent.trim()), (c) => c.charCodeAt(0));
  const ds = new Response(new Blob([bin]).stream().pipeThrough(new DecompressionStream('gzip')));
  return new Uint8Array(await ds.arrayBuffer());
}

async function hasWebGPU() {
  try { return !!(navigator.gpu && (await navigator.gpu.requestAdapter())); } catch { return false; }
}

export async function loadModel(buffer, onStatus = () => {}) {
  onStatus('解压运行时…');
  const wasm = await loadWasmBinary();
  if (wasm) ort.env.wasm.wasmBinary = wasm;
  ort.env.wasm.numThreads = 1;
  ort.env.wasm.proxy = false;
  const bytes = new Uint8Array(buffer);
  const order = (await hasWebGPU()) ? ['webgpu', 'wasm'] : ['wasm'];
  let lastErr;
  for (const ep of order) {
    try {
      onStatus(`创建会话（${ep === 'webgpu' ? 'WebGPU' : 'WASM'}）…`);
      const s = await ort.InferenceSession.create(bytes, { executionProviders: [ep], graphOptimizationLevel: 'all' });
      // 冒烟测试：WebGPU 个别算子失败则回退
      session = s; backend = ep === 'webgpu' ? 'WebGPU' : 'WASM';
      await runIds([[CLS, vocab.get('中'), SEP]]);
      return backend;
    } catch (e) { lastErr = e; session = null; console.warn(ep, e); }
  }
  throw lastErr;
}

export const ready = () => !!session;

async function runIds(seqs) {
  const L = Math.max(...seqs.map((s) => s.length)), B = seqs.length;
  const ids = new BigInt64Array(B * L), mask = new BigInt64Array(B * L), tt = new BigInt64Array(B * L);
  seqs.forEach((s, b) => s.forEach((v, j) => { ids[b * L + j] = BigInt(v); mask[b * L + j] = 1n; }));
  const feeds = {};
  for (const name of session.inputNames) {
    const data = name === 'input_ids' ? ids : name === 'attention_mask' ? mask : tt;
    feeds[name] = new ort.Tensor('int64', data, [B, L]);
  }
  const out = await session.run(feeds);
  const t = out[session.outputNames[0]];
  return { logits: t.data, L, C: t.dims[2] };
}

function softmaxArg(d, o, C) {
  let m = -1e9, k = 0;
  for (let i = 0; i < C; i++) if (d[o + i] > m) { m = d[o + i]; k = i; }
  let sum = 0;
  for (let i = 0; i < C; i++) sum += Math.exp(d[o + i] - m);
  return [k, 1 / sum];
}

// 对多行文本做 NER，返回 [{text, type, score, line, s, e}]；按 ~400 token 打包相邻行
export async function recognize(lines, onProgress = () => {}) {
  const items = lines.map((t, i) => ({ i, t, tok: tokenize(t) })).filter((x) => x.tok.length);
  const windows = [];
  let w = null;
  for (const it of items) {
    for (let k = 0; k < it.tok.length; k += 400) {
      const part = it.tok.slice(k, k + 400);
      if (!w || w.n + part.length + 1 > 440) { w = { n: 0, segs: [] }; windows.push(w); }
      w.segs.push({ line: it.i, tok: part }); w.n += part.length + 1;
    }
  }
  const res = [];
  for (let wi = 0; wi < windows.length; wi++) {
    const win = windows[wi];
    const seq = [CLS], meta = [null];
    for (const sg of win.segs) { for (const t of sg.tok) { seq.push(t.id); meta.push({ line: sg.line, t }); } seq.push(SEP); meta.push(null); }
    const { logits, L, C } = await runIds([seq]);
    let span = null;
    const close = () => {
      if (span) {
        const text = lines[span.line].slice(span.s, span.e);
        res.push({ text, type: span.type, score: span.sum / span.n, line: span.line, s: span.s, e: span.e });
      }
      span = null;
    };
    for (let j = 0; j < L; j++) {
      const m = meta[j];
      if (!m) { close(); continue; }
      const [k, p] = softmaxArg(logits, j * C, C);
      const lab = LABELS[k];
      if (lab === 'O') { close(); continue; }
      const [bio, type] = [lab[0], lab.slice(2)];
      if (bio === 'B' || bio === 'S' || !span || span.type !== type || span.line !== m.line) { close(); span = { type, line: m.line, s: m.t.s, e: m.t.e, sum: 0, n: 0 }; }
      span.e = m.t.e; span.sum += p; span.n++;
      if (bio === 'E' || bio === 'S') close();
    }
    close();
    onProgress((wi + 1) / windows.length);
  }
  return res;
}
