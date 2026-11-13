// NER 推理线程：每个线程一个独立的 onnxruntime 会话。
// file:// 页面没有跨源隔离（无 SharedArrayBuffer），WASM 无法多线程，所以靠多个 Worker 并行。
import * as ort from 'onnxruntime-web/webgpu';

let session = null;
const keep = []; // 持有 GPU 对象的强引用，避免被回收后出现 "A valid external Instance reference no longer exists"

async function infer(ids) {
  const L = ids.length;
  const a = new BigInt64Array(L), mask = new BigInt64Array(L).fill(1n), tt = new BigInt64Array(L);
  for (let i = 0; i < L; i++) a[i] = BigInt(ids[i]);
  const feeds = {};
  for (const name of session.inputNames) feeds[name] = new ort.Tensor('int64', name === 'input_ids' ? a : name === 'attention_mask' ? mask : tt, [1, L]);
  const out = await session.run(feeds);
  const t = out[session.outputNames[0]], C = t.dims[2], d = t.data;
  const lab = new Uint8Array(L), prob = new Float32Array(L);
  for (let j = 0; j < L; j++) { // 每个位置 argmax + softmax 概率，只回传这两项（比整张 logits 小 ~70 倍）
    const o = j * C;
    let m = -Infinity, k = 0;
    for (let i = 0; i < C; i++) if (d[o + i] > m) { m = d[o + i]; k = i; }
    let sum = 0;
    for (let i = 0; i < C; i++) sum += Math.exp(d[o + i] - m);
    lab[j] = k; prob[j] = 1 / sum;
  }
  t.dispose?.();
  return { lab, prob };
}

self.onmessage = async ({ data: m }) => {
  try {
    if (m.t === 'init') {
      if (m.wasm) ort.env.wasm.wasmBinary = m.wasm;
      ort.env.wasm.numThreads = 1;
      ort.env.wasm.proxy = false;
      if (m.ep === 'webgpu') {
        const adapter = await navigator.gpu?.requestAdapter({ powerPreference: 'high-performance' });
        if (!adapter) throw new Error('没有可用的 WebGPU 适配器');
        keep.push(adapter);
      }
      const bytes = new Uint8Array(await m.model.arrayBuffer());
      session = await ort.InferenceSession.create(bytes, { executionProviders: [m.ep], graphOptimizationLevel: 'all' });
      await infer([101, 704, 102]); // 冒烟测试（[CLS] 中 [SEP]）
      self.postMessage({ t: 'ready' });
    } else if (m.t === 'run') {
      const { lab, prob } = await infer(m.ids);
      self.postMessage({ t: 'done', id: m.id, lab, prob }, [lab.buffer, prob.buffer]);
    }
  } catch (e) {
    self.postMessage({ t: m.t === 'run' ? 'fail' : 'error', id: m.id, msg: String(e?.message || e) });
  }
};
