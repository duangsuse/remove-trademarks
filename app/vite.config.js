import { defineConfig } from 'vite';
import { viteSingleFile } from 'vite-plugin-singlefile';
import { readFileSync } from 'node:fs';
import { gzipSync } from 'node:zlib';

// 构建后把 ort 的 wasm（gzip+base64）塞进 HTML，运行时用 DecompressionStream 解压
const injectWasm = () => ({
  name: 'inject-ort-wasm',
  enforce: 'post',
  generateBundle(_, bundle) {
    const wasm = readFileSync(new URL('./node_modules/onnxruntime-web/dist/ort-wasm-simd-threaded.asyncify.wasm', import.meta.url));
    const b64 = gzipSync(wasm, { level: 9 }).toString('base64');
    for (const f of Object.values(bundle)) {
      if (f.type === 'asset' && f.fileName.endsWith('.html')) {
        f.source = String(f.source).replace('</body>', () => `<script id="ort-wasm-gz" type="text/plain">${b64}</script></body>`);
      }
    }
  },
});

// 去掉 ort 对 wasm 的 new URL() 引用，避免被 singlefile 再内联两份
const stripOrtWasmUrl = () => ({
  name: 'strip-ort-wasm-url',
  enforce: 'pre',
  transform(code, id) {
    if (!id.includes('onnxruntime-web')) return null;
    const out = code.replace(/new URL\(("ort-wasm[^"]*\.wasm"),import\.meta\.url\)/g, '({href:$1})');
    return out === code ? null : { code: out, map: null };
  },
});

export default defineConfig({
  plugins: [stripOrtWasmUrl(), viteSingleFile(), injectWasm()],
  build: { target: 'es2022', chunkSizeWarningLimit: 99999, assetsInlineLimit: 1e9 },
});
