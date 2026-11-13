// md -> txt 投影：去掉转义反斜杠、粗斜体 * 后的纯文字，带回 md 坐标的映射。
// 规则 / 黑名单 / NER 都在纯文本上匹配（「张**三**」也能命中「张三」），命中区间再换回 md 坐标（中间的标记一并包进去）
import { diffLines, diffChars } from 'diff';

// 去掉 Markdown 标记（转义反斜杠、粗斜体 *）后的文字，像 innerText
export const unmark = (s) => s.replace(/\\([\s\S])|\*+/g, (m, c) => c ?? '');

// text：纯文字；toMd(s, e)：纯文本区间 -> md 区间；at(pos)：md 位置 -> 纯文本位置；slice(s, e)：md 区间 -> 纯文字
let viewOf = { md: null, v: null };
export function plainView(md) {
  if (viewOf.md === md) return viewOf.v;
  const out = [], idx = new Int32Array(md.length + 1), esc = new Uint8Array(md.length + 1); // esc[k]：text[k] 在 md 里带转义反斜杠
  for (let i = 0; i < md.length; i++) {
    const c = md[i];
    if (c === '\\' && i + 1 < md.length) { esc[out.length] = 1; idx[out.length] = ++i; out.push(md[i]); }
    else if (c !== '*') { idx[out.length] = i; out.push(c); }
  }
  const n = out.length;
  idx[n] = md.length;
  const at = (pos) => { let a = 0, b = n; while (a < b) { const m = (a + b) >> 1; if (idx[m] < pos) a = m + 1; else b = m; } return a; };
  const v = {
    text: out.join(''),
    toMd: (s, e) => [idx[s] - esc[s], e > s ? idx[e - 1] + 1 : idx[s] - esc[s]], // 以转义字符开头时连反斜杠一起
    at,
    slice: (s, e) => v.text.slice(at(s), at(e)),
  };
  viewOf = { md, v };
  return v;
}
export const plainOf = (md, s, e) => plainView(md).slice(s, e);

// 对 md 应用区间替换 [{s, e, text}]（已按 s 排序、不重叠）。区间里的粗斜体标记原样留下，
// 否则「张**三**」命中「张**三」时会剩下落单的 **
export function spliceMd(md, reps) {
  let out = '', p = 0;
  for (const x of reps) { out += md.slice(p, x.s) + x.text + md.slice(x.s, x.e).replace(/\\[\s\S]|[^*]/g, ''); p = x.e; }
  return out + md.slice(p);
}

// 第 a~b 行（从 1 起，含 b）的 md 区间 [s, e]，可作 editText 的 range
export function lineRange(md, a, b = a) {
  const starts = [0];
  for (let i = md.indexOf('\n'); i >= 0; i = md.indexOf('\n', i + 1)) starts.push(i + 1);
  const s = starts[Math.max(1, a) - 1] ?? md.length, e = b < starts.length ? starts[b] - 1 : md.length;
  return [s, Math.max(s, e)];
}

// a -> b 的差异，作为 a 上的区间替换 [{s, e, text}]（先按行再按字，超长文本也不慢）
export function textDiff(a, b) {
  const out = [];
  const put = (s, e, text) => { const l = out.at(-1); if (l && l.e === s) { l.e = e; l.text += text; } else if (e > s || text) out.push({ s, e, text }); };
  let pos = 0;
  const chars = (x, y) => {
    for (const p of diffChars(x, y)) {
      if (p.added) put(pos, pos, p.value);
      else if (p.removed) { put(pos, pos + p.value.length, ''); pos += p.value.length; }
      else pos += p.value.length;
    }
  };
  const parts = diffLines(a, b);
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i], nx = parts[i + 1];
    if (p.removed && nx?.added) { chars(p.value, nx.value); i++; }
    else if (p.added) put(pos, pos, p.value);
    else if (p.removed) { put(pos, pos + p.value.length, ''); pos += p.value.length; }
    else pos += p.value.length;
  }
  return out;
}

// 用函数编辑纯文本投影：fn(纯文字) -> 新纯文字，差异换回 md 区间 [{s, e, text}]（交给 DocxProjection.replace）。
// range=[s, e]（md 坐标，如选区或 lineRange）时只把这一段交给 fn
export function plainEdits(md, fn, range = [0, md.length]) {
  const v = plainView(md), a = v.at(range[0]), b = v.at(range[1]);
  const before = v.text.slice(a, b), after = fn(before);
  if (typeof after !== 'string') throw new TypeError('editText：fn 须返回字符串');
  return textDiff(before, after).map(({ s, e, text }) => {
    const [ms, me] = v.toMd(a + s, a + e);
    return { s: ms, e: me, text };
  });
}
