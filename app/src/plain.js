// md -> txt 投影：去掉转义反斜杠、粗斜体 * 后的纯文字，带回 md 坐标的映射。
// 规则 / 黑名单 / NER 都在纯文本上匹配（「张**三**」也能命中「张三」），命中区间再换回 md 坐标（中间的标记一并包进去）
import { diffLines, diffChars } from 'diff';
import { readLines, MARK } from './markup.js';

// 单元格内换行在 md 里写作 <br>，纯文本里是 1 个字符：不能是 \n（会让纯文本行号与 md 对不上），用行分隔符，\s 能匹配、. 不跨它
export const BR = '\u2028';

// 去掉 Markdown 标记（转义反斜杠、粗斜体 *）后的文字，像 innerText
export const unmark = (s) => s.replace(/\\([\s\S])|\*+|<br>/g, (m, c) => (m === '<br>' ? BR : c ?? ''));

// text：纯文字；lines：逐行结构（readLines，按 md 判定再换成纯文本坐标——转义掉的 | / --- 等不会被当成结构）；toMd(s, e)：纯文本区间 -> md 区间；at(pos)：md 位置 -> 纯文本位置；slice(s, e)：md 区间 -> 纯文字
// marks：投影用的结构标记（DocxProjection 的 opts.marks），决定 lines 怎么分
let viewOf = { md: null, marks: null, v: null };
export function plainView(md, marks = MARK) {
  if (viewOf.md === md && viewOf.marks === marks) return viewOf.v;
  const out = [], idx = new Int32Array(md.length + 1), end = new Int32Array(md.length + 1), esc = new Uint8Array(md.length + 1); // esc[k]：text[k] 在 md 里带转义反斜杠
  for (let i = 0; i < md.length; i++) {
    const c = md[i];
    if (c === '\\' && i + 1 < md.length) { esc[out.length] = 1; idx[out.length] = i + 1; end[out.length] = i + 2; out.push(md[++i]); }
    else if (c === '<' && md.startsWith('<br>', i)) { idx[out.length] = i; end[out.length] = i + 4; out.push(BR); i += 3; }
    else if (c !== '*') { idx[out.length] = i; end[out.length] = i + 1; out.push(c); }
  }
  const n = out.length;
  idx[n] = md.length;
  const at = (pos) => { let a = 0, b = n; while (a < b) { const m = (a + b) >> 1; if (idx[m] < pos) a = m + 1; else b = m; } return a; };
  const v = {
    text: out.join(''),
    get lines() { return (this._lines ||= readLines(md, marks).map((l) => { const s = at(l.s); return { ...l, s, text: this.text.slice(s, at(l.s + l.text.length)) }; })); },
    toMd: (s, e) => [idx[s] - esc[s], e > s ? end[e - 1] : idx[s] - esc[s]], // 以转义字符开头时连反斜杠一起
    at,
    slice: (s, e) => v.text.slice(at(s), at(e)),
  };
  viewOf = { md, marks, v };
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
export function plainEdits(md, fn, range = [0, md.length], marks = MARK) {
  const v = plainView(md, marks), a = v.at(range[0]), b = v.at(range[1]);
  const before = v.text.slice(a, b), after = fn(before);
  if (typeof after !== 'string') throw new TypeError('editText：fn 须返回字符串');
  return textDiff(before, after).map(({ s, e, text }) => {
    const [ms, me] = v.toMd(a + s, a + e);
    return { s: ms, e: me, text };
  });
}
