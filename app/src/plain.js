// md -> txt 投影：去掉转义反斜杠、粗斜体 * 后的纯文字，带回 md 坐标的映射。
// 规则 / 黑名单 / NER 都在纯文本上匹配（「张**三**」也能命中「张三」），命中区间再换回 md 坐标（中间的标记一并包进去）
import { diffLines, diffChars, diffArrays } from 'diff';
import { mdLines, MARK } from './markup.js';

// 单元格内换行在 md 里写作 <br>，纯文本里是 1 个字符：不能是 \n（会让纯文本行号与 md 对不上），用行分隔符，\s 能匹配、. 不跨它
export const BR = '\u2028';

// 去掉 Markdown 标记（转义反斜杠、粗斜体 *）后的文字，像 innerText
export const unmark = (s) => s.replace(/\\([\s\S])|\*+|<br>/g, (m, c) => (m === '<br>' ? BR : c ?? ''));

// text：纯文字；lines：逐行结构（mdLines，按 md 判定再换成纯文本坐标——转义掉的 | / --- 等不会被当成结构）；toMd(s, e)：纯文本区间 -> md 区间；fromMd(pos)：md 位置 -> 纯文本位置
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
  const fromMd = (pos) => { let a = 0, b = n; while (a < b) { const m = (a + b) >> 1; if (idx[m] < pos) a = m + 1; else b = m; } return a; };
  const v = {
    text: out.join(''),
    get lines() { return (this._lines ||= mdLines(md, marks).map((l) => { const s = fromMd(l.s); return { ...l, s, text: this.text.slice(s, fromMd(l.s + l.text.length)) }; })); },
    toMd: (s, e) => [idx[s] - esc[s], e > s ? end[e - 1] : idx[s] - esc[s]], // 以转义字符开头时连反斜杠一起
    fromMd,
  };
  viewOf = { md, marks, v };
  return v;
}
// md 区间 -> 纯文字
export const plainOf = (md, s, e) => { const v = plainView(md); return v.text.slice(v.fromMd(s), v.fromMd(e)); };

// 对 md 应用区间替换 [{s, e, text}]（已按 s 排序、不重叠）。区间里的粗斜体标记原样留下，
// 否则「张**三**」命中「张**三」时会剩下落单的 **
export function spliceMd(md, reps) {
  let out = '', p = 0;
  for (const x of reps) { out += md.slice(p, x.s) + x.text + md.slice(x.s, x.e).replace(/\\[\s\S]|[^*]/g, ''); p = x.e; }
  return out + md.slice(p);
}

// md 每个字的粗斜体状态（位 1 粗、位 2 斜），按 * 标记开合推算；换行处归零。
// 投影在格式切换处写「收 + 开」两个标记（如 粗→斜 写 ** + *），所以连着 n 个 * 翻转的是 [无, 斜, 粗, 粗斜, 粗, 斜][n % 6]
const FLIP = [0, 2, 1, 3, 1, 2];
export function emphasisOf(md) {
  const out = new Uint8Array(md.length);
  let f = 0;
  for (let i = 0; i < md.length; i++) {
    const c = md[i];
    if (c === '\\' && i + 1 < md.length) { out[i] = out[i + 1] = f; i++; }
    else if (c === '*') { let j = i; while (md[j] === '*') j++; f ^= FLIP[(j - i) % 6]; for (; i < j; i++) out[i] = f; i--; }
    else { if (c === '\n') f = 0; out[i] = f; }
  }
  return out;
}

// 第 a~b 行（从 1 起，含 b）的 md 区间 [s, e]，可作 editText 的 range
export function lineRange(md, a, b = a) {
  const starts = [0];
  for (let i = md.indexOf('\n'); i >= 0; i = md.indexOf('\n', i + 1)) starts.push(i + 1);
  const s = starts[Math.max(1, a) - 1] ?? md.length, e = b < starts.length ? starts[b] - 1 : md.length;
  return [s, Math.max(s, e)];
}

// a -> b 的差异，作为 a 上的区间替换 [{s, e, text, at}]：at 是 text 在 b 里的起点；相邻的删除 + 插入合成一段替换。
// 先按行再按字；超长的改动段先去掉公共首尾，仍太长就整段替换。
// tokens(s)：把改动段切成词（如整张图片一个词）再比，默认逐字；same(x, y)：一删一插的两个词还能逐字比（如同一张图片只改了说明）
export function diffSpans(a, b, { tokens = null, same = () => false } = {}) {
  const out = [];
  let pos = 0, bp = 0;
  const put = (s, e, text) => { const l = out.at(-1); if (l && l.e === s) { l.e = e; l.text += text; } else if (e > s || text) out.push({ s, e, text, at: bp }); };
  const del = (n) => { put(pos, pos + n, ''); pos += n; };
  const ins = (t) => { put(pos, pos, t); bp += t.length; };
  const chars = (x, y) => {
    if (x.length * y.length > 4e6) {
      let i = 0, j = 0;
      while (i < x.length && i < y.length && x[i] === y[i]) i++;
      while (j < x.length - i && j < y.length - i && x[x.length - 1 - j] === y[y.length - 1 - j]) j++;
      pos += i; bp += i;
      x = x.slice(i, x.length - j); y = y.slice(i, y.length - j);
      if (x.length * y.length > 4e6) { if (x) del(x.length); if (y) ins(y); pos += j; bp += j; return; }
      chars(x, y); pos += j; bp += j; return;
    }
    if (!tokens) { walk(diffChars(x, y).map((p) => ({ ...p, value: [p.value] }))); return; }
    walk(diffArrays(tokens(x), tokens(y)), (r, d) => r.length === 1 && d.length === 1 && same(r[0], d[0]) && (byChar(r[0], d[0]), true));
  };
  const byChar = (x, y) => walk(diffChars(x, y).map((p) => ({ ...p, value: [p.value] })));
  // parts 的 value 是词数组；refine(删的词, 插的词) 返回真表示已自行处理
  function walk(parts, refine) {
    for (let i = 0; i < parts.length; i++) {
      const p = parts[i], v = p.value.join('');
      if (refine && p.removed && parts[i + 1]?.added && refine(p.value, parts[i + 1].value)) { i++; continue; }
      if (p.added) ins(v);
      else if (p.removed) del(v.length);
      else { pos += v.length; bp += v.length; }
    }
  }
  // 按行比出改动块再逐字比。夹在两个改动块之间的空行不当锚点（md 里每两段之间都有空行），并进同一块，
  // 否则拆段「一二三四」->「一二 / 三四」会被比成删掉「三四」再插入一遍
  const parts = diffLines(a, b);
  let ga = '', gb = '';
  const flush = () => { if (!ga) { if (gb) ins(gb); } else if (!gb) del(ga.length); else chars(ga, gb); ga = gb = ''; };
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i], nx = parts[i + 1];
    if (p.added) gb += p.value;
    else if (p.removed) ga += p.value;
    else if ((ga || gb) && /^\n+$/.test(p.value) && (nx?.added || nx?.removed)) { ga += p.value; gb += p.value; }
    else { flush(); pos += p.value.length; bp += p.value.length; }
  }
  flush();
  return out;
}

// 用函数编辑纯文本投影：fn(纯文字) -> 新纯文字，差异换回 md 区间 [{s, e, text}]（交给 DocxProjection.splice）。
// range=[s, e]（md 坐标，如选区或 lineRange）时只把这一段交给 fn
export function plainToSpans(md, fn, range = [0, md.length], marks = MARK) {
  const v = plainView(md, marks), a = v.fromMd(range[0]), b = v.fromMd(range[1]);
  const before = v.text.slice(a, b), after = fn(before);
  if (typeof after !== 'string') throw new TypeError('mapPlain：fn 须返回字符串');
  return diffSpans(before, after).map(({ s, e, text }) => {
    const [ms, me] = v.toMd(a + s, a + e);
    return { s: ms, e: me, text };
  });
}
