// Word <-> Markdown 1:1 投影；差异以修订（w:del / w:ins）写回
// 用法：const d = await DocxProjection.load(buf); d.md / d.anchor(i)；d.replace(spans) 或 d.applyText(newMd)；d.dropImages / d.wipe；await d.save()
import JSZip from 'jszip';
import { diffLines, diffChars } from 'diff';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS = 'http://www.w3.org/XML/1998/namespace';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const V_NS = 'urn:schemas-microsoft-com:vml';
const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const isW = (n, name) => n.nodeType === 1 && n.namespaceURI === NS && n.localName === name;
const kids = (n, name) => (n ? [...n.childNodes].filter((c) => isW(c, name)) : []);
const kid = (n, name) => (n ? [...n.childNodes].find((c) => isW(c, name)) : undefined);
const val = (el) => el?.getAttributeNS(NS, 'val') ?? null;
const on = (el) => !!el && !['0', 'false', 'off'].includes(val(el) ?? '1');
const inFallback = (el) => { for (let n = el.parentNode; n && n.nodeType === 1; n = n.parentNode) if (n.localName === 'Fallback') return true; return false; };
// root 下第一个 ns:name，不进文本框（文本框里的图片属于文本框自己的段落）
function own(root, ns, name) {
  for (const el of root.getElementsByTagNameNS(ns, name)) {
    let n = el.parentNode;
    while (n && n !== root && n.localName !== 'txbxContent') n = n.parentNode;
    if (n === root) return el;
  }
  return null;
}

// Markdown 里的结构标记（detect.js 据此识别特殊表格 / 其他文本区）
export const MARK = { special: ':::特殊表格', end: ':::', others: '---\n\n## 其他文本' };
// 本工具写入的审阅修订的作者；重新载入时据此先「拒绝」掉
export const TOOL_AUTHOR = '脱敏助手';

// ---------- 样式 / 编号 ----------
// 返回 look(styleId, f)：沿 basedOn 链找第一个 f(style) 不为空的值
function styleLookup(doc) {
  const map = new Map();
  for (const s of doc ? doc.getElementsByTagNameNS(NS, 'style') : []) {
    const pPr = kid(s, 'pPr');
    map.set(s.getAttributeNS(NS, 'styleId'), { name: (val(kid(s, 'name')) || '').toLowerCase(), based: val(kid(s, 'basedOn')), outline: val(kid(pPr, 'outlineLvl')), numPr: kid(pPr, 'numPr'), rPr: kid(s, 'rPr') });
  }
  return (id, f) => {
    for (let k = 0; id && k < 16; k++) {
      const s = map.get(id);
      if (!s) return undefined;
      const v = f(s);
      if (v != null) return v;
      id = s.based;
    }
    return undefined;
  };
}

// 中文 Word 的标题样式 ID 是 "1"/"2"…，要看样式名（heading n）或大纲级别
function headingLevel(p, style) {
  const pPr = kid(p, 'pPr');
  const lvl = (o) => (o != null && /^\d$/.test(o) ? +o + 1 : undefined); // 大纲级别 9 = 正文
  const st = val(kid(pPr, 'pStyle'));
  const lv = lvl(val(kid(pPr, 'outlineLvl')))
    ?? style(st, (s) => { const m = /^(?:heading|标题)\s*([1-9])$/.exec(s.name); return m ? +m[1] : lvl(s.outline); })
    ?? (/^(?:heading|标题)\s*([1-6])$/i.exec(st || '')?.[1] | 0);
  return lv >= 1 && lv <= 6 ? lv : 0;
}

const CN = '零一二三四五六七八九';
const cnNum = (n) => (n < 10 ? CN[n] : n < 20 ? '十' + (n % 10 ? CN[n % 10] : '') : n < 100 ? CN[(n / 10) | 0] + '十' + (n % 10 ? CN[n % 10] : '') : String(n));
const letters = (n) => { let s = ''; for (; n > 0; n = ((n - 1) / 26) | 0) s = String.fromCharCode(97 + ((n - 1) % 26)) + s; return s; };
const roman = (n) => { let s = ''; for (const [v, r] of [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]) for (; n >= v; n -= v) s += r; return s; };
const NUM_FMT = {
  decimal: String, decimalZero: (n) => String(n).padStart(2, '0'), lowerLetter: letters, upperLetter: (n) => letters(n).toUpperCase(),
  lowerRoman: roman, upperRoman: (n) => roman(n).toUpperCase(), chineseCounting: cnNum, chineseCountingThousand: cnNum, chineseLegalSimplified: cnNum, japaneseCounting: cnNum,
  ideographTraditional: (n) => '甲乙丙丁戊己庚辛壬癸'[n - 1] ?? String(n), ideographZodiac: (n) => '子丑寅卯辰巳午未申酉戌亥'[n - 1] ?? String(n),
  decimalEnclosedCircle: (n) => (n >= 1 && n <= 20 ? String.fromCodePoint(0x245f + n) : String(n)), decimalEnclosedCircleChinese: (n) => (n >= 1 && n <= 20 ? String.fromCodePoint(0x245f + n) : String(n)),
  decimalEnclosedParen: (n) => (n >= 1 && n <= 20 ? String.fromCodePoint(0x2473 + n) : `(${n})`), decimalFullWidth: (n) => String(n).replace(/\d/g, (d) => String.fromCharCode(0xff10 + +d)),
};

// 自动编号：按文档顺序计数，label(p) 返回 {ilvl, text}（项目符号为 '-'）或 null
class Lists {
  constructor(doc, style) {
    this.style = style; this.abs = new Map(); this.nums = new Map(); this.cnt = new Map();
    if (!doc) return;
    for (const a of doc.getElementsByTagNameNS(NS, 'abstractNum')) {
      const lv = new Map();
      for (const l of kids(a, 'lvl')) lv.set(+l.getAttributeNS(NS, 'ilvl'), { fmt: val(kid(l, 'numFmt')) || 'decimal', text: val(kid(l, 'lvlText')) ?? '', start: +(val(kid(l, 'start')) ?? 1) });
      this.abs.set(a.getAttributeNS(NS, 'abstractNumId'), lv);
    }
    for (const n of doc.getElementsByTagNameNS(NS, 'num')) {
      const start = new Map();
      for (const o of kids(n, 'lvlOverride')) { const so = val(kid(o, 'startOverride')); if (so != null) start.set(+o.getAttributeNS(NS, 'ilvl'), +so); }
      this.nums.set(n.getAttributeNS(NS, 'numId'), { abs: val(kid(n, 'abstractNumId')), start });
    }
  }

  label(p) {
    const pPr = kid(p, 'pPr'), np = kid(pPr, 'numPr');
    let numId = val(kid(np, 'numId')), ilvl = val(kid(np, 'ilvl'));
    if (numId == null) {
      const snp = this.style(val(kid(pPr, 'pStyle')), (s) => s.numPr);
      if (snp) { numId = val(kid(snp, 'numId')); ilvl ??= val(kid(snp, 'ilvl')); }
    }
    const num = numId && numId !== '0' ? this.nums.get(numId) : null;
    const lvls = num && this.abs.get(num.abs);
    ilvl = +(ilvl || 0);
    const lv = lvls?.get(ilvl);
    if (!lv) return null;
    const key = num.start.size ? 'n' + numId : 'a' + num.abs; // 有起始值覆盖的编号单独计数，否则同一抽象编号连续
    let c = this.cnt.get(key);
    if (!c) this.cnt.set(key, (c = []));
    const startOf = (k) => num.start.get(k) ?? lvls.get(k)?.start ?? 1;
    c[ilvl] = c[ilvl] == null ? startOf(ilvl) : c[ilvl] + 1;
    c.length = ilvl + 1; // 下级重新计数
    if (lv.fmt === 'bullet') return { ilvl, text: '-' };
    if (lv.fmt === 'none') return null;
    const text = lv.text.replace(/%([1-9])/g, (_, k) => { const i = k - 1, f = NUM_FMT[lvls.get(i)?.fmt] || String; return f(c[i] ?? startOf(i)); }).trim();
    return text ? { ilvl, text } : null;
  }
}

function runFmt(r, style) {
  const p = kid(r, 'rPr');
  if (!p) return '';
  const flag = (n) => { const el = kid(p, n); return el ? on(el) : !!style(val(kid(p, 'rStyle')), (s) => { const e = kid(s.rPr, n); return e ? on(e) : undefined; }); };
  return (flag('b') ? 'b' : '') + (flag('i') ? 'i' : '');
}

// ---------- Markdown 构建 ----------
class MD {
  constructor() { this.out = []; this.pid = []; this.off = []; this.ed = []; }
  push(str, pid, off, ed) { for (let i = 0; i < str.length; i++) { this.out.push(str[i]); this.pid.push(pid); this.off.push(off); this.ed.push(ed); } }
  s(str) { this.push(str, -1, 0, 0); }
  sp(str, para, off) { this.push(str, para.idx, off, 0); }
  cut(n) { this.out.length = this.pid.length = this.off.length = this.ed.length = n; }
  get length() { return this.out.length; }
}
const marker = (f) => (f === 'bi' ? '***' : f === 'b' ? '**' : '*');
const ESC = new Set(['*', '\\', '|']);
const AMBIG = /^(?:#|\||:::|---|- )/; // 行首会被误认成标题 / 表格 / 分隔线 / 列表的正文，前面加 \
const head3 = (atoms) => atoms.slice(0, 3).map((a) => a.ch).join('');

// 文本单元（链接地址、图片说明、图表文字…）：逐字可编辑
function emitUnit(md, u, lineStart = false) {
  if (lineStart && AMBIG.test(head3(u.atoms))) md.push('\\', u.idx, 0, 0);
  u.atoms.forEach((a, i) => { if (ESC.has(a.ch)) md.push('\\', u.idx, i, 0); md.push(a.ch, u.idx, i, 1); });
}

function emitObj(md, para, i, o) {
  md.sp('![', para, i);
  if (o.alt) emitUnit(md, o.alt); else md.sp(o.label, para, i);
  md.sp('](' + o.path, para, i + 1);
  if (o.title) { md.sp(' "', para, i + 1); emitUnit(md, o.title); md.sp('"', para, i + 1); }
  md.sp(')', para, i + 1);
}

function emitPara(md, para, prefix = '', lineStart = false) {
  const { atoms } = para;
  md.s(prefix);
  if (lineStart && !prefix && !atoms[0].link && AMBIG.test(head3(atoms))) md.sp('\\', para, 0);
  // 空白 / 制表 / 换行不单独成粗斜体段：沿用前一个（开头则取后一个）非空白的格式
  const eff = atoms.map((a) => (/^\s$/.test(a.ch) ? null : a.fmt));
  let carry = eff.find((f) => f !== null) ?? '';
  for (let i = 0; i < eff.length; i++) { if (eff[i] === null) eff[i] = carry; else carry = eff[i]; }
  let fmt = '', link = null;
  const closeLink = (i) => { md.sp('](', para, i); emitUnit(md, link.unit); md.sp(')', para, i); };
  atoms.forEach((a, i) => {
    if (a.link !== link) { // 链接边界：先收起粗斜体，避免标记交叉
      if (fmt) md.sp(marker(fmt), para, i);
      fmt = '';
      if (link) closeLink(i);
      if (a.link) md.sp('[', para, i);
      link = a.link;
    }
    if (eff[i] !== fmt) {
      if (fmt) md.sp(marker(fmt), para, i);
      if (eff[i]) md.sp(marker(eff[i]), para, i);
      fmt = eff[i];
    }
    if (a.obj) emitObj(md, para, i, a.obj);
    else if (a.edit) {
      if (ESC.has(a.ch)) md.sp('\\', para, i);
      md.push(a.ch, para.idx, i, 1);
    } else md.sp(a.ch === '\n' ? '<br>' : a.ch, para, i);
  });
  if (fmt) md.sp(marker(fmt), para, atoms.length);
  if (link) closeLink(atoms.length);
}

function plainTable(tbl) {
  const rows = kids(tbl, 'tr');
  if (!rows.length) return null;
  const grid = [];
  let n = -1;
  for (const tr of rows) {
    const trPr = kid(tr, 'trPr');
    if (trPr && (kid(trPr, 'gridBefore') || kid(trPr, 'gridAfter'))) return null;
    const cells = [];
    for (const c of tr.childNodes) {
      if (c.nodeType !== 1) continue;
      if (isW(c, 'tc')) cells.push(c);
      else if (isW(c, 'sdt')) return null;
    }
    if (!cells.length || (n >= 0 && cells.length !== n)) return null;
    n = cells.length;
    const row = [];
    for (const tc of cells) {
      const pr = kid(tc, 'tcPr');
      if (pr && (kid(pr, 'vMerge') || kid(pr, 'hMerge') || (kid(pr, 'gridSpan') && +val(kid(pr, 'gridSpan')) > 1))) return null;
      const ps = kids(tc, 'p');
      const other = [...tc.childNodes].some((x) => x.nodeType === 1 && x.namespaceURI === NS && !['tcPr', 'p'].includes(x.localName));
      if (other || ps.length !== 1) return null;
      row.push(ps[0]);
    }
    grid.push(row);
  }
  return grid;
}

// 本工具上一次写入的修订先「拒绝」掉再投影：删除线里的原文还原成正常文字，插入的内容去掉。
// 否则 Ctrl+Enter 写入审阅修订后，重新载入 A 时原文已经藏在 w:del 里，投影看不到旧内容、黑名单就扫不出来了；
// 同时重复写入审阅不会叠加。别的作者的修订保持原样。旧版作者名 ss 只认本工具用过的 id 段（>= 900000）。返回处理的数量。
const ours = (x) => { const a = x.getAttributeNS(NS, 'author'); return a === TOOL_AUTHOR || (a === 'ss' && +x.getAttributeNS(NS, 'id') >= 900000); };
function rejectOurs(doc) {
  let n = 0;
  const byTag = (t) => [...doc.getElementsByTagNameNS(NS, t)].filter(ours);
  for (const x of byTag('ins')) { if (x.parentNode && x.parentNode.localName !== 'rPr') { x.remove(); n++; } }
  for (const x of byTag('del')) {
    if (!x.parentNode || x.parentNode.localName === 'rPr') continue;
    for (const dt of [...x.getElementsByTagNameNS(NS, 'delText')]) dt.replaceWith(mkText(doc, 't', dt.textContent));
    while (x.firstChild) x.parentNode.insertBefore(x.firstChild, x);
    x.remove(); n++;
  }
  return n;
}

async function loadXml(zip, path, warnings) {
  const f = path && zip.file(path);
  if (!f) return null;
  const doc = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) { warnings.push(`${path} 解析失败，其中的文字没有处理`); return null; }
  return doc;
}
const relsPathOf = (path) => path.replace(/([^/]*)$/, '_rels/$1.rels');
const resolvePath = (base, target) => decodeURIComponent(new URL(target, 'http://x/' + base).pathname.slice(1));
const baseName = (path) => path.split('/').pop().replace(/\.xml$/, '');
const TEXT_PARTS = { header: '页眉', footer: '页脚', footnotes: '脚注', endnotes: '尾注' };
const HYPER = /^(\s*HYPERLINK\s+")([^"]+)(")/i;
const INLINE_WRAP = new Set(['smartTag', 'sdt', 'sdtContent', 'dir', 'bdo', 'customXml']);

// ---------- 编辑记录 ----------
class Edits extends Map {
  of(pid) { let e = this.get(pid); if (!e) this.set(pid, (e = { del: new Set(), ins: new Map() })); return e; }
  del(pid, off) { this.of(pid).del.add(off); }
  ins(pid, at, text) { const e = this.of(pid); e.ins.set(at, (e.ins.get(at) || '') + text); }
}
const unmark = (s) => s.replace(/\\([\s\S])|\*+/g, (m, c) => c ?? '').replace(/<br>/g, '\n');

// diff -> 每段编辑。newMap：新文本也是某个 docx 的投影时传入——插入文字只取其中「可编辑字符」，转义反斜杠 / 粗斜体标记 / 表格竖线 / 换行等结构字符不会混进来
export function computeEdits(origMd, newMd, map, newMap = null, edits = new Edits()) {
  const stats = { del: 0, ins: 0, skipped: 0 };
  const del = (a, b) => {
    let any = false;
    for (let i = a; i < b; i++) if (map.pid[i] >= 0 && map.ed[i]) { edits.del(map.pid[i], map.off[i]); any = true; }
    if (any) stats.del++; else stats.skipped++;
  };
  const ins = (p, text, np, raw) => {
    if (newMap) { let t = ''; for (let i = 0; i < text.length; i++) if (newMap.ed[np + i]) t += text[i]; text = t; }
    else text = (raw ? unmark(text) : text).replace(/\r\n?/g, '\n');
    if (!text) return;
    let pid = -1, at = 0;
    if (p > 0 && map.pid[p - 1] >= 0) { pid = map.pid[p - 1]; at = map.off[p - 1] + map.ed[p - 1]; }
    else if (p < map.pid.length && map.pid[p] >= 0) { pid = map.pid[p]; at = map.off[p]; }
    if (pid < 0) { stats.skipped++; return; }
    edits.ins(pid, at, text);
    stats.ins++;
  };
  if (origMd === newMd) return { byPara: edits, stats };
  const walkChars = (a, b, pos, np) => {
    if (!b) { del(pos, pos + a.length); return; }
    if (!a) { ins(pos, b, np); return; }
    if (a.length * b.length > 4e6) { // 超长行：先去掉公共首尾，仍太长才整段替换（插入文字去掉 Markdown 标记）
      let i = 0, j = 0;
      while (i < a.length && i < b.length && a[i] === b[i]) i++;
      while (j < a.length - i && j < b.length - i && a[a.length - 1 - j] === b[b.length - 1 - j]) j++;
      a = a.slice(i, a.length - j); b = b.slice(i, b.length - j); pos += i; np += i;
      if (!a || !b || a.length * b.length > 4e6) { if (a) del(pos, pos + a.length); if (b) ins(pos + a.length, b, np, true); return; }
    }
    for (const part of diffChars(a, b)) {
      const n = part.value.length;
      if (part.removed) { del(pos, pos + n); pos += n; }
      else if (part.added) { ins(pos, part.value, np); np += n; }
      else { pos += n; np += n; }
    }
  };
  const parts = diffLines(origMd, newMd);
  let pos = 0, np = 0;
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    if (p.removed) {
      const nx = parts[i + 1];
      if (nx && nx.added) { walkChars(p.value, nx.value, pos, np); pos += p.value.length; np += nx.value.length; i++; }
      else { del(pos, pos + p.value.length); pos += p.value.length; }
    } else if (p.added) { ins(pos, p.value, np); np += p.value.length; }
    else { pos += p.value.length; np += p.value.length; }
  }
  return { byPara: edits, stats };
}

// ---------- 写入修订 ----------
function mkEl(doc, name, attrs = {}) {
  const el = doc.createElementNS(NS, 'w:' + name);
  for (const [k, v] of Object.entries(attrs)) el.setAttributeNS(NS, 'w:' + k, v);
  return el;
}
function mkText(doc, name, text) {
  const t = mkEl(doc, name);
  t.setAttributeNS(XMLNS, 'xml:space', 'preserve');
  t.textContent = text;
  return t;
}

// 返回因落在别人修订里而丢掉的插入数
function applyParaEdits(para, ed, ctx) {
  const { atoms } = para;
  const doc = para.el.ownerDocument;
  const need = new Set();
  const lastIdx = atoms.length - 1;
  for (const i of ed.del) if (atoms[i]?.edit) need.add(atoms[i].r);
  for (const at of ed.ins.keys()) need.add(atoms[Math.min(at, lastIdx)].r);
  const rprClones = new Map();
  for (const a of atoms) if (!rprClones.has(a.r)) rprClones.set(a.r, kid(a.r, 'rPr')?.cloneNode(true) || null);
  // 插入文字的格式：替换（该位置的字被删）取被删字的格式，否则取前一个字的
  const rprOfAtom = (gi) => rprClones.get(atoms[ed.del.has(gi) && atoms[gi]?.edit ? gi : Math.max(0, Math.min(gi - 1, lastIdx))].r);
  const meta = () => ({ id: String(ctx.nextId++), author: ctx.author, date: ctx.date });
  const mkRun = (r, rpr, child) => {
    const nr = r.cloneNode(false);
    if (rpr) nr.appendChild(rpr.cloneNode(true));
    nr.appendChild(child);
    return nr;
  };
  // 插入文本里的 \n 写成 <w:br/>（点即换行）
  const insNode = (r, gi, text) => {
    const runs = [];
    text.split('\n').forEach((part, i) => {
      if (i) runs.push(mkRun(r, rprOfAtom(gi), mkEl(doc, 'br')));
      if (part) runs.push(mkRun(r, rprOfAtom(gi), mkText(doc, 't', part)));
    });
    if (!ctx.track) return runs;
    const w = mkEl(doc, 'ins', meta());
    runs.forEach((x) => w.appendChild(x));
    return w;
  };
  const delNode = (r, rpr, text) => {
    if (!ctx.track) return null;
    const w = mkEl(doc, 'del', meta());
    w.appendChild(mkRun(r, rpr, mkText(doc, 'delText', text)));
    return w;
  };
  let dropped = 0;
  for (const r of need) {
    const parent = r.parentNode;
    if (!parent || isW(parent, 'ins')) { for (const at of ed.ins.keys()) if (atoms[Math.min(at, lastIdx)].r === r) dropped++; continue; }
    const rpr = rprClones.get(r);
    const byItem = new Map();
    atoms.forEach((a, gi) => { if (a.r === r) { if (!byItem.has(a.item)) byItem.set(a.item, []); byItem.get(a.item).push([a, gi]); } });
    const out = [];
    const push = (n) => { if (Array.isArray(n)) out.push(...n); else if (n) out.push(n); };
    for (const c of [...r.childNodes]) {
      if (isW(c, 'rPr')) continue;
      const list = byItem.get(c);
      if (!list) { push(mkRun(r, rpr, c.cloneNode(true))); continue; }
      if (!isW(c, 't')) {
        for (const [, gi] of list) {
          if (ed.ins.has(gi)) push(insNode(r, gi, ed.ins.get(gi)));
          push(mkRun(r, rpr, c.cloneNode(true)));
          if (gi === lastIdx && ed.ins.has(atoms.length)) push(insNode(r, atoms.length, ed.ins.get(atoms.length)));
        }
        continue;
      }
      let buf = '', bufDel = false;
      const flush = () => {
        if (!buf) return;
        push(bufDel ? delNode(r, rpr, buf) : mkRun(r, rpr, mkText(doc, 't', buf)));
        buf = '';
      };
      for (const [a, gi] of list) {
        if (ed.ins.has(gi)) { flush(); push(insNode(r, gi, ed.ins.get(gi))); }
        const d = ed.del.has(gi) && a.edit;
        if (d !== bufDel) { flush(); bufDel = d; }
        buf += a.ch;
        if (gi === lastIdx && ed.ins.has(atoms.length)) { flush(); push(insNode(r, atoms.length, ed.ins.get(atoms.length))); }
      }
      flush();
    }
    for (const n of out) parent.insertBefore(n, r);
    parent.removeChild(r);
  }
  return dropped;
}

function applyTextEdits(u, ed) {
  let s = '';
  u.atoms.forEach((a, i) => { s += ed.ins.get(i) ?? ''; if (!ed.del.has(i)) s += a.ch; });
  s = (s + (ed.ins.get(u.atoms.length) ?? '')).replace(/\n/g, ' ');
  u.set(s);
  return s;
}

// 文本框 mc:AlternateContent：投影 / 编辑的是 Choice（wps），Fallback（VML）里还有一份原文——改过的照 Choice 同步过去
function syncFallbacks(edited) {
  const acs = new Set();
  for (const el of edited) for (let n = el.parentNode; n && n.nodeType === 1; n = n.parentNode) if (n.namespaceURI === MC_NS && n.localName === 'Choice') acs.add(n.parentNode);
  const depth = (n) => { let d = 0; for (; n; n = n.parentNode) d++; return d; };
  const outer = (root) => [...root.getElementsByTagNameNS(NS, 'txbxContent')].filter((tb) => { for (let n = tb.parentNode; n && n !== root; n = n.parentNode) if (n.localName === 'txbxContent') return false; return true; });
  for (const ac of [...acs].sort((a, b) => depth(b) - depth(a))) { // 先内层后外层
    const pick = (name) => [...ac.childNodes].find((x) => x.namespaceURI === MC_NS && x.localName === name);
    const choice = pick('Choice'), fb = pick('Fallback');
    if (!choice || !fb) continue;
    const src = outer(choice), dst = outer(fb);
    if (src.length !== dst.length) { fb.remove(); continue; } // 对不上就去掉 Fallback（Word 2010 起只读 Choice）
    src.forEach((s, i) => {
      while (dst[i].firstChild) dst[i].firstChild.remove();
      for (const c of s.childNodes) dst[i].appendChild(c.cloneNode(true));
    });
  }
}

// SmartArt：data 是数据，drawing 是 Word 缓存的排版结果，按原文逐个同步；返回没找到的个数
function mirrorSmartArt(parts, map) {
  if (!parts.length) return 0;
  const hit = new Set();
  for (const part of parts) for (const t of part.doc.getElementsByTagNameNS(A_NS, 't')) {
    if (!map.has(t.textContent)) continue;
    hit.add(t.textContent); t.textContent = map.get(t.textContent); part.dirty = true;
  }
  return [...map.keys()].filter((k) => !hit.has(k)).length;
}

function maxId(parts) {
  let m = 0;
  for (const p of parts) for (const el of p.doc.getElementsByTagName('*')) { const v = el.getAttributeNS(NS, 'id'); if (v && /^\d+$/.test(v)) m = Math.max(m, +v); }
  return m;
}

// ---------- 投影 ----------
export class DocxProjection {
  zip = null;
  parts = new Map(); // path -> {path, doc, dirty}（含 .rels）
  textParts = []; // 正文 + 页眉页脚脚注尾注
  units = []; // 段落 {kind:'para', el, part, atoms} 或文本单元 {kind:'text', part, atoms, set}
  md = '';
  map = null; // md 每个字符 -> {pid 单元, off 单元内位置, ed 可编辑}
  warnings = [];
  stats = { del: 0, ins: 0, skipped: 0 };
  #edits = new Edits();
  #mirrors = [];
  #drop = [];
  #wipe = false;
  #saved = false;

  static async load(buffer) {
    const d = new DocxProjection();
    await d.#load(buffer);
    return d;
  }

  // md 第 i 个字符对应的位置；结构字符（Markdown 标记、表格竖线等）editable 为 false
  anchor(i) {
    const pid = this.map.pid[i];
    if (pid == null || pid < 0) return null;
    const unit = this.units[pid];
    return { unit, atom: unit.atoms[this.map.off[i]] ?? null, editable: this.map.ed[i] === 1 };
  }

  // 按 md 区间替换：区间内可编辑字符全删，text 插在第一个可编辑字符处（用被删字的格式）
  replace(spans) {
    const { pid, off, ed } = this.map, st = { del: 0, ins: 0, skipped: 0 };
    for (const { s, e, text = '' } of spans) {
      let first = -1;
      for (let i = s; i < e; i++) if (ed[i]) { if (first < 0) first = i; this.#edits.del(pid[i], off[i]); }
      if (first < 0) { st.skipped++; continue; }
      st.del++;
      if (text) { this.#edits.ins(pid[first], off[first], text); st.ins++; }
    }
    return this.#count(st);
  }

  // 与新文本做 diff；from：新文本来自另一个 DocxProjection 时传入，结构字符不会写进正文
  applyText(newMd, { from = null } = {}) {
    return this.#count(computeEdits(this.md, newMd, this.map, from?.map, this.#edits).stats);
  }

  dropImages(paths) { this.#drop.push(...paths); return this; }
  wipe() { this.#wipe = true; return this; }

  // track：写成审阅修订（作者 author）；type 同 JSZip.generateAsync。DOM 会被改写，只能调用一次
  async save({ track = false, author = TOOL_AUTHOR, type = 'blob' } = {}) {
    if (this.#saved) throw new Error('DocxProjection.save() 只能调用一次');
    this.#saved = true;
    const warnings = [...this.warnings], stats = { ...this.stats };
    const ctx = { author, date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), nextId: maxId(this.textParts) + 1, track };
    const edited = [], mirror = new Map();
    let untracked = 0, failed = 0;
    for (const [pid, ed] of this.#edits) {
      const u = this.units[pid];
      if (u.kind === 'text') {
        if (track) { untracked++; continue; } // 属性 / 关系里的文字没有修订语法，写了就清不掉
        const v = applyTextEdits(u, ed);
        if (u.mirror) mirror.set(u.orig, v);
        u.part.dirty = true;
        continue;
      }
      const backup = u.el.cloneNode(true);
      try {
        const n = applyParaEdits(u, ed, ctx);
        stats.ins = Math.max(0, stats.ins - n); stats.skipped += n;
        edited.push(u.el); u.part.dirty = true;
      } catch (e) {
        console.error(e); failed++;
        u.el.replaceWith(backup); // 半途失败的段落整段还原
      }
    }
    if (untracked) warnings.push(`${untracked} 处链接地址 / 图片说明 / 图表文字无法写成修订，未改动`);
    if (failed) { warnings.push(`${failed} 段写入失败（详见控制台），这些段落保持原样`); stats.skipped += failed; }
    syncFallbacks(edited);
    const miss = mirror.size ? mirrorSmartArt(this.#mirrors, mirror) : 0;
    if (miss) warnings.push(`SmartArt 缓存里有 ${miss} 处文字没对上，可能还残留原文（在 Word 里点一下该 SmartArt 会重新排版）`);
    const wiped = new Set(this.#wipe ? this.textParts : []);
    for (const part of wiped) cleanDom(part.doc);
    for (const part of this.parts.values()) {
      if (!part || !(part.dirty || wiped.has(part))) continue;
      let xml = new XMLSerializer().serializeToString(part.doc);
      if (wiped.has(part)) xml = xml.replace(FP_ATTR, '');
      this.zip.file(part.path, /^<\?xml/.test(xml) ? xml : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml);
    }
    for (const path of this.#drop) {
      const ext = path.split('.').pop().toLowerCase();
      const key = ext === 'jpeg' ? 'jpg' : ext;
      if (PLACEHOLDER[key] && this.zip.file(path)) this.zip.file(path, b64(PLACEHOLDER[key]));
    }
    if (this.#wipe) await wipeFingerprints(this.zip);
    const out = await this.zip.generateAsync({ type, compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    return { [type === 'blob' ? 'blob' : 'data']: out, stats, warnings };
  }

  #count(st) { for (const k in st) this.stats[k] += st[k]; return st; }

  async #load(buffer) {
    const zip = (this.zip = await JSZip.loadAsync(buffer));
    const { parts, warnings, units } = this;
    const load = async (path) => {
      if (!path) return null;
      if (!parts.has(path)) { const doc = await loadXml(zip, path, warnings); parts.set(path, doc ? { path, doc, dirty: false } : null); }
      return parts.get(path);
    };
    const relsOf = new Map(); // 部件 path -> Map(rId -> rel)
    const rels = async (path) => {
      if (relsOf.has(path)) return relsOf.get(path);
      const m = new Map();
      relsOf.set(path, m);
      const rp = await load(relsPathOf(path));
      for (const el of rp ? rp.doc.getElementsByTagNameNS(PKG_NS, 'Relationship') : []) {
        const external = el.getAttribute('TargetMode') === 'External', target = el.getAttribute('Target') || '';
        m.set(el.getAttribute('Id'), { el, part: rp, type: (el.getAttribute('Type') || '').split('/').pop(), external, target, path: external ? target : resolvePath(path, target) });
      }
      return m;
    };
    const mainPath = [...(await rels('')).values()].find((r) => r.type === 'officeDocument')?.path || 'word/document.xml';
    const main = await load(mainPath);
    const mainRels = await rels(mainPath);
    const ofType = (t) => [...new Set([...mainRels.values()].filter((r) => r.type === t && !r.external).map((r) => r.path))].sort((a, b) => a.localeCompare(b, 'en', { numeric: true }));
    const style = styleLookup((await load(ofType('styles')[0]))?.doc);
    const lists = new Lists((await load(ofType('numbering')[0]))?.doc, style);
    const others = [];
    for (const t in TEXT_PARTS) for (const p of ofType(t)) { const part = await load(p); if (part) others.push({ type: t, part }); }
    this.textParts = [main, ...others.map((o) => o.part)].filter(Boolean);
    for (const part of this.textParts) { await rels(part.path); if (rejectOurs(part.doc)) part.dirty = true; }

    const unit = (u) => { u.idx = units.length; units.push(u); return u; };
    const textUnit = (part, text, set, extra) => unit({ kind: 'text', part, atoms: [...text].map((ch) => ({ ch, edit: true })), set, ...extra });
    const found = []; // 正文里发现的文本框 / 图表 / SmartArt，放到「其他文本」区
    const linkUnits = new Map();
    const relLink = (h, part) => {
      const r = relsOf.get(part.path)?.get(h.getAttributeNS(R_NS, 'id'));
      if (!r?.external) return null; // 文档内锚点（目录等）不显示成链接
      if (!linkUnits.has(r.el)) linkUnits.set(r.el, { unit: textUnit(r.part, r.target, (v) => r.el.setAttribute('Target', v)) });
      return linkUnits.get(r.el);
    };
    const fieldLink = (get, set, part) => {
      const m = HYPER.exec(get());
      return m ? { unit: textUnit(part, m[2], (v) => set(get().replace(HYPER, (_, a, _u, c) => a + v + c))) } : null;
    };
    const object = (d, part) => {
      const rel = (id) => (id ? relsOf.get(part.path)?.get(id) : null);
      let label = '', target = null, embed = '';
      const blip = own(d, A_NS, 'blip'), vimg = own(d, V_NS, 'imagedata'), gd = own(d, A_NS, 'graphicData');
      if (blip) target = rel(blip.getAttributeNS(R_NS, 'embed') || blip.getAttributeNS(R_NS, 'link'));
      else if (vimg) target = rel(vimg.getAttributeNS(R_NS, 'id'));
      else if (gd) {
        const uri = gd.getAttribute('uri') || '';
        if (uri.endsWith('/chart')) { label = '图表'; embed = 'chart'; target = rel(own(gd, C_NS, 'chart')?.getAttributeNS(R_NS, 'id')); }
        else if (uri.endsWith('/diagram')) { label = 'SmartArt'; embed = 'smartart'; target = rel(own(gd, DGM_NS, 'relIds')?.getAttributeNS(R_NS, 'dm')); }
      }
      if (!target) return null;
      if (embed && !target.external) found.push({ embed, path: target.path });
      const docPr = own(d, WP_NS, 'docPr'), cNvPr = own(d, PIC_NS, 'cNvPr');
      const attr = (name) => {
        const v = docPr?.getAttribute(name);
        return v ? textUnit(part, v, (nv) => { docPr.setAttribute(name, nv); if (cNvPr?.getAttribute(name) === v) cNvPr.setAttribute(name, nv); }) : null;
      };
      return { label, path: target.external ? target.target : target.path.replace(/^word\//, ''), alt: attr('descr'), title: attr('title') };
    };

    // 段落原子：每个字一个；制表 / 换行 / 图片等是不可编辑的原子
    const collect = (p, part) => {
      const atoms = [], fields = []; // fields：复杂域栈 {instr: [instrText], link}
      let link = null;
      const cur = () => fields.findLast((f) => f.link)?.link ?? link;
      const addRun = (r, locked) => {
        const fmt = runFmt(r, style);
        for (const c of r.childNodes) {
          if (c.nodeType !== 1) continue;
          const base = { r, item: c, fmt, locked, link: cur() };
          const obj = (d) => { const o = d && object(d, part); if (o) atoms.push({ ...base, ch: '￼', obj: o, edit: false }); };
          if (c.namespaceURI === MC_NS && c.localName === 'AlternateContent') {
            const ch = [...c.childNodes].find((x) => x.localName === 'Choice');
            obj(ch && [...ch.childNodes].find((x) => isW(x, 'drawing')));
            continue;
          }
          if (c.namespaceURI !== NS) continue;
          switch (c.localName) {
            case 't': for (const ch of c.textContent) atoms.push({ ...base, ch, edit: !locked }); break;
            case 'tab': atoms.push({ ...base, ch: '\t', edit: false }); break;
            case 'br': case 'cr': atoms.push({ ...base, ch: '\n', edit: false }); break;
            case 'noBreakHyphen': atoms.push({ ...base, ch: '-', edit: false }); break;
            case 'drawing': case 'pict': obj(c); break;
            case 'instrText': fields.at(-1)?.instr.push(c); break;
            case 'fldChar': {
              const t = c.getAttributeNS(NS, 'fldCharType');
              if (t === 'begin') fields.push({ instr: [], link: null });
              else if (t === 'end') fields.pop();
              else if (t === 'separate' && fields.length) {
                const f = fields.at(-1);
                f.link = fieldLink(() => f.instr.map((n) => n.textContent).join(''), (v) => { f.instr.forEach((n, i) => { n.textContent = i ? '' : v; }); }, part);
              }
              break;
            }
          }
        }
      };
      const walk = (n, locked) => {
        for (const c of n.childNodes) {
          if (c.nodeType !== 1 || c.namespaceURI !== NS) continue;
          const ln = c.localName;
          if (ln === 'r') addRun(c, locked);
          else if (ln === 'ins' || ln === 'moveTo') walk(c, true);
          else if (ln === 'hyperlink' || ln === 'fldSimple') {
            const prev = link;
            link = (ln === 'hyperlink' ? relLink(c, part) : fieldLink(() => c.getAttributeNS(NS, 'instr') || '', (v) => c.setAttributeNS(NS, 'w:instr', v), part)) || link;
            walk(c, locked);
            link = prev;
          } else if (INLINE_WRAP.has(ln)) walk(c, locked);
        }
      };
      walk(p, false);
      return atoms;
    };

    const md = new MD();
    const mkPara = (p, part) => { const atoms = collect(p, part); return atoms.length ? unit({ kind: 'para', el: p, part, atoms }) : null; };
    const listPrefix = (lab, indent) => (lab ? (indent ? '  '.repeat(lab.ilvl) : '') + lab.text + ' ' : '');
    const scanTextboxes = (p, part) => {
      for (const tb of p.getElementsByTagNameNS(NS, 'txbxContent')) {
        if (inFallback(tb)) continue;
        let outer = true;
        for (let n = tb.parentNode; n && n.nodeType === 1; n = n.parentNode) if (n.localName === 'txbxContent') { outer = false; break; }
        if (outer) found.push({ title: '文本框', container: tb, part });
      }
    };
    let altChunk = 0;
    const walkBlocks = (container, part) => {
      for (const c of container.childNodes) {
        if (c.nodeType !== 1 || c.namespaceURI !== NS) continue;
        if (c.localName === 'p') {
          scanTextboxes(c, part);
          const lab = lists.label(c); // 空段落也要计数
          const para = mkPara(c, part);
          if (!para) continue;
          const h = headingLevel(c, style);
          emitPara(md, para, h ? '#'.repeat(h) + ' ' + listPrefix(lab) : listPrefix(lab, true), true);
          md.s('\n\n');
        } else if (c.localName === 'sdt') {
          const sc = kid(c, 'sdtContent');
          if (sc) walkBlocks(sc, part);
        } else if (c.localName === 'customXml') walkBlocks(c, part);
        else if (c.localName === 'altChunk') altChunk++;
        else if (c.localName === 'tbl') {
          const grid = plainTable(c);
          if (grid) {
            grid.forEach((row, ri) => {
              md.s('| ');
              row.forEach((p, ci) => {
                scanTextboxes(p, part);
                const lab = lists.label(p), para = mkPara(p, part);
                if (ci) md.s(' | ');
                if (para) emitPara(md, para, listPrefix(lab));
              });
              md.s(' |\n');
              if (ri === 0) md.s('|' + ' --- |'.repeat(row.length) + '\n');
            });
            md.s('\n');
          } else {
            const start = md.length;
            md.s(MARK.special + '（1格1行）\n');
            const inner = [];
            for (const tr of kids(c, 'tr')) for (const tc of tr.childNodes) {
              if (isW(tc, 'tc')) inner.push(tc);
              else if (isW(tc, 'sdt')) inner.push(...kids(kid(tc, 'sdtContent'), 'tc'));
            }
            const before = md.length;
            for (const tc of inner) walkCell(tc, part);
            if (md.length === before) md.cut(start);
            else md.s(MARK.end + '\n\n');
          }
        }
      }
    };
    // 特殊表格单元格：每段一行，嵌套表格就地展平
    const walkCell = (tc, part) => {
      for (const c of tc.childNodes) {
        if (c.nodeType !== 1 || c.namespaceURI !== NS) continue;
        if (c.localName === 'p') {
          scanTextboxes(c, part);
          const lab = lists.label(c), para = mkPara(c, part);
          if (para) { emitPara(md, para, listPrefix(lab), true); md.s('\n'); }
        } else if (c.localName === 'tbl') {
          for (const tr of kids(c, 'tr')) for (const t2 of kids(tr, 'tc')) walkCell(t2, part);
        } else if (c.localName === 'sdt' || c.localName === 'customXml') { const sc = c.localName === 'sdt' ? kid(c, 'sdtContent') : c; if (sc) walkCell(sc, part); }
      }
    };

    if (main) { const body = main.doc.getElementsByTagNameNS(NS, 'body')[0]; if (body) walkBlocks(body, main); }
    else warnings.push(`找不到正文 ${mainPath}`);

    // 其他文本放末尾：页眉页脚 / 脚注尾注，然后是正文里发现的文本框 / 图表 / SmartArt（处理中再发现的接着排）
    const queue = [];
    for (const { type, part } of others) {
      const title = TEXT_PARTS[type] + (type === 'header' || type === 'footer' ? ' ' + baseName(part.path) : '');
      if (type === 'footnotes' || type === 'endnotes') {
        for (const fn of part.doc.documentElement.childNodes) {
          if (fn.nodeType !== 1 || ['separator', 'continuationSeparator', 'continuationNotice'].includes(fn.getAttributeNS(NS, 'type'))) continue;
          queue.push({ title, container: fn, part });
        }
      } else queue.push({ title, container: part.doc.documentElement, part });
    }
    queue.push(...found.splice(0));
    const seen = new Set();
    const sectionOf = async (x) => {
      if (x.container) return x;
      if (seen.has(x.path)) return null;
      seen.add(x.path);
      const part = await load(x.path);
      if (!part) return null;
      const nodeUnit = (el, mirror) => textUnit(part, el.textContent, (v) => { el.textContent = v; }, mirror ? { mirror, orig: el.textContent } : {});
      if (x.embed === 'chart') {
        const els = [...part.doc.getElementsByTagName('*')].filter((el) => el.textContent.trim() && ((el.namespaceURI === A_NS && el.localName === 't') || (el.namespaceURI === C_NS && el.localName === 'v' && ['strCache', 'strLit'].includes(el.parentNode?.parentNode?.localName))));
        return { title: '图表 ' + baseName(x.path), units: els.map((el) => nodeUnit(el)) };
      }
      const els = [...part.doc.getElementsByTagNameNS(A_NS, 't')].filter((el) => el.textContent.trim());
      return { title: 'SmartArt ' + baseName(x.path), units: els.map((el) => nodeUnit(el, true)) };
    };
    let first = true;
    for (let i = 0; i < queue.length; i++) {
      const sec = await sectionOf(queue[i]);
      if (sec) {
        const start = md.length;
        if (first) md.s(MARK.others + '（页眉页脚 / 脚注 / 文本框 / 图表）\n\n');
        md.s('### ' + sec.title + '\n\n');
        const h = md.length;
        if (sec.container) walkBlocks(sec.container, sec.part);
        else if (sec.units.length) { for (const u of sec.units) { emitUnit(md, u, true); md.s('\n'); } md.s('\n'); }
        if (md.length === h) md.cut(start); else first = false;
      }
      queue.push(...found.splice(0));
    }
    for (const m of relsOf.values()) for (const r of m.values()) if (r.type === 'diagramDrawing' && !r.external) { const p = await load(r.path); if (p && !this.#mirrors.includes(p)) this.#mirrors.push(p); }

    // 投影不到的内容：列出来提醒
    if (altChunk) warnings.push(`正文里有 ${altChunk} 处嵌入的外部文档（altChunk），其中文字没有处理`);
    const emb = Object.keys(zip.files).filter((p) => /^word\/embeddings\/[^/]+$/.test(p));
    if (emb.length) warnings.push(`有 ${emb.length} 个内嵌对象（Excel / Visio / 图表数据等，word/embeddings/），里面的内容原样保留`);
    const math = main ? main.doc.getElementsByTagNameNS(M_NS, 'oMath').length : 0;
    if (math) warnings.push(`有 ${math} 处公式，公式里的文字没有投影`);
    const vec = Object.keys(zip.files).filter((p) => /^word\/media\/[^/]+\.(?:svg|emf|wmf|tiff?)$/i.test(p));
    if (vec.length) warnings.push(`有 ${vec.length} 张 svg / emf 等格式的图片不能替换成占位图（svg 常与同名 png 成对，只换 png 时 Word 仍显示 svg）`);

    this.md = md.out.join('');
    this.map = { pid: Int32Array.from(md.pid), off: Int32Array.from(md.off), ed: Uint8Array.from(md.ed) };
  }
}

// ---------- 图片 / 不可见元素 ----------
const PLACEHOLDER = {
  png: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGM4e/YsAATSAmjiBBWmAAAAAElFTkSuQmCC',
  gif: 'R0lGODdhAQABAIEAAM3NzQAAAAAAAAAAACwAAAAAAQABAAAIBAABBAQAOw==',
  bmp: 'Qk06AAAAAAAAADYAAAAoAAAAAQAAAAEAAAABABgAAAAAAAQAAADEDgAAxA4AAAAAAAAAAAAAzc3NAA==',
  jpg: '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD02iiigD//2Q==',
};
const b64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };

// word/media/* 与 docProps/thumbnail.*；replaceable 表示本工具能换成占位小图
export async function listImages(zip) {
  const out = [];
  for (const path of Object.keys(zip.files)) {
    const m = /^word\/media\/([^/]+)\.([A-Za-z0-9]+)$/.exec(path);
    if (!m || zip.files[path].dir) continue;
    const ext = m[2].toLowerCase();
    const bytes = await zip.file(path).async('uint8array');
    const mime = MIME[ext];
    out.push({ path, name: m[1] + '.' + ext, ext, size: bytes.length, replaceable: !!mime, url: mime ? URL.createObjectURL(new Blob([bytes], { type: mime })) : '' });
  }
  return out.sort((a, b) => a.path.localeCompare(b.path, 'en', { numeric: true }));
}

// ---------- 不可见指纹抹除：只留正文、格式和图片 ----------
// 参考微软「检查文档」可删除的隐藏数据和个人信息：文档属性/作者、批注、修订、隐藏文字、自定义 XML、缩略图、会话 ID(rsid/paraId) 等
const FP_PARTS = [
  /^word\/(?:comments\w*|people)\.xml$/, /^word\/_rels\/(?:comments\w*|people)\.xml\.rels$/,
  /^customXml\//, /^docProps\/(?:custom\.xml|thumbnail\..+)$/, /^word\/glossary\//, /^word\/printerSettings\//,
];
const FP_ATTR = /\s(?:w|w14|w15):(?:rsid[A-Za-z]*|paraId|textId)="[^"]*"/g;
const MIN_CORE = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>';
const MIN_APP = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office Word</Application></Properties>';

async function dropParts(zip, test) {
  const gone = new Set(Object.keys(zip.files).filter((p) => test(p)));
  if (!gone.size) return;
  for (const p of gone) zip.remove(p);
  for (const rp of Object.keys(zip.files)) {
    if (!rp.endsWith('.rels')) continue;
    const dir = rp.replace(/(^|\/)_rels\/[^/]*$/, '$1');
    const x = await zip.file(rp).async('string');
    zip.file(rp, x.replace(/<Relationship\b[^>]*\/>/g, (rel) => {
      const m = /Target="([^"]*)"/.exec(rel);
      if (!m || /TargetMode="External"/.test(rel)) return rel;
      return gone.has(decodeURIComponent(new URL(m[1], 'http://x/' + dir).pathname.slice(1))) ? '' : rel;
    }));
  }
  const ct = zip.file('[Content_Types].xml');
  if (ct) zip.file('[Content_Types].xml', (await ct.async('string')).replace(/<Override\b[^>]*PartName="\/([^"]*)"[^>]*\/>/g, (o, p) => (gone.has(p) ? '' : o)));
}

// 正文类 XML：批注锚点、修订历史（接受全部）、隐藏文字、拼写标记
function cleanDom(d) {
  const q = (n) => [...d.getElementsByTagNameNS(NS, n)];
  for (const n of ['proofErr', 'commentRangeStart', 'commentRangeEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd',
    'rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'tcPrChange', 'trPrChange', 'tblGridChange', 'tblPrExChange', 'numberingChange']) q(n).forEach((x) => x.remove());
  q('commentReference').forEach((x) => x.parentNode?.parentNode && x.parentNode.remove());
  const gb = q('bookmarkStart').filter((x) => x.getAttributeNS(NS, 'name') === '_GoBack');
  const ids = new Set(gb.map((x) => x.getAttributeNS(NS, 'id')));
  gb.forEach((x) => x.remove());
  q('bookmarkEnd').filter((x) => ids.has(x.getAttributeNS(NS, 'id'))).forEach((x) => x.remove());
  [...q('del'), ...q('moveFrom')].forEach((x) => x.remove());
  for (const x of [...q('ins'), ...q('moveTo')]) {
    if (!x.parentNode) continue;
    if (['rPr', 'pPr'].includes(x.parentNode.localName)) { x.remove(); continue; }
    while (x.firstChild) x.parentNode.insertBefore(x.firstChild, x);
    x.remove();
  }
  q('r').filter((r) => { const p = kid(r, 'rPr'); return p && on(kid(p, 'vanish')); }).forEach((r) => r.remove());
}

async function wipeFingerprints(zip) {
  await dropParts(zip, (p) => FP_PARTS.some((re) => re.test(p)));
  if (zip.file('docProps/core.xml')) zip.file('docProps/core.xml', MIN_CORE);
  if (zip.file('docProps/app.xml')) zip.file('docProps/app.xml', MIN_APP);
  for (const p of Object.keys(zip.files)) {
    if (!/^word\/[^/]+\.xml$/.test(p)) continue;
    let x = await zip.file(p).async('string');
    const y = x.replace(FP_ATTR, '').replace(/<w:rsid\b[^>]*\/>/g, '')
      .replace(/<w:rsids>[\s\S]*?<\/w:rsids>/, '').replace(/<w:docVars>[\s\S]*?<\/w:docVars>/, '').replace(/<w:attachedTemplate\b[^>]*\/>/, '');
    if (y !== x) zip.file(p, y);
  }
}
