// Word <-> Markdown 1:1 投影；差异以修订（w:del / w:ins）写回。三层：docx ⇄ md（本文件）⇄ txt（plain.js）
// 用法：const d = await DocxProjection.load(buf, opts); d.md / d.text / d.locate(i) / await d.media()；
//       改动：d.splice(spans)（md 区间）/ d.diffMd(newMd)（整份 md）/ d.mapPlain(fn, { range })（纯文本）——都归到 splice；
//       d.blankImages / d.wipeMetadata；await d.save({ revisions })
// 不依赖脱敏逻辑；结构标记、修订作者的默认值来自 markup.js，界面文字见 I18N，都可经 opts 覆盖
import JSZip from 'jszip';
import { PROJECTION } from './markup.js';
import { plainView, plainToSpans, diffSpans, emphasisOf, BR } from './plain.js';

const NS = {
  WORD: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  XML: 'http://www.w3.org/XML/1998/namespace',
  R: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  PKG: 'http://schemas.openxmlformats.org/package/2006/relationships',
  MC: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  A: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  WP: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  PIC: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  C: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  DGM: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  V: 'urn:schemas-microsoft-com:vml',
};
// 正文以外要投影的部件（主文档关系的类型），标题见 I18N.parts
const TEXT_PARTS = ['header', 'footer', 'footnotes', 'endnotes'];
const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

// 界面文字，opts.i18n 按组覆盖：{ parts: { header: 'Header' }, warn: { noMain: (p) => … } }
export const I18N = {
  parts: { header: '页眉', footer: '页脚', footnotes: '脚注', endnotes: '尾注', textbox: '文本框', chart: '图表', smartart: 'SmartArt', link: '链接地址' },
  comment: (what, from, to) => `${what}：「${from}」→「${to}」`, // 审阅模式下没有修订语法的改动（图表 / SmartArt 文字等）写成批注
  notes: { oMath: '公式', oMathPara: '公式', ruby: '注音', altChunk: '嵌入的外部文档', object: 'OLE 对象', t: '没投影到的文字', comment: '批注' }, // unhandled 里各标签的说明
  warn: {
    parseFailed: (path) => `${path} 解析失败，其中的文字没有处理`,
    noMain: (path) => `找不到正文 ${path}`,
    unhandled: (list) => '以下元素没有投影，其中的文字没有处理：\n' + list.map(({ tag, count, note, sample }) => `    <${tag}> ×${count}` + (note ? `（${note}${sample ? '，如「' + (sample.length > 20 ? sample.slice(0, 20) + '…' : sample) + '」' : ''}）` : '')).join('\n'),
    embeddings: (n) => `有 ${n} 个内嵌对象（Excel / Visio / 图表数据等，word/embeddings/），里面的内容原样保留`,
    vectors: (n) => `有 ${n} 张 svg / emf 等格式的图片不能替换成占位图（svg 常与同名 png 成对，只换 png 时 Word 仍显示 svg）`,
    newImages: (n) => `${n} 处新增图片被忽略（图片只能删除、移动或复制文档里已有的：路径须与原图相同，且在同一部分——正文 / 页眉等不能互相搬）`,
    untracked: (n) => `${n} 处页眉页脚里的图表 / SmartArt / 链接文字已直接改动（那里加不了批注）`,
    failed: (n) => `${n} 段写入失败（详见控制台），这些段落保持原样`,
    tableLost: (n) => `${n} 张表格没有找到可以放回的位置，没写入`,
    unmerged: (n) => `${n} 处段落合并没写入（后面不是段落，或会删掉分节符）`,
    foreign: (n) => `${n} 处插入落在别人的修订里，没写入`,
    smartArt: (n) => `SmartArt 缓存里有 ${n} 处文字没对上，可能还残留原文（在 Word 里点一下该 SmartArt 会重新排版）`,
    savedTwice: 'DocxProjection.save() 只能调用一次',
  },
};
const mergeI18n = (o = {}) => Object.fromEntries(Object.entries(I18N).map(([k, v]) => [k, typeof v === 'function' ? o[k] ?? v : { ...v, ...o[k] }]));

// 1×1 占位图（dropImages），按扩展名
const PLACEHOLDER = {
  png: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGM4e/YsAATSAmjiBBWmAAAAAElFTkSuQmCC',
  gif: 'R0lGODdhAQABAIEAAM3NzQAAAAAAAAAAACwAAAAAAQABAAAIBAABBAQAOw==',
  bmp: 'Qk06AAAAAAAAADYAAAAoAAAAAQAAAAEAAAABABgAAAAAAAQAAADEDgAAxA4AAAAAAAAAAAAAzc3NAA==',
  jpg: '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD02iiigD//2Q==',
};
const b64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };

const isW = (n, name) => n.nodeType === 1 && n.namespaceURI === NS.WORD && n.localName === name;
const kids = (n, name) => (n ? [...n.childNodes].filter((c) => isW(c, name)) : []);
const kid = (n, name) => (n ? [...n.childNodes].find((c) => isW(c, name)) : undefined);
const val = (el) => el?.getAttributeNS(NS.WORD, 'val') ?? null;
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


// ---------- 样式 / 编号 ----------
// 返回 look(styleId, f)：沿 basedOn 链找第一个 f(style) 不为空的值
function styleLookup(doc) {
  const map = new Map();
  for (const s of doc ? doc.getElementsByTagNameNS(NS.WORD, 'style') : []) {
    const pPr = kid(s, 'pPr');
    map.set(s.getAttributeNS(NS.WORD, 'styleId'), { name: (val(kid(s, 'name')) || '').toLowerCase(), based: val(kid(s, 'basedOn')), outline: val(kid(pPr, 'outlineLvl')), numPr: kid(pPr, 'numPr'), rPr: kid(s, 'rPr') });
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
    for (const a of doc.getElementsByTagNameNS(NS.WORD, 'abstractNum')) {
      const lv = new Map();
      for (const l of kids(a, 'lvl')) lv.set(+l.getAttributeNS(NS.WORD, 'ilvl'), { fmt: val(kid(l, 'numFmt')) || 'decimal', text: val(kid(l, 'lvlText')) ?? '', start: +(val(kid(l, 'start')) ?? 1) });
      this.abs.set(a.getAttributeNS(NS.WORD, 'abstractNumId'), lv);
    }
    for (const n of doc.getElementsByTagNameNS(NS.WORD, 'num')) {
      const start = new Map();
      for (const o of kids(n, 'lvlOverride')) { const so = val(kid(o, 'startOverride')); if (so != null) start.set(+o.getAttributeNS(NS.WORD, 'ilvl'), +so); }
      this.nums.set(n.getAttributeNS(NS.WORD, 'numId'), { abs: val(kid(n, 'abstractNumId')), start });
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
// 每个 md 字符记：pid 所属单元、off 单元内原子位置、ed 可编辑、aft「紧跟这个字之后插入」落在单元内哪个位置
// （可编辑字是 off + 1；不可编辑的原子如 <br> 也是 off + 1；它前面的标记、转义是 off，即插在原子之前）；
// brk：段落之间的换行代表哪个段落的段落标记（删掉它 = 与下一段合并），其余为 -1
class MD {
  constructor(ambig) { this.ambig = ambig; this.out = []; this.pid = []; this.off = []; this.ed = []; this.aft = []; this.brk = []; this.imgs = []; this.tbls = []; }
  push(str, pid, off, ed, aft = off + ed, brk = -1) { for (let i = 0; i < str.length; i++) { this.out.push(str[i]); this.pid.push(pid); this.off.push(off); this.ed.push(ed); this.aft.push(aft); this.brk.push(brk); } }
  s(str) { this.push(str, -1, 0, 0); }
  sp(str, para, off, aft = off) { this.push(str, para.idx, off, 0, aft); }
  mark(str, para) { this.push(str, -1, 0, 0, 0, para.idx); } // 段落标记
  cut(n) { this.out.length = this.pid.length = this.off.length = this.ed.length = this.aft.length = this.brk.length = n; this.imgs = this.imgs.filter((x) => x.e <= n); this.tbls = this.tbls.filter((x) => x.e <= n); }
  trim(n) { while (n > 0 && this.out[n - 1] === '\n') n--; return n; } // 去掉末尾换行后的位置
  get length() { return this.out.length; }
}
const marker = (f) => (f === 'bi' ? '***' : f === 'b' ? '**' : '*');
const ESC = new Set(['*', '\\', '|']);
const head = (atoms) => atoms.slice(0, 16).map((a) => a.ch).join(''); // 行首够长，能和 marks.ambig 里最长的标记比较

// 文本单元（链接地址、图片说明、图表文字…）：逐字可编辑
function emitUnit(md, u, lineStart = false) {
  if (lineStart && md.ambig.test(head(u.atoms))) md.push('\\', u.idx, 0, 0);
  u.atoms.forEach((a, i) => { if (ESC.has(a.ch)) md.push('\\', u.idx, i, 0); md.push(a.ch, u.idx, i, 1); });
}

// 图片 / 图表 / SmartArt：整段 ![alt](路径) 记进 md.imgs。只能整体删除或整体移动（见 Edits.dropObj / IMG_TOKEN），不能新增
function emitObj(md, para, i, o, locked) {
  const s = md.length;
  md.sp('![', para, i);
  if (o.alt) emitUnit(md, o.alt); else md.sp(o.label, para, i);
  md.sp('](' + o.path, para, i + 1);
  if (o.title) { md.sp(' "', para, i + 1); emitUnit(md, o.title); md.sp('"', para, i + 1); }
  md.sp(')', para, i + 1);
  md.imgs.push({ s, e: md.length, pid: para.idx, off: i, path: o.path, file: o.file, part: para.part, locked: !!locked });
}

function emitPara(md, para, prefix = '', lineStart = false) {
  const { atoms } = para;
  if (prefix) md.sp(prefix, para, 0); // 标题 # / 编号前缀：结构字符，但行首插入能挂到本段开头
  if (lineStart && !prefix && !atoms[0].link && md.ambig.test(head(atoms))) md.sp('\\', para, 0);
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
    if (a.obj) emitObj(md, para, i, a.obj, a.locked);
    else if (a.edit) {
      if (ESC.has(a.ch)) md.sp('\\', para, i);
      md.push(a.ch, para.idx, i, 1);
    } else md.sp(a.ch === '\n' ? '<br>' : a.ch, para, i, i + 1); // 软换行等不可编辑原子：紧跟其后插入的字落在它后面
  });
  if (fmt) md.sp(marker(fmt), para, atoms.length);
  if (link) closeLink(atoms.length);
}

// 表格的行 / 行里的格：穿过内容控件（w:sdt > w:sdtContent）与 w:customXml 包装（表单的「重复节」会把整行包在 sdt 里）
function unwrap(n, name) {
  const out = [];
  for (const c of n?.childNodes || []) {
    if (isW(c, name)) out.push(c);
    else if (isW(c, 'sdt')) out.push(...unwrap(kid(c, 'sdtContent'), name));
    else if (isW(c, 'customXml')) out.push(...unwrap(c, name));
  }
  return out;
}

function plainTable(tbl) {
  const rows = unwrap(tbl, 'tr');
  if (!rows.length) return null;
  const grid = [];
  let n = -1;
  for (const tr of rows) {
    const trPr = kid(tr, 'trPr');
    if (trPr && (kid(trPr, 'gridBefore') || kid(trPr, 'gridAfter'))) return null;
    const cells = unwrap(tr, 'tc');
    if (!cells.length || (n >= 0 && cells.length !== n)) return null;
    n = cells.length;
    const row = [];
    for (const tc of cells) {
      const pr = kid(tc, 'tcPr');
      if (pr && (kid(pr, 'vMerge') || kid(pr, 'hMerge') || (kid(pr, 'gridSpan') && +val(kid(pr, 'gridSpan')) > 1))) return null;
      const ps = kids(tc, 'p');
      const other = [...tc.childNodes].some((x) => x.nodeType === 1 && x.namespaceURI === NS.WORD && !['tcPr', 'p'].includes(x.localName));
      if (other || ps.length !== 1) return null;
      row.push(ps[0]);
    }
    grid.push(row);
  }
  return grid;
}

// 本工具上一次写入的修订先「拒绝」掉再投影：删除线里的原文还原成正常文字，插入的内容去掉。
// 否则 Ctrl+Enter 写入审阅修订后，重新载入 A 时原文已经藏在 w:del 里，投影看不到旧内容、黑名单就扫不出来了；
// 同时重复写入审阅不会叠加。只认作者 author，别的作者的修订保持原样。返回处理的数量。
function rejectOurs(doc, author) {
  let n = 0;
  const byTag = (t) => [...doc.getElementsByTagNameNS(NS.WORD, t)].filter((x) => x.getAttributeNS(NS.WORD, 'author') === author);
  for (const x of byTag('ins')) { // 整行是我们插入的（见 markInserted）：整行撤掉，表格没有行了连表格一起撤掉
    if (x.parentNode?.localName !== 'trPr') continue;
    const tr = x.parentNode.parentNode, tbl = tr.parentNode;
    tr.remove(); n++;
    if (tbl && isW(tbl, 'tbl') && !kids(tbl, 'tr').length) tbl.remove();
  }
  for (const x of byTag('ins')) { if (x.parentNode && x.parentNode.localName !== 'rPr') { x.remove(); n++; } }
  for (const x of byTag('del')) {
    if (!x.parentNode || x.parentNode.localName === 'rPr') continue;
    for (const dt of [...x.getElementsByTagNameNS(NS.WORD, 'delText')]) dt.replaceWith(mkText(doc, 't', dt.textContent));
    while (x.firstChild) x.parentNode.insertBefore(x.firstChild, x);
    x.remove(); n++;
  }
  // 改链接地址时插入的新链接：里面的插入撤掉后只剩空壳
  for (const t of ['hyperlink', 'fldSimple']) for (const x of [...doc.getElementsByTagNameNS(NS.WORD, t)]) if (![...x.childNodes].some((c) => c.nodeType === 1)) { x.remove(); n++; }
  // 格式修订：换回 rPrChange 里记的原格式
  for (const x of byTag('rPrChange')) {
    const rPr = x.parentNode, old = kid(x, 'rPr');
    while (rPr.firstChild) rPr.removeChild(rPr.firstChild);
    for (const c of [...(old?.childNodes || [])]) rPr.appendChild(c);
    if (!rPr.childNodes.length) rPr.remove();
    n++;
  }
  // 段落标记：自己删的去掉删除标记；自己插入的（拆段 / 新段）与下一段合并回去，用下一段的段落属性（拆段时原来的标记留在后一段）
  const marks = (t) => byTag(t).filter((x) => x.parentNode?.localName === 'rPr' && x.parentNode.parentNode?.localName === 'pPr');
  for (const x of marks('del')) { x.remove(); n++; }
  for (const x of marks('ins').reverse()) {
    const p = x.parentNode.parentNode.parentNode, np = nextBlock(p);
    x.remove(); n++;
    if (!isW(p, 'p') || !np || !isW(np, 'p')) continue;
    kid(p, 'pPr')?.remove();
    const npPr = kid(np, 'pPr');
    if (npPr) p.insertBefore(npPr, p.firstChild);
    for (const c of [...np.childNodes]) p.appendChild(c);
    np.remove();
  }
  return n;
}

async function loadXml(zip, path, onFail) {
  const f = path && zip.file(path);
  if (!f) return null;
  const doc = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) { onFail(path); return null; }
  return doc;
}
const relsPathOf = (path) => path.replace(/([^/]*)$/, '_rels/$1.rels');
const resolvePath = (base, target) => decodeURIComponent(new URL(target, 'http://x/' + base).pathname.slice(1));
const baseName = (path) => path.split('/').pop().replace(/\.xml$/, '');
const HYPER = /^(\s*HYPERLINK\s+")([^"]+)(")/i;
const INLINE_WRAP = new Set(['smartTag', 'sdt', 'sdtContent', 'dir', 'bdo', 'customXml']);
// 投影不了的元素：不写投影，只按标签名列进 unhandled 清单。删除的修订 / 域代码 / 文本框（另行处理）里的文字不算
const NO_TEXT = new Set(['del', 'moveFrom', 'delText', 'delInstrText', 'instrText', 'txbxContent', 'rPr', 'pPr']);
const hasText = (el) => [...el.childNodes].some((c) => (c.nodeType === 3 ? c.data.trim() : c.nodeType === 1 && !NO_TEXT.has(c.localName) && hasText(c)));

// ---------- 编辑记录 ----------
// 插入文本里的 ![..](路径)：不是新增图片，而是移动（放回被删掉的同路径图片）或复制（克隆文档里已有的同路径图片）。文本里换成 OBJ，路径按顺序记进 refs，保存时再配对
const OBJ = '\uFFFC';
const PS = '\u2029'; // 插入文字里的分段（splice 把 \n 换成它，\u2028 换回 \n 即软换行）
const IMG_TOKEN = /!\[(?:\\[\s\S]|[^\]\\])*\]\(([^\s)"]+)(?: "(?:\\[\s\S]|[^"\\])*")?\)/g;
// diffMd 逐段比较时整张图片算一个词（不和别的图片逐字错配）；同一路径的才逐字比（改说明）
const imgTokens = (s) => { const out = []; let p = 0; for (const m of s.matchAll(IMG_TOKEN)) { out.push(...s.slice(p, m.index), m[0]); p = m.index + m[0].length; } out.push(...s.slice(p)); return out; };
const imgPath = (tok) => (tok.length > 1 && tok.startsWith('![') ? new RegExp(IMG_TOKEN.source).exec(tok)?.[1] : null);
const IMG_DIFF = { tokens: imgTokens, same: (x, y) => imgPath(x) != null && imgPath(x) === imgPath(y) };
class Edits extends Map {
  dropped = []; // 被整体删除的图片 { pid, off, path, part }，没被配对到插入处的就是真删除
  of(pid) { let e = this.get(pid); if (!e) this.set(pid, (e = { del: new Set(), ins: new Map(), drop: new Set(), refs: new Map(), nodes: new Map(), merge: false, fmt: new Map() })); return e; }
  del(pid, off) { this.of(pid).del.add(off); }
  ins(pid, at, text) {
    const e = this.of(pid), refs = e.refs.get(at) || [];
    text = text.replace(IMG_TOKEN, (_, path) => { refs.push(path); return OBJ; });
    if (refs.length) e.refs.set(at, refs);
    e.ins.set(at, (e.ins.get(at) || '') + text);
  }
  merge(pid) { this.of(pid).merge = true; } // 删掉段落标记：与下一段合并
  fmt(pid, off, bits) { this.of(pid).fmt.set(off, bits); } // 改粗斜体：位 1 粗、位 2 斜
  rows = new Map(); // 要整行删除的 w:tr -> 所在部件
  tbls = new Map(); // 要整表删除的 w:tbl -> 所在部件
  rowIns = []; // 要插入的行 [{after（本文档里的 w:tr）, el（来自另一个文档的 w:tr）}]
  tblIns = []; // 要整表插入的 [{pos（md 坐标）, el（来自另一个文档的 w:tbl）}]
  dropObj(img) { const e = this.of(img.pid); if (!e.drop.has(img.off)) { e.drop.add(img.off); this.dropped.push(img); } }
}
// md 区间 [a, b) 整个包住的图片（只包住一半的不算：结构字符不可编辑）
const coveredImgs = (images, a, b) => images.filter((x) => !x.locked && x.s >= a && x.e <= b);

// md 位置 p 处插入：挂到前一个字之后，否则挂到后一个字之前；结构位置（不属于任何单元）返回 null
function insertAnchor(map, p) {
  if (p > 0 && map.pid[p - 1] >= 0) return [map.pid[p - 1], map.aft[p - 1]];
  if (p < map.pid.length && map.pid[p] >= 0) return [map.pid[p], map.off[p]];
  return null;
}

// diffMd 插入的文字只取新 md 里的「内容」：来自另一个投影时取它的可编辑字符和整张图片（图片是移动 / 复制），
// 否则按纯文本投影去掉转义与粗斜体标记。返回 (at, n) -> 新 md [at, at + n) 里的内容，写法同 splice 的 text：\n 分段、\u2028 软换行
function contentOf(newMd, from, drop = []) {
  const keep = new Array(newMd.length).fill('');
  if (from) {
    for (let i = 0; i < newMd.length; i++) if (from.ed[i] || from.brk[i] >= 0) keep[i] = newMd[i];
    for (const x of from.images) for (let i = x.s; i < x.e; i++) keep[i] = newMd[i];
  } else {
    const v = plainView(newMd);
    for (let k = 0; k < v.text.length; k++) {
      const [ms, me] = v.toMd(k, k + 1), c = v.text[k];
      keep[c === BR ? ms : me - 1] = c;
    }
  }
  for (const [a, b] of [...drop].reverse()) keep.splice(a, b - a);
  return (at, n) => keep.slice(at, at + n).join('').replace(/\r\n?/g, '\n');
}

// ---------- 写入修订 ----------
function mkEl(doc, name, attrs = {}) {
  const el = doc.createElementNS(NS.WORD, 'w:' + name);
  for (const [k, v] of Object.entries(attrs)) el.setAttributeNS(NS.WORD, 'w:' + k, v);
  return el;
}
function mkText(doc, name, text) {
  const t = mkEl(doc, name);
  t.setAttributeNS(NS.XML, 'xml:space', 'preserve');
  t.textContent = text;
  return t;
}

// rPr 子元素的规定顺序（不按顺序 Word 可能报文件损坏）；不认识的（w14: 等扩展）排在 rPrChange 之前
const RPR_ORDER = ['ins', 'del', 'moveFrom', 'moveTo', 'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
const rprRank = (n) => (n.namespaceURI === NS.WORD && RPR_ORDER.includes(n.localName) ? RPR_ORDER.indexOf(n.localName) : n.localName === 'rPrChange' ? 1e9 : 1e8);
function putRPr(rPr, el) {
  const r = rprRank(el);
  rPr.insertBefore(el, [...rPr.childNodes].find((c) => c.nodeType === 1 && rprRank(c) > r) || null);
}
// 把 run 的粗斜体从 cur 改成 bits（位 1 粗、位 2 斜，只动变了的位），返回新 rPr；track 时附 w:rPrChange 记下原格式。
// 别人已经记过 rPrChange 的不再加（原格式以他们记的为准）
function withFmt(doc, rpr, bits, cur, meta, track) {
  const n = rpr ? rpr.cloneNode(true) : mkEl(doc, 'rPr');
  const had = kid(n, 'rPrChange');
  for (const [name, bit] of [['b', 1], ['i', 2]]) {
    if (!((bits ^ cur) & bit)) continue;
    kid(n, name)?.remove();
    putRPr(n, bits & bit ? mkEl(doc, name) : mkEl(doc, name, { val: '0' }));
  }
  if (track && !had) {
    const old = rpr ? rpr.cloneNode(true) : mkEl(doc, 'rPr');
    for (const c of [...old.childNodes]) if (c.nodeType === 1 && ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange'].includes(c.localName)) c.remove();
    const ch = mkEl(doc, 'rPrChange', meta());
    ch.appendChild(old);
    n.appendChild(ch);
  }
  return n;
}

// 段落标记的格式（空段落里插字时用）：pPr > rPr 去掉其中的修订标记
function markRPr(p) {
  const r = kid(kid(p, 'pPr'), 'rPr')?.cloneNode(true);
  if (!r) return null;
  for (const c of [...r.childNodes]) if (c.nodeType === 1 && ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange'].includes(c.localName)) c.remove();
  return r.childNodes.length ? r : null;
}

// 段落标记的修订：pPr > rPr 里的 w:ins / w:del（排在 rPr 最前；rPr 排在 pPr 的 sectPr / pPrChange 之前）
function markPara(p, kind, ctx) {
  const doc = p.ownerDocument;
  let pPr = kid(p, 'pPr');
  if (!pPr) { pPr = mkEl(doc, 'pPr'); p.insertBefore(pPr, p.firstChild); }
  let rPr = kid(pPr, 'rPr');
  if (!rPr) { rPr = mkEl(doc, 'rPr'); pPr.insertBefore(rPr, kid(pPr, 'sectPr') || kid(pPr, 'pPrChange') || null); }
  const ins = kid(rPr, 'ins');
  rPr.insertBefore(mkEl(doc, kind, { id: String(ctx.nextId++), author: ctx.author, date: ctx.date }), kind === 'del' && ins ? ins.nextSibling : rPr.firstChild);
}

// 在 marker 处把段落 p 拆成两段，返回后一段。后一段沿用原来的段落标记（含分节符）；前一段的标记是新插入的（审阅模式下标成插入）。
// marker 可以在超链接等容器里：沿途的容器各拆成两份
const SPLIT_NS = 'urn:x-docx-projection:split';
function splitPara(p, marker, ctx) {
  const np = p.cloneNode(false);
  for (const a of [...np.attributes]) if (a.localName === 'paraId' || a.localName === 'textId') np.removeAttributeNode(a);
  const pPr = kid(p, 'pPr');
  if (pPr) np.appendChild(pPr.cloneNode(true));
  let cur = marker, right = null;
  while (cur.parentNode !== p) {
    const parent = cur.parentNode, clone = parent.cloneNode(false);
    while (cur.nextSibling) clone.appendChild(cur.nextSibling);
    if (right) clone.insertBefore(right, clone.firstChild);
    right = clone.firstChild ? clone : null;
    cur = parent;
  }
  if (right) np.appendChild(right);
  while (cur.nextSibling) np.appendChild(cur.nextSibling);
  marker.remove();
  kid(pPr, 'sectPr')?.remove();
  if (ctx.track) markPara(p, 'ins', ctx);
  p.parentNode.insertBefore(np, p.nextSibling);
  ctx.created?.push(np);
  return np;
}

// root 里还没进修订的 run 各自包进 w:ins / w:del（删除时 w:t / w:instrText 换成 w:delText / w:delInstrText）
function trackRuns(root, kind, ctx) {
  const doc = root.ownerDocument;
  for (const r of [...root.getElementsByTagNameNS(NS.WORD, 'r')]) {
    let inside = false;
    for (let n = r.parentNode; n && n !== root.parentNode; n = n.parentNode) if (isW(n, 'ins') || isW(n, 'del')) inside = true;
    if (inside) continue;
    if (kind === 'del') for (const x of [...r.childNodes]) if (isW(x, 't') || isW(x, 'instrText')) x.replaceWith(mkText(doc, x.localName === 't' ? 'delText' : 'delInstrText', x.textContent));
    const w = mkEl(doc, kind, ctx.meta());
    r.parentNode.insertBefore(w, r);
    w.appendChild(r);
  }
}

// 改链接地址（审阅模式）：原链接整个标成删除，后面跟一份用 setNew 改了地址、整个标成插入的副本
function reviseLink(h, setNew, ctx) {
  const neu = h.cloneNode(true);
  setNew(neu);
  for (const x of [...h.getElementsByTagNameNS(NS.WORD, 'ins')]) if (x.getAttributeNS(NS.WORD, 'author') === ctx.author) x.remove();
  trackRuns(h, 'del', ctx);
  for (const x of [...neu.getElementsByTagNameNS(NS.WORD, 'del')]) x.remove();
  for (const x of [...neu.getElementsByTagNameNS(NS.WORD, 'ins')]) { while (x.firstChild) x.parentNode.insertBefore(x.firstChild, x); x.remove(); }
  trackRuns(neu, 'ins', ctx);
  h.parentNode.insertBefore(neu, h.nextSibling);
}

// 部件 from 里引用包内路径 to 的相对写法
const relTarget = (from, to) => {
  const a = from.split('/').slice(0, -1), b = to.split('/');
  let i = 0;
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
  return '../'.repeat(a.length - i) + b.slice(i).join('/');
};
const closestP = (n) => { for (; n && n.nodeType === 1; n = n.parentNode) if (isW(n, 'p')) return n; return null; };

// 段落之间可以跳过的范围标记（书签、批注范围等）
const RANGE_MARKS = new Set(['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'permStart', 'permEnd', 'proofErr', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']);
const nextBlock = (n) => { for (n = n.nextSibling; n; n = n.nextSibling) if (n.nodeType === 1 && !(n.namespaceURI === NS.WORD && RANGE_MARKS.has(n.localName))) return n; return null; };

// 删掉段落 el 的段落标记：与后面第一个有内容的段落合并（中间空段落的标记一并删）。合并后的段落用后一段的格式（Word 也是这样）。
// 后面不是段落（表格、正文结尾）或要删掉分节符时不合并，返回 false
function mergeMark(el, ctx) {
  const chain = [el];
  let n = nextBlock(el);
  while (n && isW(n, 'p') && !n.getElementsByTagNameNS(NS.WORD, 'r').length) { chain.push(n); n = nextBlock(n); }
  if (!n || !isW(n, 'p') || chain.some((p) => kid(kid(p, 'pPr'), 'sectPr'))) return false;
  if (ctx.track) { for (const p of chain) markPara(p, 'del', ctx); return true; }
  const pPr = kid(n, 'pPr'), at = pPr ? pPr.nextSibling : n.firstChild;
  for (const p of chain) { for (const c of [...p.childNodes]) if (!isW(c, 'pPr')) n.insertBefore(c, at); p.remove(); }
  return true;
}

// 克隆来的整张表标成插入：每行 trPr 加 w:ins，每个段落标记标插入，每个 run 包进 w:ins；原文档里已被删除的 run 直接丢掉
function markInserted(tbl, ctx) {
  const doc = tbl.ownerDocument;
  for (const d of [...tbl.getElementsByTagNameNS(NS.WORD, 'del')]) if (d.parentNode?.localName !== 'rPr' && d.parentNode?.localName !== 'trPr') d.remove();
  for (const x of [...tbl.getElementsByTagNameNS(NS.WORD, 'ins')]) {
    if (x.parentNode?.localName === 'rPr' || x.parentNode?.localName === 'trPr') { x.remove(); continue; }
    while (x.firstChild) x.parentNode.insertBefore(x.firstChild, x);
    x.remove();
  }
  for (const tr of [...(isW(tbl, 'tr') ? [tbl] : []), ...tbl.getElementsByTagNameNS(NS.WORD, 'tr')]) {
    let pr = kid(tr, 'trPr');
    if (!pr) { pr = mkEl(doc, 'trPr'); tr.insertBefore(pr, [...tr.childNodes].find((c) => c.nodeType === 1 && !isW(c, 'tblPrEx')) || null); }
    pr.insertBefore(mkEl(doc, 'ins', ctx.meta()), kid(pr, 'del') || kid(pr, 'trPrChange') || null);
  }
  for (const p of [...tbl.getElementsByTagNameNS(NS.WORD, 'p')]) markPara(p, 'ins', ctx);
  for (const r of [...tbl.getElementsByTagNameNS(NS.WORD, 'r')]) {
    const w = mkEl(doc, 'ins', ctx.meta());
    r.parentNode.insertBefore(w, r);
    w.appendChild(r);
  }
}

// 整行 / 整表删除：审阅模式给每行的 trPr 加 w:del（行里的文字已标成删除），否则直接移除；表格一行不剩就连表格一起移除
function dropRows({ rows, tbls }, ctx) {
  const all = new Map(rows);
  for (const [tbl, part] of tbls) for (const tr of unwrap(tbl, 'tr')) all.set(tr, part);
  const owner = (n) => { while (n && !isW(n, 'tbl')) n = n.parentNode; return n; };
  const gone = (el) => {
    const up = el.parentNode;
    el.remove();
    if (up && isW(up, 'tc') && !kids(up, 'p').length) up.appendChild(mkEl(up.ownerDocument, 'p')); // 单元格不能没有段落
  };
  const tables = new Map();
  for (const [tr, part] of all) {
    part.dirty = true;
    if (!ctx.track) { const t = owner(tr); if (t) tables.set(t, part); tr.remove(); continue; }
    const doc = tr.ownerDocument;
    let pr = kid(tr, 'trPr');
    if (!pr) { pr = mkEl(doc, 'trPr'); tr.insertBefore(pr, [...tr.childNodes].find((c) => c.nodeType === 1 && !isW(c, 'tblPrEx')) || null); }
    if (!kid(pr, 'del')) pr.insertBefore(mkEl(doc, 'del', ctx.meta()), kid(pr, 'trPrChange') || null);
  }
  if (ctx.track) return;
  for (const [t, part] of tables) if (!unwrap(t, 'tr').length) tbls.set(t, part);
  for (const [t, part] of tbls) { part.dirty = true; if (t.parentNode) gone(t); }
}

// 返回 { dropped：因落在别人修订里而丢掉的插入数, frags：拆段后的各段（第一段是原段落） }
function applyParaEdits(para, ed, ctx) {
  const { atoms } = para;
  const doc = para.el.ownerDocument;
  const need = new Set();
  const lastIdx = atoms.length - 1;
  for (const i of ed.del) if (atoms[i]?.edit) need.add(atoms[i].r);
  for (const i of ed.drop) need.add(atoms[i].r);
  for (const i of ed.fmt.keys()) if (atoms[i]?.edit) need.add(atoms[i].r);
  if (lastIdx >= 0) for (const at of ed.ins.keys()) need.add(atoms[Math.min(at, lastIdx)].r);
  const rprClones = new Map();
  for (const a of atoms) if (!rprClones.has(a.r)) rprClones.set(a.r, kid(a.r, 'rPr')?.cloneNode(true) || null);
  // 插入文字的格式：替换（该位置的字被删）取被删字的格式，否则取前一个字的；空段落取段落标记的
  const rprOfAtom = (gi) => (lastIdx < 0 ? markRPr(para.el) : rprClones.get(atoms[ed.del.has(gi) && atoms[gi]?.edit ? gi : Math.max(0, Math.min(gi - 1, lastIdx))].r));
  const meta = () => ({ id: String(ctx.nextId++), author: ctx.author, date: ctx.date });
  const mkRun = (r, rpr, child) => {
    const nr = r.cloneNode(false);
    if (rpr) nr.appendChild(rpr.cloneNode(true));
    nr.appendChild(child);
    return nr;
  };
  // 插入文本里的 \n 写成 <w:br/>、\t 写成 <w:tab/>；OBJ 处放回被移动的图片（ed.nodes，已在 save 里配好对）；
  // PS 处放一个拆段标记，整段改完后再在那里拆开（标记不进 w:ins）
  const insNode = (r, gi, text) => {
    const out = [], objs = [...(ed.nodes.get(gi) || [])];
    let runs = [];
    const wrap = () => {
      if (!runs.length) return;
      if (ctx.track) { const w = mkEl(doc, 'ins', meta()); runs.forEach((x) => w.appendChild(x)); out.push(w); } else out.push(...runs);
      runs = [];
    };
    for (const part of text.split(/([\n\t\uFFFC\u2029])/)) {
      if (part === PS) { wrap(); out.push(doc.createElementNS(SPLIT_NS, 'split')); }
      else if (part === '\n') runs.push(mkRun(r, rprOfAtom(gi), mkEl(doc, 'br')));
      else if (part === '\t') runs.push(mkRun(r, rprOfAtom(gi), mkEl(doc, 'tab')));
      else if (part === OBJ) { const n = objs.shift(); if (n) runs.push(n); }
      else if (part) runs.push(mkRun(r, rprOfAtom(gi), mkText(doc, 't', part)));
    }
    wrap();
    return out;
  };
  const delNode = (r, rpr, text) => {
    if (!ctx.track) return null;
    const w = mkEl(doc, 'del', meta());
    w.appendChild(mkRun(r, rpr, typeof text !== 'string' ? text : mkText(doc, 'delText', text))); // 制表符 / 图片整个 run 包进 w:del
    return w;
  };
  let dropped = 0;
  if (lastIdx < 0 && ed.ins.has(0)) for (const n of insNode(mkEl(doc, 'r'), 0, ed.ins.get(0))) para.el.appendChild(n); // 空段落（空单元格）
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
        for (const [a, gi] of list) {
          if (ed.ins.has(gi)) push(insNode(r, gi, ed.ins.get(gi)));
          if ((a.edit && ed.del.has(gi)) || ed.drop.has(gi)) push(delNode(r, rpr, c.cloneNode(true))); // 制表符 / 图片：删除（审阅模式下是删除线）
          else push(mkRun(r, rpr, c.cloneNode(true)));
          if (gi === lastIdx && ed.ins.has(atoms.length)) push(insNode(r, atoms.length, ed.ins.get(atoms.length)));
        }
        continue;
      }
      let buf = '', bufKey = ''; // 'd' 删除，'f' + 目标位 + 现有位 改格式，'' 原样
      const flush = () => {
        if (!buf) return;
        push(bufKey === 'd' ? delNode(r, rpr, buf) : mkRun(r, bufKey ? withFmt(doc, rpr, +bufKey[1], +bufKey[2], meta, ctx.track) : rpr, mkText(doc, 't', buf)));
        buf = '';
      };
      for (const [a, gi] of list) {
        if (ed.ins.has(gi)) { flush(); push(insNode(r, gi, ed.ins.get(gi))); }
        const k = a.edit && ed.del.has(gi) ? 'd' : a.edit && ed.fmt.has(gi) ? 'f' + ed.fmt.get(gi) + ((a.fmt.includes('b') ? 1 : 0) | (a.fmt.includes('i') ? 2 : 0)) : '';
        if (k !== bufKey) { flush(); bufKey = k; }
        buf += a.ch;
        if (gi === lastIdx && ed.ins.has(atoms.length)) { flush(); push(insNode(r, atoms.length, ed.ins.get(atoms.length))); }
      }
      flush();
    }
    for (const n of out) parent.insertBefore(n, r);
    parent.removeChild(r);
  }
  const frags = [para.el];
  for (const m of [...para.el.getElementsByTagNameNS(SPLIT_NS, 'split')]) frags.push(splitPara(frags.at(-1), m, ctx));
  return { dropped, frags };
}

// 文本单元改完后的文字
function textAfter(u, ed) {
  let s = '';
  u.atoms.forEach((a, i) => { s += ed.ins.get(i) ?? ''; if (!ed.del.has(i)) s += a.ch; });
  return (s + (ed.ins.get(u.atoms.length) ?? '')).replace(/[\n\u2029]/g, ' ').replace(/\uFFFC/g, ''); // 属性文字里放不了图片、不能分段
}

// 文本框 mc:AlternateContent：投影 / 编辑的是 Choice（wps），Fallback（VML）里还有一份原文——改过的照 Choice 同步过去
function syncFallbacks(edited) {
  const acs = new Set();
  for (const el of edited) for (let n = el.parentNode; n && n.nodeType === 1; n = n.parentNode) if (n.namespaceURI === NS.MC && n.localName === 'Choice') acs.add(n.parentNode);
  const depth = (n) => { let d = 0; for (; n; n = n.parentNode) d++; return d; };
  const outer = (root) => [...root.getElementsByTagNameNS(NS.WORD, 'txbxContent')].filter((tb) => { for (let n = tb.parentNode; n && n !== root; n = n.parentNode) if (n.localName === 'txbxContent') return false; return true; });
  for (const ac of [...acs].sort((a, b) => depth(b) - depth(a))) { // 先内层后外层
    const pick = (name) => [...ac.childNodes].find((x) => x.namespaceURI === NS.MC && x.localName === name);
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
  for (const part of parts) for (const t of part.doc.getElementsByTagNameNS(NS.A, 't')) {
    if (!map.has(t.textContent)) continue;
    hit.add(t.textContent); t.textContent = map.get(t.textContent); part.dirty = true;
  }
  return [...map.keys()].filter((k) => !hit.has(k)).length;
}

function maxId(parts) {
  let m = 0;
  for (const p of parts) for (const el of p.doc.getElementsByTagName('*')) { const v = el.getAttributeNS(NS.WORD, 'id'); if (v && /^\d+$/.test(v)) m = Math.max(m, +v); }
  return m;
}

// ---------- 投影 ----------
export class DocxProjection {
  md = '';
  tables = []; // 表格在 md 里的区间 [{s, e, full（含后面的空行）, el, part, rows: [{s, e, el}]}]（特殊表格没有 rows，只能整表删）
  images = []; // 图片 / 图表 / SmartArt 在 md 里的整段 [{s, e, pid, off, path（md 里的写法）, file（包内路径，外链为 null）, part, locked}]
  warnings = [];
  unhandled = []; // 投影不了的元素 [{tag: 'm:oMath', count}]，也会汇总进 warnings
  #opts = null; // { marks, author, i18n }
  #zip = null;
  #parts = new Map(); // path -> {path, doc, dirty}（含 .rels）
  #textParts = []; // 正文 + 页眉页脚脚注尾注
  #units = []; // 段落 {kind:'para', el, part, atoms} 或文本单元 {kind:'text', part, atoms, set}
  #map = null; // md 每个字符 -> {pid 单元, off 单元内位置, ed 可编辑, aft 紧跟其后插入的位置}（见 MD）
  #stats = { del: 0, ins: 0, skipped: 0, fmt: 0 }; // 各次编辑累计，save 再按实际写入修正成 report
  #edits = new Edits();
  #mirrors = [];
  #swap = new Map(); // 媒体文件 -> 新内容
  #mainPath = '';
  #types = []; // 新加的部件要登记进 [Content_Types].xml 的 <Override>
  #wipe = false;
  #saved = false;

  // opts（默认见 markup.js PROJECTION）：marks 结构标记（makeMarks）；author save({revisions}) 的修订作者，载入时先拒绝掉这个作者的修订；
  // i18n 界面文字（见 I18N，按组覆盖）
  static async load(buffer, opts = {}) {
    const d = new DocxProjection();
    d.#opts = { ...PROJECTION, ...opts, i18n: mergeI18n(opts.i18n) };
    await d.#load(buffer);
    return d;
  }

  // md -> txt 投影（去掉转义与粗斜体标记），mapPlain 的 fn 拿到的就是它（或其中一段）
  get text() { return plainView(this.md, this.#opts.marks).text; }

  // md 第 i 个字符对应的位置；结构字符（Markdown 标记、表格竖线等）editable 为 false
  locate(i) {
    const pid = this.#map.pid[i];
    if (pid == null || pid < 0) return null;
    const unit = this.#units[pid];
    return { unit, atom: unit.atoms[this.#map.off[i]] ?? null, editable: this.#map.ed[i] === 1 };
  }

  // 三种编辑最终都走这里，返回本次的 {del, ins, skipped}。坐标都基于载入时的投影：多次调用在原文上叠加，后一次看不到前一次的结果

  // 按 md 区间替换 [{s, e, text}]：区间内可编辑字符全删，text 插在第一个被删的字处、用它的格式；s === e 为纯插入。
  // text 按纯文本写：\n 分段（连着几个算一个），\u2028 软换行，\t 制表。区间里段落之间的换行被删 = 与下一段合并
  splice(spans) {
    const map = this.#map, { pid, off, ed, brk } = map, st = { del: 0, ins: 0, skipped: 0, fmt: 0 };
    for (const { s, e, text: raw = '' } of spans) {
      let first = -1;
      const gone = coveredImgs(this.images, s, e);
      for (const x of gone) this.#edits.dropObj(x);
      for (let i = s; i < e; i++) if (ed[i] && !gone.some((x) => i >= x.s && i < x.e)) { if (first < 0) first = i; this.#edits.del(pid[i], off[i]); }
      const merged = this.#mergesIn(s, e);
      for (const p of merged) this.#edits.merge(p);
      let rows = 0; // 区间整个盖住的行 / 表格：连结构一起删（只清空文字会留下空表格，行号也对不上）
      for (const t of this.tables) {
        if (s <= t.s && e >= t.e) { this.#edits.tbls.set(t.el, t.part); rows++; continue; }
        for (const r of t.rows) if (s <= r.s && e >= r.e) { this.#edits.rows.set(r.el, t.part); rows++; }
      }
      if (first >= 0 || gone.length || merged.length || rows) st.del++; else if (e > s) st.skipped++; // 只有结构字符：删不掉
      let text = raw.replace(/\r\n?/g, '\n').replace(/\n+/g, PS).replaceAll(BR, '\n');
      if (!text) continue;
      let at = first >= 0 ? [pid[first], off[first]] : insertAnchor(map, s);
      if (!at && s > 0 && brk[s - 1] >= 0) { // 两段之间、后面没有可挂的段落（如正文末尾）：接在上一段末尾另起一段
        at = [brk[s - 1], this.#units[brk[s - 1]].atoms.length];
        text = PS + text.replace(/\u2029+$/, '');
      }
      if (at) { this.#edits.ins(at[0], at[1], text); st.ins++; } else st.skipped++;
    }
    return this.#count(st);
  }

  // 区间 [s, e) 里被删的段落标记 -> 要与下一段合并的段落。删的是「上一段的标记 + 下一段 Q 的全部文字」（如 \n\nQ）时改成删 Q 自己的标记：
  // 段落格式存在段落标记里，这样上一段保持原样，Q 整段消失
  #mergesIn(s, e) {
    const { pid, ed, brk } = this.#map, out = [];
    for (let i = s; i < e; i++) {
      const P = brk[i];
      if (P < 0 || out.includes(P) || out.includes(-P - 1)) continue;
      let k = i;
      while (k < brk.length && (brk[k] === P || (pid[k] < 0 && brk[k] < 0))) k++;
      const Q = pid[k] ?? -1;
      let q = k, any = false, all = true;
      for (; q < brk.length && brk[q] < 0; q++) if (ed[q]) { any = true; if (q >= e) all = false; }
      const swap = Q >= 0 && brk[q] === Q && q >= e && any && all;
      out.push(swap ? Q : P);
    }
    return out;
  }

  // 拿改过的整份 md 做 diff，差异作为区间替换交给 splice；没改的字若粗斜体变了，记成格式修改。
  // from：新 md 来自另一个 DocxProjection 时传入，只取它的可编辑字符（和整张图片）
  diffMd(newMd, { from = null } = {}) {
    // 来自另一个投影时，新 md 里有、本文档里没有的整张表格：摘出来不参与逐字比较（否则表格里的字会和别处相同的字错配、散落成插入文字），
    // 之后整表克隆进来（审阅模式下是插入的行）。判据：摘掉它之后差异更小
    const drop = [], miss = [], rowsBefore = this.#edits.rowIns.length;
    if (from && from.tables.length > this.tables.length) {
      const cost = (m) => diffSpans(this.md, m, IMG_DIFF).reduce((n, x) => n + (x.e - x.s) + x.text.length, 0);
      let cur = newMd, now = cost(cur), cut = 0;
      for (const t of from.tables) {
        if (t.part.path !== from.#mainPath || this.md.includes(newMd.slice(t.s, t.e))) continue;
        const a = t.s - cut, b = t.full - cut, next = cur.slice(0, a) + cur.slice(b), c = cost(next);
        if (c >= now) continue;
        drop.push([t.s, t.full]); miss.push(t);
        cur = next; now = c; cut += t.full - t.s;
      }
    }
    // 两边都有的表格里，新 md 有而本文档整行缺失的行：同样摘出来，之后整行克隆回对应的位置（插入的行）
    if (from) {
      const mine = (ts, f) => ts.filter((t) => t.part.path === f.#mainPath && !miss.includes(t)), ta = mine(from.tables, from), tb = mine(this.tables, this);
      if (ta.length === tb.length) ta.forEach((T, i) => this.#missingRows(T, tb[i], drop, newMd));
    }
    drop.sort((x, y) => x[0] - y[0]);
    const kept = (arr) => (drop.length ? Array.from(arr).filter((_, i) => !drop.some(([a, b]) => i >= a && i < b)) : arr);
    const B = from && { ed: kept(from.#map.ed), brk: kept(from.#map.brk) };
    const content = contentOf(newMd, from && { ed: from.#map.ed, brk: from.#map.brk, images: from.images }, drop);
    for (const [a, b] of [...drop].reverse()) newMd = newMd.slice(0, a) + newMd.slice(b);
    const diffs = diffSpans(this.md, newMd, IMG_DIFF);
    for (const t of miss) { // 新 md 位置 at -> 本文档 md 位置：落在某处替换的文字里就取它的起点
      const at = t.s - drop.reduce((n, [a, b]) => (b <= t.s ? n + b - a : n), 0);
      let delta = 0, pos = -1;
      for (const x of diffs) {
        if (at <= x.at) break;
        if (at < x.at + x.text.length) { pos = x.s; break; }
        delta += x.e - x.s - x.text.length;
      }
      this.#edits.tblIns.push({ pos: pos >= 0 ? pos : at + delta, el: t.el });
    }
    const spans = [], st = { del: 0, ins: 0, skipped: 0, fmt: 0 };
    for (const x of diffs) {
      if (/^\**$/.test(this.md.slice(x.s, x.e)) && /^\**$/.test(x.text)) continue; // 只动了粗斜体标记：下面按格式比
      const text = content(x.at, x.text.length);
      if (x.text && !text && x.e === x.s) st.skipped++; // 只插入了结构字符（如表格竖线）
      spans.push({ s: x.s, e: x.e, text });
    }
    // 两边都有的字：比粗斜体（空白不比，投影里空白沿用相邻字的格式）
    const { pid, off, ed } = this.#map, fa = emphasisOf(this.md), fb = emphasisOf(newMd);
    let prev = 0, delta = 0, run = false;
    const same = (a, b) => {
      for (let i = a; i < b; i++) {
        const j = i + delta, hit = ed[i] && (!B || B.ed[j]) && !/\s/.test(this.md[i]) && fa[i] !== fb[j];
        if (hit) { this.#edits.fmt(pid[i], off[i], fb[j]); if (!run) st.fmt++; }
        if (ed[i]) run = hit;
      }
    };
    for (const x of diffs) { same(prev, x.s); run = false; delta += x.text.length - (x.e - x.s); prev = x.e; }
    same(prev, this.md.length);
    this.#count(st);
    const r = this.splice(spans);
    r.skipped += st.skipped; r.fmt = st.fmt;
    const rowsIn = this.#edits.rowIns.length - rowsBefore;
    r.ins += miss.length + rowsIn; this.#stats.ins += miss.length + rowsIn;
    return r;
  }

  // 用函数改纯文本投影：fn(text) -> 新 text，差异交给 splice。range=[s, e]（md 坐标，如 lineRange(d.md, 11, 22)）时 fn 只拿到这一段
  mapPlain(fn, { range } = {}) {
    return this.splice(plainToSpans(this.md, fn, range, this.#opts.marks));
  }

  // 媒体文件 word/media/*：[{path, name, ext, size, bytes, mime, replaceable, uses: [{s, e}]}]；replaceable 表示 blankImages 能换成占位小图，uses 是它在 md 里的各处。须在 save 之前调用
  async media() {
    const out = [], zip = this.#zip;
    for (const path of Object.keys(zip.files)) {
      const m = /^word\/media\/([^/]+)\.([A-Za-z0-9]+)$/.exec(path);
      if (!m || zip.files[path].dir) continue;
      const ext = m[2].toLowerCase(), bytes = await zip.file(path).async('uint8array'), mime = MIME[ext] || '';
      const uses = this.images.filter((x) => x.file === path).map(({ s, e }) => ({ s, e }));
      out.push({ path, name: m[1] + '.' + ext, ext, size: bytes.length, bytes, mime, replaceable: !!(mime && PLACEHOLDER[ext === 'jpeg' ? 'jpg' : ext]), uses });
    }
    return out.sort((a, b) => a.path.localeCompare(b.path, 'en', { numeric: true }));
  }

  // 图片换成 bytes：非审阅模式直接换掉媒体文件（所有引用一起换）；审阅模式下正文里每处都是「原图删除 + 指向新媒体文件的副本插入」。
  // list：[{ path: 'word/media/image1.png', bytes }]；返回正文里涉及的处数
  replaceImages(list) {
    for (const { path, bytes } of list) this.#swap.set(path, bytes);
    return this.images.filter((x) => this.#swap.has(x.file)).length;
  }
  // 换成 1×1 占位图（svg / emf 等换不了的跳过）
  blankImages(paths) {
    return this.replaceImages(paths.flatMap((path) => { const ext = path.split('.').pop().toLowerCase(), b = PLACEHOLDER[ext === 'jpeg' ? 'jpg' : ext]; return b ? [{ path, bytes: b64(b) }] : []; }));
  }
  wipeMetadata() { this.#wipe = true; return this; } // 抹掉作者、批注、修订记录、rsid 等不可见指纹

  // revisions：写成审阅修订（作者 author）；type 同 JSZip.generateAsync。DOM 会被改写，只能调用一次。
  // 返回 { blob | data, report: {del, ins, skipped}（按实际写入修正过的累计数）, warnings }
  async save({ revisions = false, author = this.#opts.author, type = 'blob' } = {}) {
    if (this.#saved) throw new Error(this.#opts.i18n.warn.savedTwice);
    this.#saved = true;
    const warnings = [...this.warnings], stats = { ...this.#stats }, t = this.#opts.i18n, track = revisions;
    const ctx = { author, date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), nextId: maxId(this.#textParts) + 1, track };
    ctx.meta = () => ({ id: String(ctx.nextId++), author: ctx.author, date: ctx.date });
    let docPr = 0; // wp:docPr 的 id（与 w:id 是两套编号）；复制出来的图片要换新 id
    for (const part of this.#textParts) for (const el of part.doc.getElementsByTagNameNS(NS.WP, 'docPr')) docPr = Math.max(docPr, +el.getAttribute('id') || 0);
    ctx.renew = (r) => { for (const el of r.getElementsByTagNameNS(NS.WP, 'docPr')) el.setAttribute('id', String(++docPr)); };
    const edited = [], mirror = new Map(), later = [];
    let untracked = 0, failed = 0, foreign = 0;
    const merges = [];
    const added = this.#pairImages(ctx);
    // 文本单元（链接地址、图片说明、图表文字…）：非审阅模式直接改；审阅模式下图片说明改成「删原图 + 插改了说明的副本」，其余等段落改完再处理
    const reimg = new Map(); // 图片原子 -> [在副本上做的改动]
    const onCopy = (a, fn) => { if (!reimg.has(a)) reimg.set(a, []); reimg.get(a).push(fn); };
    for (const [pid, ed] of this.#edits) {
      const u = this.#units[pid];
      if (u.kind !== 'text') continue;
      const v = textAfter(u, ed);
      if (!track) { u.set(v); if (u.mirror) mirror.set(u.orig, v); u.part.dirty = true; continue; }
      const a = u.attr && u.obj.atom;
      if (a && !a.locked && !this.#edits.get(a.pid)?.drop.has(a.off)) onCopy(a, (r) => u.setOn(r, v));
      else later.push([u, v]);
    }
    if (track) for (const [path, bytes] of this.#swap) { // 换图：副本指向新媒体文件（每个部件各加一条关系）
      const hits = this.images.filter((x) => x.file === path && !x.locked && !this.#edits.get(x.pid)?.drop.has(x.off));
      if (!hits.length || !this.#zip.file(path)) continue;
      const np = this.#newMedia(path, bytes), ids = new Map();
      for (const x of hits) {
        const a = this.#units[x.pid].atoms[x.off], [ns, tag, attr] = a.obj.ref;
        if (!ids.has(x.part)) ids.set(x.part, this.#addRel(this.#parts.get(relsPathOf(x.part.path)), REL_TYPE + 'image', relTarget(x.part.path, np), false, x.part.path));
        onCopy(a, (r) => own(r, ns, tag)?.setAttributeNS(NS.R, attr, ids.get(x.part)));
      }
    }
    for (const [a, fns] of reimg) {
      const ped = this.#edits.of(a.pid), r = a.r.cloneNode(false), rpr = kid(a.r, 'rPr');
      if (rpr) r.appendChild(rpr.cloneNode(true));
      r.appendChild(a.item.cloneNode(true));
      for (const fn of fns) fn(r);
      ctx.renew(r);
      ped.drop.add(a.off);
      ped.ins.set(a.off, OBJ + (ped.ins.get(a.off) || ''));
      ped.nodes.set(a.off, [r, ...(ped.nodes.get(a.off) || [])]);
    }
    for (const [pid, ed] of this.#edits) {
      const u = this.#units[pid];
      if (u.kind === 'text') continue;
      const backup = u.el.cloneNode(true);
      ctx.created = [];
      try {
        const { dropped, frags } = applyParaEdits(u, ed, ctx);
        stats.ins = Math.max(0, stats.ins - dropped); stats.skipped += dropped; foreign += dropped;
        edited.push(...frags); u.part.dirty = true;
        if (ed.merge) merges.push([pid, frags.at(-1)]);
      } catch (e) {
        console.error(e); failed++;
        for (const x of ctx.created) x.remove();
        u.el.replaceWith(backup); // 半途失败的段落整段还原
      }
    }
    dropRows(this.#edits, ctx);
    this.#insertRows(ctx);
    const lost = this.#insertTables(ctx);
    if (lost) { stats.ins = Math.max(0, stats.ins - lost); stats.skipped += lost; warnings.push(t.warn.tableLost(lost)); }
    let unmerged = 0;
    for (const [, el] of merges.sort((a, b) => a[0] - b[0])) if (!mergeMark(el, ctx)) unmerged++; // 按文档顺序合并
    // 审阅模式下的链接地址：整个链接删 + 插；图表 / SmartArt / 复杂域链接没有修订语法：直接改，在所在段落加批注写明原文
    const notes = new Map(), mainDoc = this.#parts.get(this.#mainPath)?.doc;
    for (const [u, v] of later) {
      if (u.link) {
        const id = this.#addRel(u.link.rels, u.link.type, v, true);
        for (const h of [...u.link.part.doc.getElementsByTagNameNS(NS.WORD, 'hyperlink')]) if (h.getAttributeNS(NS.R, 'id') === u.link.id) reviseLink(h, (el) => el.setAttributeNS(NS.R, 'r:id', id), ctx);
        u.link.part.dirty = true;
        continue;
      }
      if (u.field) {
        reviseLink(u.field, (el) => el.setAttributeNS(NS.WORD, 'w:instr', (el.getAttributeNS(NS.WORD, 'instr') || '').replace(HYPER, (_, a, _u, c) => a + v + c)), ctx);
        u.part.dirty = true;
        continue;
      }
      const from = u.atoms.map((a) => a.ch).join('');
      u.set(v); u.part.dirty = true;
      if (u.mirror) mirror.set(u.orig, v);
      if (u.attr) continue; // 图片已删掉 / 在别人的修订里：直接改说明
      const p = closestP(u.anchor || u.obj?.atom?.item);
      if (!p || p.ownerDocument !== mainDoc) { untracked++; continue; }
      if (!notes.has(p)) notes.set(p, []);
      notes.get(p).push(t.comment(u.anchor ? t.parts.link : u.label, from, v));
    }
    for (const [p, lines] of notes) await this.#addComment(p, lines, ctx);
    if (unmerged) { warnings.push(t.warn.unmerged(unmerged)); stats.skipped += unmerged; }
    if (foreign) warnings.push(t.warn.foreign(foreign));
    if (added) { warnings.push(t.warn.newImages(added)); stats.ins = Math.max(0, stats.ins - added); stats.skipped += added; }
    if (untracked) warnings.push(t.warn.untracked(untracked));
    if (failed) { warnings.push(t.warn.failed(failed)); stats.skipped += failed; }
    syncFallbacks(edited);
    const miss = mirror.size ? mirrorSmartArt(this.#mirrors, mirror) : 0;
    if (miss) warnings.push(t.warn.smartArt(miss));
    const wiped = new Set(this.#wipe ? this.#textParts : []);
    for (const part of wiped) cleanDom(part.doc);
    for (const part of this.#parts.values()) {
      if (!part || !(part.dirty || wiped.has(part))) continue;
      let xml = new XMLSerializer().serializeToString(part.doc);
      if (wiped.has(part)) xml = xml.replace(FP_ATTR, '');
      this.#zip.file(part.path, /^<\?xml/.test(xml) ? xml : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml);
    }
    if (!track) for (const [path, bytes] of this.#swap) if (this.#zip.file(path)) this.#zip.file(path, bytes);
    if (this.#types.length) {
      const f = this.#zip.file('[Content_Types].xml');
      let x = (await f.async('string')).replace(/<Types([^>]*)\/>/, '<Types$1></Types>');
      for (const [part, type] of this.#types) if (!x.includes(`PartName="/${part}"`)) x = x.replace('</Types>', `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
      this.#zip.file('[Content_Types].xml', x);
    }
    if (this.#wipe) await wipeParts(this.#zip);
    const out = await this.#zip.generateAsync({ type, compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    return { [type === 'blob' ? 'blob' : 'data']: out, report: stats, warnings };
  }

  // 插入文本里的图片引用：优先配对到被删掉的同路径图片（同一部件内，按出现顺序）= 移动；
  // 没有可配的就复制文档里已有的同路径图片（关系 / 媒体文件现成，克隆 drawing 即可）；都没有的是新增，忽略。返回被忽略的个数。没被放回的被删图片就是真删除
  #pairImages(ctx) {
    const objRun = (img) => {
      const a = this.#units[img.pid].atoms[img.off], r = a.r.cloneNode(false), rpr = kid(a.r, 'rPr');
      if (rpr) r.appendChild(rpr.cloneNode(true));
      r.appendChild(a.item.cloneNode(true));
      return r;
    };
    const gone = this.#edits.dropped.map((img) => ({ img, r: objRun(img), used: false }));
    let ignored = 0;
    const { renew } = ctx;
    for (const [pid, ed] of this.#edits) {
      for (const [at, paths] of ed.refs) {
        const u = this.#units[pid], nodes = [];
        for (const path of paths) {
          const same = (x) => x.path === path && x.part === u.part;
          const move = u.kind === 'para' && gone.find((x) => !x.used && same(x.img));
          if (move) { move.used = true; if (ctx.track) renew(move.r); nodes.push(move.r); continue; } // 审阅模式下原图留在删除线里，副本换新 id
          const src = u.kind === 'para' && this.images.find(same);
          if (!src) { ignored++; nodes.push(null); continue; }
          const r = objRun(src); renew(r); nodes.push(r); // 复制：原图还在，必须换新 id
        }
        ed.nodes.set(at, nodes);
      }
    }
    return ignored;
  }

  // 整表插入（见 diffMd）：克隆到 pos 前面最近的块之后（文档开头则放在后面第一个块之前）。审阅模式下每行、每个段落标记、每个 run 都标成插入。返回放不下的个数
  #insertTables(ctx) {
    const map = this.#map, main = this.#parts.get(this.#mainPath), last = new Map();
    let lost = 0;
    const unitAt = (i) => (map.pid[i] >= 0 ? map.pid[i] : map.brk[i] >= 0 ? map.brk[i] : -1);
    const block = (i) => {
      const u = this.#units[i];
      if (u?.kind !== 'para' || u.part !== main) return null;
      let b = u.el;
      while (b.parentNode && !isW(b.parentNode, 'body')) b = b.parentNode;
      return b.parentNode ? b : null;
    };
    for (const { pos, el } of [...this.#edits.tblIns].sort((a, b) => a.pos - b.pos)) {
      let after = true, b = null;
      for (let i = Math.min(pos, map.pid.length) - 1; i >= 0 && !b; i--) { const k = unitAt(i); if (k >= 0) b = block(k); }
      if (!b) { after = false; for (let i = pos; i < map.pid.length && !b; i++) { const k = unitAt(i); if (k >= 0) b = block(k); } }
      if (!b) { lost++; continue; }
      const tbl = main.doc.importNode(el, true);
      if (ctx.track) markInserted(tbl, ctx);
      const prev = last.get(b);
      if (after) b.parentNode.insertBefore(tbl, (prev || b).nextSibling); else b.parentNode.insertBefore(tbl, b);
      if (after) last.set(b, tbl);
      main.dirty = true;
    }
    return lost;
  }

  // 表格 T（新 md 里的）比本文档的对应表格 U 多出的整行：按相似度把 U 的每一行对到 T 的一行（保持顺序，表头对表头），没被对到的就是缺的。
  // 缺的行的 md 区间记进 drop，克隆来源记进 rowIns。只处理列数相同、没有被删表头的普通表格
  #missingRows(T, U, drop, newMd) {
    const n = T.rows.length, m = U.rows.length;
    if (!m || n <= m || !T.cols || T.cols !== U.cols || n * m > 4000) return;
    // 按格比：整格相同记 1（标签格一般原样保留），否则按逐字差异记 0~0.5。整行逐字比会偏向短行
    const cells = (md, r) => md.slice(r.s, r.e).split('\n')[0].slice(2, -2).split(' | ');
    const sim = (a, b) => {
      const x = cells(newMd, T.rows[a]), y = cells(this.md, U.rows[b]);
      return x.reduce((sum, c, k) => {
        const d = y[k] ?? '';
        if (c === d) return sum + 1;
        return sum + 0.5 * (1 - diffSpans(c, d).reduce((n, z) => n + (z.e - z.s) + z.text.length, 0) / (c.length + d.length || 1));
      }, 0);
    };
    const dp = Array.from({ length: n }, () => new Array(m).fill(-Infinity)), pick = Array.from({ length: n }, () => new Array(m).fill(false));
    dp[0][0] = 0; // 表头对表头
    for (let i = 1; i < n; i++) for (let j = 0; j <= Math.min(i, m - 1); j++) {
      const skip = dp[i - 1][j], take = j ? dp[i - 1][j - 1] + sim(i, j) : -Infinity;
      if (take > skip) { dp[i][j] = take; pick[i][j] = true; } else dp[i][j] = skip;
    }
    if (dp[n - 1][m - 1] === -Infinity) return;
    const at = new Array(n).fill(-1);
    for (let i = n - 1, j = m - 1; i >= 0 && j >= 0; i--) if (pick[i][j] || (i === 0 && j === 0)) at[i] = j--;
    let prev = 0;
    for (let i = 1; i < n; i++) {
      if (at[i] >= 0) { prev = at[i]; continue; }
      drop.push([T.rows[i].s, T.rows[i].e + 1]);
      this.#edits.rowIns.push({ after: U.rows[prev].el, el: T.rows[i].el });
    }
  }

  // 整行插入（见 #missingRows）：克隆到对应行之后，同一行后面的按原顺序排。审阅模式下标成插入
  #insertRows(ctx) {
    const last = new Map();
    for (const { after, el } of this.#edits.rowIns) {
      const row = after.ownerDocument.importNode(el, true);
      if (ctx.track) markInserted(row, ctx);
      after.parentNode.insertBefore(row, (last.get(after) || after).nextSibling);
      last.set(after, row);
      const part = this.#textParts.find((x) => x.doc === after.ownerDocument);
      if (part) part.dirty = true;
    }
  }

  #count(st) { for (const k in st) this.#stats[k] += st[k]; return st; }

  // 媒体文件 path 的新版本另存一份（名字加 _new），返回包内路径
  #newMedia(path, bytes) {
    let np = path.replace(/(\.[^./]+)?$/, '_new$1');
    for (let n = 2; this.#zip.file(np); n++) np = path.replace(/(\.[^./]+)?$/, `_new${n}$1`);
    this.#zip.file(np, bytes);
    const mime = MIME[np.split('.').pop().toLowerCase()];
    if (mime) this.#types.push([np, mime]);
    return np;
  }

  // 在关系部件 rels（{path, doc}，没有则按 owner 部件新建）里加一条关系，返回新 Id
  #addRel(rels, type, target, external = false, owner = this.#mainPath) {
    if (!rels) {
      const path = relsPathOf(owner);
      rels = this.#parts.get(path);
      if (!rels) { rels = { path, doc: new DOMParser().parseFromString(`<Relationships xmlns="${NS.PKG}"/>`, 'application/xml'), dirty: true }; this.#parts.set(path, rels); }
    }
    const ids = new Set([...rels.doc.getElementsByTagNameNS(NS.PKG, 'Relationship')].map((r) => r.getAttribute('Id')));
    let n = 1;
    while (ids.has('rIdDi' + n)) n++;
    const el = rels.doc.createElementNS(NS.PKG, 'Relationship');
    el.setAttribute('Id', 'rIdDi' + n);
    el.setAttribute('Type', type);
    el.setAttribute('Target', target);
    if (external) el.setAttribute('TargetMode', 'External');
    rels.doc.documentElement.appendChild(el);
    rels.dirty = true;
    return 'rIdDi' + n;
  }

  // 给正文段落 p 整段加一条批注（lines 每行一段）；没有批注部件就新建
  async #addComment(p, lines, ctx) {
    const relsPath = relsPathOf(this.#mainPath), rels = this.#parts.get(relsPath);
    let rel = rels && [...rels.doc.getElementsByTagNameNS(NS.PKG, 'Relationship')].find((r) => r.getAttribute('Type') === REL_TYPE + 'comments');
    let part = rel && this.#parts.get(resolvePath(this.#mainPath, rel.getAttribute('Target')));
    if (rel && !part) { const path = resolvePath(this.#mainPath, rel.getAttribute('Target')); const doc = await loadXml(this.#zip, path, () => {}); if (doc) this.#parts.set(path, (part = { path, doc, dirty: false })); }
    if (!part) {
      const path = this.#mainPath.replace(/[^/]*$/, 'comments.xml');
      part = { path, doc: new DOMParser().parseFromString(`<w:comments xmlns:w="${NS.WORD}"/>`, 'application/xml'), dirty: true };
      this.#parts.set(path, part);
      this.#addRel(rels, REL_TYPE + 'comments', 'comments.xml');
      this.#types.push([path, 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml']);
    }
    const doc = p.ownerDocument, cdoc = part.doc, id = String(ctx.nextId++);
    const c = cdoc.createElementNS(NS.WORD, 'w:comment');
    for (const [k, v] of [['id', id], ['author', ctx.author], ['date', ctx.date]]) c.setAttributeNS(NS.WORD, 'w:' + k, v);
    for (const line of lines) { const cp = mkEl(cdoc, 'p'), r = mkEl(cdoc, 'r'); r.appendChild(mkText(cdoc, 't', line)); cp.appendChild(r); c.appendChild(cp); }
    cdoc.documentElement.appendChild(c);
    part.dirty = true;
    const pPr = kid(p, 'pPr');
    p.insertBefore(mkEl(doc, 'commentRangeStart', { id }), pPr ? pPr.nextSibling : p.firstChild);
    p.appendChild(mkEl(doc, 'commentRangeEnd', { id }));
    const ref = mkEl(doc, 'r');
    ref.appendChild(mkEl(doc, 'commentReference', { id }));
    p.appendChild(ref);
  }

  async #load(buffer) {
    const zip = (this.#zip = await JSZip.loadAsync(buffer));
    const { warnings } = this, parts = this.#parts, units = this.#units, { marks, author, i18n: t } = this.#opts;
    const load = async (path) => {
      if (!path) return null;
      if (!parts.has(path)) { const doc = await loadXml(zip, path, (p) => warnings.push(t.warn.parseFailed(p))); parts.set(path, doc ? { path, doc, dirty: false } : null); }
      return parts.get(path);
    };
    const relsOf = new Map(); // 部件 path -> Map(rId -> rel)
    const rels = async (path) => {
      if (relsOf.has(path)) return relsOf.get(path);
      const m = new Map();
      relsOf.set(path, m);
      const rp = await load(relsPathOf(path));
      for (const el of rp ? rp.doc.getElementsByTagNameNS(NS.PKG, 'Relationship') : []) {
        const external = el.getAttribute('TargetMode') === 'External', target = el.getAttribute('Target') || '';
        m.set(el.getAttribute('Id'), { el, part: rp, type: (el.getAttribute('Type') || '').split('/').pop(), external, target, path: external ? target : resolvePath(path, target) });
      }
      return m;
    };
    const mainPath = (this.#mainPath = [...(await rels('')).values()].find((r) => r.type === 'officeDocument')?.path || 'word/document.xml');
    const main = await load(mainPath);
    const mainRels = await rels(mainPath);
    const ofType = (t) => [...new Set([...mainRels.values()].filter((r) => r.type === t && !r.external).map((r) => r.path))].sort((a, b) => a.localeCompare(b, 'en', { numeric: true }));
    const style = styleLookup((await load(ofType('styles')[0]))?.doc);
    const lists = new Lists((await load(ofType('numbering')[0]))?.doc, style);
    const others = [];
    for (const type of TEXT_PARTS) for (const p of ofType(type)) { const part = await load(p); if (part) others.push({ type, part }); }
    this.#textParts = [main, ...others.map((o) => o.part)].filter(Boolean);
    for (const part of this.#textParts) { await rels(part.path); if (rejectOurs(part.doc, author)) part.dirty = true; }

    const unit = (u) => { u.idx = units.length; units.push(u); return u; };
    const textUnit = (part, text, set, extra) => unit({ kind: 'text', part, atoms: [...text].map((ch) => ({ ch, edit: true })), set, ...extra });
    const unhandled = new Map(); // 标签名 -> 次数
    const strays = new Set(), bump = (tag) => unhandled.set(tag, (unhandled.get(tag) || 0) + 1);
    const stray = (el) => { if (!NO_TEXT.has(el.localName) && (el.localName === 'altChunk' || hasText(el))) { strays.add(el); bump(el.nodeName); } };
    const found = []; // 正文里发现的文本框 / 图表 / SmartArt，放到「其他文本」区
    const linkUnits = new Map();
    const relLink = (h, part) => {
      const r = relsOf.get(part.path)?.get(h.getAttributeNS(NS.R, 'id'));
      if (!r?.external) return null; // 文档内锚点（目录等）不显示成链接
      if (!linkUnits.has(r.el)) linkUnits.set(r.el, { unit: textUnit(r.part, r.target, (v) => r.el.setAttribute('Target', v), { link: { rels: r.part, id: h.getAttributeNS(NS.R, 'id'), type: r.el.getAttribute('Type'), part } }) });
      return linkUnits.get(r.el);
    };
    // field：w:fldSimple（审阅模式下整个换成删 + 插）；anchor：复杂域的 separate（只能直接改 + 批注）
    const fieldLink = (get, set, part, extra) => {
      const m = HYPER.exec(get());
      return m ? { unit: textUnit(part, m[2], (v) => set(get().replace(HYPER, (_, a, _u, c) => a + v + c)), extra) } : null;
    };
    const object = (d, part) => {
      const rel = (id) => (id ? relsOf.get(part.path)?.get(id) : null);
      let label = '', target = null, embed = '', ref = null; // ref：图片指向媒体的属性（换图时改在副本上）
      const blip = own(d, NS.A, 'blip'), vimg = own(d, NS.V, 'imagedata'), gd = own(d, NS.A, 'graphicData');
      if (blip) { ref = blip.getAttributeNS(NS.R, 'embed') ? [NS.A, 'blip', 'r:embed'] : [NS.A, 'blip', 'r:link']; target = rel(blip.getAttributeNS(NS.R, ref[2].slice(2))); }
      else if (vimg) { ref = [NS.V, 'imagedata', 'r:id']; target = rel(vimg.getAttributeNS(NS.R, 'id')); }
      else if (gd) {
        const uri = gd.getAttribute('uri') || '';
        if (uri.endsWith('/chart')) { label = t.parts.chart; embed = 'chart'; target = rel(own(gd, NS.C, 'chart')?.getAttributeNS(NS.R, 'id')); }
        else if (uri.endsWith('/diagram')) { label = t.parts.smartart; embed = 'smartart'; target = rel(own(gd, NS.DGM, 'relIds')?.getAttributeNS(NS.R, 'dm')); }
      }
      if (!target) return null;
      // 图片说明 / 标题：setOn(root, 新值) 改 root 里的 docPr（及同值的 pic:cNvPr），审阅模式下改在图片副本上
      const attr = (name) => {
        const v = own(d, NS.WP, 'docPr')?.getAttribute(name);
        if (!v) return null;
        const setOn = (root, nv) => { const dp = own(root, NS.WP, 'docPr'), cn = own(root, NS.PIC, 'cNvPr'); if (cn && cn.getAttribute(name) === dp?.getAttribute(name)) cn.setAttribute(name, nv); dp?.setAttribute(name, nv); };
        return textUnit(part, v, (nv) => setOn(d, nv), { attr: name, setOn });
      };
      const o = { label, path: target.external ? target.target : target.path.replace(/^word\//, ''), file: target.external ? null : target.path, ref, alt: attr('descr'), title: attr('title') };
      for (const u of [o.alt, o.title]) if (u) u.obj = o;
      if (embed && !target.external) found.push({ embed, path: target.path, obj: o });
      return o;
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
          const obj = (d) => { const o = d && object(d, part); if (o) { atoms.push({ ...base, ch: '\uFFFC', obj: o, edit: false }); o.atom = atoms.at(-1); } };
          if (c.namespaceURI === NS.MC && c.localName === 'AlternateContent') {
            const ch = [...c.childNodes].find((x) => x.localName === 'Choice');
            obj(ch && [...ch.childNodes].find((x) => isW(x, 'drawing')));
            continue;
          }
          if (c.namespaceURI !== NS.WORD) { stray(c); continue; }
          switch (c.localName) {
            case 't': for (const ch of c.textContent) atoms.push({ ...base, ch, edit: !locked }); break;
            case 'tab': atoms.push({ ...base, ch: '\t', edit: !locked }); break;
            case 'br': case 'cr': atoms.push({ ...base, ch: '\n', edit: false }); break;
            case 'noBreakHyphen': atoms.push({ ...base, ch: '-', edit: false }); break;
            case 'drawing': case 'pict': obj(c); break;
            case 'instrText': fields.at(-1)?.instr.push(c); break;
            case 'fldChar': {
              const t = c.getAttributeNS(NS.WORD, 'fldCharType');
              if (t === 'begin') fields.push({ instr: [], link: null });
              else if (t === 'end') fields.pop();
              else if (t === 'separate' && fields.length) {
                const f = fields.at(-1);
                f.link = fieldLink(() => f.instr.map((n) => n.textContent).join(''), (v) => { f.instr.forEach((n, i) => { n.textContent = i ? '' : v; }); }, part, { anchor: c });
              }
              break;
            }
            default: stray(c);
          }
        }
      };
      const walk = (n, locked) => {
        for (const c of n.childNodes) {
          if (c.nodeType !== 1) continue;
          if (c.namespaceURI !== NS.WORD) { stray(c); continue; }
          const ln = c.localName;
          if (ln === 'r') addRun(c, locked);
          else if (ln === 'ins' || ln === 'moveTo') walk(c, true);
          else if (ln === 'hyperlink' || ln === 'fldSimple') {
            const prev = link;
            link = (ln === 'hyperlink' ? relLink(c, part) : fieldLink(() => c.getAttributeNS(NS.WORD, 'instr') || '', (v) => c.setAttributeNS(NS.WORD, 'w:instr', v), part, { field: c })) || link;
            walk(c, locked);
            link = prev;
          } else if (INLINE_WRAP.has(ln)) walk(c, locked);
          else stray(c);
        }
      };
      walk(p, false);
      return atoms;
    };

    const md = new MD(marks.ambig);
    const mkPara = (p, part) => {
      const atoms = collect(p, part);
      if (!atoms.length) return null;
      const u = unit({ kind: 'para', el: p, part, atoms });
      atoms.forEach((a, i) => { a.pid = u.idx; a.off = i; });
      return u;
    };
    const listPrefix = (lab, indent) => (lab ? (indent ? '  '.repeat(lab.ilvl) : '') + lab.text + ' ' : '');
    // 段落 p 里的文本框（只取直属 p 的一层；文本框里再套的，等遍历到外框里的段落时再找）
    const scanTextboxes = (p, part) => {
      for (const tb of p.getElementsByTagNameNS(NS.WORD, 'txbxContent')) {
        if (inFallback(tb)) continue;
        let outer = true;
        for (let n = tb.parentNode; n && n !== p; n = n.parentNode) if (n.localName === 'txbxContent') { outer = false; break; }
        if (outer) found.push({ title: t.parts.textbox, container: tb, part });
      }
    };
    const walkBlocks = (container, part) => {
      for (const c of container.childNodes) {
        if (c.nodeType !== 1) continue;
        if (c.namespaceURI !== NS.WORD) { stray(c); continue; }
        if (c.localName === 'p') {
          scanTextboxes(c, part);
          const lab = lists.label(c); // 空段落也要计数
          const para = mkPara(c, part);
          if (!para) continue;
          const h = headingLevel(c, style);
          emitPara(md, para, h ? '#'.repeat(h) + ' ' + listPrefix(lab) : listPrefix(lab, true), true);
          md.mark('\n\n', para);
        } else if (c.localName === 'sdt') {
          const sc = kid(c, 'sdtContent');
          if (sc) walkBlocks(sc, part);
        } else if (c.localName === 'customXml') walkBlocks(c, part);
        else if (c.localName === 'tbl') {
          const grid = plainTable(c);
          if (grid) {
            const trs = unwrap(c, 'tr'), rec = { s: md.length, e: 0, el: c, part, cols: grid[0].length, rows: [] };
            grid.forEach((row, ri) => {
              const rs = md.length;
              row.forEach((p, ci) => {
                scanTextboxes(p, part);
                const lab = lists.label(p), para = mkPara(p, part);
                md.s(ci ? ' |' : '|');
                if (para) { md.s(' '); emitPara(md, para, listPrefix(lab)); }
                else md.push(' ', unit({ kind: 'para', el: p, part, atoms: [] }).idx, 0, 0, 0); // 空单元格：竖线后的空格挂着它，在这里输入能写进去
              });
              md.s(' |\n');
              if (ri === 0) md.s('|' + ' --- |'.repeat(row.length) + '\n');
              rec.rows.push({ s: rs, e: md.trim(md.length), el: trs[ri] });
            });
            rec.e = md.trim(md.length);
            md.s('\n');
            rec.full = md.length;
            md.tbls.push(rec);
          } else {
            const start = md.length;
            md.s(marks.specialBegin);
            const inner = unwrap(c, 'tr').flatMap((tr) => unwrap(tr, 'tc'));
            const before = md.length;
            for (const tc of inner) walkCell(tc, part);
            if (md.length === before) md.cut(start);
            else { md.s(marks.specialEnd); md.tbls.push({ s: start, e: md.trim(md.length), full: md.length, el: c, part, cols: 0, rows: [] }); }
          }
        } else stray(c);
      }
    };
    // 特殊表格单元格：每段一行，嵌套表格就地展平
    const walkCell = (tc, part) => {
      for (const c of tc.childNodes) {
        if (c.nodeType !== 1) continue;
        if (c.namespaceURI !== NS.WORD) { stray(c); continue; }
        if (c.localName === 'p') {
          scanTextboxes(c, part);
          const lab = lists.label(c), para = mkPara(c, part);
          if (para) { emitPara(md, para, listPrefix(lab), true); md.mark('\n', para); }
        } else if (c.localName === 'tbl') {
          for (const tr of unwrap(c, 'tr')) for (const t2 of unwrap(tr, 'tc')) walkCell(t2, part);
        } else if (c.localName === 'sdt' || c.localName === 'customXml') { const sc = c.localName === 'sdt' ? kid(c, 'sdtContent') : c; if (sc) walkCell(sc, part); }
        else stray(c);
      }
    };

    if (main) { const body = main.doc.getElementsByTagNameNS(NS.WORD, 'body')[0]; if (body) walkBlocks(body, main); }
    else warnings.push(t.warn.noMain(mainPath));

    // 其他文本放末尾：页眉页脚 / 脚注尾注，然后是正文里发现的文本框 / 图表 / SmartArt（处理中再发现的接着排）
    const queue = [];
    for (const { type, part } of others) {
      const title = t.parts[type] + (type === 'header' || type === 'footer' ? ' ' + baseName(part.path) : '');
      if (type === 'footnotes' || type === 'endnotes') {
        for (const fn of part.doc.documentElement.childNodes) {
          if (fn.nodeType !== 1 || ['separator', 'continuationSeparator', 'continuationNotice'].includes(fn.getAttributeNS(NS.WORD, 'type'))) continue;
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
      const nodeUnit = (el, mirror) => textUnit(part, el.textContent, (v) => { el.textContent = v; }, { obj: x.obj, label: x.embed === 'chart' ? t.parts.chart : t.parts.smartart, ...(mirror ? { mirror, orig: el.textContent } : {}) });
      if (x.embed === 'chart') {
        const els = [...part.doc.getElementsByTagName('*')].filter((el) => el.textContent.trim() && ((el.namespaceURI === NS.A && el.localName === 't') || (el.namespaceURI === NS.C && el.localName === 'v' && ['strCache', 'strLit'].includes(el.parentNode?.parentNode?.localName))));
        return { title: t.parts.chart + ' ' + baseName(x.path), units: els.map((el) => nodeUnit(el)) };
      }
      const els = [...part.doc.getElementsByTagNameNS(NS.A, 't')].filter((el) => el.textContent.trim());
      return { title: t.parts.smartart + ' ' + baseName(x.path), units: els.map((el) => nodeUnit(el, true)) };
    };
    let first = true;
    for (let i = 0; i < queue.length; i++) {
      const sec = await sectionOf(queue[i]);
      if (sec) {
        const start = md.length;
        if (first) md.s(marks.othersHead);
        md.s(marks.sectionHead(sec.title));
        const h = md.length;
        if (sec.container) walkBlocks(sec.container, sec.part);
        else if (sec.units.length) { for (const u of sec.units) { emitUnit(md, u, true); md.s('\n'); } md.s('\n'); }
        if (md.length === h) md.cut(start); else first = false;
      }
      queue.push(...found.splice(0));
    }
    for (const m of relsOf.values()) for (const r of m.values()) if (r.type === 'diagramDrawing' && !r.external) { const p = await load(r.path); if (p && !this.#mirrors.includes(p)) this.#mirrors.push(p); }

    // 兜底：文本部件里没进投影的文字（不在任何段落原子里，也不在删除修订、Fallback、上面已报过的元素里），以及批注，都按标签报进 unhandled
    const covered = new Set(), samples = new Map();
    for (const u of units) if (u.kind === 'para') for (const a of u.atoms) covered.add(a.item);
    const skip = (el) => { for (let n = el.parentNode; n && n.nodeType === 1; n = n.parentNode) if (strays.has(n) || n.localName === 'Fallback' || isW(n, 'del') || isW(n, 'moveFrom')) return true; return false; };
    for (const part of this.#textParts) for (const [ns, tag] of [[NS.WORD, 'w:t'], [NS.A, 'a:t']]) for (const el of part.doc.getElementsByTagNameNS(ns, 't')) {
      if (covered.has(el) || !el.textContent.trim() || skip(el)) continue;
      bump(tag);
      if (!samples.has(tag)) samples.set(tag, el.textContent.trim());
    }
    for (const p of ofType('comments')) for (const c of (await load(p))?.doc.getElementsByTagNameNS(NS.WORD, 'comment') || []) if (c.textContent.trim()) bump('w:comment');

    // 投影不到的内容：列出来提醒
    this.unhandled = [...unhandled].map(([tag, count]) => (samples.has(tag) ? { tag, count, sample: samples.get(tag) } : { tag, count }));
    if (unhandled.size) warnings.push(t.warn.unhandled(this.unhandled.map((u) => ({ ...u, note: t.notes[u.tag.split(':').pop()] }))));
    const emb = Object.keys(zip.files).filter((p) => /^word\/embeddings\/[^/]+$/.test(p));
    if (emb.length) warnings.push(t.warn.embeddings(emb.length));
    const vec = Object.keys(zip.files).filter((p) => /^word\/media\/[^/]+\.(?:svg|emf|wmf|tiff?)$/i.test(p));
    if (vec.length) warnings.push(t.warn.vectors(vec.length));

    this.images = md.imgs;
    this.tables = md.tbls;
    this.md = md.out.join('');
    this.#map = { pid: Int32Array.from(md.pid), off: Int32Array.from(md.off), ed: Uint8Array.from(md.ed), aft: Int32Array.from(md.aft), brk: Int32Array.from(md.brk) };
  }
}

// ---------- 图片 / 不可见元素 ----------

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
  const q = (n) => [...d.getElementsByTagNameNS(NS.WORD, n)];
  for (const n of ['proofErr', 'commentRangeStart', 'commentRangeEnd', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd',
    'rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'tcPrChange', 'trPrChange', 'tblGridChange', 'tblPrExChange', 'numberingChange']) q(n).forEach((x) => x.remove());
  q('commentReference').forEach((x) => x.parentNode?.parentNode && x.parentNode.remove());
  const gb = q('bookmarkStart').filter((x) => x.getAttributeNS(NS.WORD, 'name') === '_GoBack');
  const ids = new Set(gb.map((x) => x.getAttributeNS(NS.WORD, 'id')));
  gb.forEach((x) => x.remove());
  q('bookmarkEnd').filter((x) => ids.has(x.getAttributeNS(NS.WORD, 'id'))).forEach((x) => x.remove());
  [...q('del'), ...q('moveFrom')].forEach((x) => x.remove());
  for (const x of [...q('ins'), ...q('moveTo')]) {
    if (!x.parentNode) continue;
    if (['rPr', 'pPr'].includes(x.parentNode.localName)) { x.remove(); continue; }
    while (x.firstChild) x.parentNode.insertBefore(x.firstChild, x);
    x.remove();
  }
  q('r').filter((r) => { const p = kid(r, 'rPr'); return p && on(kid(p, 'vanish')); }).forEach((r) => r.remove());
}

async function wipeParts(zip) {
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
