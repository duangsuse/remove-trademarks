// 规则执行：扫描 Markdown → 候选项；选区解析；替换；章节。规则本身见 rules.js
import { RULES, ABBR_MAX, ABBR_RE, NAME_LABELS, NAME_VALUE, NAME_STOP, NER_GROUPS, CHAPTER_FALLBACK } from './rules.js';

function findAll(md, text, add) {
  if (!text) return;
  let i = md.indexOf(text);
  while (i >= 0) { add(i, i + text.length); i = md.indexOf(text, i + text.length); }
}

// 黑名单框的行：普通词，或行号模式 #L11~L222 / #L11
export function parseBlackLines(lines) {
  const words = [], ranges = [];
  for (const l of lines) {
    const t = l.trim();
    if (!t) continue;
    const m = /^#L(\d+)(?:\s*~\s*L?(\d+))?$/i.exec(t);
    if (m) ranges.push([+m[1], +(m[2] || m[1])]);
    else words.push(t);
  }
  return { words, ranges };
}

export function parseCustom(text) {
  const out = [];
  for (const ln of text.split('\n')) {
    const t = ln.trim();
    if (!t) continue;
    const i = t.indexOf('::');
    const name = i >= 0 ? t.slice(0, i).trim() : '自订', src = i >= 0 ? t.slice(i + 2).trim() : t;
    try { out.push({ name, re: new RegExp(src, 'gu') }); } catch { /* 无效正则忽略 */ }
  }
  return out;
}

// 含姓名的表格：标签格的下一格 / 特殊表格中标签行的下一行
function tableNames(md, put) {
  let pos = 0, inSpecial = false, label = false, ln = 0;
  const strip = (s) => s.replace(/\*+/g, '').replace(/[：:\s]/g, '');
  for (const line of md.split('\n')) {
    const start = pos;
    pos += line.length + 1; ln++;
    if (line.startsWith(':::特殊表格')) { inSpecial = true; label = false; continue; }
    if (line === ':::') { inSpecial = false; label = false; continue; }
    if (line.startsWith('| ') && line.endsWith(' |')) {
      if (/^\| (?:--- \| )+$|^\|(?: --- \|)+$/.test(line)) continue;
      let p = start + 2, prevLabel = false;
      for (const cell of line.slice(2, -2).split(' | ')) {
        const t = strip(cell);
        if (prevLabel && NAME_VALUE.test(t) && !NAME_STOP.test(t)) { const i = cell.indexOf(t); put(t, p + i, p + i + t.length); }
        prevLabel = NAME_LABELS.test(t);
        p += cell.length + 3;
      }
    } else if (inSpecial) {
      const t = strip(line);
      if (label && NAME_VALUE.test(t) && !NAME_STOP.test(t)) put(t, start, start + line.length, ln); // 特殊表格 1 格 1 行：按行号删，和章节一样
      label = NAME_LABELS.test(t);
    }
  }
}

// 模型常见的无效碎片：只剩公司后缀（"限公司"、"技有限公司"）、括号不成对（"中国）有限公司"）
const SUFFIX_ONLY = /^.{0,1}?(?:股份)?(?:有限责任|有限)?公司$/;
const ORGISH = new Set(['ORG', 'FAC', 'GPE', 'LOC']), NUM_TYPES = new Set(['DATE', 'MONEY', 'PERCENT']);
const junkNer = (t) => SUFFIX_ONLY.test(t) || !/[A-Za-z\u4e00-\u9fff]{2}/.test(t) || (t.match(/[（(]/g) || []).length !== (t.match(/[）)]/g) || []).length; // 另：至少要有连续 2 个字母/汉字（过滤 "137"、".0"）

// 返回 Map(key -> {key,text,cat,en,cls,warn,type,grp,score,occ:[[s,e]]})，不受勾选影响（勾选由调用方过滤，便于显示 ({num})）
export function scan(md, st) {
  const items = new Map();
  const put = (key, text, meta, s, e) => {
    let it = items.get(key);
    if (!it) items.set(key, (it = { key, text, ...meta, occ: [] }));
    it.occ.push([s, e]);
  };
  for (const [id, r] of Object.entries(RULES)) {
    r.re.lastIndex = 0;
    for (const m of md.matchAll(r.re)) {
      if (r.check && !r.check(m[0])) continue;
      const k = r.lead ? r.lead(m[0]) : 0, text = m[0].slice(k), at = m.index + k;
      put(id + ':' + text, text, { cat: id, en: r.en, cls: r.cls, warn: r.warn }, at, at + text.length);
    }
  }
  // 「…开发联合南通东华软件有限公司」：联合/合作 可能是名称本身的一部分，所以只有当后半段本身就是文中另一处识别出的公司全称（>=8 字）时才截
  const names = new Set([...items.values()].filter((it) => it.cat === 'company').map((it) => it.text));
  for (const it of [...items.values()]) {
    if (it.cat !== 'company') continue;
    for (const m of it.text.matchAll(/联合|合作|与|和|及/g)) {
      const rest = it.text.slice(m.index + m[0].length);
      if (m.index < 2 || rest.length < 8 || !names.has(rest)) continue;
      items.delete(it.key);
      for (const [s, e] of it.occ) put('company:' + rest, rest, { cat: 'company', en: it.en, cls: it.cls, warn: it.warn }, e - rest.length, e);
      break;
    }
  }
  for (const c of st.custom) {
    c.re.lastIndex = 0;
    for (const m of md.matchAll(c.re)) if (m[0]) put('cus:' + m[0], m[0], { cat: 'custom', en: 'custom', cls: 'tag', label: c.name }, m.index, m.index + m[0].length);
  }
  for (const w of st.blacklist) findAll(md, w, (s, e) => put('bl:' + w, w, { cat: 'black', en: 'black', cls: 'black' }, s, e));
  // 行号模式 #L11~L222：整行（含该章节全部段落）删除
  if (st.ranges.length) {
    const starts = [0];
    for (let i = md.indexOf('\n'); i >= 0; i = md.indexOf('\n', i + 1)) starts.push(i + 1);
    for (const [a, b] of st.ranges) {
      if (a < 1 || a > starts.length) continue;
      const bb = Math.min(b, starts.length), s = starts[a - 1], e = bb < starts.length ? starts[bb] - 1 : md.length;
      if (e > s) put('rng:' + a + '~' + bb, `#L${a}~L${bb}`, { cat: 'black', en: 'black', cls: 'black', del: true }, s, e);
    }
  }
  ABBR_RE.lastIndex = 0;
  const abbr = new Set();
  for (const m of md.matchAll(ABBR_RE)) abbr.add(m[1]);
  for (const w of abbr) findAll(md, w, (s, e) => put('abbr:' + w, w, { cat: 'abbr', en: 'abbr', cls: 'tag' }, s, e));
  tableNames(md, (t, s, e, ln) => (ln
    ? put('tbl:#L' + ln, '#L' + ln, { cat: 'tblname', en: 'tblname', cls: 'warn', warn: true, del: true, show: t }, s, e)
    : put('tbl:' + t, t, { cat: 'tblname', en: 'tblname', cls: 'warn', warn: true }, s, e)));
  // 模型只做兜底召回：规则 / 黑名单已识别的公司、地址里的碎片（"件有限公司"、"东华软"）不再重复列出
  const anchors = [...items.values()].filter((it) => it.cat === 'company' || it.cat === 'addr' || it.cat === 'black').map((it) => it.text);
  const cats = st.cats || {}, minScore = st.nerMin || 0, nerAll = st.ner.filter((n) => n.score >= minScore);
  // NER 校验正则：公司／地址正则命中处没有任何 机构/设施/地名 实体与之重叠 → 标成存疑（warn，默认不选）
  if (cats.ner_verify) {
    const orgish = nerAll.filter((n) => ORGISH.has(n.type));
    const ovl = (it) => orgish.some((n) => n.text.includes(it.text) || it.text.includes(n.text));
    for (const it of items.values()) if ((it.cat === 'company' || it.cat === 'addr') && !ovl(it)) { it.warn = true; it.cls = 'warn'; it.weak = true; }
  }
  // NER 人名放大：张三丰 → 张总 / 张工 / 小张 / 老张 …（同姓的称呼也一并标出）
  if (cats.ner_alias) {
    for (const n of nerAll) {
      if (n.type !== 'PERSON' || n.text.length < 2 || n.text.length > 4 || /[^\u4e00-\u9fff]/.test(n.text)) continue;
      const sn = n.text[0], re = new RegExp(`(?:[小老]${sn}(?![\u4e00-\u9fff])|${sn}(?:总|工|老师|经理|主任|博士|教授|局长|处长|董事长|先生|女士|书记|院长|所长|部长))`, 'g');
      for (const m of md.matchAll(re)) {
        if (items.has('ner:' + m[0])) continue;
        put('ner:' + m[0], m[0], { cat: 'ner', en: 'ner_person', cls: 'person', type: 'PERSON', grp: 'person', score: n.score, alias: true }, m.index, m.index + m[0].length);
      }
    }
  }
  // NER 日期/数量/金额：补正则漏掉的（已被其他条目包含的不再列）
  if (cats.ner_num) {
    const have = [...items.values()].map((it) => it.text);
    for (const n of nerAll) {
      if (!NUM_TYPES.has(n.type) || n.text.length < 2 || !/[0-9０-９一二三四五六七八九十百千万亿零〇]/.test(n.text)) continue;
      const key = 'ner:' + n.text;
      if (items.has(key) || have.some((h) => h.includes(n.text))) continue;
      findAll(md, n.text, (s, e) => put(key, n.text, { cat: 'ner', en: 'ner_num', cls: 'cand', type: n.type, grp: 'cand', score: n.score }, s, e));
    }
  }
  for (const n of nerAll) {
    const grp = NER_GROUPS[n.type];
    if (!grp || /[*|\n\\]/.test(n.text) || n.text.length < 2) continue;
    const key = 'ner:' + n.text;
    if (items.has('bl:' + n.text) || items.has(key)) continue;
    if (junkNer(n.text) || anchors.some((a) => a.includes(n.text))) continue;
    const abbr = grp === 'cand' && [...n.text].length <= ABBR_MAX; // 机构/产品类 <=3 字算简称
    findAll(md, n.text, (s, e) => put(key, n.text, { cat: 'ner', en: abbr ? 'abbr' : 'ner_' + grp, cls: abbr ? 'tag' : grp, type: n.type, grp: abbr ? 'abbr' : grp, score: n.score }, s, e));
  }
  for (const it of items.values()) it.occ.sort((a, b) => a[0] - b[0]);
  return items;
}

export function parseRules(text) {
  const rules = [];
  for (const ln of text.split('\n')) {
    const i = ln.indexOf('=>');
    if (i < 0) continue;
    let re = ln.slice(0, i).trim();
    const rep = ln.slice(i + 2).trim(), pre = /^::([^:]+)::\s*/.exec(re); // ::cname:: 前缀：只作用于该类别（cat / en / NER 类型）
    if (pre) re = re.slice(pre[0].length);
    if (!re) continue;
    try { rules.push({ test: new RegExp('^(?:' + re + ')$'), re: new RegExp(re, 'g'), rep, only: pre ? pre[1].trim() : '' }); } catch { /* 忽略无效规则 */ }
  }
  return rules;
}

// 替换模板：{表达式} 以 s（命中文本本身）、$m（正则匹配数组，$m[0] 为整段，$m[1] 为第 1 组）、kind（类别显示名）为变量求值，如 {s[0]}姓人员、{kind}{__di.ruleNth(kind, s)}；另支持 $1
function expand(rep, m, kind) {
  const s = m[0];
  return rep.replace(/\{([^{}]+)\}/g, (all, expr) => {
    try { return String(new Function('s', '$m', 'kind', '__di', `return (${expr})`)(s, m, kind, globalThis.__di)); } catch { return all; }
  }).replace(/\$(\d)/g, (_, k) => m[+k] ?? '');
}
export function replacement(text, ovr, rules, ctx = { names: [], kind: '' }) {
  if (ovr !== undefined && ovr !== '') return ovr === '∅' ? '' : ovr;
  for (const r of rules) {
    if (r.only && !ctx.names.includes(r.only)) continue;
    if (!r.test.test(text)) continue;
    r.re.lastIndex = 0;
    return text.replace(r.re, (...a) => (a[0] === '' ? '' : expand(r.rep, a.slice(0, a.findIndex((x) => typeof x === 'number')), ctx.kind))); // 空匹配（.* 在末尾多匹配一次）不展开
  }
  return '**';
}

// 选中的出现位置（去重叠：先出现、较长优先）
export function selection(items, isOn) {
  const all = [];
  for (const it of items) if (isOn(it)) for (const [s, e] of it.occ) all.push({ s, e, it });
  all.sort((a, b) => a.s - b.s || b.e - a.e);
  const out = [];
  let end = -1;
  for (const x of all) if (x.s >= end) { out.push(x); end = x.e; }
  return out;
}

// 章节：优先 # 标题，没有则按编号回退；范围到下一个同级或更高级标题（或"其他文本"区）之前
export function chapters(md) {
  const cut = md.indexOf('\n---\n\n## 其他文本');
  const body = cut >= 0 ? md.slice(0, cut + 1) : md;
  const lines = [];
  let pos = 0;
  for (const l of body.split('\n')) { lines.push([l, pos]); pos += l.length + 1; }
  let heads = [];
  for (const [l, s] of lines) { const m = /^(#{1,6}) (.+)$/.exec(l); if (m) heads.push({ level: m[1].length, title: m[2].replace(/\*+/g, ''), s }); }
  if (heads.length < 3) { // 样式标题太少（常见于手工排版的申报书）：叠加按编号识别
    const have = new Set(heads.map((h) => h.s));
    let inSpecial = false;
    for (const [l, s] of lines) {
      if (l.startsWith(':::特殊表格')) inSpecial = true;
      else if (l === ':::') inSpecial = false;
      if (inSpecial || have.has(s)) continue;
      if (!l || l.startsWith('|') || l.startsWith(':::') || l.length > 40 || /[。；;，,]$/.test(l)) continue;
      const plain = l.replace(/\*+/g, '');
      for (const [re, lv] of CHAPTER_FALLBACK) { const m = re.exec(plain); if (m) { heads.push({ level: lv(m), title: plain, s }); break; } }
    }
    heads.sort((a, b) => a.s - b.s);
  }
  const end = cut >= 0 ? cut + 1 : md.length;
  heads.forEach((h, i) => {
    let e = end;
    for (let j = i + 1; j < heads.length; j++) if (heads[j].level <= h.level) { e = heads[j].s; break; }
    h.e = e;
  });
  return heads;
}
