// 规则执行：扫描 Markdown → 候选项；选区；替换；章节。规则本身见 rules.js，投影结构见 markup.js，输入框格式见 boxes.js
import { RULES, ABBR_MAX, ABBR_RE, NAME_LABELS, NAME_VALUE, NAME_STOP, NER_GROUPS, CHAPTER_FALLBACK } from './rules.js';
import { readLines } from './markup.js';
import { lineRule } from './boxes.js';
import { plainView, unmark } from './plain.js';

// 送给 NER 的行：与 md 行一一对应；结构行置空，表格竖线换成空格
export const nerLines = (md) => plainView(md).lines.map((l) => (l.structural ? '' : l.text.replace(/\|/g, ' ')));

function findAll(md, text, add) {
  if (!text) return;
  let i = md.indexOf(text);
  while (i >= 0) { add(i, i + text.length); i = md.indexOf(text, i + text.length); }
}

// 含姓名的表格：标签格的下一格 / 特殊表格中标签行的下一行
function tableNames(lines, put) {
  let label = false;
  const strip = (s) => s.replace(/[：:\s]/g, '');
  for (const { text: line, s: start, n: ln, kind } of lines) {
    if (kind === 'open' || kind === 'close') { label = false; continue; }
    if (kind === 'table') {
      let p = start + 2, prevLabel = false;
      for (const cell of line.slice(2, -2).split(' | ')) {
        const t = strip(cell);
        if (prevLabel && NAME_VALUE.test(t) && !NAME_STOP.test(t)) { const a = cell.search(/[^：:\s]/), z = cell.search(/[：:\s]*$/); put(t, p + a, p + z); } // 「张 三」：整段（含中间空白）都算命中，不能按 t 的长度截
        prevLabel = NAME_LABELS.test(t);
        p += cell.length + 3;
      }
    } else if (kind === 'cell') {
      const t = strip(line);
      if (label && NAME_VALUE.test(t) && !NAME_STOP.test(t)) put(t, start, start + line.length, ln); // 特殊表格 1 格 1 行：按行号删，和章节一样
      label = NAME_LABELS.test(t);
    }
  }
}

// 模型常见的无效碎片：只剩公司后缀（"限公司"、"技有限公司"）、括号不成对（"中国）有限公司"）
const SUFFIX_ONLY = /^.{0,1}?(?:股份)?(?:有限责任|有限)?公司$/;
const ORGISH = new Set(['ORG', 'FAC', 'GPE', 'LOC']);
const junkNer = (t) => SUFFIX_ONLY.test(t) || !/[A-Za-z\u4e00-\u9fff]{2}/.test(t) || (t.match(/[（(]/g) || []).length !== (t.match(/[）)]/g) || []).length; // 另：至少要有连续 2 个字母/汉字（过滤 "137"、".0"）

// st：{ cats, blacklist: [词], ranges: [[a, b]], custom: [{name, re}], ner: [{text, type, score}], nerMin }，都可省略
// 返回 Map(key -> {key,text,cat,en,cls,warn,type,grp,score,occ:[[s,e]]})，不受勾选影响（勾选由调用方过滤，便于显示 ({num})）
// 匹配在纯文本视图上做，occ 最后换成 md 坐标
export function scan(md0, st0 = {}) {
  const st = { blacklist: [], ranges: [], custom: [], ner: [], ...st0 };
  const view = plainView(md0), md = view.text;
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
    for (const m of md.matchAll(c.re)) if (m[0]) put('cus:' + m[0], m[0], { cat: 'custom', en: 'custom', cls: 'custom', label: c.name }, m.index, m.index + m[0].length);
  }
  for (const w0 of st.blacklist) { const w = unmark(w0); findAll(md, w, (s, e) => put('bl:' + w, w, { cat: 'black', en: 'black', cls: 'black' }, s, e)); }
  // 行号模式 #L11~L222：整行（含该章节全部段落）删除
  if (st.ranges.length) {
    const starts = [0];
    for (let i = md.indexOf('\n'); i >= 0; i = md.indexOf('\n', i + 1)) starts.push(i + 1);
    for (const [a, b] of st.ranges) {
      if (a < 1 || a > starts.length) continue;
      const bb = Math.min(b, starts.length), s = starts[a - 1], e = bb < starts.length ? starts[bb] - 1 : md.length;
      if (e > s) put('rng:' + a + '~' + bb, lineRule.dump(a, bb), { cat: 'black', en: 'black', cls: 'black', del: true }, s, e);
    }
  }
  ABBR_RE.lastIndex = 0;
  const abbr = new Set();
  for (const m of md.matchAll(ABBR_RE)) abbr.add(m[1]);
  for (const w of abbr) findAll(md, w, (s, e) => put('abbr:' + w, w, { cat: 'abbr', en: 'abbr', cls: 'abbr' }, s, e));
  tableNames(view.lines, (t, s, e, ln) => (ln
    ? put('tbl:' + lineRule.dump(ln), lineRule.dump(ln), { cat: 'tblname', en: 'tblname', cls: 'warn', warn: true, del: true, show: t }, s, e)
    : put('tbl:' + t, t, { cat: 'tblname', en: 'tblname', cls: 'warn', warn: true }, s, e)));
  // 模型只做兜底召回：规则 / 黑名单已识别的公司、地址里的碎片（"件有限公司"、"东华软"）不再重复列出
  const anchors = [...items.values()].filter((it) => it.cat === 'company' || it.cat === 'addr' || it.cat === 'black').map((it) => it.text);
  const cats = st.cats || {}, minScore = st.nerMin || 0, nerAll = st.ner.filter((n) => n.score >= minScore);
  // NER 校验正则：公司／地址正则命中处没有任何 机构/设施/地名 实体与之重叠 → 标成存疑（warn，默认不选）
  // 另：正则从左边界一路匹配过来，常把前面的话也吞进去（「委托北京某某有限公司」）——若某 NER 实体正好是命中文本的真后缀，则把左端截到实体处
  if (cats.ner_verify) {
    const orgish = nerAll.filter((n) => ORGISH.has(n.type) && n.text.length >= 4 && !junkNer(n.text));
    for (const it of [...items.values()]) {
      if (it.cat !== 'company' && it.cat !== 'addr') continue;
      const n = orgish.filter((x) => x.text.length < it.text.length && it.text.endsWith(x.text)).sort((a, b) => b.text.length - a.text.length)[0];
      if (!n) continue;
      items.delete(it.key);
      for (const [, e] of it.occ) put(it.cat + ':' + n.text, n.text, { cat: it.cat, en: it.en, cls: it.cls, warn: it.warn }, e - n.text.length, e);
    }
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
  // NER 日期 / 金额 / 百分比：补正则漏掉的（已被其他条目包含的不再列）。百分比随「日期和百分比」开关；进人名地名面板（grp: place）
  const numOn = { DATE: cats.ner_date, MONEY: cats.ner_money, PERCENT: cats.date }, numEn = { DATE: 'ner_date', MONEY: 'ner_money', PERCENT: 'date' };
  if (numOn.DATE || numOn.MONEY || numOn.PERCENT) {
    const have = [...items.values()].map((it) => it.text);
    for (const n of nerAll) {
      if (!numOn[n.type] || n.text.length < 2 || !/[0-9０-９一二三四五六七八九十百千万亿零〇]/.test(n.text)) continue;
      const key = 'ner:' + n.text;
      if (items.has(key) || have.some((h) => h.includes(n.text))) continue;
      findAll(md, n.text, (s, e) => put(key, n.text, { cat: 'ner', en: numEn[n.type], cls: 'date', type: n.type, grp: 'place', score: n.score }, s, e));
    }
  }
  for (const n of nerAll) {
    const grp = NER_GROUPS[n.type];
    if (!grp || /[*|\n\\]/.test(n.text) || n.text.length < 2) continue;
    const key = 'ner:' + n.text;
    if (items.has('bl:' + n.text) || items.has(key)) continue;
    if (junkNer(n.text) || anchors.some((a) => a.includes(n.text))) continue;
    const abbr = grp === 'cand' && [...n.text].length <= ABBR_MAX; // 机构/产品类 <=3 字算简称
    findAll(md, n.text, (s, e) => put(key, n.text, { cat: 'ner', en: abbr ? 'abbr' : 'ner_' + grp, cls: abbr ? 'abbr' : grp, type: n.type, grp: abbr ? 'abbr' : grp, score: n.score }, s, e));
  }
  for (const it of items.values()) it.occ = it.occ.map(([s, e]) => view.toMd(s, e)).sort((a, b) => a[0] - b[0]);
  return items;
}

// scan 的异步版：ner 是函数 (lines) => Promise<[{text, type, score}]>（如 ner.recognize），不传则只用规则 / 黑名单。
// range=[s, e]（md 坐标）时只送这段给 NER，规则照常全文扫描
export async function detect(md, st = {}, { ner, range } = {}) {
  if (!ner) return scan(md, st);
  const lines = nerLines(md);
  if (range) { const a = md.slice(0, range[0]).split('\n').length - 1, b = md.slice(0, range[1]).split('\n').length; lines.forEach((_, i) => { if (i < a || i >= b) lines[i] = ''; }); }
  return scan(md, { ...st, ner: [...(st.ner || []), ...(await ner(lines))] });
}

// 替换模板：{表达式} 只认白名单写法（不执行任意代码）：s（命中文本）、$m（匹配数组）、kind（类别显示名），可接 [n] / .length / .slice(a,b)；
// nth() 或 __di.ruleNth(kind, s) 为同类内按首次出现编号。如 {s[0]}姓人员、{kind}{nth()}；另支持 $1
const TPL = /^\s*(?:(s|kind|\$m)((?:\[\d+\]|\.length|\.slice\(\s*-?\d+\s*(?:,\s*-?\d+\s*)?\))*)|(?:__di\.ruleNth|nth)\(\s*(?:kind\s*,\s*s\s*)?\))\s*$/;
function evalTpl(expr, vars, nth) {
  const m = TPL.exec(expr);
  if (!m) return null;
  if (!m[1]) return nth ? String(nth(vars.kind, vars.s)) : null;
  let v = vars[m[1]];
  for (const op of m[2].matchAll(/\[(\d+)\]|\.length|\.slice\(([^)]*)\)/g)) {
    if (op[1] !== undefined) v = v?.[+op[1]];
    else if (op[0] === '.length') v = v?.length;
    else v = v?.slice?.(...op[2].split(',').map(Number));
  }
  return v == null ? '' : String(v);
}
function expand(rep, m, kind, nth) {
  const vars = { s: m[0], $m: m, kind };
  // 一遍完成：展开结果里的 $1 等不会再被当成分组引用
  return rep.replace(/\{([^{}]+)\}|\$(\d)/g, (all, expr, k) => (expr !== undefined ? evalTpl(expr, vars, nth) ?? all : m[+k] ?? ''));
}
// ctx：names 该条目的类别名与原文（供 ::类别:: / ::原文:: 前缀匹配），kind 显示名，nth(kind, s) 编号函数
export function replacement(text, rules, { names = [], kind = '', nth = null } = {}) {
  for (const r of rules) {
    if (r.only && !names.includes(r.only)) continue;
    if (r.whole) return expand(r.rep, [text], kind, nth); // 「::原文:: 替换」
    if (!r.test.test(text)) continue;
    r.re.lastIndex = 0;
    return text.replace(r.re, (...a) => (a[0] === '' ? '' : expand(r.rep, a.slice(0, a.findIndex((x) => typeof x === 'number')), kind, nth))); // 空匹配（.* 在末尾多匹配一次）不展开
  }
  return '**';
}

// 选中的出现位置：重叠的并成一段（「张三」与「三丰公司」→「张三丰公司」，不会漏出半截），替换方式取其中最长的命中
export function selection(items, isOn) {
  const all = [];
  for (const it of items) if (isOn(it)) for (const [s, e] of it.occ) all.push({ s, e, it });
  all.sort((a, b) => a.s - b.s || b.e - a.e);
  const out = [];
  for (const x of all) {
    const last = out.at(-1);
    if (!last || x.s >= last.e) { out.push({ ...x }); continue; }
    if (x.e - x.s > last.e - last.s) last.it = x.it;
    last.e = Math.max(last.e, x.e);
  }
  return out;
}

// 章节：优先 # 标题，没有则按编号回退；范围到下一个同级或更高级标题（或"其他文本"区）之前
export function chapters(md) {
  const all = readLines(md), lines = all.filter((l) => !l.others);
  let heads = [];
  for (const { text, s } of lines) { const m = /^(#{1,6}) (.+)$/.exec(text); if (m) heads.push({ level: m[1].length, title: unmark(m[2]), s }); }
  if (heads.length < 3) { // 样式标题太少（常见于手工排版的申报书）：叠加按编号识别
    const have = new Set(heads.map((h) => h.s));
    for (const { text, s, kind } of lines) {
      if (kind !== 'text' || have.has(s) || !text || text.length > 40 || /[。；;，,]$/.test(text)) continue;
      const plain = unmark(text);
      for (const [re, lv] of CHAPTER_FALLBACK) { const m = re.exec(plain); if (m) { heads.push({ level: lv(m), title: plain, s }); break; } }
    }
    heads.sort((a, b) => a.s - b.s);
  }
  const end = all.find((l) => l.others)?.s ?? md.length;
  heads.forEach((h, i) => {
    let e = end;
    for (let j = i + 1; j < heads.length; j++) if (heads[j].level <= h.level) { e = heads[j].s; break; }
    h.e = e;
  });
  return heads;
}
