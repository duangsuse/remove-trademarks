import './style.css';
import { DocxProjection } from './docx.js';
import { scan, replacement, selection, chapters, nerLines } from './detect.js';
import { plainOf, unmark, spliceMd, lineRange } from './plain.js';
import { partnerOf, mdLines } from './markup.js';
import { BlackBox, lineRule, parseRules, parseCustom } from './boxes.js';
import { MAIN, EXTRA, TOOLS, TYPE_NAME, DEFAULT_REPLACE_RULES, HIDE_WHEN_OFF, TAG_CATS, ENT_CATS, TITLE_RE, ABBR_MAX, RULES, FINGERPRINT_URL, ABBR_RE, NAME_LABELS, NAME_VALUE } from './rules.js';
import * as ner from './ner.js';
import * as fs from './fs.js';

const $ = (id) => document.getElementById(id);
const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const mb = (n) => (n / 1048576).toFixed(2) + ' MB';
const store = { get: (k) => { try { return localStorage.getItem(k); } catch { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch { /* 无存储 */ } } };

// ---------- 主题：点 logo 切换 ----------
const root = document.documentElement, THEMES = ['light', 'dark', 'omarchy'];
root.dataset.theme = store.get('theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
$('logo').onclick = () => { root.dataset.theme = THEMES[(THEMES.indexOf(root.dataset.theme) + 1) % THEMES.length]; store.set('theme', root.dataset.theme); };
// 高亮类别（rules.js 的 cls 与 top1~3）：底色取各主题的 --hl-<类别>
const HL = ['tag', 'warn', 'place', 'cand', 'black', 'phone', 'mail', 'person', 'company', 'addr', 'date', 'net', 'abbr', 'custom', 'top1', 'top2', 'top3', 'frame'];
document.head.append(Object.assign(document.createElement('style'), { textContent: HL.map((k) => `::highlight(${k}), .badge.${k}, .row.${k} { background-color: var(--hl-${k}); }`).join('\n') }));

// URL 带 #editor：关闭脱敏功能，当作纯文本投影编辑器（Ctrl+S → 新版/原名.docx）；左栏只剩章节，右栏仅在粘贴时更新
// hash：#editor#文件名 / #文件名（选中的文档记在这里，刷新后自动打开）
const hashArg = decodeURIComponent(location.hash.slice(1)).split('#');
const EDITOR = hashArg[0] === 'editor';
const hashFile = (EDITOR ? hashArg.slice(1) : hashArg).join('#');
const setHash = (name) => history.replaceState(null, '', '#' + (EDITOR ? 'editor#' : '') + encodeURIComponent(name));
addEventListener('hashchange', () => location.reload());
document.body.classList.toggle('ed', EDITOR);

// ---------- 左右栏宽度：拖动分隔条，双击恢复默认 ----------
const layout = document.querySelector('.layout');
for (const side of ['l', 'r']) { const w = store.get('w' + side); if (w) layout.style.setProperty(`--${side}w`, w + 'px'); }
layout.addEventListener('pointerdown', (e) => {
  const g = e.target.closest('.gutter'); if (!g) return;
  e.preventDefault(); g.setPointerCapture(e.pointerId); g.classList.add('drag');
  const side = g.dataset.side, box = layout.getBoundingClientRect();
  g.onpointermove = (ev) => {
    const w = Math.round(Math.min(Math.max(side === 'l' ? ev.clientX - box.left : box.right - ev.clientX, 180), box.width * 0.45));
    layout.style.setProperty(`--${side}w`, w + 'px'); store.set('w' + side, w);
  };
  g.onpointerup = () => { g.onpointermove = g.onpointerup = null; g.classList.remove('drag'); };
});
layout.addEventListener('dblclick', (e) => {
  const g = e.target.closest('.gutter'); if (!g) return;
  layout.style.removeProperty(`--${g.dataset.side}w`); store.set('w' + g.dataset.side, '');
});

// ---------- 状态 ----------
// 脱敏的唯一来源是左栏两个大框：黑名单（词 / 行号 #L11~L222）与「删除时（替换）规则」。
// 右侧面板、正文里的删除只是在编辑这两个框；正文 #ta 只读。
const state = { cats: { custom: true }, tools: {}, blacklist: [], ranges: [], custom: [], ner: [], on: {}, nerMin: 0.5 };
for (const x of [...MAIN, ...EXTRA]) state.cats[x.id] = x.def !== false;
for (const t of TOOLS) state[t.scope][t.id] = t.def;
const doc = { entries: [], dir: null, cur: null, buffer: null, parsed: null, orig: '', images: [], order: [], keep: new Set(), outHandle: null };
let pendingModel = null, pendingDir = null; // 上次的模型 / 文件夹句柄，权限待点击后请求

const SPECIAL_CODE = {
  img: '图片：替换为 1×1 占位图（勾选缩略图可保留）', meta: '.docx 作者/公司属性、缩略图、批注、修订记录、隐藏文字等',
  black: '黑名单词汇：左栏黑名单框的词 / #L行号 规则', abbr: () => `以下简称“…” → ${ABBR_RE}`,
  tblname: () => `标签格：${NAME_LABELS}\n姓名值：${NAME_VALUE}\n特殊表格中按行号删除`,
  ner_person: '离线 NER：PERSON', ner_place: '离线 NER：GPE / LOC', ner_cand: '离线 NER：ORG / FAC / PRODUCT / NORP（> 3 字）',
  ner_verify: '公司／地址正则命中处，若没有任何 NER 的 ORG / FAC / GPE / LOC 实体与之重叠 → 标为存疑（红，默认不选）',
  ner_alias: '对 NER 识别出的人名（姓 X）额外标出 X总 / X工 / X老师 / X经理… / 小X / 老X',
  ner_date: '离线 NER：DATE（含数字，正则已覆盖的不再列）；PERCENT 随「日期和百分比」开关生效',
  ner_money: '离线 NER：MONEY（含数字，正则已覆盖的不再列）',
};
function ruleCode(id) { // 悬停预览：该开关背后的规则代码
  const lines = Object.entries(RULES).filter(([, r]) => r.en === id).map(([k, r]) => `${k}: ${r.re}`);
  const sp = SPECIAL_CODE[id];
  if (sp) lines.push(typeof sp === 'function' ? sp() : sp);
  return lines.join('\n') || id;
}
const row = (x, scope) => {
  const tail = x.info ? `<a class="info" href="${FINGERPRINT_URL}" target="_blank" rel="noopener" title="微软：删除隐藏数据和个人信息">ⓘ</a>`
    : scope === 'cats' && !TOOLS.some((t) => t.id === x.id) ? `<span class="n" data-n="${x.id}"></span>` : '';
  return `<label title="${esc(ruleCode(x.id))}"><input type="checkbox" data-s="${scope}" data-k="${x.id}" ${state[scope][x.id] ? 'checked' : ''}><span>${x.name}${tail}${x.note ? ` <span class="n">- ${x.note}</span>` : ''}</span></label>`;
};
const OPEN_ORDER = ['company', 'patent', 'cname', 'contact', 'ner_cand'];
function renderLeft() {
  const rest = EXTRA.filter((x) => !x.grp);
  $('catList').innerHTML = MAIN.filter((x) => !x.grp).map((x) => row(x, 'cats')).join('');
  $('openList').innerHTML = [...MAIN, ...EXTRA].filter((x) => x.grp === 'open').sort((a, b) => OPEN_ORDER.indexOf(a.id) - OPEN_ORDER.indexOf(b.id)).map((x) => row(x, 'cats')).join('');
  $('extraList').innerHTML = rest.map((x) => row(x, 'cats')).join('');
  $('moreN').textContent = `(${rest.length})`;
  $('toolList').innerHTML = TOOLS.map((t) => row(t, t.scope)).join('') +
    `<label title="低于该置信度的 NER 结果不进右栏"><span>NER 置信度 ≥ <b id="nerMinV" class="mono">${state.nerMin.toFixed(2)}</b></span><input id="nerMin" type="range" min="0" max="0.99" step="0.01" value="${state.nerMin}" style="appearance:auto;width:90px;height:auto;border:0;border-radius:0"></label>`;
}
renderLeft();
$('rules').value = DEFAULT_REPLACE_RULES;

document.addEventListener('input', (e) => {
  if (e.target.id !== 'nerMin') return;
  state.nerMin = +e.target.value; $('nerMinV').textContent = state.nerMin.toFixed(2);
  clearTimeout(tmMin);
  tmMin = setTimeout(() => { // 重新过滤：低于阈值而不再出现的 NER 条目，连同黑名单框里自动同步的行一并去掉
    run();
    for (const [k, m] of memo) if (m.ner && !items.has(k)) { memo.delete(k); delete state.on[k]; }
    syncBlack();
  }, 150);
});
let tmMin = 0;
document.addEventListener('change', (e) => {
  const t = e.target, k = t.dataset?.k;
  if (!k) return;
  state[t.dataset.s][k] = t.checked;
  if (t.dataset.s === 'cats') { // 左栏只是右栏的开关：同类条目一并勾选 / 取消
    for (const [key, m] of memo) if (m.en === k) state.on[key] = t.checked;
    for (const it of items.values()) if (it.en === k && it.cat !== 'black') state.on[it.key] = t.checked;
  }
  if (k === 'img') resetKeep();
  refresh(true);
});

// ---------- 黑名单框：=== 以上手填，以下随右侧面板同步 ----------
const blackEl = $('black');
let lastGen = new Set();
const base = (it) => !(it.cat === 'ner' && it.grp === 'cand') && !it.weak; // 机构/产品候选（>3 字）默认不选
const memo = new Map(); // key -> {text,en,base}：见过的条目；文中不再出现（比如已脱敏）也保留，黑名单框不随文本变化
const memoOn = (key, m) => (key in state.on ? state.on[key] : !!state.cats[m.en] && m.base);
function remember() {
  for (const it of items.values()) if (it.cat !== 'black') memo.set(it.key, { text: it.text, en: it.en, base: base(it), ner: it.cat === 'ner' });
}
const black = () => BlackBox.load(blackEl.value);
blackEl.value = BlackBox.dump({ own: [''] });
function readHead() {
  const { words, ranges } = black();
  state.blacklist = words; state.ranges = ranges;
}
function syncBlack() {
  const { own: head, entries } = black();
  const own = new Set(entries), gen = [], seen = new Set();
  for (const [key, m] of memo) {
    if (!memoOn(key, m) || own.has(m.text) || seen.has(m.text)) continue;
    seen.add(m.text); gen.push(m.text);
  }
  lastGen = seen;
  const headText = head.join('\n');
  const next = BlackBox.dump({ own: head, synced: gen.slice(0, 800) });
  if (next === blackEl.value) return;
  if (document.activeElement === blackEl && blackEl.selectionStart > headText.length) return; // 正在编辑 === 以下，失焦后再同步
  const [a, b, t] = [blackEl.selectionStart, blackEl.selectionEnd, blackEl.scrollTop];
  blackEl.value = next;
  if (document.activeElement === blackEl) blackEl.setSelectionRange(Math.min(a, headText.length), Math.min(b, headText.length));
  blackEl.scrollTop = t;
}
blackEl.oninput = () => refresh();
blackEl.onchange = () => { // 编辑了 === 以下：删行=取消选中，新增行=选中同名项，无同名项则并入黑名单
  const { entries, synced } = black();
  const want = new Set(synced), own = new Set(entries);
  const known = new Set();
  for (const [key, m] of memo) {
    known.add(m.text);
    if (lastGen.has(m.text) && !want.has(m.text)) state.on[key] = false;
    else if (!lastGen.has(m.text) && want.has(m.text)) state.on[key] = true;
  }
  const add = [...want].filter((w) => !lastGen.has(w) && !own.has(w) && !known.has(w));
  if (add.length) blackEl.value = BlackBox.dump({ own: [...entries, ...add] });
  run();
};
$('custom').oninput = () => { state.custom = parseCustom($('custom').value); refresh(); };
$('rules').oninput = () => refresh();

// 对黑名单框的结构化修改（右侧面板 / 正文删除 / 章节删除 都走这里）
const boxAdd = (text) => { blackEl.value = BlackBox.add(blackEl.value, text); };
const boxRemove = (text) => { blackEl.value = BlackBox.remove(blackEl.value, text); };
function setOn(it, on) {
  if (it.cat === 'black') { if (!on) boxRemove(it.text); } else state.on[it.key] = on;
}

// ---------- 撤销 / 重做：只管黑名单 + 删除时替换规则 ----------
const undoS = [], redoS = [];
let lastPush = 0;
const snap = () => ({ own: black().own, on: { ...state.on }, rules: $('rules').value });
function checkpoint(force) {
  const now = Date.now();
  if (!force && now - lastPush < 800) { lastPush = now; return; } // 连续输入合并为一步
  undoS.push(snap()); if (undoS.length > 200) undoS.shift();
  redoS.length = 0; lastPush = now;
}
function restore(sn) {
  blackEl.value = BlackBox.dump({ own: sn.own }); state.on = { ...sn.on }; $('rules').value = sn.rules;
  lastPush = 0; run();
}
function undo() { if (!undoS.length) return status('没有可撤回的操作（只记录黑名单与替换规则）'); redoS.push(snap()); restore(undoS.pop()); status('已撤回'); }
function redo() { if (!redoS.length) return; undoS.push(snap()); restore(redoS.pop()); status('已重做'); }
const edit = (fn) => { checkpoint(true); fn(); run(); };
blackEl.addEventListener('beforeinput', () => checkpoint(false));
$('rules').addEventListener('beforeinput', () => checkpoint(false));

// ---------- 识别 / 渲染 ----------
const ta = $('ta');
// #ta 是只读 div（#editor 模式为 plaintext-only 可编辑），下面补上 textarea 的 value / selectionStart / End / setSelectionRange，其余代码不用改
let lastSel = [0, 0];
const offOf = (node, off) => { const r = document.createRange(); r.setStart(ta, 0); r.setEnd(node, off); return r.toString().length; };
function syncSel() {
  const g = getSelection();
  if (g.rangeCount && ta.contains(g.anchorNode) && ta.contains(g.focusNode)) { const r = g.getRangeAt(0); lastSel = [offOf(r.startContainer, r.startOffset), offOf(r.endContainer, r.endOffset)]; }
}
document.addEventListener('selectionchange', syncSel);
function nodeAt(off) {
  const w = document.createTreeWalker(ta, NodeFilter.SHOW_TEXT);
  let n, acc = 0, last = null;
  while ((n = w.nextNode())) { last = n; if (off <= acc + n.length) return [n, off - acc]; acc += n.length; }
  return last ? [last, last.length] : [ta, 0];
}
function rangeOf(s, e) { const r = document.createRange(); r.setStart(...nodeAt(s)); r.setEnd(...nodeAt(e)); return r; }
Object.defineProperties(ta, {
  value: { get: () => ta.textContent, set: (v) => { ta.textContent = v; lastSel = [0, 0]; } },
  selectionStart: { get: () => lastSel[0] },
  selectionEnd: { get: () => lastSel[1] },
});
ta.setSelectionRange = (s, e) => { const g = getSelection(); g.removeAllRanges(); g.addRange(rangeOf(s, e)); lastSel = [s, e]; };
// 特殊表格的边界行（点一下会跳到另一头）；#editor 模式也要画
function frameLines(by) {
  for (const l of mdLines(ta.value)) if (l.kind === 'open' || l.kind === 'close') (by.frame ||= new Highlight()).add(rangeOf(l.s, l.s + l.text.length));
}
// 高亮：每个类别一个 Highlight（样式见 ::highlight(...)）；未选中（不会被打码）的统一用 off（下划线）
function paintHighlights(occ) {
  CSS.highlights.clear();
  const by = {};
  let p = 0;
  for (const o of occ.sort((a, b) => a.s - b.s || b.e - a.e)) {
    if (o.s < p) continue;
    p = o.e;
    const on = isOn(o.it), ranked = top.has(o.it.key), r = rangeOf(o.s, o.e);
    (by[on || ranked ? cls(o.it) : 'off'] ||= new Highlight()).add(r); // 前 3 名不论选没选都整段着色，未选再加下划线
    if (ranked && !on) (by.off ||= new Highlight()).add(r);
  }
  frameLines(by);
  for (const k in by) CSS.highlights.set(k, by[k]);
}
let items = new Map(), live = [], sel = [], tm = 0, chs = [];
const enabled = (it) => !!state.cats[it.en];
const listed = (it) => enabled(it) || !HIDE_WHEN_OFF.has(it.en);
const isOn = (it) => (it.key in state.on ? state.on[it.key] : enabled(it) && base(it));
// 配色：按规则类别（rules.js 的 cls）；NER 识别出的实体按出现次数前 3 各有一色；称谓代词（XX经理）红色
const cls = (it) => top.get(it.key) || (it.warn || TITLE_RE.test(it.text) ? 'warn' : it.cls);
const rankTag = (it) => (top.has(it.key) ? `<b class="rank">#${top.get(it.key).slice(3)}</b>` : '');
const top = new Map();
function rankNer() {
  top.clear();
  [...items.values()].filter((it) => it.cat === 'ner').sort((a, b) => b.occ.length - a.occ.length).slice(0, 3).forEach((it, i) => top.set(it.key, 'top' + (i + 1)));
}

function refresh(now) { clearTimeout(tm); tm = setTimeout(run, now ? 0 : 160); }
function run() {
  clearTimeout(tm);
  if (EDITOR) { CSS.highlights.clear(); const by = {}; frameLines(by); for (const k in by) CSS.highlights.set(k, by[k]); renderChapters(); return; } // 编辑器模式：只维护章节
  readHead();
  const md = ta.value;
  items = scan(md, state);
  rankNer();
  live = [...items.values()].filter(listed);
  sel = selection(live, isOn);
  const occ = [];
  for (const it of live) for (const [s, e] of it.occ) occ.push({ s, e, it });
  occ.sort((a, b) => a.s - b.s || b.e - a.e);
  paintHighlights(occ);
  updateCounts();
  renderPanels();
  renderChapters();
  remember();
  syncBlack();
  updatePill();
}

const rules = () => parseRules($('rules').value);
const nthMap = new Map();
const ruleNth = (kind, s) => { // 同类（kind）内按首次出现顺序编号，同一文本编号相同
  const key = Array.isArray(s) ? s[0] : String(s);
  if (!nthMap.has(kind)) nthMap.set(kind, new Map());
  const m = nthMap.get(kind);
  if (!m.has(key)) m.set(key, m.size + 1);
  return m.get(key);
};
const replOf = (it, text) => (it.del ? '' : replacement(text, rules(), { names: [it.cat, it.en, it.type, it.text], kind: kindOf(it), nth: ruleNth }));

function updateCounts() {
  const cnt = {};
  for (const it of items.values()) cnt[it.en] = (cnt[it.en] || 0) + 1;
  cnt.img = doc.images.length;
  ruleOccs = {};
  for (const x of [...MAIN, ...EXTRA]) if (x.grp || EXTRA.includes(x)) ruleOccs[x.id] = ruleOcc(x.id);
  for (const el of document.querySelectorAll('[data-n]')) if (!ruleOccs[el.dataset.n]) el.textContent = ` (${cnt[el.dataset.n] || 0})`;
  paintLoc();
}

// 面板归属：实体清单(ent) / 人名地名(place) / 简称·电话·专利号(tag) / 自动识别(cand，规则类置信度 0)
function panelOf(it) {
  if (TITLE_RE.test(it.text) && it.cat !== 'black') return 'tag'; // 称谓代词（XX经理）
  if (it.cat === 'ner') return it.grp === 'cand' ? 'cand' : it.grp === 'abbr' ? 'tag' : 'place';
  if (it.cat === 'black') return it.del || [...it.text].length > ABBR_MAX ? 'ent' : 'tag';
  if (TAG_CATS.has(it.cat)) return 'tag';
  if (ENT_CATS.has(it.cat)) return 'ent';
  return 'cand';
}
const kindOf = (it) => (it.cat === 'tblname' && it.del ? '含姓名表格 · 行号' : it.cat === 'black' ? (it.del ? '行号' : '手动') : it.cat === 'ner' ? TYPE_NAME[it.type] || it.type : RULES[it.cat]?.name || it.label || '规则');
// 两种条目视图，开关都由 data-key 所在元素的点击统一处理（见 rightEl click）
let renames = new Map(); // 原文 -> 替换：#rules 里的「::原文:: => 替换」行，renderPanels 时重读
const renTag = (it) => (renames.has(it.text) ? `<span class="rn">→ ${esc(renames.get(it.text))}</span>` : '');
const chip = (it) => `<span class="badge ${cls(it)} ${isOn(it) ? 'on' : 'off'}" data-key="${esc(it.key)}" title="${esc(it.text)}">${rankTag(it)}${esc(it.show || it.text)}${renTag(it)}<small>${it.occ.length}</small></span>`;
const scoreOf = (it) => `<span class="score">${it.cat === 'ner' ? it.score.toFixed(2) : '0'}</span>`;
const rowOf = (it, extra = '') => `<div class="row ${isOn(it) ? 'on' : ''} ${top.get(it.key) || ''}" data-key="${esc(it.key)}"><input type="checkbox" ${isOn(it) ? 'checked' : ''}><div class="t">${rankTag(it)}${esc(it.show ? it.show + ' ' : '')}${esc(it.text)}${renTag(it)}<small>${kindOf(it)} · ${it.occ.length}次</small></div>${extra}${locBtn(it)}<button class="ren" data-ren="${esc(it.key)}" title="重命名：替换成别的文字，记在替换规则框里">✎</button></div>`;
// 四张卡片：#<id>List 放条目，#cnt-<id> 显示 选中 OF 总数；只在视图、空提示、置信度上有差别。条目一律按出现次数从多到少
const PANELS = [
  { id: 'ent', item: (it) => rowOf(it), empty: () => '暂无。点下方 Add，或在正文选中文字按 Delete' },
  { id: 'tag', item: chip, empty: () => '无' },
  { id: 'place', item: chip, empty: () => (ner.ready() ? '无' : '未载入模型，仅规则识别可用') },
  { id: 'cand', item: (it) => rowOf(it, scoreOf(it)), empty: () => (ner.ready() ? '无备选' : '请先载入离线模型') },
];
const panels = Object.fromEntries(PANELS.map((p) => [p.id, []]));
function renderPanels() {
  renames = new Map(rules().filter((r) => r.whole).reverse().map((r) => [r.only, r.rep])); // 重复的取最上面一行
  for (const k in panels) panels[k] = [];
  for (const it of live) panels[panelOf(it)].push(it);
  for (const p of PANELS) {
    const list = panels[p.id];
    list.sort((a, b) => b.occ.length - a.occ.length || (b.score || 0) - (a.score || 0)); // 次数相同按置信度
    $(p.id + 'List').innerHTML = list.map(p.item).join('') || `<div class="empty">${p.empty()}</div>`;
    $('cnt-' + p.id).textContent = list.length ? `${list.filter(isOn).length} OF ${list.length}` : '0'; // 按规则（条目）数，不是匹配数
  }
  $('sum').textContent = `${sel.length} of ${live.reduce((a, i) => a + i.occ.length, 0)}`;
}
// 定位到下一个出现：全部以 #ta 的光标 / 选区为准（无独立状态）——剩余 = 起点 >= 选区末尾 的出现；点一次选中下一个，到头后变「回顶」
const caret = () => ta.selectionEnd;
const remainOf = (occ) => occ.filter((o) => o[0] >= caret()).length;
// 全局只有一个查找目标 locT：点 / 右键只是把它挪到鼠标下临时显示 (剩N / 回顶)；其余位置一律显示总数
const locT = { kind: '', id: '' };
const isT = (kind, id) => locT.kind === kind && locT.id === id;
const locLabel = (it) => { if (!isT('item', it.key)) return `${it.occ.length} ↓`; const n = remainOf(it.occ); return n ? `剩${n} ↓` : '回首 ↥'; };
const locBtn = (it) => (it.occ.length ? `<button class="loc" data-loc="${esc(it.key)}" title="定位到下一个出现（到头后回第一处）">${locLabel(it)}</button>` : '<button class="loc" disabled>—</button>');
// 把 md 位置 off 滚到视口高度的 at 处（默认 1/3）
function scrollToOff(off, smooth, at = 1 / 3) {
  const top = rangeOf(off, Math.min(off + 1, ta.value.length)).getBoundingClientRect().top - ta.getBoundingClientRect().top + ta.scrollTop;
  ta.scrollTo({ top: Math.max(0, top - ta.clientHeight * at), behavior: smooth ? 'smooth' : 'auto' });
}
function gotoOcc(s, e, smooth) {
  ta.focus(); ta.setSelectionRange(s, e);
  scrollToOff(s, smooth);
}
// 特殊表格：选中（或点在）开头行上，平滑滚到收尾行（放在视口下方，表格内容在上面）；在收尾行上则滚回开头
function jumpPair() {
  syncSel();
  const s = ta.selectionStart, to = partnerOf(ta.value, s, ta.selectionEnd);
  if (to >= 0) scrollToOff(to, true, to > s ? 2 / 3 : 1 / 3);
}
ta.addEventListener('mouseup', jumpPair);
ta.addEventListener('keyup', (e) => { if (e.key.startsWith('Arrow') || e.key === 'Home' || e.key === 'End') jumpPair(); });
let lastTarget = '';
function locate(occ) { // 平滑滚动只在切换查找目标时用；同一目标连着点则直接跳
  const sig = locT.kind + locT.id, smooth = sig !== lastTarget;
  lastTarget = sig;
  const o = occ.find((x) => x[0] >= caret()) || occ[0]; // 到头后回第一处出现（不是文档顶）；无出现则无操作
  if (o) gotoOcc(o[0], o[1], smooth);
}
// 右键查找的目标：芯片面板 → 最近右键的那一项；「更多规则」→ 被右键过的规则。尾部显示 (剩N) / (回首)
let ruleOccs = {};
function ruleOcc(en) { // 更多规则在右栏已过滤的文本上运行：被其他类别已选中（将被打码）的位置不再计入
  const others = sel.filter((x) => x.it.en !== en), out = [];
  for (const it of items.values()) if (it.en === en) for (const o of it.occ) if (!others.some((x) => x.s < o[1] && o[0] < x.e)) out.push(o);
  return out.sort((a, b) => a[0] - b[0]);
}
const tail = (occ) => { const n = remainOf(occ); return n ? `(剩${n})` : '(回首)'; };
let paintRaf = 0;
function paintLoc() {
  for (const b of document.querySelectorAll('[data-loc]')) { const it = items.get(b.dataset.loc); if (it) b.textContent = locLabel(it); }
  const cur = locT.kind === 'item' ? items.get(locT.id) : null;
  for (const el of document.querySelectorAll('[data-lp]')) el.textContent = cur && panelOf(cur) === el.dataset.lp ? tail(cur.occ) : '';
  for (const x of [...MAIN, ...EXTRA]) { if (!ruleOccs[x.id]) continue; const el = document.querySelector(`[data-n="${x.id}"]`); if (el) el.textContent = ' ' + (isT('rule', x.id) ? tail(ruleOccs[x.id] || []) : `(${(ruleOccs[x.id] || []).length})`); }
}
document.addEventListener('selectionchange', () => { cancelAnimationFrame(paintRaf); paintRaf = requestAnimationFrame(paintLoc); });
const rightEl = document.querySelector('.right');
rightEl.addEventListener('click', (e) => {
  const all = e.target.closest('[data-all]');
  if (all) { const list = panels[all.dataset.all]; if (list.length) { const to = !list.every(isOn); edit(() => list.forEach((it) => setOn(it, to))); } return; }
  const l = e.target.closest('[data-loc]');
  if (l) { const it = items.get(l.dataset.loc); if (it) { locT.kind = 'item'; locT.id = it.key; locate(it.occ); paintLoc(); } return; }
  const lp = e.target.closest('[data-lp]');
  if (lp) { const it = locT.kind === 'item' ? items.get(locT.id) : null; if (it && panelOf(it) === lp.dataset.lp) { locate(it.occ); paintLoc(); } return; }
  const r = e.target.closest('[data-ren]');
  if (r) {
    const it = items.get(r.dataset.ren); if (!it) return;
    if (it.del) return status('行号规则是整段删除，没有替换文字可改');
    const cur = isOn(it) ? renames.get(it.text) ?? replOf(it, it.text) : it.text; // 已勾选才显示替换结果，否则照原文
    const t = prompt(`把「${it.text}」替换为（留空 = 不再打码）`, cur);
    const to = t?.trim();
    if (to === undefined || to === cur) return;
    edit(() => { setOn(it, !!to); if (to) $('rules').value = `::${it.text}:: => ${to}\n` + $('rules').value; }); // 新的写在最上面，自上而下取第一条
    return;
  }
  // 切换开关：芯片 / 行（含复选框本身，键盘空格也会触发 click）统一在这里；随后 run() 重绘，复选框状态以 isOn 为准
  const k = e.target.closest('[data-key]');
  if (k && !e.target.closest('button')) { const it = items.get(k.dataset.key); if (it) edit(() => setOn(it, !isOn(it))); }
});
for (const box of [$('rules'), blackEl]) box.addEventListener('click', () => { // 单击（无选区）自动选中整行
  const { value: v, selectionStart: a, selectionEnd: z } = box; if (a !== z) return;
  const s = v.lastIndexOf('\n', a - 1) + 1, e = v.indexOf('\n', a);
  box.setSelectionRange(s, e < 0 ? v.length : e);
});
rightEl.addEventListener('contextmenu', (e) => { // 芯片没有定位按钮：右键查找下一个，(剩/回顶) 显示在面板名尾部
  const b = e.target.closest('.badge'); if (!b) return;
  const it = items.get(b.dataset.key); if (!it) return;
  e.preventDefault();
  locT.kind = 'item'; locT.id = it.key;
  locate(it.occ); paintLoc();
});

for (const box of [$('extraList'), $('openList')]) box.addEventListener('contextmenu', (e) => {
  const inp = e.target.closest('label')?.querySelector('input'); if (!inp) return;
  e.preventDefault();
  const en = inp.dataset.k;
  ruleOccs[en] = ruleOcc(en); locT.kind = 'rule'; locT.id = en;
  locate(ruleOccs[en]); paintLoc();
});

// ---------- 正文只读：选中后删除 / 拖选 = 往黑名单加规则 ----------
// 多行选择：整行用行号模式 #L11~L22，行内片段用词；单行选择一律用词
function rulesFromSelection(s, e) {
  const lines = ta.value.split('\n'), out = [], full = [];
  let pos = 0;
  const hits = [];
  lines.forEach((ln, i) => {
    const ls = pos, le = pos + ln.length; pos = le + 1;
    const os = Math.max(s, ls), oe = Math.min(e, le);
    if (oe > os) hits.push({ i, ln, os: os - ls, oe: oe - ls, whole: os <= ls && oe >= le });
  });
  const word = (h) => unmark(h.ln.slice(h.os, h.oe)).trim(); // 黑名单存纯文字，不带 ** 等标记
  if (hits.length === 1) { const t = word(hits[0]); return t ? [t] : []; }
  for (const h of hits) { if (h.whole) full.push(h.i + 1); else { const t = word(h); if (t) out.push(t); } }
  for (let k = 0; k < full.length;) { let j = k; while (full[j + 1] === full[j] + 1) j++; out.unshift(lineRule.dump(full[k], full[j])); k = j + 1; }
  return out;
}
function addFromSelection() {
  if (EDITOR) return false;
  const s = ta.selectionStart, e = ta.selectionEnd;
  if (e <= s) return false;
  const rs = rulesFromSelection(s, e);
  if (!rs.length) return false;
  edit(() => rs.forEach(boxAdd));
  status(`已加入黑名单：${rs.length > 1 ? rs.length + ' 条' : rs[0].slice(0, 24)}（Ctrl+Z 撤回）`);
  ta.setSelectionRange(e, e);
  return true;
}
ta.addEventListener('keydown', (e) => {
  if (!EDITOR && (e.key === 'Backspace' || e.key === 'Delete') && ta.selectionEnd > ta.selectionStart) { e.preventDefault(); addFromSelection(); }
});
ta.addEventListener('cut', () => { addFromSelection(); });
ta.addEventListener('mouseup', () => {
  if (EDITOR || !state.tools.drag) return;
  syncSel();
  if (ta.selectionEnd > ta.selectionStart) addFromSelection();
});
ta.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') { e.preventDefault(); ta.setSelectionRange(0, ta.value.length); }
});

$('addMan').onclick = () => {
  const s = ta.selectionStart, e = ta.selectionEnd;
  if (e > s) { const rs = rulesFromSelection(s, e); if (rs.length) { edit(() => rs.forEach(boxAdd)); status(`已添加：${rs[0].slice(0, 24)}${rs.length > 1 ? ` 等 ${rs.length} 条` : ''}`); return; } } // 有选区：一键加入
  const box = $('addBox');
  if (box.querySelector('input')) return;
  const inp = document.createElement('input');
  inp.placeholder = '输入要打码的文字，回车添加';
  box.appendChild(inp); inp.focus();
  inp.onkeydown = (ev) => {
    if (ev.key === 'Escape') inp.remove();
    if (ev.key === 'Enter' && inp.value.trim()) { const t = inp.value.trim(); inp.remove(); edit(() => boxAdd(t)); }
  };
};
$('addMan').onmousedown = (e) => e.preventDefault(); // 点按钮不让文本框失去选区

// ---------- 章节悬浮框：左键跳转，右键整章删除 ----------
function renderChapters() {
  const box = $('chapters');
  box.hidden = !EDITOR && !state.tools.chapter;
  if (box.hidden) return;
  chs = chapters(ta.value);
  const min = chs.reduce((a, c) => Math.min(a, c.level), 9);
  box.innerHTML = `<div class="ch-h">章节 ${chs.length} · ${EDITOR ? '点击跳转' : '左键跳转 · 右键整章加入黑名单'}</div>` + (chs.map((c, i) => `<div class="ch" data-i="${i}" style="padding-left:${6 + (c.level - min) * 12}px" title="${esc(c.title)}">${esc(c.title)}</div>`).join('') || '<div class="empty" style="padding:4px 6px">未识别到章节</div>');
}
$('chapters').addEventListener('click', (e) => {
  const c = chs[e.target.closest('.ch')?.dataset.i]; if (!c) return;
  gotoOcc(c.s, c.s);
});
$('chapters').addEventListener('contextmenu', (e) => {
  if (EDITOR) return;
  const el = e.target.closest('.ch'); if (!el) return;
  e.preventDefault();
  const c = chs[el.dataset.i]; if (!c) return;
  const lineNo = (pos) => ta.value.slice(0, pos).split('\n').length;
  const rule = lineRule.dump(lineNo(c.s), lineNo(Math.max(c.s, c.e - 1)));
  edit(() => boxAdd(rule));
  status(`章节「${c.title.slice(0, 16)}」已加入黑名单 ${rule}（Ctrl+Z 撤回）`);
});
$('chapters').addEventListener('mouseover', (e) => e.target.closest('.ch')?.classList.add('del'));
$('chapters').addEventListener('mouseout', (e) => e.target.closest('.ch')?.classList.remove('del'));

// ---------- 图片保留工具 ----------
function resetKeep() {
  doc.keep = new Set(state.cats.img ? [] : doc.images.filter((i) => i.replaceable).map((i) => i.path));
  doc.order = doc.images.map((i) => i.path);
  renderGallery();
}
function renderGallery() {
  const by = new Map(doc.images.map((i) => [i.path, i]));
  const ord = [...doc.order.filter((p) => doc.keep.has(p)), ...doc.order.filter((p) => !doc.keep.has(p))];
  $('gallery').innerHTML = ord.map((p) => {
    const i = by.get(p), tip = `${i.name} · ${(i.size / 1024).toFixed(0)} KB`;
    if (!i.replaceable) return `<button class="thumb na" title="${esc(tip)}（${i.ext} 暂不能替换，原样保留）">${i.ext}</button>`;
    return `<button class="thumb ${doc.keep.has(p) ? 'keep' : ''}" data-p="${esc(p)}" title="${esc(tip)}"><img src="${i.url}" alt=""></button>`;
  }).join('');
}
// 悬停预览（始终原色，大图）
const pv = document.createElement('img');
pv.style.cssText = 'position:fixed;z-index:20;max-width:min(520px,60vw);max-height:70vh;border:1px solid var(--line);background:var(--panel);pointer-events:none;display:none';
document.body.appendChild(pv);
$('gallery').addEventListener('mouseover', (e) => {
  const b = e.target.closest('.thumb[data-p]'); if (!b) return;
  pv.src = b.querySelector('img').src; pv.style.display = 'block';
  const r = b.getBoundingClientRect();
  pv.style.left = r.right + 12 + 'px'; pv.style.top = Math.max(8, Math.min(r.top, innerHeight - 320)) + 'px';
});
$('gallery').addEventListener('mouseout', (e) => { if (e.target.closest('.thumb')) pv.style.display = 'none'; });
$('gallery').addEventListener('contextmenu', (e) => { // 右键：在正文里定位该图片（到头后回第一处）
  const b = e.target.closest('.thumb[data-p]'); if (!b) return;
  e.preventDefault();
  const occ = (doc.images.find((i) => i.path === b.dataset.p)?.uses || []).map((x) => [x.s, x.e]);
  if (!occ.length) return status('这张图片在正文里没有出现位置（可能在页眉页脚 / 母版里）');
  locT.kind = 'img'; locT.id = b.dataset.p;
  locate(occ);
});
$('gallery').addEventListener('click', (e) => {
  const b = e.target.closest('.thumb[data-p]'); if (!b) return;
  const p = b.dataset.p;
  if (doc.keep.has(p)) doc.keep.delete(p);
  else { doc.keep.add(p); doc.order = [p, ...doc.order.filter((x) => x !== p)]; }
  renderGallery();
});

// ---------- 输出：A →（Ctrl+S）→ 脱敏/A.docx；A 与 脱敏/A →（Ctrl+Enter 审阅面板）→ 修订写入 脱敏/A（原文 A 不动） ----------
function status(t) { if (tip) statusText = t; else $('status').textContent = t; }
function redactions() { // 黑名单 + 替换规则 → md 区间替换 [{s, e, text}]
  const md = ta.value;
  nthMap.clear();
  return sel.map((x) => ({ s: x.s, e: x.e, text: replOf(x.it, plainOf(md, x.s, x.e)) }));
}
function finalMd() { // 预览：对原文应用 黑名单 + 替换规则
  return EDITOR ? ta.value : spliceMd(ta.value, redactions());
}
const alertList = (title, ws) => { if (ws.length) alert(title + '\n\n' + ws.map((w) => '• ' + w).join('\n')); };
// 输出到 原文件夹/脱敏/原名.docx，编辑器模式写 原文件夹/新版/原名.docx；审阅面板按 新版 ?? 脱敏 ?? 手选 找 B
const OUT_DIRS = ['新版', '脱敏'], OUT_DIR = EDITOR ? OUT_DIRS[0] : OUT_DIRS[1];
const outDir = async (create, name = OUT_DIR) => doc.dir && (await fs.subdir(doc.dir, name, create));
const DOCX_TYPE = [{ description: 'Word', accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] } }];
let busy = false;
async function guarded(fn) {
  if (!doc.parsed) return status('请先打开文件：点右上 ▲▼ 或把文件夹拖进来');
  if (busy) return;
  busy = true; $('apply').disabled = $('append').disabled = true;
  try { await fn(); } catch (e) { if (e.name !== 'AbortError') { console.error(e); status('失败：' + e.message); } }
  busy = false; $('apply').disabled = $('append').disabled = false;
}

// 完整应用方案：直接替换（无修订），图片 / 不可见指纹一并处理，写到 脱敏/原名.docx；可反复执行（每次都从原文重新生成）
const applyAll = () => guarded(async () => {
  const granted = !!doc.dir && (await fs.allowed(doc.dir, true, true)); // 需在用户手势内先拿权限
  const dropImages = EDITOR ? [] : doc.images.filter((i) => i.replaceable && !doc.keep.has(i.path)).map((i) => i.path);
  const d = await DocxProjection.load(doc.buffer);
  if (EDITOR) d.diffMd(ta.value); else d.splice(redactions());
  d.blankImages(dropImages);
  if (!EDITOR && state.cats.meta) d.wipeMetadata();
  const { blob, report: stats, warnings } = await d.save();
  const fh = granted ? await fs.child(await outDir(true), doc.cur.name, true) : await fs.pickSave(doc.cur.name, DOCX_TYPE);
  await fs.write(fh, blob);
  doc.outHandle = fh;
  const mh = granted && !EDITOR ? await saveManual() : null; // 没有文件夹权限（另存为）时不写
  reviewed.delete(doc.cur.name); markTitle();
  if (granted && !reviewed.size) (await outDir(false))?.removeEntry(WARN_FILE).catch(() => {}); // 本会话没有别的带修订文档了，撤掉外发警示
  alertList(`已写入 ${fh.name}，但有以下问题：`, warnings);
  if (EDITOR) return status(`已保存 ${fh.name}：${stats.del} 删 / ${stats.ins} 插` + (stats.skipped ? `，${stats.skipped} 处无法映射` : ''));
  status(`已生成 ${fh.name}：${sel.length} 处脱敏（${stats.del} 删 / ${stats.ins} 插）` + (dropImages.length ? `，${dropImages.length} 张图片已替换` : '') + (state.cats.meta ? '，不可见指纹已抹除' : '') + (stats.skipped ? `，${stats.skipped} 处无法映射` : '') + (mh ? ` · 手动修正/${mh.name}` : '') + ' · Ctrl ⏎ 可生成审阅修订');
});

// 手动修正/{}.txt：左栏两个框的内容（=== 以下就是右栏状态）。Ctrl+S 时写出，打开文档 / NER 完成时读回
const MANUAL_DIR = '手动修正';
const manualName = (ent = doc.cur) => ent.name.replace(/\.docx$/i, '') + '.txt';
async function saveManual() {
  const fh = await fs.child(await fs.subdir(doc.dir, MANUAL_DIR, true), manualName(), true);
  await fs.write(fh, [...$('rules').value.split('\n').filter((l) => l.includes('=>')), blackEl.value].join('\n')); // 含 => 的行是规则，其余是黑名单框
  return fh;
}
let saved = null; // 载入的 手动修正：右栏当时选中的文字（=== 以下 + 手填词）
async function loadManual(ent) {
  saved = null;
  if (EDITOR || !doc.dir || !(await fs.allowed(doc.dir))) return;
  try {
    const dir = await fs.subdir(doc.dir, MANUAL_DIR), fh = dir && (await fs.child(dir, manualName(ent)));
    if (!fh) return;
    const ls = (await (await fs.file(fh)).text()).split('\n'), black = ls.filter((l) => !l.includes('=>')).join('\n'), { entries, synced } = BlackBox.load(black);
    blackEl.value = black; $('rules').value = ls.filter((l) => l.includes('=>')).join('\n');
    saved = new Set([...entries, ...synced]);
  } catch (e) { console.error(e); }
}
// 存档里没有的文字（识别出来的新条目）一律不勾选；用户已手动改过的不动。有变化返回真
function restoreOn() {
  if (!saved) return false;
  let n = 0;
  for (const it of items.values()) {
    if (it.cat === 'black' || it.key in state.on) continue;
    const on = saved.has(it.text);
    if (on !== isOn(it)) { state.on[it.key] = on; n++; }
  }
  return n > 0;
}

// 审阅面板：对比 A（原文）与 脱敏/A，把 B -> A 的差异作为修订写入 脱敏/A；原文 A 一个字节都不改。不看黑名单，也可当 diff 用。
// 写入后 B 里带着原文（修订），标题标红，直到下次 Ctrl+S 重新生成 B
const WARN_FILE = 'DO_NOT_SEND-有文件含敏感信息.inf'; // 写入修订后放在 B 的文件夹里，提醒别外发
const reviewed = new Set(); // 写过修订、还没重新 Ctrl+S 的文档名
const markTitle = () => $('title').classList.toggle('tainted', !!doc.cur && reviewed.has(doc.cur.name));
const reviewToB = () => guarded(async () => {
  if (doc.dir && !(await fs.allowed(doc.dir, true, true))) throw new Error('没有写入文件夹的权限'); // 须在按键 / 点击里先拿权限
  let hB = null, bDir = null;
  for (const name of OUT_DIRS) { const od = await outDir(false, name); if ((hB = od && (await fs.child(od, doc.cur.name)))) { bDir = od; break; } }
  hB ||= doc.outHandle;
  if (!hB) { status(`未找到 ${OUT_DIRS.map((d) => d + '/' + doc.cur.name).join(' 或 ')}：请先 Ctrl+S 生成，或在弹出的窗口里选择要对比的文件`); hB = await fs.pickFile(DOCX_TYPE); }
  if (!(await fs.allowed(hB, true, true))) throw new Error(`没有写入 ${hB.name} 的权限`);
  const a = await DocxProjection.load(await fs.read(doc.cur.h)); // 现读 A：保留打开后在 Word 里做的修改
  const b = await DocxProjection.load(await fs.read(hB)); // 上次写入 B 的修订载入时已撤掉，回到 Ctrl+S 生成时的样子
  const diff = b.diffMd(a.md, { from: a });
  // 图片：A 里内容不同的媒体文件（如 B 里换成的占位图），在 B 里写成「占位图删除 + 原图插入」
  const am = new Map((await a.media()).map((x) => [x.path, x.bytes]));
  const sameBytes = (x, y) => x.length === y.length && x.every((v, i) => v === y[i]);
  const swapped = (await b.media()).filter((x) => am.has(x.path) && !sameBytes(x.bytes, am.get(x.path))).map((x) => ({ path: x.path, bytes: am.get(x.path) }));
  const imgs = b.replaceImages(swapped);
  if (!diff.del && !diff.ins && !diff.fmt && !imgs) return status(diff.skipped ? `${hB.name} 与原文有 ${diff.skipped} 处差异无法写成修订，未写入` : `${hB.name} 与原文没有差异，未写入`);
  const { blob, report: stats, warnings } = await b.save({ revisions: true });
  await fs.write(hB, blob);
  reviewed.add(doc.cur.name); markTitle();
  const warnDir = bDir || doc.dir; // B 所在文件夹（手选的文件取不到父目录，退回原文件夹）
  if (warnDir) await fs.write(await fs.child(warnDir, WARN_FILE, true), `${WARN_FILE}\n\n本文件夹里的 ${hB.name} 以修订形式带着原文（敏感信息），不能外发。\n重新 Ctrl+S 生成后确认无误，再手动删除本文件。\n`);
  alertList(`已写入 ${hB.name}，但有以下问题：\n`, warnings);
  status(`${stats.del} 处删除 / ${stats.ins} 处插入` + (stats.fmt ? ` / ${stats.fmt} 处格式` : '') + (imgs ? ` / ${imgs} 处图片` : '') + (stats.skipped ? `，${stats.skipped} 处无法映射` : ''));
});
$('apply').onclick = applyAll;
$('append').onclick = reviewToB;
addEventListener('keydown', (e) => {
  if (!e.ctrlKey && !e.metaKey && !e.altKey && e.key.toLowerCase() === 't' && !EDITOR && !e.target.closest?.('input:not([type=checkbox]), textarea:not([readonly]), [contenteditable]')) { // T：开关章节浮框
    const cb = document.querySelector('[data-k="chapter"]'); cb.checked = !cb.checked; cb.dispatchEvent(new Event('change', { bubbles: true })); return;
  }
  if (!(e.ctrlKey || e.metaKey)) return;
  const k = e.key.toLowerCase();
  if (EDITOR && k !== 's' && e.key !== 'Enter') return; // 编辑器模式：撤销交给文本框自带
  if (e.key === 'Enter') { e.preventDefault(); reviewToB(); }
  else if (k === 's') { e.preventDefault(); applyAll(); }
  else if (k === 'z' || k === 'y') {
    if (e.target.closest?.('#addBox, #custom')) return;
    e.preventDefault();
    (k === 'y' || e.shiftKey) ? redo() : undo();
  }
});

// 右上角 badge：选区总数；每次变化显示增量，增量不是 1（说明匹配到了别处或没匹配到）时暂时变黄
let lastN = null, flash = 0, flashing = false, nerInfo = '', tip = '', statusText = null;
$('pill').onclick = () => { if (doc.cur) openDoc(doc.entries.indexOf(doc.cur)); };
const PILL_TIP = '黑名单 + 替换规则命中的位置总数 · 点击：关闭并重新打开当前文档';
let pillBase = '';
// pill 只显示计数（+ NER 信息）；控件的 title 悬停提示改在底部状态栏显示
function renderPill() {
  const pill = $('pill');
  pill.textContent = pillBase + (nerInfo && !flashing ? ` · ${nerInfo}` : '');
}
function setNerInfo(t) { nerInfo = t; renderPill(); }
function updatePill() {
  const pill = $('pill'), n = sel.length;
  pill.dataset.tip = PILL_TIP;
  if (lastN !== null && n !== lastN) {
    const d = n - lastN;
    pillBase = `${d > 0 ? '+' : ''}${d} 次删除`;
    pill.classList.toggle('hot', d !== 1);
    flashing = true; clearTimeout(flash);
    flash = setTimeout(() => { flashing = false; pill.classList.remove('hot'); pillBase = `总共${lastN}次删除`; renderPill(); }, 1000);
  } else if (!flashing) pillBase = `总共${n}次删除`;
  lastN = n;
  renderPill();
}
// 悬停提示：把 title 挪到 data-tip（去掉原生气泡），mouseover / mouseout 时在底部状态栏显示 / 还原
function setTip(t) {
  const st = $('status');
  if (t && !tip) statusText = st.textContent;
  if (!t && !tip) return;
  tip = t;
  st.textContent = t || statusText;
}
let tipTm = 0, tipWant = '';
function showTip(t) { // 长提示（含换行或 >= 35 字）：悬停 1 秒才显示（同一个提示等待期间不重新计时）
  if (t === tipWant) return;
  tipWant = t; clearTimeout(tipTm);
  if (t.includes('\n') || [...t].length >= 35) { setTip(''); tipTm = setTimeout(() => setTip(t), 1000); } else setTip(t);
}
addEventListener('mouseover', (e) => {
  const el = e.target.closest?.('[title], [data-tip]');
  if (el?.hasAttribute('title')) { el.dataset.tip = el.title; el.removeAttribute('title'); }
  showTip(el?.dataset.tip || '');
});
addEventListener('mouseout', (e) => { if (!e.relatedTarget) showTip(''); });
// ---------- 文件：右上 ▲▼ 选择 / 打开文件夹 ----------
const shortName = (name) => {
  const stem = name.replace(/\.docx$/i, '');
  let head = stem.split(/\s+/)[0];
  if (head.length > 10) head = head.slice(0, 10);
  return head === stem ? name : head + '….docx';
};
function setFileInfo(name, size) { $('fileInfo').textContent = name ? `${shortName(name)} (${mb(size)})` : '打开文件夹…'; }

async function openDoc(i, auto = false, noNer = false) {
  if (!auto) batch = false; // 手动切换文档即中止批量
  const ent = doc.entries[i]; if (!ent) return;
  status('读取中…');
  const file = await fs.file(ent.h);
  doc.buffer = await file.arrayBuffer(); doc.cur = ent;
  const p = await DocxProjection.load(doc.buffer);
  doc.parsed = p; doc.orig = p.md;
  doc.images.forEach((x) => x.url && URL.revokeObjectURL(x.url));
  doc.images = (await p.media()).map((i) => ({ ...i, bytes: null, url: i.mime ? URL.createObjectURL(new Blob([i.bytes], { type: i.mime })) : '' }));
  resetKeep();
  ta.value = p.md; ta.setSelectionRange(0, 0); ta.scrollTop = 0; state.ner = []; state.on = {}; memo.clear(); 
  doc.outHandle = null; lastN = null; undoS.length = redoS.length = 0;
  await loadManual(ent); // 有存档就以存档为准，否则保留黑名单词（行号规则属于旧文档故丢弃）
  if (!saved) blackEl.value = BlackBox.dump({ own: black().words });
  $('title').textContent = ent.name.replace(/\.docx$/i, ''); markTitle();
  setFileInfo(ent.name, file.size); setHash(ent.name);
  run(); if (restoreOn()) run();
  status(saved ? '已载入 · 已读回 手动修正/' + manualName() : `已载入 · Ctrl S → ${OUT_DIR}/{}.docx · Ctrl ⏎ 审阅面板 → ${OUT_DIR}/{}.docx`);
  alertList(`${ent.name}：以下内容没有处理`, p.warnings);
  if (noNer) { nerAbort?.abort(); setNerInfo(''); $('bannerText').textContent = '已切换文档（未做 NER）：点右上 ▲▼ 重新选中本文档即可识别。'; }
  else if (ner.ready()) runNer();
}
let nerAbort = null;
const fmtSec = (n) => (n >= 90 ? `${Math.round(n / 60)} 分钟` : `${Math.max(1, Math.round(n))} 秒`);
async function runNer() {
  if (!doc.parsed) return;
  nerAbort?.abort(); // 切换文档 / 重新识别时，丢弃上一轮
  const ac = nerAbort = new AbortController();
  const lines = nerLines(doc.orig);
  const base = EDITOR ? '纯文本投影编辑器：直改docx内文本，Ctrl+S 保存为 新版/原名.docx；右栏用于手工修订' : '已提取纯文本部分，普通表格按行对应，复杂表格1格1行。请确保处理敏感文字。【右键点规则/专利号即可在全文中定位】';
  $('bannerText').textContent = '离线 NER 识别中…'; setNerInfo('NER 识别中…');
  const t0 = performance.now();
  try {
    const res = await ner.recognize(lines, (f, i) => {
      const eta = i.etaSec > 1 && f > 0.03 ? ` · 剩约 ${fmtSec(i.etaSec)}` : '';
      setNerInfo(`NER ${(f * 100) | 0}%${eta} · ${ner.backend}${i.tokPerSec ? ` · ${Math.round(i.tokPerSec)} token/s` : ''}`);
      status(`NER ${(f * 100) | 0}%${eta}`);
    }, { signal: ac.signal, onStatus: status });
    if (ac.signal.aborted) return;
    state.ner = res;
    $('bannerText').textContent = base;
    setNerInfo(`NER ${((performance.now() - t0) / 1000).toFixed(1)}s · ${ner.backend === '待启动' ? '缓存' : ner.backend}`);
    status('NER 完成');
    fs.remember('nercache', ner.cacheExport());
  } catch (e) {
    if (e.name === 'AbortError' || ac.signal.aborted) return;
    console.error(e); $('bannerText').textContent = 'NER 失败：' + e.message; setNerInfo(''); batch = false;
  }
  EDITOR ? scanPanels() : run();
  if (!EDITOR && restoreOn()) run(); // NER 条目这时才出现，存档里没有的不勾选
  if (!EDITOR) await nextBatch();
}
// 批量缓存：NER 完成时若文件列表 > 4 项，问一次是否把其余文件也依次跑完（结果进 NER 缓存），全自动逐个打开
let batch = false, batchAsked = false;
async function nextBatch() {
  const i = doc.entries.indexOf(doc.cur), left = doc.entries.length - i - 1;
  if (!batchAsked && doc.entries.length > 4 && left > 0) {
    batchAsked = true;
    batch = confirm(`文件列表有 ${doc.entries.length} 项，是否对其余 ${left} 项文件批量缓存 NER 结果？\n（将自动逐个打开并识别，可随时手动切换文件中止）`);
  }
  if (!batch) return;
  if (left > 0) return openDoc(i + 1, true);
  batch = false; status('批量缓存完成');
}

// 模型：未载入时右上角 ⚠️
const warn = $('modelWarn');
async function loadModelFile(h, name) {
  const file = await fs.file(h);
  warn.hidden = false; warn.textContent = '⚠️ 离线模型载入中…';
  try {
    const be = await ner.loadModel(file); // 只登记文件；推理线程在第一次识别时才拉起
    ner.cacheImport(await fs.recall('nercache')); // 缓存按内容寻址，模型哈希对得上才会启用
    warn.hidden = true; $('banner').title = `${name} (${mb(file.size)}) · ${be}`;
    status(`模型就绪 · ${be}`);
    if (h.getFile) fs.remember('model', h);
    if (doc.parsed) runNer(); else run();
  } catch (e) { console.error(e); warn.textContent = '⚠️ 模型载入失败：' + e.message; }
}
warn.onclick = async () => {
  try {
    if (pendingModel && (await fs.allowed(pendingModel, false, true))) return loadModelFile(pendingModel, pendingModel.name);
    const h = await fs.pickFile([{ description: 'ONNX', accept: { 'application/octet-stream': ['.onnx'] } }]);
    await loadModelFile(h, h.name);
  } catch (e) { if (e.name !== 'AbortError') status(e.message); }
};

// 返回文件夹里是否有模型
async function openFolder(dir) {
  doc.dir = dir; doc.entries = []; pendingDir = null; batch = batchAsked = false; let model = null;
  fs.remember('dir', dir);
  for (const { name, h } of await fs.list(dir)) {
    if (/\.docx$/i.test(name) && !name.startsWith('~$') && !/(?:_脱敏|-红化)\.docx$/.test(name)) doc.entries.push({ name, h });
    else if (/\.onnx$/i.test(name) && (!model || /quant|q8/i.test(name))) model = { name, h };
  }
  doc.entries.sort((a, b) => a.name.localeCompare(b.name, 'zh'));
  if (model && !ner.ready()) loadModelFile(model.h, model.name);
  if (doc.entries.length) await openDoc(Math.max(0, doc.entries.findIndex((e) => e.name === hashFile)));
  else status('该文件夹没有 .docx');
  return !!model;
}
async function pickFolder() {
  try { await handleItems([await fs.pickDir()]); } catch (e) { if (e.name !== 'AbortError') status(e.message); }
}
async function resumeFolder() { // 继续上次的文件夹：点击时请求权限（Chromium 可选「每次访问都允许」），被拒则改为选择
  try { if (await fs.allowed(pendingDir, true, true)) return await openFolder(pendingDir); } catch (e) { console.error(e); }
  pickFolder();
}
const menu = $('fileMenu');
async function openMenu() {
  menu.innerHTML = doc.entries.map((e, i) => `<div data-i="${i}" class="${e === doc.cur ? 'cur' : ''}">${esc(e.name)}<small>${e.size == null ? '' : ' · ' + mb(e.size)}</small></div>`).join('') + '<div class="act" data-act="dir">📂 打开文件夹…</div>';
  menu.hidden = false;
  for (const [i, e] of doc.entries.entries()) { // 大小按需读取并缓存
    if (e.size == null) e.size = await Promise.resolve(fs.file(e.h)).then((f) => f.size, () => 0);
    const el = menu.querySelector(`[data-i="${i}"] small`); if (el) el.textContent = ' · ' + mb(e.size);
  }
}
$('fileSel').addEventListener('click', (e) => {
  const it = e.target.closest('#fileMenu > div');
  if (it) { menu.hidden = true; if (it.dataset.act) pickFolder(); else openDoc(+it.dataset.i); return; }
  if (!doc.entries.length) return pendingDir ? resumeFolder() : pickFolder();
  menu.hidden ? openMenu() : (menu.hidden = true);
});
addEventListener('click', (e) => { if (!e.target.closest('#fileSel')) menu.hidden = true; });
// 文件选择框聚焦时 ↑↓：上一个 / 下一个文档，只切换不做 NER
const step = (d) => { const i = doc.entries.indexOf(doc.cur) + d; if (doc.cur && doc.entries[i]) { menu.hidden = true; openDoc(i, false, true); } };
$('fileSel').addEventListener('keydown', (e) => { if (e.key === 'ArrowUp' || e.key === 'ArrowDown') { e.preventDefault(); step(e.key === 'ArrowUp' ? -1 : 1); } });
$('fileSel').addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('fileSel').click(); } });

async function handleItems(handles) {
  const files = [];
  for (const h of handles) { if (h.kind === 'directory') { await openFolder(h); return; } files.push(h); }
  let last = -1;
  for (const h of files) {
    if (/\.onnx$/i.test(h.name)) await loadModelFile(h, h.name);
    else if (/\.docx$/i.test(h.name)) { doc.entries.push({ name: h.name, h }); last = doc.entries.length - 1; }
  }
  if (last >= 0) await openDoc(last);
}
const dm = $('dropMask');
let dragN = 0;
addEventListener('dragenter', (e) => { e.preventDefault(); dragN++; dm.classList.add('on'); });
addEventListener('dragleave', () => { if (--dragN <= 0) { dragN = 0; dm.classList.remove('on'); } });
addEventListener('dragover', (e) => e.preventDefault());
addEventListener('drop', async (e) => {
  e.preventDefault(); dragN = 0; dm.classList.remove('on');
  const hs = [];
  for (const it of e.dataTransfer.items) if (it.kind === 'file') { const h = await it.getAsFileSystemHandle?.(); if (h) hs.push(h); }
  if (hs.length) handleItems(hs).catch((err) => { console.error(err); status(err.message); });
});

// ---------- 找回上次的文件夹、模型：已有「每次访问都允许」的权限就直接打开，否则等点击 ----------
(async () => {
  warn.hidden = false;
  let hasModel = false;
  const dir = await fs.recall('dir');
  if (dir && !doc.entries.length) {
    pendingDir = dir;
    if (await fs.allowed(dir, true).catch(() => false)) hasModel = await openFolder(dir).catch((e) => { console.error(e); return false; });
    else $('fileInfo').textContent = `📂 继续上次：${dir.name}`;
  }
  const h = await fs.recall('model');
  if (h && !hasModel) {
    pendingModel = h;
    if (await fs.allowed(h).catch(() => false)) loadModelFile(h, h.name);
    else warn.textContent = `⚠️ 离线模型未载入（点击载入 ${h.name}）`;
  }
})();
if (!('showDirectoryPicker' in window)) status('需要基于 Chromium 的浏览器（File System Access API）');
// F12：__di.rules() 查看；__di.rules({ mobile: { re: /…/g }, 新规则: { name, re }, X: null /*删除*/, MAIN: [...], EXTRA: [...], TOOLS: [...], REPLACE: '替换规则文本', TAG_CATS: [...], ENT_CATS: [...], HIDE_WHEN_OFF: [...] }) 动态替换预制规则并重新识别
function rulesApi(patch) {
  const lists = { MAIN, EXTRA, TOOLS }, sets = { TAG_CATS, ENT_CATS, HIDE_WHEN_OFF };
  for (const [k, v] of Object.entries(patch || {})) {
    if (lists[k]) lists[k].splice(0, lists[k].length, ...v);
    else if (sets[k]) { sets[k].clear(); v.forEach((x) => sets[k].add(x)); }
    else if (k === 'REPLACE') $('rules').value = v;
    else if (v === null) delete RULES[k];
    else {
      const r = (RULES[k] ||= { name: k, cls: 'custom' });
      Object.assign(r, v);
      if (typeof r.re === 'string') r.re = new RegExp(r.re, 'g');
      else if (r.re && !r.re.global) r.re = new RegExp(r.re.source, r.re.flags + 'g');
      r.en ||= k;
      if (![...MAIN, ...EXTRA].some((x) => x.id === r.en)) EXTRA.push({ id: r.en, name: r.name, def: true });
    }
  }
  for (const x of [...MAIN, ...EXTRA]) state.cats[x.id] ??= x.def !== false;
  for (const t of TOOLS) state[t.scope][t.id] ??= t.def;
  renderLeft(); run();
  return { RULES, MAIN, EXTRA, TOOLS, TAG_CATS, ENT_CATS, HIDE_WHEN_OFF, REPLACE: $('rules').value };
}

// ---------- #editor 模式 ----------
if (EDITOR) {
  document.title = '文本投影编辑器';
  $('title').textContent = '未打开文件';
  ta.contentEditable = 'plaintext-only';
  ta.dataset.ph = '把文件夹（含 .docx）拖到这里，或点右上 ▲▼ 打开文件夹；直接编辑文本，Ctrl+S 保存到 新版/原名.docx';
  $('bannerText').textContent = '纯文本投影编辑器：脱敏功能已关闭，直接修改文本，Ctrl+S 保存为 新版/原名.docx（只改动有差异的段落）。';
  $('apply').innerHTML = '保存修改<small>（到 新版/原名.docx）</small>';
  $('apply').title = 'Ctrl+S：把修改后的文本写回 .docx，保存为 新版/原名.docx';
  $('status').textContent = 'Ctrl S → 新版/ · Ctrl ⏎ 审阅面板 → 新版/ · 右栏仅在粘贴时更新';
  document.querySelector('.left').prepend($('chapters'));
  ta.addEventListener('input', () => refresh());
  ta.addEventListener('paste', (e) => { // 右栏：只在粘贴时更新——对粘贴内容做 NER，再对全文重新识别
    const text = e.clipboardData?.getData('text') || '';
    setTimeout(async () => {
      if (text && ner.ready()) {
        try {
          const add = await ner.recognize(nerLines(text), () => {}, { onStatus: status });
          const have = new Set(state.ner.map((n) => n.type + n.text));
          for (const n of add) if (!have.has(n.type + n.text)) state.ner.push(n);
        } catch (err) { console.error(err); status('NER 失败：' + err.message); }
      }
      scanPanels();
    }, 0);
  });
}
function scanPanels() {
  items = scan(ta.value, state);
  rankNer();
  live = [...items.values()].filter(listed);
  sel = selection(live, isOn);
  updateCounts(); renderPanels();
}

window.__di = { ruleNth, rules: rulesApi, openDoc, state, ta, run, ner, fs, DocxProjection, lineRange, doc, boxAdd, edit, undo, redo, finalMd, redactions, applyAll, reviewToB, get items() { return items; }, get sel() { return sel; } };
run();
