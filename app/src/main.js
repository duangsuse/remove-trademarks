import './style.css';
import { parseDocx, computeEdits, exportDocx, listImages } from './docx.js';
import { scan, parseRules, parseCustom, parseBlackLines, replacement, selection, chapters } from './detect.js';
import { MAIN, EXTRA, TOOLS, TYPE_NAME, DEFAULT_REPLACE_RULES, HIDE_WHEN_OFF, TAG_CATS, ENT_CATS, TITLE_RE, ABBR_MAX, RULES, FINGERPRINT_URL } from './rules.js';
import * as ner from './ner.js';

const $ = (id) => document.getElementById(id);
const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const mb = (n) => (n / 1048576).toFixed(2) + ' MB';
const store = { get: (k) => { try { return localStorage.getItem(k); } catch { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch { /* 无存储 */ } } };

// ---------- 主题：点 logo 切换 ----------
const root = document.documentElement;
root.dataset.theme = store.get('theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
$('logo').onclick = () => { root.dataset.theme = root.dataset.theme === 'dark' ? 'light' : 'dark'; store.set('theme', root.dataset.theme); };

// ---------- 状态 ----------
// 脱敏的唯一来源是左栏两个大框：黑名单（词 / 行号 #L11~L222）与「删除时（替换）规则」。
// 右侧面板、正文里的删除只是在编辑这两个框；正文 #ta 只读。
const state = { cats: { custom: true }, tools: {}, blacklist: [], ranges: [], custom: [], ner: [], on: {} };
for (const x of [...MAIN, ...EXTRA]) state.cats[x.id] = x.def !== false;
for (const t of TOOLS) state[t.scope][t.id] = t.def;
const doc = { entries: [], dir: null, cur: null, buffer: null, parsed: null, orig: '', images: [], order: [], keep: new Set(), outHandle: null };
let pendingModel = null;

const row = (x, scope) => {
  const tail = x.info ? `<a class="info" href="${FINGERPRINT_URL}" target="_blank" rel="noopener" title="微软：删除隐藏数据和个人信息">ⓘ</a>`
    : scope === 'cats' && !TOOLS.some((t) => t.id === x.id) ? `<span class="n" data-n="${x.id}"></span>` : '';
  return `<label><input type="checkbox" data-s="${scope}" data-k="${x.id}" ${state[scope][x.id] ? 'checked' : ''}><span>${x.name}${tail}${x.note ? ` <span class="n">- ${x.note}</span>` : ''}</span></label>`;
};
$('catList').innerHTML = MAIN.map((x) => row(x, 'cats')).join('');
$('extraList').innerHTML = EXTRA.map((x) => row(x, 'cats')).join('');
$('moreN').textContent = `(${EXTRA.length})`;
$('toolList').innerHTML = TOOLS.map((t) => row(t, t.scope)).join('');
$('rules').value = DEFAULT_REPLACE_RULES;

document.addEventListener('change', (e) => {
  const t = e.target, k = t.dataset?.k;
  if (!k) return;
  state[t.dataset.s][k] = t.checked;
  if (k === 'img') resetKeep();
  refresh(true);
});

// ---------- 黑名单框：=== 以上手填，以下随右侧面板同步 ----------
const SEP = '===', blackEl = $('black');
let lastGen = new Set();
const base = (it) => !(it.cat === 'ner' && it.grp === 'cand'); // 机构/产品候选（>3 字）默认不选
const memo = new Map(); // key -> {text,en,base}：见过的条目；文中不再出现（比如已脱敏）也保留，黑名单框不随文本变化
const memoOn = (key, m) => (key in state.on ? state.on[key] : !!state.cats[m.en] && m.base);
function remember() {
  for (const it of items.values()) if (it.cat !== 'black') memo.set(it.key, { text: it.text, en: it.en, base: base(it) });
}
const splitBlack = () => {
  const lines = blackEl.value.split('\n'), i = lines.findIndex((l) => l.trim() === SEP);
  return { head: i < 0 ? lines : lines.slice(0, i), tail: i < 0 ? [] : lines.slice(i + 1) };
};
const words = (ls) => ls.map((l) => l.trim()).filter(Boolean);
blackEl.value = '\n' + SEP;
function readHead() {
  const { words: w, ranges } = parseBlackLines(splitBlack().head);
  state.blacklist = w; state.ranges = ranges;
}
function syncBlack() {
  const { head } = splitBlack();
  const own = new Set(words(head)), gen = [], seen = new Set();
  for (const [key, m] of memo) {
    if (!memoOn(key, m) || own.has(m.text) || seen.has(m.text)) continue;
    seen.add(m.text); gen.push(m.text);
  }
  lastGen = seen;
  const headText = head.join('\n');
  const next = headText + '\n' + SEP + (gen.length ? '\n' + gen.slice(0, 800).join('\n') : '');
  if (next === blackEl.value) return;
  if (document.activeElement === blackEl && blackEl.selectionStart > headText.length) return; // 正在编辑 === 以下，失焦后再同步
  const [a, b, t] = [blackEl.selectionStart, blackEl.selectionEnd, blackEl.scrollTop];
  blackEl.value = next;
  if (document.activeElement === blackEl) blackEl.setSelectionRange(Math.min(a, headText.length), Math.min(b, headText.length));
  blackEl.scrollTop = t;
}
blackEl.oninput = () => refresh();
blackEl.onchange = () => { // 编辑了 === 以下：删行=取消选中，新增行=选中同名项，无同名项则并入黑名单
  const { head, tail } = splitBlack();
  const want = new Set(words(tail)), own = new Set(words(head));
  const known = new Set();
  for (const [key, m] of memo) {
    known.add(m.text);
    if (lastGen.has(m.text) && !want.has(m.text)) state.on[key] = false;
    else if (!lastGen.has(m.text) && want.has(m.text)) state.on[key] = true;
  }
  const add = [...want].filter((w) => !lastGen.has(w) && !own.has(w) && !known.has(w));
  if (add.length) blackEl.value = [...words(head), ...add].join('\n') + '\n' + SEP;
  run();
};
$('custom').oninput = () => { state.custom = parseCustom($('custom').value); refresh(); };
$('rules').oninput = () => refresh();

// 对黑名单框的结构化修改（右侧面板 / 正文删除 / 章节删除 都走这里）
const sameLine = (a, b) => a === b || (() => { const x = parseBlackLines([a]), y = parseBlackLines([b]); return x.ranges.length && JSON.stringify(x.ranges) === JSON.stringify(y.ranges); })();
function boxAdd(text) {
  const { head } = splitBlack();
  if (words(head).some((l) => sameLine(l, text))) return;
  blackEl.value = [...words(head), text].join('\n') + '\n' + SEP;
}
function boxRemove(text) {
  const { head } = splitBlack();
  blackEl.value = words(head).filter((l) => !sameLine(l, text)).join('\n') + '\n' + SEP;
}
function setOn(it, on) {
  if (it.cat === 'black') { if (!on) boxRemove(it.text); } else state.on[it.key] = on;
}

// ---------- 撤销 / 重做：只管黑名单 + 删除时替换规则 ----------
const undoS = [], redoS = [];
let lastPush = 0;
const snap = () => ({ head: splitBlack().head.join('\n'), on: { ...state.on }, rules: $('rules').value });
function checkpoint(force) {
  const now = Date.now();
  if (!force && now - lastPush < 800) { lastPush = now; return; } // 连续输入合并为一步
  undoS.push(snap()); if (undoS.length > 200) undoS.shift();
  redoS.length = 0; lastPush = now;
}
function restore(sn) {
  blackEl.value = sn.head + '\n' + SEP; state.on = { ...sn.on }; $('rules').value = sn.rules;
  lastPush = 0; run();
}
function undo() { if (!undoS.length) return status('没有可撤回的操作（只记录黑名单与替换规则）'); redoS.push(snap()); restore(undoS.pop()); status('已撤回'); }
function redo() { if (!redoS.length) return; undoS.push(snap()); restore(redoS.pop()); status('已重做'); }
const edit = (fn) => { checkpoint(true); fn(); run(); };
blackEl.addEventListener('beforeinput', () => checkpoint(false));
$('rules').addEventListener('beforeinput', () => checkpoint(false));

// ---------- 识别 / 渲染 ----------
const ta = $('ta'), hl = $('hl');
let items = new Map(), live = [], sel = [], tm = 0, chs = [];
const enabled = (it) => !!state.cats[it.en];
const listed = (it) => enabled(it) || !HIDE_WHEN_OFF.has(it.en);
const isOn = (it) => (it.key in state.on ? state.on[it.key] : enabled(it) && base(it));
const cls = (it) => (it.warn || TITLE_RE.test(it.text) ? 'warn' : it.cls); // 称谓代词（XX经理）红色

function refresh(now) { clearTimeout(tm); tm = setTimeout(run, now ? 0 : 160); }
function run() {
  clearTimeout(tm);
  readHead();
  const md = ta.value;
  items = scan(md, state);
  live = [...items.values()].filter(listed);
  sel = selection(live, isOn);
  const occ = [];
  for (const it of live) for (const [s, e] of it.occ) occ.push({ s, e, it });
  occ.sort((a, b) => a.s - b.s || b.e - a.e);
  let html = '', p = 0;
  for (const o of occ) {
    if (o.s < p) continue;
    html += esc(md.slice(p, o.s)) + `<mark class="${cls(o.it)} ${isOn(o.it) ? '' : 'off'}">${esc(md.slice(o.s, o.e))}</mark>`;
    p = o.e;
  }
  hl.innerHTML = html + esc(md.slice(p)) + '\n ';
  hl.scrollTop = ta.scrollTop;
  updateCounts();
  renderPanels();
  renderChapters();
  remember();
  syncBlack();
  updatePill();
}
ta.addEventListener('scroll', () => { hl.scrollTop = ta.scrollTop; });

const rules = () => parseRules($('rules').value);
const replOf = (it, text) => (it.del ? '' : replacement(text, undefined, rules()));

function updateCounts() {
  const cnt = {};
  for (const it of items.values()) cnt[it.en] = (cnt[it.en] || 0) + 1;
  cnt.img = doc.images.length;
  for (const el of document.querySelectorAll('[data-n]')) el.textContent = ` (${cnt[el.dataset.n] || 0})`;
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
const panels = { ent: [], place: [], tag: [], cand: [] };
const chip = (it) => `<span class="badge ${cls(it)} ${isOn(it) ? '' : 'off'}" data-key="${esc(it.key)}">${esc(it.text)}<small>${it.occ.length}</small></span>`;
const kindOf = (it) => (it.cat === 'black' ? (it.del ? '行号' : '手动') : it.cat === 'ner' ? TYPE_NAME[it.type] || it.type : RULES[it.cat]?.name || it.label || '规则');
const rowOf = (it, extra = '') => `<div class="row"><input type="checkbox" data-key="${esc(it.key)}" ${isOn(it) ? 'checked' : ''}><div class="t">${esc(it.text)}<small>${kindOf(it)} · ${it.occ.length}次</small></div>${extra}${locBtn(it)}<button class="ren" data-ren="${esc(it.key)}" title="重命名（留空即删除）">✎</button></div>`;
function renderPanels() {
  for (const k in panels) panels[k] = [];
  for (const it of live) panels[panelOf(it)].push(it);
  panels.cand.sort((a, b) => (b.score || 0) - (a.score || 0) || b.occ.length - a.occ.length);
  const count = (list) => { if (!list.length) return '0'; const t = list.reduce((a, i) => a + i.occ.length, 0), o = list.filter(isOn).reduce((a, i) => a + i.occ.length, 0); return `${o} OF ${t}`; };
  $('entList').innerHTML = panels.ent.map((it) => rowOf(it)).join('') || '<div class="empty">暂无。点下方 Add，或在正文选中文字按 Delete</div>';
  $('cnt-man').textContent = count(panels.ent);
  $('placeList').innerHTML = panels.place.map(chip).join('') || `<span class="empty">${ner.ready() ? '无' : '未载入模型，仅规则识别可用'}</span>`;
  $('cnt-place').textContent = count(panels.place);
  $('tagList').innerHTML = panels.tag.map(chip).join('') || '<span class="empty">无</span>';
  $('cnt-tag').textContent = count(panels.tag);
  $('candList').innerHTML = panels.cand.map((it) => rowOf(it, `<span class="score">${it.cat === 'ner' ? it.score.toFixed(2) : '0'}</span>`)).join('') || `<div class="empty">${ner.ready() ? '无备选' : '请先载入离线模型'}</div>`;
  $('cnt-cand').textContent = count(panels.cand);
  $('sum').textContent = `${sel.length} of ${live.reduce((a, i) => a + i.occ.length, 0)}`;
}
// 定位到下一个出现：剩 N ↓ ；到头后变「回顶 ↥」
const visited = {};
const remaining = (it) => it.occ.filter((o) => o[0] > (visited[it.key] ?? -1)).length;
const locBtn = (it) => {
  if (!it.occ.length) return '<button class="loc" disabled>—</button>';
  const n = remaining(it);
  return `<button class="loc" data-loc="${esc(it.key)}" title="${n ? '定位到下一个出现' : '回到文首'}">${n ? `剩${n} ↓` : '回顶 ↥'}</button>`;
};
function gotoOcc(s, e) {
  ta.focus(); ta.setSelectionRange(s, e);
  ta.scrollTop = Math.max(0, yOf(s) - 80);
}
const rightEl = document.querySelector('.right');
rightEl.addEventListener('click', (e) => {
  const all = e.target.closest('[data-all]');
  if (all) { const list = panels[all.dataset.all]; if (list.length) { const to = !list.every(isOn); edit(() => list.forEach((it) => setOn(it, to))); } return; }
  const l = e.target.closest('[data-loc]');
  if (l) {
    const it = items.get(l.dataset.loc); if (!it) return;
    const o = it.occ.find((x) => x[0] > (visited[it.key] ?? -1));
    if (o) { visited[it.key] = o[0]; gotoOcc(o[0], o[1]); } else { visited[it.key] = -1; ta.scrollTop = 0; ta.setSelectionRange(0, 0); hl.scrollTop = 0; }
    renderPanels(); return;
  }
  const r = e.target.closest('[data-ren]');
  if (r) {
    const it = items.get(r.dataset.ren); if (!it) return;
    const t = prompt('重命名（留空即删除）', it.text);
    if (t === null) return;
    edit(() => { setOn(it, false); if (t.trim()) boxAdd(t.trim()); });
    return;
  }
  const b = e.target.closest('.badge'); if (b) { const it = items.get(b.dataset.key); edit(() => setOn(it, !isOn(it))); }
});
rightEl.addEventListener('change', (e) => {
  const t = e.target;
  if (t.type === 'checkbox' && t.dataset.key) { const it = items.get(t.dataset.key); if (it) edit(() => setOn(it, t.checked)); }
});

// ---------- 正文只读：选中后删除 / 拖选 / 点高亮 = 往黑名单加规则 ----------
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
  if (hits.length === 1) { const t = hits[0].ln.slice(hits[0].os, hits[0].oe).trim(); return t ? [t] : []; }
  for (const h of hits) { if (h.whole) full.push(h.i + 1); else { const t = h.ln.slice(h.os, h.oe).trim(); if (t) out.push(t); } }
  for (let k = 0; k < full.length;) { let j = k; while (full[j + 1] === full[j] + 1) j++; out.unshift(j > k ? `#L${full[k]}~L${full[j]}` : `#L${full[k]}`); k = j + 1; }
  return out;
}
function addFromSelection() {
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
  if ((e.key === 'Backspace' || e.key === 'Delete') && ta.selectionEnd > ta.selectionStart) { e.preventDefault(); addFromSelection(); }
});
ta.addEventListener('cut', () => { addFromSelection(); });
ta.addEventListener('mouseup', () => {
  const s = ta.selectionStart, e = ta.selectionEnd;
  if (e > s) { if (state.tools.drag) addFromSelection(); return; }
  if (!state.tools.click) return;
  for (const it of live) for (const [os, oe] of it.occ) if (s >= os && s < oe) {
    if (isOn(it)) { status(`「${it.text.slice(0, 20)}」已在黑名单`); return; }
    edit(() => { if (it.cat === 'black') return; boxAdd(it.text); });
    status(`已加入黑名单：${it.text.slice(0, 24)}`);
    return;
  }
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
function yOf(pos) {
  const w = document.createTreeWalker(hl, NodeFilter.SHOW_TEXT);
  let n, acc = 0;
  while ((n = w.nextNode())) {
    const l = n.nodeValue.length;
    if (pos <= acc + l) {
      const r = document.createRange(), o = pos - acc;
      r.setStart(n, Math.min(o, l)); r.setEnd(n, Math.min(o + 1, l));
      return r.getBoundingClientRect().top - hl.getBoundingClientRect().top + hl.scrollTop;
    }
    acc += l;
  }
  return 0;
}
function renderChapters() {
  const box = $('chapters');
  box.hidden = !state.tools.chapter;
  if (!state.tools.chapter) return;
  chs = chapters(ta.value);
  const min = chs.reduce((a, c) => Math.min(a, c.level), 9);
  box.innerHTML = `<div class="ch-h">章节 ${chs.length} · 左键跳转 · 右键整章加入黑名单</div>` + (chs.map((c, i) => `<div class="ch" data-i="${i}" style="padding-left:${6 + (c.level - min) * 12}px" title="${esc(c.title)}">${esc(c.title)}</div>`).join('') || '<div class="empty" style="padding:4px 6px">未识别到章节</div>');
}
$('chapters').addEventListener('click', (e) => {
  const c = chs[e.target.closest('.ch')?.dataset.i]; if (!c) return;
  ta.focus(); ta.setSelectionRange(c.s, c.s); ta.scrollTop = Math.max(0, yOf(c.s) - 30);
});
$('chapters').addEventListener('contextmenu', (e) => {
  const el = e.target.closest('.ch'); if (!el) return;
  e.preventDefault();
  const c = chs[el.dataset.i]; if (!c) return;
  const lineNo = (pos) => ta.value.slice(0, pos).split('\n').length;
  const rule = `#L${lineNo(c.s)}~L${lineNo(Math.max(c.s, c.e - 1))}`;
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
$('gallery').addEventListener('click', (e) => {
  const b = e.target.closest('.thumb[data-p]'); if (!b) return;
  const p = b.dataset.p;
  if (doc.keep.has(p)) doc.keep.delete(p);
  else { doc.keep.add(p); doc.order = [p, ...doc.order.filter((x) => x !== p)]; }
  renderGallery();
});

// ---------- 输出：A →（Ctrl+S）→ A_脱敏.docx；A 与 A_脱敏 →（Ctrl+Enter 审阅面板）→ 修订写入 A ----------
function status(t) { $('status').textContent = t; }
function finalMd() { // 对原文应用 黑名单 + 替换规则
  const md = ta.value;
  let out = '', p = 0;
  for (const x of sel) { out += md.slice(p, x.s) + replOf(x.it, md.slice(x.s, x.e)); p = x.e; }
  return out + md.slice(p);
}
const outName = () => doc.cur.name.replace(/\.docx$/i, '') + '_脱敏.docx';
const DOCX_TYPE = [{ description: 'Word', accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] } }];
const write = async (fh, blob) => { const w = await fh.createWritable(); await w.write(blob); await w.close(); };
let busy = false;
async function guarded(fn) {
  if (!doc.parsed) return status('请先打开文件：点右上 ▼ 或把文件夹拖进来');
  if (busy) return;
  busy = true; $('apply').disabled = $('append').disabled = true;
  try { await fn(); } catch (e) { if (e.name !== 'AbortError') { console.error(e); status('失败：' + e.message); } }
  busy = false; $('apply').disabled = $('append').disabled = false;
}

// 完整应用方案：直接替换（无修订），图片 / 不可见指纹一并处理，写到 原名_脱敏.docx；可反复执行（每次都从原文重新生成）
const applyAll = () => guarded(async () => {
  const granted = doc.dir ? (await doc.dir.requestPermission?.({ mode: 'readwrite' })) === 'granted' : false; // 需在用户手势内先拿权限
  const dropImages = doc.images.filter((i) => i.replaceable && !doc.keep.has(i.path)).map((i) => i.path);
  const { blob, stats } = await exportDocx(doc.buffer, finalMd(), { track: false, dropImages, wipe: !!state.cats.meta });
  const fh = granted ? await doc.dir.getFileHandle(outName(), { create: true }) : await showSaveFilePicker({ suggestedName: outName(), types: DOCX_TYPE });
  await write(fh, blob);
  doc.outHandle = fh;
  status(`已生成 ${fh.name}：${sel.length} 处脱敏（${stats.del} 删 / ${stats.ins} 插）` + (dropImages.length ? `，${dropImages.length} 张图片已替换` : '') + (state.cats.meta ? '，不可见指纹已抹除' : '') + (stats.skipped ? `，${stats.skipped} 处无法映射` : '') + ' · Ctrl ⏎ 可生成审阅修订');
});

// 审阅面板：只读取并对比 A（已载入的原文）与 A_脱敏，差异作为修订写入 A；不看黑名单，也可当 diff 用
const reviewToA = () => guarded(async () => {
  if ((await doc.cur.h.requestPermission?.({ mode: 'readwrite' })) !== 'granted') throw new Error('没有写入原文件的权限');
  let hB = null;
  if (doc.dir) hB = await doc.dir.getFileHandle(outName()).catch(() => null);
  hB ||= doc.outHandle;
  if (!hB) { status(`未找到 ${outName()}：请先 Ctrl+S 生成，或在弹出的窗口里选择要对比的文件`); [hB] = await showOpenFilePicker({ types: DOCX_TYPE }); }
  const pb = await parseDocx(await (await hB.getFile()).arrayBuffer());
  const { blob, stats } = await exportDocx(doc.buffer, pb.md, { track: true, author: 'ss', newMap: pb.map });
  if (!stats.del && !stats.ins) return status(`${doc.cur.name} 与 ${hB.name} 的文字没有差异，未写入`);
  await write(doc.cur.h, blob);
  status(`已把 ${hB.name} 的差异写入 ${doc.cur.name}：${stats.del} 处删除 / ${stats.ins} 处插入` + (stats.skipped ? `，${stats.skipped} 处无法映射` : '') + ' · 请在 Word 审阅面板检查后保存');
});
$('apply').onclick = applyAll;
$('append').onclick = reviewToA;
addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  const k = e.key.toLowerCase();
  if (e.key === 'Enter') { e.preventDefault(); reviewToA(); }
  else if (k === 's') { e.preventDefault(); applyAll(); }
  else if (k === 'z' || k === 'y') {
    if (e.target.closest?.('#addBox, #custom')) return;
    e.preventDefault();
    (k === 'y' || e.shiftKey) ? redo() : undo();
  }
});

// 右上角 badge：选区总数；每次变化显示增量，增量不是 1（说明匹配到了别处或没匹配到）时暂时变黄
let lastN = null, flash = 0, flashing = false;
function updatePill() {
  const pill = $('pill'), n = sel.length;
  pill.title = '黑名单 + 替换规则命中的位置总数';
  if (lastN !== null && n !== lastN) {
    const d = n - lastN;
    pill.innerHTML = `自加载共 ${n} 次删除<sup>${d > 0 ? '+' : ''}${d}</sup>`;
    pill.classList.toggle('hot', d !== 1);
    flashing = true; clearTimeout(flash);
    flash = setTimeout(() => { flashing = false; pill.classList.remove('hot'); pill.textContent = `自加载共 ${lastN} 次删除`; }, 3000);
  } else if (!flashing) pill.textContent = `自加载共 ${n} 次删除`;
  lastN = n;
}

// ---------- 文件：右上 ▼ 选择 / 打开文件夹 ----------
const shortName = (name) => {
  const stem = name.replace(/\.docx$/i, '');
  let head = stem.split(/\s+/)[0];
  if (head.length > 10) head = head.slice(0, 10);
  return head === stem ? name : head + '….docx';
};
function setFileInfo(name, size) { $('fileInfo').textContent = name ? `${shortName(name)} (${mb(size)})` : '打开文件夹…'; }

async function openDoc(i) {
  const ent = doc.entries[i]; if (!ent) return;
  status('读取中…');
  const file = await ent.h.getFile();
  doc.buffer = await file.arrayBuffer(); doc.cur = ent;
  const p = await parseDocx(doc.buffer);
  doc.parsed = p; doc.orig = p.md;
  doc.images.forEach((x) => x.url && URL.revokeObjectURL(x.url));
  doc.images = await listImages(p.zip);
  resetKeep();
  ta.value = p.md; state.ner = []; state.on = {}; memo.clear(); for (const k in visited) delete visited[k];
  doc.outHandle = null; lastN = null; undoS.length = redoS.length = 0; blackEl.value = [...words(splitBlack().head).filter((l) => !parseBlackLines([l]).ranges.length), ''].join('\n') + SEP; // 保留黑名单词，行号规则属于旧文档故丢弃
  $('title').textContent = ent.name.replace(/\.docx$/i, '');
  setFileInfo(ent.name, file.size);
  run(); status('已载入 · Ctrl S → _脱敏.docx · Ctrl ⏎ 审阅面板 → A');
  if (ner.ready()) runNer();
}
async function runNer() {
  if (!doc.parsed) return;
  const raw = doc.orig.split('\n');
  const lines = raw.map((l) => (/^(:::|---|\|[ -|]+$)/.test(l) ? '' : l.replace(/\|/g, ' ').replace(/\*+/g, '')));
  $('bannerText').textContent = '离线 NER 识别中…';
  const t0 = performance.now();
  try {
    state.ner = await ner.recognize(lines, (f) => status(`NER ${(f * 100) | 0}%`));
    $('bannerText').textContent = `已提取纯文本部分，普通表格按行对应，复杂表格1格1行。请确保处理敏感文字。（NER ${((performance.now() - t0) / 1000).toFixed(1)}s · ${ner.backend}）`;
    status('NER 完成');
  } catch (e) { console.error(e); $('bannerText').textContent = 'NER 失败：' + e.message; }
  run();
}

// 模型：未载入时右上角 ⚠️
const warn = $('modelWarn');
async function loadModelFile(h, name) {
  const file = h.getFile ? await h.getFile() : h;
  warn.hidden = false; warn.textContent = '⚠️ 离线模型载入中…';
  try {
    const be = await ner.loadModel(await file.arrayBuffer(), (s) => status(s));
    warn.hidden = true; $('banner').title = `${name} (${mb(file.size)}) · ${be}`;
    status(`模型就绪 · ${be}`);
    if (h.getFile) saveHandle('model', h);
    if (doc.parsed) runNer(); else run();
  } catch (e) { console.error(e); warn.textContent = '⚠️ 模型载入失败：' + e.message; }
}
warn.onclick = async () => {
  try {
    if (pendingModel) { if ((await pendingModel.requestPermission({ mode: 'read' })) === 'granted') return loadModelFile(pendingModel, pendingModel.name); }
    const [h] = await showOpenFilePicker({ types: [{ description: 'ONNX', accept: { 'application/octet-stream': ['.onnx'] } }] });
    await loadModelFile(h, h.name);
  } catch (e) { if (e.name !== 'AbortError') status(e.message); }
};

async function openFolder(dir) {
  doc.dir = dir; doc.entries = []; let model = null;
  for await (const [name, h] of dir.entries()) {
    if (h.kind !== 'file') continue;
    if (/\.docx$/i.test(name) && !name.startsWith('~$') && !/(?:_脱敏|-红化)\.docx$/.test(name)) doc.entries.push({ name, h });
    else if (/\.onnx$/i.test(name) && (!model || /quant|q8/i.test(name))) model = { name, h };
  }
  doc.entries.sort((a, b) => a.name.localeCompare(b.name, 'zh'));
  if (model && !ner.ready()) loadModelFile(model.h, model.name);
  if (doc.entries.length) await openDoc(0);
  else status('该文件夹没有 .docx');
}
async function pickFolder() {
  try { await handleItems([await showDirectoryPicker({ mode: 'readwrite' })]); } catch (e) { if (e.name !== 'AbortError') status(e.message); }
}
const menu = $('fileMenu');
function openMenu() {
  menu.innerHTML = doc.entries.map((e, i) => `<div data-i="${i}" class="${e === doc.cur ? 'cur' : ''}">${esc(e.name)}</div>`).join('') + '<div class="act" data-act="dir">📂 打开文件夹…</div>';
  menu.hidden = false;
}
$('fileSel').addEventListener('click', (e) => {
  const it = e.target.closest('#fileMenu > div');
  if (it) { menu.hidden = true; if (it.dataset.act) pickFolder(); else openDoc(+it.dataset.i); return; }
  if (!doc.entries.length) return pickFolder();
  menu.hidden ? openMenu() : (menu.hidden = true);
});
addEventListener('click', (e) => { if (!e.target.closest('#fileSel')) menu.hidden = true; });
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

// ---------- 记住上次模型文件句柄 ----------
function idb() { return new Promise((res, rej) => { const r = indexedDB.open('di', 1); r.onupgradeneeded = () => r.result.createObjectStore('h'); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); }
async function saveHandle(k, v) { try { const db = await idb(); db.transaction('h', 'readwrite').objectStore('h').put(v, k); } catch { /* 忽略 */ } }
async function lastHandle(k) { try { const db = await idb(); return await new Promise((res) => { const q = db.transaction('h').objectStore('h').get(k); q.onsuccess = () => res(q.result); q.onerror = () => res(null); }); } catch { return null; } }
(async () => {
  warn.hidden = false;
  const h = await lastHandle('model');
  if (h) {
    pendingModel = h;
    if ((await h.queryPermission?.({ mode: 'read' })) === 'granted') loadModelFile(h, h.name);
    else warn.textContent = `⚠️ 离线模型未载入（点击载入 ${h.name}）`;
  }
})();
if (!('showDirectoryPicker' in window)) status('需要基于 Chromium 的浏览器（File System Access API）');
window.__di = { openDoc, state, ta, run, ner, parseDocx, exportDocx, computeEdits, doc, boxAdd, edit, undo, redo, finalMd, applyAll, reviewToA, get items() { return items; }, get sel() { return sel; } };
run();
