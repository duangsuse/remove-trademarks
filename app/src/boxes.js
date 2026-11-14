// 左栏输入框的文本格式：读（load / parse）与写（dump）放在一起，调用方不再自己拼分隔符

// 行号规则 #L11~L22 / #L11（黑名单框、章节右键、含姓名表格）
export const lineRule = {
  parse: (t) => { const m = /^#L(\d+)(?:\s*~\s*L?(\d+))?$/i.exec(t.trim()); return m ? [+m[1], +(m[2] || m[1])] : null; },
  dump: (a, b = a) => (a === b ? `#L${a}` : `#L${a}~L${b}`),
};

// 黑名单框：=== 以上是手填的条目（词或行号规则），以下随右侧面板同步
const SEP = '===';
const entries = (ls) => ls.map((l) => l.trim()).filter(Boolean);
const sameEntry = (a, b) => { if (a === b) return true; const x = lineRule.parse(a), y = lineRule.parse(b); return !!x && !!y && x[0] === y[0] && x[1] === y[1]; };
export const BlackBox = {
  // own：手填区原样的行（含空行，便于保留光标）；entries：手填条目；synced：同步区条目；words / ranges：条目拆成词与行号区间
  load(text) {
    const lines = text.split('\n'), i = lines.findIndex((l) => l.trim() === SEP);
    const own = i < 0 ? lines : lines.slice(0, i), es = entries(own), words = [], ranges = [];
    for (const t of es) { const r = lineRule.parse(t); if (r) ranges.push(r); else words.push(t); }
    return { own, entries: es, synced: i < 0 ? [] : entries(lines.slice(i + 1)), words, ranges };
  },
  dump: ({ own, synced = [] }) => [...own, SEP, ...synced].join('\n'),
  // 加 / 删一个手填条目（行号规则按区间比较）；同步区清空，由调用方重新同步
  add(text, entry) { const { entries: es } = BlackBox.load(text); return es.some((x) => sameEntry(x, entry)) ? text : BlackBox.dump({ own: [...es, entry] }); },
  remove: (text, entry) => BlackBox.dump({ own: BlackBox.load(text).entries.filter((x) => !sameEntry(x, entry)) }),
};

// 替换规则框：每行「[::类别::] 正则 => 替换」，类别前缀只作用于该类别（cat / en / NER 类型）
// =>前面没有正则（「::原文:: => 替换」）：整段换成「替换」，前缀可以是条目原文——右栏 ✎ 重命名写的就是这种行
export function parseRules(text) {
  const rules = [];
  for (const ln of text.split('\n')) {
    const i = ln.indexOf('=>');
    if (i < 0) continue;
    let re = ln.slice(0, i).trim();
    const rep = ln.slice(i + 2).trim(), pre = /^::(.+?)::\s*/.exec(re);
    if (pre) re = re.slice(pre[0].length);
    if (!re) { if (pre) rules.push({ whole: true, only: pre[1].trim(), rep }); continue; }
    try { rules.push({ test: new RegExp('^(?:' + re + ')$'), re: new RegExp(re, 'g'), rep, only: pre ? pre[1].trim() : '' }); } catch { /* 忽略无效规则 */ }
  }
  return rules;
}

// 自订规则框：每行「名称 :: 正则」，省略名称为「自订」
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
