// Markdown 投影的结构：docx.js 用 makeMarks 的 dump 部分写，detect.js / main.js 用 readLines 读。结构只在这里定义一次
const reEsc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

// 特殊表格（合并单元格等，1 格 1 行）夹在 open / close 两行之间；others 之后是页眉页脚 / 脚注 / 文本框 / 图表，每节以 section 开头
export function makeMarks(o = {}) {
  const c = {
    open: '<details>:::', openNote: '特殊表格（1格1行）', close: ':::</details>',
    others: '---\n\n## 其他文本', othersNote: '（页眉页脚 / 脚注 / 文本框 / 图表）', section: '### ',
    ...o,
  };
  return {
    ...c,
    // dump：docx.js 原样写进投影
    specialBegin: c.open + c.openNote + '\n',
    specialEnd: c.close + '\n\n',
    othersHead: c.others + c.othersNote + '\n\n',
    sectionHead: (title) => c.section + title + '\n\n',
    // 行首会被误认成标题 / 表格 / 分隔线 / 列表 / 特殊表格边界的正文，投影时前面加 \
    ambig: new RegExp('^(?:#|\\||---|- |:::|' + [c.open, c.close].map(reEsc).join('|') + ')'),
  };
}
export const MARK = makeMarks();

// docx.js 的默认选项：结构标记；save({track}) 的修订作者（载入时先拒绝掉这个作者以前写的修订）
export const PROJECTION = { marks: MARK, author: '脱敏助手' };

// load：逐行标出结构 [{text, s, n, kind, others}]。n 从 1 起；others 为真表示已在「其他文本」区。
// kind：open / close（特殊表格边界）、cell（特殊表格里的行）、table / tableSep（普通表格行 / 分隔行）、rule（---）、text。
// structural：不是正文的行（不送 NER、不当章节）
export function readLines(text, m = MARK) {
  const cut = text.indexOf('\n' + m.others), out = [];
  let s = 0, n = 0, special = false;
  for (const line of text.split('\n')) {
    let kind = 'text';
    if (line.startsWith(m.open)) { kind = 'open'; special = true; }
    else if (line === m.close) { kind = 'close'; special = false; }
    else if (special) kind = 'cell';
    else if (/^\|[ |-]+$/.test(line)) kind = 'tableSep';
    else if (line.startsWith('| ') && line.endsWith(' |')) kind = 'table';
    else if (line.startsWith('---')) kind = 'rule';
    out.push({ text: line, s, n: ++n, kind, structural: kind === 'open' || kind === 'close' || kind === 'tableSep' || kind === 'rule', others: cut >= 0 && s > cut });
    s += line.length + 1;
  }
  return out;
}

// 特殊表格的开头行与收尾行互为另一头：[s, e] 落在其中一行内时返回另一头的行首位置，否则 -1
export function partnerOf(text, s, e = s, m = MARK) {
  const lines = readLines(text, m), i = lines.findLastIndex((l) => l.s <= s), l = lines[i];
  if (!l || e > l.s + l.text.length) return -1;
  if (l.kind === 'open') return lines.slice(i + 1).find((x) => x.kind === 'close')?.s ?? -1;
  if (l.kind === 'close') return lines.slice(0, i).findLast((x) => x.kind === 'open')?.s ?? -1;
  return -1;
}
