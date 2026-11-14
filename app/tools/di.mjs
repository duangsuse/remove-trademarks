#!/usr/bin/env node
// 命令行试用投影 API（Node 里用 jsdom 提供 DOMParser）。用法见 README.md 或 node tools/di.mjs -h
import { readFileSync, writeFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import { JSDOM } from 'jsdom';

const { window } = new JSDOM('');
globalThis.DOMParser = window.DOMParser;
globalThis.XMLSerializer = window.XMLSerializer;
const { DocxProjection } = await import('../src/docx.js');
const { scan, selection, replacement } = await import('../src/detect.js');
const { lineRule, parseRules } = await import('../src/boxes.js');
const { lineRange, plainOf } = await import('../src/plain.js');
const { DEFAULT_REPLACE_RULES } = await import('../src/rules.js');

const HELP = `node tools/di.mjs <命令> A.docx [参数]
  md     A.docx                         打印 Markdown 投影（重定向到 .md 文件再编辑）
  txt    A.docx                         打印纯文本投影
  apply  A.docx A.md                    把改过的 .md 写回，只动有差异的字
  edit   A.docx -e "s => s.replaceAll('甲', '乙')"   用 JS 函数改纯文本投影
  redact A.docx -w 张三 -w 李四 [-c company,contact] [-r ".*公司 => 某公司"]   黑名单 / 规则脱敏
通用：-L 11~22   只作用于这些行（可多次；edit 用第一个）
      -o 输出.docx（默认 A_新版.docx，redact 为 A_脱敏.docx）   --track 写成审阅修订`;

const { values: o, positionals: [cmd, src, mdFile] } = parseArgs({
  allowPositionals: true,
  options: {
    e: { type: 'string' }, w: { type: 'string', multiple: true, default: [] }, L: { type: 'string', multiple: true, default: [] },
    c: { type: 'string', default: '' }, r: { type: 'string', multiple: true, default: [] }, o: { type: 'string' }, track: { type: 'boolean' }, h: { type: 'boolean' },
  },
});
if (o.h || !cmd || !src) { console.log(HELP); process.exit(o.h ? 0 : 1); }

const d = await DocxProjection.load(readFileSync(src));
for (const w of d.warnings) console.error('⚠ ' + w);
const ranges = o.L.map((l) => lineRule.parse('#L' + l.replace(/^#?L?/i, ''))).filter(Boolean);

switch (cmd) {
  case 'md': process.stdout.write(d.md); process.exit(0);
  case 'txt': process.stdout.write(d.text); process.exit(0);
  case 'apply': d.diffMd(readFileSync(mdFile, 'utf8')); break;
  case 'edit': {
    const fn = (0, eval)(o.e); // 自己机器上跑自己写的函数
    d.mapPlain(fn, { range: ranges[0] && lineRange(d.md, ...ranges[0]) });
    break;
  }
  case 'redact': {
    const cats = new Set(['black', ...o.c.split(',').filter(Boolean)]);
    const items = scan(d.md, { blacklist: o.w, ranges });
    const rules = parseRules(o.r.length ? o.r.join('\n') : DEFAULT_REPLACE_RULES);
    const sel = selection([...items.values()], (it) => cats.has(it.en));
    const nums = new Map(), nth = (kind, s) => { const m = nums.get(kind) ?? nums.set(kind, new Map()).get(kind); if (!m.has(s)) m.set(s, m.size + 1); return m.get(s); }; // 同类按首次出现编号
    d.splice(sel.map((x) => {
      const s = plainOf(d.md, x.s, x.e);
      return { s: x.s, e: x.e, text: x.it.del ? '' : replacement(s, rules, { names: [x.it.cat, x.it.en, x.it.text], kind: x.it.cat, nth }) };
    }));
    break;
  }
  default: console.log(HELP); process.exit(1);
}
const out = o.o || src.replace(/\.docx$/i, '') + (cmd === 'redact' ? '_脱敏.docx' : '_新版.docx');
const { data, report: stats, warnings } = await d.save({ type: 'uint8array', revisions: o.track });
writeFileSync(out, data);
for (const w of warnings.filter((w) => !d.warnings.includes(w))) console.error('⚠ ' + w);
console.error(`→ ${out}：${stats.del} 删 / ${stats.ins} 插` + (stats.skipped ? `，${stats.skipped} 处无法映射（结构字符）` : ''));
