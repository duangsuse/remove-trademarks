// Word <-> Markdown 1:1 投影；差异以修订（w:del / w:ins）写回
import JSZip from 'jszip';
import { diffLines, diffChars } from 'diff';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS = 'http://www.w3.org/XML/1998/namespace';
const isW = (n, name) => n.nodeType === 1 && n.namespaceURI === NS && n.localName === name;
const kids = (n, name) => [...n.childNodes].filter((c) => isW(c, name));
const kid = (n, name) => [...n.childNodes].find((c) => isW(c, name));
const on = (el) => !!el && !['0', 'false', 'off'].includes(el.getAttributeNS(NS, 'val') ?? '1');
const inFallback = (el) => { for (let n = el.parentNode; n && n.nodeType === 1; n = n.parentNode) if (n.localName === 'Fallback') return true; return false; };

// ---------- 段落原子 ----------
function runFmt(r) {
  const p = kid(r, 'rPr');
  return p ? (on(kid(p, 'b')) ? 'b' : '') + (on(kid(p, 'i')) ? 'i' : '') : '';
}

function collectAtoms(p) {
  const atoms = [];
  const addRun = (r, locked) => {
    const fmt = runFmt(r);
    for (const c of r.childNodes) {
      if (c.nodeType !== 1 || c.namespaceURI !== NS) continue;
      const base = { r, item: c, fmt, locked };
      if (c.localName === 't') {
        let i = 0;
        for (const ch of c.textContent) { atoms.push({ ...base, ch, edit: !locked, s: i, e: i + ch.length }); i += ch.length; }
      } else if (c.localName === 'tab') atoms.push({ ...base, ch: '\t', edit: false });
      else if (c.localName === 'br' || c.localName === 'cr') atoms.push({ ...base, ch: '\n', edit: false });
      else if (c.localName === 'noBreakHyphen') atoms.push({ ...base, ch: '-', edit: false });
    }
  };
  const walk = (n, locked) => {
    for (const c of n.childNodes) {
      if (c.nodeType !== 1 || c.namespaceURI !== NS) continue;
      const ln = c.localName;
      if (ln === 'r') addRun(c, locked);
      else if (ln === 'ins' || ln === 'moveTo') walk(c, true);
      else if (['hyperlink', 'smartTag', 'sdt', 'sdtContent', 'fldSimple', 'dir', 'bdo'].includes(ln)) walk(c, locked);
    }
  };
  walk(p, false);
  return atoms;
}

function headingLevel(p) {
  const pPr = kid(p, 'pPr');
  if (!pPr) return 0;
  const st = kid(pPr, 'pStyle')?.getAttributeNS(NS, 'val') || '';
  const m = /^(?:heading|标题)\s*([1-6])$/i.exec(st);
  if (m) return +m[1];
  const ol = kid(pPr, 'outlineLvl')?.getAttributeNS(NS, 'val');
  if (ol !== undefined && ol !== null && /^[0-5]$/.test(ol)) return +ol + 1;
  return 0;
}

// ---------- Markdown 构建 ----------
class MD {
  constructor() { this.out = []; this.pid = []; this.off = []; this.ed = []; }
  push(str, pid, off, ed) { for (let i = 0; i < str.length; i++) { this.out.push(str[i]); this.pid.push(pid); this.off.push(off); this.ed.push(ed); } }
  s(str) { this.push(str, -1, 0, 0); }
  sp(str, para, off) { this.push(str, para.idx, off, 0); }
  e(str, para, off) { this.push(str, para.idx, off, 1); }
  get length() { return this.out.length; }
}
const marker = (f) => (f === 'bi' ? '***' : f === 'b' ? '**' : '*');

function emitPara(md, para, prefix = '') {
  md.s(prefix);
  let fmt = '';
  // 空白 / 制表 / 换行不单独成粗斜体段：沿用前一个（开头则取后一个）非空白的格式
  const eff = para.atoms.map((a) => (/^\s$/.test(a.ch) ? null : a.fmt));
  let carry = eff.find((f) => f !== null) ?? '';
  for (let i = 0; i < eff.length; i++) { if (eff[i] === null) eff[i] = carry; else carry = eff[i]; }
  para.atoms.forEach((a, i) => {
    if (eff[i] !== fmt) {
      if (fmt) md.sp(marker(fmt), para, i);
      if (eff[i]) md.sp(marker(eff[i]), para, i);
      fmt = eff[i];
    }
    if (a.edit) {
      if (a.ch === '*' || a.ch === '\\' || a.ch === '|') md.sp('\\', para, i);
      md.e(a.ch, para, i);
    } else md.sp(a.ch === '\n' ? '<br>' : a.ch, para, i);
  });
  if (fmt) md.sp(marker(fmt), para, para.atoms.length);
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
      if (pr && (kid(pr, 'vMerge') || kid(pr, 'hMerge') || (kid(pr, 'gridSpan') && +kid(pr, 'gridSpan').getAttributeNS(NS, 'val') > 1))) return null;
      const ps = kids(tc, 'p');
      const other = [...tc.childNodes].some((x) => x.nodeType === 1 && x.namespaceURI === NS && !['tcPr', 'p'].includes(x.localName));
      if (other || ps.length !== 1) return null;
      row.push(ps[0]);
    }
    grid.push(row);
  }
  return grid;
}

const TOOL_AUTHOR = 'ss';
// 本工具（作者 ss）上一次写入的修订先「拒绝」掉再投影：删除线里的原文还原成正常文字，插入的内容去掉。
// 否则 Ctrl+Enter 写入审阅修订后，重新载入 A 时原文已经藏在 w:del 里，投影看不到旧内容、黑名单就扫不出来了；
// 同时重复写入审阅不会叠加。别的作者的修订保持原样。返回处理的数量。
function rejectOurs(doc) {
  let n = 0;
  const byTag = (t) => [...doc.getElementsByTagNameNS(NS, t)].filter((x) => x.getAttributeNS(NS, 'author') === TOOL_AUTHOR);
  for (const x of byTag('ins')) { if (x.parentNode && x.parentNode.localName !== 'rPr') { x.remove(); n++; } }
  for (const x of byTag('del')) {
    if (!x.parentNode || x.parentNode.localName === 'rPr') continue;
    for (const dt of [...x.getElementsByTagNameNS(NS, 'delText')]) {
      const t = doc.createElementNS(NS, 'w:t');
      t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      t.textContent = dt.textContent;
      dt.replaceWith(t);
    }
    while (x.firstChild) x.parentNode.insertBefore(x.firstChild, x);
    x.remove(); n++;
  }
  return n;
}

export async function parseDocx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const names = Object.keys(zip.files).filter((n) => /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(n));
  const parts = [];
  for (const path of names) {
    const xml = await zip.file(path).async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) continue;
    parts.push({ path, doc, dirty: rejectOurs(doc) > 0 });
  }
  const paras = [];
  const md = new MD();
  const queue = []; // 其他文本：{title, container}
  const mkPara = (p, part) => { const para = { idx: paras.length, el: p, part, atoms: collectAtoms(p) }; paras.push(para); return para; };
  const hasText = (p) => collectAtoms(p).length > 0;

  const scanTextboxes = (p, part) => {
    for (const tb of p.getElementsByTagNameNS(NS, 'txbxContent')) {
      if (inFallback(tb)) continue;
      let outer = true;
      for (let n = tb.parentNode; n && n.nodeType === 1; n = n.parentNode) if (n.localName === 'txbxContent') { outer = false; break; }
      if (outer) queue.push({ title: '文本框', container: tb, part });
    }
  };

  const walkBlocks = (container, part) => {
    for (const c of container.childNodes) {
      if (c.nodeType !== 1 || c.namespaceURI !== NS) continue;
      if (c.localName === 'p') {
        scanTextboxes(c, part);
        if (!hasText(c)) continue;
        const h = headingLevel(c);
        emitPara(md, mkPara(c, part), h ? '#'.repeat(h) + ' ' : '');
        md.s('\n\n');
      } else if (c.localName === 'sdt') {
        const sc = kid(c, 'sdtContent');
        if (sc) walkBlocks(sc, part);
      } else if (c.localName === 'tbl') {
        const grid = plainTable(c);
        if (grid) {
          grid.forEach((row, ri) => {
            md.s('| ');
            row.forEach((p, ci) => {
              scanTextboxes(p, part);
              if (ci) md.s(' | ');
              if (hasText(p)) emitPara(md, mkPara(p, part));
            });
            md.s(' |\n');
            if (ri === 0) md.s('|' + ' --- |'.repeat(row.length) + '\n');
          });
          md.s('\n');
        } else {
          const start = md.length;
          md.s(':::特殊表格（1格1行）\n');
          const inner = [];
          const cellsWalk = (tbl) => {
            for (const tr of kids(tbl, 'tr')) for (const tc of tr.childNodes) {
              if (isW(tc, 'tc')) inner.push(tc);
              else if (isW(tc, 'sdt')) { const sc = kid(tc, 'sdtContent'); if (sc) for (const x of kids(sc, 'tc')) inner.push(x); }
            }
          };
          cellsWalk(c);
          const before = md.length;
          for (const tc of inner) walkCell(tc, part);
          if (md.length === before) { md.out.length = start; md.pid.length = start; md.off.length = start; md.ed.length = start; }
          else md.s(':::\n\n');
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
        if (hasText(c)) { emitPara(md, mkPara(c, part)); md.s('\n'); }
      } else if (c.localName === 'tbl') {
        for (const tr of kids(c, 'tr')) for (const t2 of kids(tr, 'tc')) walkCell(t2, part);
      } else if (c.localName === 'sdt') { const sc = kid(c, 'sdtContent'); if (sc) walkCell(sc, part); }
    }
  };

  const main = parts.find((p) => p.path === 'word/document.xml');
  if (main) { const body = main.doc.getElementsByTagNameNS(NS, 'body')[0]; if (body) walkBlocks(body, main); }

  // 其他文本放末尾
  const others = [];
  for (const part of parts) {
    if (part === main) continue;
    const nm = part.path.replace('word/', '').replace('.xml', '');
    const title = /header/.test(nm) ? '页眉 ' + nm : /footer/.test(nm) ? '页脚 ' + nm : nm === 'footnotes' ? '脚注' : '尾注';
    if (nm === 'footnotes' || nm === 'endnotes') {
      for (const fn of part.doc.documentElement.childNodes) {
        if (fn.nodeType !== 1 || ['separator', 'continuationSeparator', 'continuationNotice'].includes(fn.getAttributeNS(NS, 'type'))) continue;
        others.push({ title, container: fn, part });
      }
    } else others.push({ title, container: part.doc.documentElement, part });
  }
  const endStart = md.length;
  const sections = [...others];
  // 文本框来自正文，先排队后处理（含其内再嵌套的）
  const flushQ = () => { while (queue.length) sections.push(queue.shift()); };
  let first = true;
  const emitSection = (sec) => {
    const before = md.length;
    const mark = paras.length;
    const hdr = '### ' + sec.title + '\n\n';
    md.s(hdr);
    const h = md.length;
    walkBlocks(sec.container, sec.part);
    if (md.length === h) { md.out.length = before; md.pid.length = before; md.off.length = before; md.ed.length = before; return false; }
    return true;
  };
  flushQ();
  const emitted = [];
  for (let i = 0; i < sections.length; i++) {
    const sec = sections[i];
    const start = md.length;
    if (first) { md.s('---\n\n## 其他文本（页眉页脚 / 脚注 / 文本框）\n\n'); }
    const ok = emitSection(sec);
    if (!ok && first) { md.out.length = start; md.pid.length = start; md.off.length = start; md.ed.length = start; }
    else if (ok) first = false;
    flushQ();
  }

  return {
    zip, parts, paras,
    md: md.out.join(''),
    map: { pid: Int32Array.from(md.pid), off: Int32Array.from(md.off), ed: Uint8Array.from(md.ed) },
  };
}

// ---------- diff -> 每段编辑 ----------
// newMap：新文本也是某个 docx 的投影时传入——插入文字只取其中「可编辑字符」，转义反斜杠 / 粗斜体标记 / 表格竖线 / 换行等结构字符不会混进来
export function computeEdits(origMd, newMd, map, newMap = null) {
  const byPara = new Map();
  const stats = { del: 0, ins: 0, skipped: 0 };
  const ent = (pid) => { let e = byPara.get(pid); if (!e) byPara.set(pid, (e = { del: new Set(), ins: new Map() })); return e; };
  const del = (a, b) => {
    let any = false;
    for (let i = a; i < b; i++) if (map.pid[i] >= 0 && map.ed[i]) { ent(map.pid[i]).del.add(map.off[i]); any = true; }
    if (any) stats.del++; else stats.skipped++;
  };
  const ins = (p, text, np) => {
    if (newMap) { let t = ''; for (let i = 0; i < text.length; i++) if (newMap.ed[np + i]) t += text[i]; text = t; }
    else text = text.replace(/\r\n?/g, '\n');
    if (!text) return;
    let pid = -1, at = 0;
    if (p > 0 && map.pid[p - 1] >= 0) { pid = map.pid[p - 1]; at = map.off[p - 1] + map.ed[p - 1]; }
    else if (p < map.pid.length && map.pid[p] >= 0) { pid = map.pid[p]; at = map.off[p]; }
    if (pid < 0) { stats.skipped++; return; }
    const e = ent(pid);
    e.ins.set(at, (e.ins.get(at) || '') + text);
    stats.ins++;
  };
  if (origMd === newMd) return { byPara, stats };
  const walkChars = (a, b, base, nbase) => {
    let pos = base, np = nbase;
    if (!b) { del(pos, pos + a.length); return; }
    if (!a) { ins(pos, b, np); return; }
    if (a.length * b.length > 4e6) { del(pos, pos + a.length); ins(pos + a.length, b, np); return; }
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
  return { byPara, stats };
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

function applyParaEdits(para, ed, ctx) {
  const { atoms } = para;
  const doc = para.el.ownerDocument;
  const need = new Set();
  const lastIdx = atoms.length - 1;
  for (const i of ed.del) if (atoms[i]?.edit) need.add(atoms[i].r);
  for (const at of ed.ins.keys()) need.add(atoms[Math.min(at, lastIdx)].r);
  const rprSrc = (gi) => { const a = atoms[Math.max(0, Math.min(gi - 1, lastIdx))]; return kid(a.r, 'rPr'); };
  const rprClones = new Map();
  for (const a of atoms) if (!rprClones.has(a.r)) rprClones.set(a.r, kid(a.r, 'rPr')?.cloneNode(true) || null);
  const rprOfAtom = (gi) => rprClones.get(atoms[Math.max(0, Math.min(gi - 1, lastIdx))].r);
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
  for (const r of need) {
    const parent = r.parentNode;
    if (!parent || isW(parent, 'ins')) continue;
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

export async function exportDocx(buffer, newMd, { author = 'ss', dropImages = [], wipe = false, track = true, newMap = null } = {}) {
  const m = await parseDocx(buffer);
  const { byPara, stats } = computeEdits(m.md, newMd, m.map, newMap);
  const ctx = { author, date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), nextId: 900000, track };
  for (const [pid, ed] of byPara) applyParaEdits(m.paras[pid], ed, ctx);
  if (wipe) m.parts.forEach((part) => cleanDom(part.doc));
  for (const part of m.parts) {
    const touched = wipe || part.dirty || [...byPara.keys()].some((pid) => m.paras[pid].part === part);
    if (!touched) continue;
    let xml = new XMLSerializer().serializeToString(part.doc);
    if (wipe) xml = xml.replace(FP_ATTR, '');
    m.zip.file(part.path, /^<\?xml/.test(xml) ? xml : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml);
  }
  for (const path of dropImages) {
    const ext = path.split('.').pop().toLowerCase();
    const key = ext === 'jpeg' ? 'jpg' : ext;
    if (PLACEHOLDER[key] && m.zip.file(path)) m.zip.file(path, b64(PLACEHOLDER[key]));
  }
  if (wipe) await wipeFingerprints(m.zip);
  const blob = await m.zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  return { blob, stats };
}
