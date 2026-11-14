// node --test test/  —— 用 jsdom 提供 DOMParser / XMLSerializer
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { JSDOM } from 'jsdom';
import JSZip from 'jszip';

const { window } = new JSDOM('');
globalThis.DOMParser = window.DOMParser;
globalThis.XMLSerializer = window.XMLSerializer;
const { DocxProjection } = await import('../src/docx.js');
const { scan, detect, selection, replacement, chapters, nerLines } = await import('../src/detect.js');
const { parseRules, BlackBox } = await import('../src/boxes.js');
const { plainOf, spliceMd, lineRange } = await import('../src/plain.js');
const { MARK, PROJECTION, makeMarks, partnerOf } = await import('../src/markup.js');
const TOOL_AUTHOR = PROJECTION.author;

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const DRAW = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const t = (s, b) => `<w:r>${b ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${s}</w:t></w:r>`;
const p = (body, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${body}</w:p>`;
const txbx = (s) => `<w:txbxContent>${p(t(s))}</w:txbxContent>`;

async function sample(extra = '') {
  const z = new JSZip();
  z.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  z.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}officeDocument" Target="word/document.xml"/></Relationships>`);
  z.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rS" Type="${REL}styles" Target="styles.xml"/><Relationship Id="rN" Type="${REL}numbering" Target="numbering.xml"/>
    <Relationship Id="rL" Type="${REL}hyperlink" Target="mailto:zhangsan@x.com" TargetMode="External"/>
    <Relationship Id="rI" Type="${REL}image" Target="media/image1.png"/><Relationship Id="rC" Type="${REL}chart" Target="charts/chart1.xml"/>
    <Relationship Id="rH" Type="${REL}header" Target="header1.xml"/></Relationships>`);
  z.file('word/styles.xml', `<w:styles ${W}><w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/></w:style><w:style w:type="paragraph" w:styleId="3"><w:name w:val="heading 2"/><w:basedOn w:val="2"/></w:style></w:styles>`);
  z.file('word/numbering.xml', `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1、"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`);
  z.file('word/header1.xml', `<w:hdr ${W}>${p(t('页眉 张三'))}</w:hdr>`);
  z.file('word/charts/chart1.xml', `<c:chartSpace ${DRAW}><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>张三销售额</a:t></a:r></a:p></c:rich></c:tx></c:title><c:ser><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>张三公司</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>42</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:chart></c:chartSpace>`);
  z.file('word/media/image1.png', 'x');
  const num = (id, lvl) => `<w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${id}"/></w:numPr>`;
  const body = [
    p(t('项目概况'), '<w:pStyle w:val="2"/>' + num(1, 0)),
    p(t('背景'), '<w:pStyle w:val="3"/>' + num(1, 1)),
    p(t('联系人：张') + t('三', true) + t('，电话 13912345678')),
    p(t('第一项'), num(2, 0)),
    p(t('# 不是标题')),
    p(`<w:hyperlink r:id="rL">${t('写信给张三')}</w:hyperlink>`),
    p(`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK "https://zhangsan.example.com/" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${t('主页')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`),
    p(`<w:r><w:drawing><wp:inline><wp:docPr id="1" name="图片 1" descr="张三身份证"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="x.png" descr="张三身份证"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rI"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`),
    p(`<w:r><w:drawing><wp:inline><wp:docPr id="2" name="图表 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rC"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`),
    p(`<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:docPr id="3" name="文本框 1"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx>${txbx('框里的张三')}</wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox>${txbx('框里的张三')}</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`),
    `<w:tbl><w:tr><w:tc>${p(t('姓名'))}</w:tc><w:tc>${p(t('张三'))}</w:tc></w:tr><w:tr><w:tc>${p(t('a|b'))}</w:tc><w:tc>${p(t('x*y'))}</w:tc></w:tr></w:tbl>`,
  ].join('') + extra;
  z.file('word/document.xml', `<?xml version="1.0"?><w:document ${W} ${DRAW}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  return z.generateAsync({ type: 'uint8array' });
}

async function allText(data) {
  const z = await JSZip.loadAsync(data);
  let s = '';
  for (const f of Object.values(z.files)) if (!f.dir && /\.(xml|rels)$/.test(f.name)) s += await f.async('string');
  return s;
}
const st = (words) => ({ blacklist: words });
// 解析 → 修改 → 保存 → 重新解析
const roundTrip = async (d, opts) => DocxProjection.load((await d.save({ type: 'uint8array', ...opts })).data);

test('投影：标题样式、编号、链接、图片、图表、文本框、表格、转义', async () => {
  const d = await DocxProjection.load(await sample());
  const { md } = d;
  assert.match(md, /^# 一、 项目概况$/m, '中文样式 ID "2" 的标题 + 中文编号');
  assert.match(md, /^## 一\.1 背景$/m, 'basedOn 继承 + 多级编号（未设 isLgl 时 Word 也显示「一.1」）');
  assert.match(md, /^- 第一项$/m);
  assert.match(md, /^\\# 不是标题$/m, '行首 # 转义');
  assert.match(md, /\[写信给张三\]\(mailto:zhangsan@x\.com\)/);
  assert.match(md, /\[主页\]\(https:\/\/zhangsan\.example\.com\/\)/, '域代码 HYPERLINK');
  assert.match(md, /!\[张三身份证\]\(media\/image1\.png\)/);
  assert.match(md, /!\[图表\]\(charts\/chart1\.xml\)/);
  assert.match(md, /### 图表 chart1\n\n张三销售额\n张三公司\n/);
  assert.doesNotMatch(md, /\n42\n/, '数值不进文本池');
  assert.match(md, /### 文本框\n\n框里的张三/);
  assert.match(md, /### 页眉 header1\n\n页眉 张三/);
  assert.match(md, /\| a\\\|b \| x\\\*y \|/);
  assert.deepEqual(d.warnings, []);
  const i = md.indexOf('mailto:') + 7;
  assert.equal(d.locate(i).unit.kind, 'text');
  assert.equal(d.locate(md.indexOf('](mailto')).editable, false);
});

test('扫描在纯文本视图上：跨粗体也能命中，区间换回 md 坐标', async () => {
  const d = await DocxProjection.load(await sample());
  const items = scan(d.md, st(['张三']));
  const occ = items.get('bl:张三').occ;
  assert.ok(occ.some(([s, e]) => d.md.slice(s, e) === '张**三'));
  assert.equal(occ.length, 8);
  for (const [s, e] of occ) assert.equal(plainOf(d.md, s, e).replace(/\*+$/, ''), '张三');
});

test('区间替换：所有位置（含链接地址、图片说明、图表、VML 副本、页眉）都不再有原文', async () => {
  const buf = await sample();
  const d = await DocxProjection.load(buf);
  const items = scan(d.md, st(['张三', 'zhangsan']));
  const sel = selection([...items.values()], () => true);
  const stats = d.splice(sel.map((x) => ({ ...x, text: 'XX' })));
  assert.equal(stats.skipped, 0);
  const { data, warnings } = await d.save({ type: 'uint8array' });
  assert.deepEqual(warnings, []);
  const xml = await allText(data);
  assert.ok(!xml.includes('张三') && !xml.includes('zhangsan'), xml.match(/.{40}(张三|zhangsan).{40}/)?.[0]);
  const d2 = await DocxProjection.load(data);
  assert.match(d2.md, /联系人：XX，电话 XX/);
  assert.match(d2.md, /\[写信给XX\]\(mailto:XX\)/);
  assert.match(d2.md, /!\[XX身份证\]/);
});

test('替换文字沿用被删字的格式', async () => {
  const d = await DocxProjection.load(await sample());
  const s = d.md.indexOf('**三**') + 2;
  d.splice([{ s, e: s + 1, text: '某' }]);
  const d2 = await DocxProjection.load((await d.save({ type: 'uint8array' })).data);
  assert.match(d2.md, /联系人：张\*\*某\*\*，/);
});

test('审阅修订：写入后重新载入与原文一致，重复写入不叠加', async () => {
  const buf = await sample();
  const a0 = await DocxProjection.load(buf);
  const b = await DocxProjection.load(buf);
  b.splice(selection([...scan(b.md, st(['张三'])).values()], () => true).map((x) => ({ ...x, text: '某' })));
  const bData = (await b.save({ type: 'uint8array' })).data;
  const bp = await DocxProjection.load(bData);
  const a = await DocxProjection.load(buf);
  const stats = a.diffMd(bp.md, { from: bp });
  assert.ok(stats.del && stats.ins);
  const { data, warnings } = await a.save({ revisions: true, type: 'uint8array' });
  assert.ok(warnings.some((w) => /无法写成修订/.test(w)), '链接 / 图表文字不写成修订');
  const xml = await allText(data);
  assert.ok(xml.includes(`w:author="${TOOL_AUTHOR}"`));
  const again = await DocxProjection.load(data);
  assert.equal(again.md, a0.md, '载入时撤掉本工具的修订，投影回到原文');
  const a2 = await DocxProjection.load(data);
  a2.diffMd(bp.md, { from: bp });
  const data2 = (await a2.save({ revisions: true, type: 'uint8array' })).data;
  assert.equal((await allText(data2)).split(`w:author="${TOOL_AUTHOR}"`).length, xml.split(`w:author="${TOOL_AUTHOR}"`).length);
});

test('含姓名表格：姓名中间有空白时整段命中', () => {
  const md = '| 姓名 | 张 三 |\n| 联系人： | 李　四： |';
  const items = scan(md, st([]));
  assert.deepEqual(items.get('tbl:张三').occ.map(([s, e]) => md.slice(s, e)), ['张 三']);
  assert.deepEqual(items.get('tbl:李四').occ.map(([s, e]) => md.slice(s, e)), ['李　四']);
});

test('nerLines：只清空表格分隔行，纯 ASCII 的表格行照常送 NER', () => {
  assert.deepEqual(nerLines(`| --- | --- |\n| John Smith | IBM Corp |\n${MARK.open}x\n${MARK.close}\n---`), ['', '  John Smith   IBM Corp  ', '', '', '']);
});

test('spliceMd：跨粗体的命中替换后标记仍成对；以转义字符开头的命中带上反斜杠', () => {
  const md = '联系人：张**三**，记号 x\\|y。';
  const items = scan(md, st(['张三', '|y']));
  const sel = selection([...items.values()], () => true);
  assert.deepEqual(sel.map((x) => md.slice(x.s, x.e)), ['张**三', '\\|y']);
  assert.equal(spliceMd(md, sel.map((x) => ({ ...x, text: 'XX' }))), '联系人：XX****，记号 xXX。');
});

test('区间替换：以转义字符开头的命中写入 docx 后不留反斜杠', async () => {
  const d = await DocxProjection.load(await sample());
  const sel = selection([scan(d.md, st(['|b'])).get('bl:|b')], () => true);
  assert.equal(sel.length, 1);
  assert.equal(d.splice(sel.map((x) => ({ ...x, text: 'Z' }))).skipped, 0);
  const d2 = await DocxProjection.load((await d.save({ type: 'uint8array' })).data);
  assert.match(d2.md, /\| aZ \| x\\\*y \|/);
});

test('selection：重叠命中合并，不漏半截', () => {
  const a = { occ: [[10, 12]] }, b = { occ: [[11, 15]] };
  const out = selection([a, b], () => true);
  assert.equal(out.length, 1);
  assert.deepEqual([out[0].s, out[0].e, out[0].it], [10, 15, b]);
});

test('替换模板：白名单求值，不执行任意代码', () => {
  const rules = parseRules('::cname:: .* => {s[0]}姓人员\n.* => {kind}{nth()}\nX => {globalThis.x=1}');
  assert.equal(replacement('张三', rules, { names: ['cname'] }), '张姓人员');
  let n = 0;
  assert.equal(replacement('南京路', rules, { kind: '地址', nth: () => ++n }), '地址1');
  assert.equal(replacement('X', parseRules('X => {globalThis.x=1}')), '{globalThis.x=1}');
  assert.equal(globalThis.x, undefined);
  assert.equal(replacement('价$1', parseRules('.* => [{s}]$0')), '[价$1]价$1', '命中文本里的 $1 不再被展开');
});

test('重命名行「::原文:: => 替换」：=> 前没有正则，只改原文恰好相符的条目', () => {
  const rules = parseRules('::张三:: => 某甲\n::cname:: .* => 某经理\n.* => **');
  assert.equal(replacement('张三', rules, { names: ['cname', 'cname', undefined, '张三'] }), '某甲');
  assert.equal(replacement('张三', rules, { names: ['cname', 'cname', undefined, '李四'] }), '某经理');
  assert.equal(replacement('张三', rules, { names: [] }), '**');
  assert.equal(replacement('张三', parseRules('::cname:: => 某经理'), { names: ['cname'] }), '某经理', '前缀是类别名也行');
  assert.equal(replacement('https://a.b/c', parseRules('::https://a.b/c:: => 网址{nth()}'), { names: ['https://a.b/c'], kind: '网址', nth: () => 2 }), '网址2', '原文可以含冒号');
  assert.equal(replacement('张三', parseRules('::张三:: => 乙\n::张三:: => 甲'), { names: ['张三'] }), '乙', '重复的取最上面一行');
  assert.equal(parseRules('=> x').length, 0, '既没有正则也没有前缀：忽略');
});

const REAL = new URL('../../3、C3443 阀门和旋塞制造-杰林石化-2026江苏省先进级智能工厂申报书.docx', import.meta.url);
test('样例申报书：标题按样式识别，替换后可再解析', { skip: !existsSync(REAL) }, async () => {
  const buf = readFileSync(REAL);
  const d = await DocxProjection.load(buf);
  const heads = chapters(d.md).filter((h) => /^#/.test(d.md.slice(h.s, h.s + 1)));
  assert.ok(heads.length >= 40, `识别到 ${heads.length} 个样式标题`);
  const items = scan(d.md, { ...st([]), cats: {} });
  const sel = selection([...items.values()], () => true);
  d.splice(sel.map((x) => ({ ...x, text: '**' })));
  const { data, report: stats, warnings } = await d.save({ type: 'uint8array' });
  console.log('样例：', sel.length, '处', stats, warnings);
  const d2 = await DocxProjection.load(data);
  assert.ok(d2.md.length > 1000);
});

// ---------- round-trip：只经过 load / replace / applyText / editText / save 这些边界 ----------
const MERGED = `<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${p(t('基本信息'))}</w:tc></w:tr><w:tr><w:tc>${p(t('联系人'))}</w:tc><w:tc>${p(t('王五'))}</w:tc></w:tr></w:tbl>`;

test('editText：用函数改纯文本投影，保存后重新解析，纯文本 = 原纯文本经同一函数', async () => {
  const d = await DocxProjection.load(await sample());
  const fn = (s) => s.replaceAll('张三', '李四');
  const want = fn(d.text);
  const st = d.mapPlain(fn);
  assert.equal(st.skipped, 0);
  const d2 = await roundTrip(d);
  assert.equal(d2.text, want);
  assert.match(d2.md, /联系人：李四，/, '跨粗体的「张**三**」整体换掉，不留落单的 **');
});

test('editText：纯插入、纯删除', async () => {
  const d = await DocxProjection.load(await sample());
  d.mapPlain((s) => s.replace('联系人：', '联系人（已核实）：').replace('，电话 13912345678', ''));
  const d2 = await roundTrip(d);
  assert.match(d2.md, /^联系人（已核实）：张\*\*三\*\*$/m);
});

test('editText range：只改选区（第 n 行），其余行不动', async () => {
  const d = await DocxProjection.load(await sample());
  const lines = d.md.split('\n'), n = lines.findIndex((l) => l.startsWith('联系人')) + 1;
  d.mapPlain((s) => s.replaceAll('张', '某'), { range: lineRange(d.md, n) });
  const d2 = await roundTrip(d);
  assert.match(d2.md, /^联系人：某\*\*三\*\*/m);
  assert.match(d2.md, /写信给张三/, '选区外不改');
});

test('applyText：改过的 md 写回（如外部编辑器改完的 .md 文件）', async () => {
  const d = await DocxProjection.load(await sample());
  const edited = d.md.replace('第一项', '第 1 项').replace('页眉 张三', '页眉');
  d.diffMd(edited);
  const d2 = await roundTrip(d);
  assert.equal(d2.md, edited);
});

test('审阅修订 round-trip：track 写入后重新载入先撤掉自己的修订，投影回到原文', async () => {
  const buf = await sample();
  const d = await DocxProjection.load(buf);
  d.mapPlain((s) => s.replace('第一项', '第二项'));
  const again = await roundTrip(d, { revisions: true });
  assert.equal(again.md, (await DocxProjection.load(buf)).md);
});

test('特殊表格：<details>::: 开头、:::</details> 收尾；表格里的姓名按行号命中；章节不收结构行', async () => {
  const d = await DocxProjection.load(await sample(MERGED));
  const lines = d.md.split('\n'), a = lines.findIndex((l) => l.startsWith(MARK.open));
  assert.ok(a >= 0, d.md);
  assert.deepEqual(lines.slice(a, a + 5), [MARK.open + MARK.openNote, '基本信息', '联系人', '王五', MARK.close]);
  const items = scan(d.md, st([]));
  assert.equal(items.get('tbl:#L' + (a + 4)).show, '王五');
  assert.ok(!chapters(d.md).some((c) => /details|:::/.test(c.title)));
  const open = d.md.indexOf(MARK.open), close = d.md.indexOf(MARK.close);
  assert.equal(partnerOf(d.md, open + 3), close, '开头行 → 收尾行');
  assert.equal(partnerOf(d.md, close), open, '收尾行 → 开头行');
  assert.equal(partnerOf(d.md, open, close), -1, '选区跨行不跳');
});

test('正文里以结构标记开头的段落会转义，不会被认成特殊表格', async () => {
  const d = await DocxProjection.load(await sample(p(t(MARK.open.replace('<', '&lt;') + '正文'))));
  assert.ok(d.md.includes('\\' + MARK.open + '正文'));
  assert.ok(!d.md.split('\n').some((l) => l.startsWith(MARK.open)));
});

test('结构标记 / 修订作者可经 opts 覆盖（投影工具与本项目解耦）', async () => {
  const marks = makeMarks({ open: '[[表', openNote: '', close: '表]]', others: '## 附', othersNote: '' });
  const d = await DocxProjection.load(await sample(MERGED), { marks, author: 'me' });
  assert.match(d.md, /^\[\[表\n基本信息\n联系人\n王五\n表\]\]$/m);
  assert.match(d.md, /^## 附\n\n### 页眉/m);
  d.mapPlain((s) => s.replace('王五', '某'));
  const xml = await allText((await d.save({ revisions: true, type: 'uint8array' })).data);
  assert.ok(xml.includes('w:author="me"') && !xml.includes(TOOL_AUTHOR));
});

test('界面文字可经 opts.i18n 按组覆盖；自定义结构标记时 lines 按它分', async () => {
  const { plainView } = await import('../src/plain.js');
  const marks = makeMarks({ open: '[[表', openNote: '', close: '表]]' });
  const d = await DocxProjection.load(await sample(MERGED), { marks, i18n: { parts: { header: 'Header', chart: 'Chart' } } });
  assert.match(d.md, /^### Header header1$/m);
  assert.match(d.md, /^### Chart chart1$/m);
  assert.match(d.md, /^### 文本框$/m, '没覆盖的条目沿用默认');
  assert.deepEqual(plainView(d.md, marks).lines.filter((l) => l.kind === 'open' || l.kind === 'close').map((l) => l.kind), ['open', 'close']);
});

test('载入时只撤掉同一作者以前写的修订', async () => {
  const rev = (author) => p(`<w:ins w:id="1" w:author="${author}">${t('新增')}</w:ins><w:del w:id="2" w:author="${author}"><w:r><w:delText>旧文</w:delText></w:r></w:del>`);
  const ours = await DocxProjection.load(await sample(rev(TOOL_AUTHOR)));
  assert.match(ours.md, /^旧文$/m);
  const theirs = await DocxProjection.load(await sample(rev('ss')));
  assert.match(theirs.md, /^新增$/m);
});

test('公式 / 注音 / altChunk：不投影，按标签名列进未处理清单', async () => {
  const M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
  const extra = p(`${t('面积')}<m:oMath ${M}><m:r><m:t>S=πr²</m:t></m:r></m:oMath>`)
    + p(`<w:r><w:ruby><w:rt><w:r><w:t>zhāng</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>张</w:t></w:r></w:rubyBase></w:ruby></w:r>`)
    + '<w:altChunk r:id="rX"/>' + p(`<w:del w:id="1" w:author="x"><w:r><w:delText>删掉的</w:delText></w:r></w:del>${t('留下')}`);
  const d = await DocxProjection.load(await sample(extra));
  assert.deepEqual(d.unhandled, [{ tag: 'm:oMath', count: 1 }, { tag: 'w:ruby', count: 1 }, { tag: 'w:altChunk', count: 1 }]);
  const w = d.warnings.find((x) => x.startsWith('以下元素没有投影'));
  assert.match(w, /<m:oMath> ×1（公式）\n.*<w:ruby> ×1（注音）\n.*<w:altChunk> ×1/);
  assert.match(d.md, /^面积$/m);
});

test('detect：NER 作为函数参数注入（这里用假的）', async () => {
  const md = '项目由王小明负责。\n王小明的电话 13912345678';
  const fakeNer = async (lines) => lines.flatMap((l, line) => (l.includes('王小明') ? [{ text: '王小明', type: 'PERSON', score: 0.9, line }] : []));
  const items = await detect(md, {}, { ner: fakeNer });
  assert.equal(items.get('ner:王小明').occ.length, 2);
  assert.ok(items.get('mobile:13912345678'));
  const only2 = await detect(md, {}, { ner: async (lines) => (assert.deepEqual(lines, ['', '王小明的电话 13912345678']), []), range: lineRange(md, 2) });
  assert.ok(!only2.has('ner:王小明'));
});

test('黑名单框：load / dump 互逆；add / remove 按行号区间去重', () => {
  const text = BlackBox.dump({ own: ['张三', '', '#L3~L5'], synced: ['李四'] });
  const b = BlackBox.load(text);
  assert.deepEqual([b.entries, b.synced, b.words, b.ranges], [['张三', '#L3~L5'], ['李四'], ['张三'], [[3, 5]]]);
  assert.equal(BlackBox.dump(b), text);
  assert.equal(BlackBox.add(text, '#l3 ~ L5'), text, '同一区间不重复加');
  assert.deepEqual(BlackBox.load(BlackBox.add(text, '王五')).entries, ['张三', '#L3~L5', '王五']);
  assert.deepEqual(BlackBox.load(BlackBox.remove(text, '#L3~5')).entries, ['张三']);
});

test('纯文本投影：<br> 是 1 个字符（行号不变）；转义掉的结构标记不当结构', async () => {
  const { plainView, BR } = await import('../src/plain.js');
  const md = '甲<br>乙\n\\| 姓名 | x |\n\\<details>:::a\n负责人 张三';
  const v = plainView(md);
  assert.equal(v.text.split('\n').length, md.split('\n').length);
  assert.equal(v.text.split('\n')[0], '甲' + BR + '乙');
  assert.deepEqual(v.toMd(1, 2), [1, 5]); // <br> 整个映射回 md
  assert.deepEqual(v.lines.map((l) => l.kind), ['text', 'text', 'text', 'text']);
});

const tabSample = () => sample(p(t('姓名') + '<w:r><w:tab/></w:r>' + t('张三')) + p(t('末尾')));
const imgTok = '![张三身份证](media/image1.png)';

test('制表符可编辑：删除、插入，审阅模式下是删除线 / 插入', async () => {
  const d = await DocxProjection.load(await tabSample());
  assert.ok(d.md.includes('姓名\t张三'.replace(' ', '')));
  const i = d.md.indexOf('姓名\t') + 2;
  assert.equal(d.locate(i).editable, true);
  d.splice([{ s: i, e: i + 1, text: '：' }, { s: d.md.indexOf('末尾'), e: d.md.indexOf('末尾'), text: '\t' }]);
  const d2 = await roundTrip(d);
  assert.ok(d2.md.includes('姓名：张三') && d2.md.includes('\t末尾'));
  const e = await DocxProjection.load(await tabSample());
  e.splice([{ s: i, e: i + 1, text: '' }]);
  const { data } = await e.save({ type: 'uint8array', revisions: true });
  const xml = await (await JSZip.loadAsync(data)).file('word/document.xml').async('string');
  assert.match(xml, /<w:del [^>]*><w:r><w:tab\/><\/w:r><\/w:del>/);
  assert.ok((await DocxProjection.load(data)).md.includes('姓名\t张三'), '重新载入先撤销自己的修订');
});

test('图片：整段删除、移动、复制，不能凭空新增', async () => {
  const buf = await tabSample();
  const media = (xml) => (xml.match(/r:embed="rI"/g) || []).length;
  // 删除
  let d = await DocxProjection.load(buf);
  d.diffMd(d.md.replace(imgTok + '\n', '\n'));
  let { data, warnings } = await d.save({ type: 'uint8array' });
  assert.deepEqual(warnings, []);
  assert.equal(media(await (await JSZip.loadAsync(data)).file('word/document.xml').async('string')), 0);
  assert.doesNotMatch((await DocxProjection.load(data)).md, /image1\.png/);
  // 移动（alt 跟着走）：从原处剪切，贴到「末尾」行
  d = await DocxProjection.load(buf);
  d.diffMd(d.md.replace(imgTok, '').replace('末尾', '末尾' + imgTok));
  ({ data, warnings } = await d.save({ type: 'uint8array' }));
  assert.deepEqual(warnings, []);
  const md2 = (await DocxProjection.load(data)).md;
  assert.equal(md2.split(imgTok).length, 2, '仍是一张');
  assert.ok(md2.includes('末尾' + imgTok));
  // 复制：原图不动，多出一张（同一媒体文件，docPr id 不重复）
  d = await DocxProjection.load(buf);
  d.diffMd(d.md.replace('末尾', '末尾' + imgTok));
  ({ data, warnings } = await d.save({ type: 'uint8array' }));
  assert.deepEqual(warnings, []);
  assert.equal((await DocxProjection.load(data)).md.split(imgTok).length, 3, '两张');
  const dx = await (await JSZip.loadAsync(data)).file('word/document.xml').async('string');
  assert.equal(media(dx), 2);
  const dids = [...dx.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1]);
  assert.equal(new Set(dids).size, dids.length, 'docPr id 不重复');
  // 一删一贴两份：第一份是移动、第二份是复制
  d = await DocxProjection.load(buf);
  d.diffMd(d.md.replace(imgTok, '').replace('末尾', '末尾' + imgTok + imgTok));
  ({ data, warnings } = await d.save({ type: 'uint8array' }));
  assert.deepEqual(warnings, []);
  assert.equal((await DocxProjection.load(data)).md.split(imgTok).length, 3);
  // 凭空写一个：同样忽略
  d = await DocxProjection.load(buf);
  d.diffMd(d.md.replace('末尾', '末尾![x](media/other.png)'));
  ({ data, warnings } = await d.save({ type: 'uint8array' }));
  assert.match(warnings.join(), /新增图片被忽略/);
  assert.doesNotMatch((await DocxProjection.load(data)).md, /other\.png/);
});

test('图片：审阅模式下删除是删除线、移动的副本换新 id；整行删除会带走图片', async () => {
  const buf = await tabSample();
  let d = await DocxProjection.load(buf);
  d.diffMd(d.md.replace(imgTok, '').replace('末尾', '末尾' + imgTok));
  let { data } = await d.save({ type: 'uint8array', revisions: true });
  const xml = await (await JSZip.loadAsync(data)).file('word/document.xml').async('string');
  assert.match(xml, /<w:del [^>]*><w:r><w:drawing>/);
  assert.match(xml, /<w:ins [^>]*><w:r><w:drawing>/);
  const ids = [...xml.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1]);
  assert.equal(new Set(ids).size, ids.length, 'docPr id 不重复');
  assert.equal((await DocxProjection.load(data)).md.split(imgTok).length, 2);
  d = await DocxProjection.load(buf);
  const n = d.md.split('\n').findIndex((l) => l === imgTok) + 1;
  d.mapPlain((s) => s, { range: lineRange(d.md, n) });
  d.splice([{ s: d.md.indexOf(imgTok), e: d.md.indexOf(imgTok) + imgTok.length, text: '' }]);
  ({ data } = await d.save({ type: 'uint8array' }));
  assert.doesNotMatch((await DocxProjection.load(data)).md, /image1\.png/);
});

// ---------- 统一写回路径（splice / diffMd / mapPlain） ----------
const docXml = async (data) => (await (await JSZip.loadAsync(data)).file('word/document.xml').async('string'));
const bodyRuns = async (data) => [...(await docXml(data)).matchAll(/<w:r>(<w:rPr>.*?<\/w:rPr>)?<w:(t|br)\b[^>]*?(?:\/>|>([^<]*)<\/w:t>)/g)]
  .map((m) => (m[2] === 'br' ? '⏎' : (m[1]?.includes('<w:b/>') ? 'B:' : '') + m[3]));

test('软换行后插入：字落在换行之后（diffMd / mapPlain 一样）', async () => {
  const { BR } = await import('../src/plain.js');
  const buf = await sample(p(t('甲') + '<w:r><w:br/></w:r>' + t('乙')));
  const a = await DocxProjection.load(buf), b = await DocxProjection.load(buf);
  a.mapPlain((s) => s.replace(BR + '乙', BR + 'X乙'));
  b.diffMd(b.md.replace('<br>乙', '<br>X乙'));
  for (const d of [a, b]) {
    const runs = await bodyRuns((await d.save({ type: 'uint8array' })).data);
    const i = runs.indexOf('甲');
    assert.deepEqual(runs.slice(i, i + 4), ['甲', '⏎', 'X', '乙']);
  }
});

test('替换挂在第一个被删的字、用它的格式：diffMd 与 mapPlain 结果一致', async () => {
  const buf = await sample(p(t('张') + t('三', true) + t('丰')));
  const a = await DocxProjection.load(buf), b = await DocxProjection.load(buf);
  a.mapPlain((s) => s.replace('张三丰', '某丰'));
  b.diffMd(b.md.replace('张**三**丰', '某丰'));
  const ra = await bodyRuns((await a.save({ type: 'uint8array' })).data), rb = await bodyRuns((await b.save({ type: 'uint8array' })).data);
  assert.deepEqual(ra, rb);
  const i = ra.indexOf('某');
  assert.deepEqual(ra.slice(i, i + 2), ['某', '丰'], '「某」不加粗（沿用「张」）');
});

test('diffMd：插入的 Markdown 标记不当文字写进 Word', async () => {
  const d = await DocxProjection.load(await sample(p(t('尾巴'))));
  d.diffMd(d.md.replace('尾巴', '尾巴**加粗**\\*'));
  const runs = await bodyRuns((await d.save({ type: 'uint8array' })).data);
  assert.ok(runs.includes('加粗*'), runs.join('|'));
  assert.ok(!runs.some((r) => r.includes('**') || r.includes('\\')));
});

test('diffMd from 另一投影：移动的图片是图片，不会把说明文字当正文插入', async () => {
  const buf = await tabSample();
  const b = await DocxProjection.load(buf);
  b.diffMd(b.md.replace(imgTok, '').replace('末尾', '末尾' + imgTok));
  const bp = await DocxProjection.load((await b.save({ type: 'uint8array' })).data);
  const a = await DocxProjection.load(buf);
  a.diffMd(bp.md, { from: bp });
  const { data, warnings } = await a.save({ type: 'uint8array', revisions: true });
  assert.deepEqual(warnings, []);
  const xml = await docXml(data);
  assert.match(xml, /<w:del [^>]*><w:r><w:drawing>/);
  assert.match(xml, /<w:ins [^>]*><w:r><w:drawing>/);
  assert.doesNotMatch(xml, /<w:t[^>]*>[^<]*张三身份证/);
});

test('save 的 report 按实际写入修正：插进别人的修订里算 skipped', async () => {
  const d = await DocxProjection.load(await sample(p(t('前') + `<w:ins w:id="1" w:author="Bob">${t('他人')}</w:ins>` + t('后'))));
  const st = d.diffMd(d.md.replace('他人后', '他X人后'));
  assert.equal(st.ins, 1);
  const { report } = await d.save({ type: 'uint8array', revisions: true });
  assert.equal(report.ins, 0);
  assert.equal(report.skipped, 1);
});

test('diffSpans：应用后得到新文本；超长行也能算', async () => {
  const { diffSpans } = await import('../src/plain.js');
  const apply = (a, spans) => { let out = '', p = 0; for (const x of spans) { out += a.slice(p, x.s) + x.text; p = x.e; } return out + a.slice(p); };
  const long = (n, seed) => Array.from({ length: n }, (_, i) => '甲乙丙丁'[(i * seed) % 4]).join('');
  for (const [a, b] of [['联系人张三，电话 139', '联系人某，电话 139****'], ['a\nb\nc', 'a\nB\nc\nd'], [long(3000, 7) + '尾', long(3000, 5) + '尾']]) {
    const spans = diffSpans(a, b);
    assert.equal(apply(a, spans), b);
    for (const x of spans) assert.equal(b.slice(x.at, x.at + x.text.length), x.text);
  }
});

test('diffMd：同一张图片只改说明 = 改 alt；换成别的图片不会逐字错配', async () => {
  const d = await DocxProjection.load(await tabSample());
  assert.deepEqual(d.diffMd(d.md.replace(imgTok, imgTok.replace('张三', '某'))), { del: 1, ins: 1, skipped: 0 });
  const xml = await docXml((await d.save({ type: 'uint8array' })).data);
  assert.match(xml, /descr="某身份证"/);
  assert.equal((xml.match(/<w:drawing>/g) || []).length, 3, '图片本身没动');
});

// ---------- 漏投影 ----------
const tc = (body) => `<w:tc>${p(body)}</w:tc>`;
const wsp = (inner) => `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:docPr id="9" name="框"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`;

test('表格里被内容控件 / customXml 包住的行和格也投影（普通表格与特殊表格）', async () => {
  const sdtRow = `<w:sdt><w:sdtContent><w:tr>${tc(t('王五'))}${tc(t('13700001111'))}</w:tr></w:sdtContent></w:sdt>`;
  const plain = `<w:tbl><w:tr>${tc(t('姓名'))}${tc(t('电话'))}</w:tr>${sdtRow}<w:customXml><w:tr>${tc(t('赵六'))}<w:sdt><w:sdtContent>${tc(t('13800002222'))}</w:sdtContent></w:sdt></w:tr></w:customXml></w:tbl>`;
  const merged = `<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${p(t('合并'))}</w:tc></w:tr>${sdtRow}</w:tbl>`;
  const d = await DocxProjection.load(await sample(plain + merged));
  assert.match(d.md, /^\| 王五 \| 13700001111 \|$/m);
  assert.match(d.md, /^\| 赵六 \| 13800002222 \|$/m);
  assert.match(d.md, /^合并\n王五\n13700001111$/m);
  d.mapPlain((s) => s.replaceAll('王五', '某'));
  const xml = await docXml((await d.save({ type: 'uint8array' })).data);
  assert.doesNotMatch(xml, /王五/);
});

test('文本框里套的文本框也投影，改动写回内层', async () => {
  const d = await DocxProjection.load(await sample(p(t('正文') + wsp(p(t('外框') + wsp(p(t('内框张三'))))))));
  assert.match(d.md, /^内框张三$/m);
  assert.ok(!d.unhandled.some((u) => u.tag === 'w:t'));
  d.mapPlain((s) => s.replace('内框张三', '内框某'));
  const xml = await docXml((await d.save({ type: 'uint8array' })).data);
  assert.match(xml, /内框<\/w:t><\/w:r><w:r><w:t[^>]*>某</);
  assert.doesNotMatch(xml, /内框张三/);
});

test('兜底：投影之外的文字（如图形里的 a:t）和批注都报进 unhandled', async () => {
  const shape = `<w:r><w:drawing><wp:inline><wp:docPr id="7" name="形状"/><a:graphic><a:graphicData uri="urn:x-unknown"><a:p><a:r><a:t>形状里的张三</a:t></a:r></a:p></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const buf = await sample(p(t('有形状') + shape));
  const z = await JSZip.loadAsync(buf);
  z.file('word/comments.xml', `<w:comments ${W}><w:comment w:id="0" w:author="x"><w:p><w:r><w:t>批注里的张三</w:t></w:r></w:p></w:comment></w:comments>`);
  z.file('word/_rels/document.xml.rels', (await z.file('word/_rels/document.xml.rels').async('string')).replace('</Relationships>', `<Relationship Id="rCm" Type="${REL}comments" Target="comments.xml"/></Relationships>`));
  const d = await DocxProjection.load(await z.generateAsync({ type: 'uint8array' }));
  assert.deepEqual(d.unhandled.filter((u) => u.tag !== 'w:drawing'), [{ tag: 'a:t', count: 1, sample: '形状里的张三' }, { tag: 'w:comment', count: 1 }]);
  assert.match(d.warnings.join('\n'), /<a:t> ×1（没投影到的文字，如「形状里的张三」）/);
  assert.match(d.warnings.join('\n'), /<w:comment> ×1（批注）/);
});
