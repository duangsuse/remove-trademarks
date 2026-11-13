// node --test test/  —— 用 jsdom 提供 DOMParser / XMLSerializer
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { JSDOM } from 'jsdom';
import JSZip from 'jszip';

const { window } = new JSDOM('');
globalThis.DOMParser = window.DOMParser;
globalThis.XMLSerializer = window.XMLSerializer;
const { DocxProjection, TOOL_AUTHOR } = await import('../src/docx.js');
const { scan, selection, replacement, parseRules, chapters, plainOf } = await import('../src/detect.js');

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const DRAW = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const t = (s, b) => `<w:r>${b ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${s}</w:t></w:r>`;
const p = (body, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${body}</w:p>`;
const txbx = (s) => `<w:txbxContent>${p(t(s))}</w:txbxContent>`;

async function sample() {
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
  ].join('');
  z.file('word/document.xml', `<?xml version="1.0"?><w:document ${W} ${DRAW}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  return z.generateAsync({ type: 'uint8array' });
}

async function allText(data) {
  const z = await JSZip.loadAsync(data);
  let s = '';
  for (const f of Object.values(z.files)) if (!f.dir && /\.(xml|rels)$/.test(f.name)) s += await f.async('string');
  return s;
}
const st = (words) => ({ cats: {}, blacklist: words, ranges: [], custom: [], ner: [] });

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
  assert.equal(d.anchor(i).unit.kind, 'text');
  assert.equal(d.anchor(md.indexOf('](mailto')).editable, false);
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
  const stats = d.replace(sel.map((x) => ({ ...x, text: 'XX' })));
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
  d.replace([{ s, e: s + 1, text: '某' }]);
  const d2 = await DocxProjection.load((await d.save({ type: 'uint8array' })).data);
  assert.match(d2.md, /联系人：张\*\*某\*\*，/);
});

test('审阅修订：写入后重新载入与原文一致，重复写入不叠加', async () => {
  const buf = await sample();
  const a0 = await DocxProjection.load(buf);
  const b = await DocxProjection.load(buf);
  b.replace(selection([...scan(b.md, st(['张三'])).values()], () => true).map((x) => ({ ...x, text: '某' })));
  const bData = (await b.save({ type: 'uint8array' })).data;
  const bp = await DocxProjection.load(bData);
  const a = await DocxProjection.load(buf);
  const stats = a.applyText(bp.md, { from: bp });
  assert.ok(stats.del && stats.ins);
  const { data, warnings } = await a.save({ track: true, type: 'uint8array' });
  assert.ok(warnings.some((w) => /无法写成修订/.test(w)), '链接 / 图表文字不写成修订');
  const xml = await allText(data);
  assert.ok(xml.includes(`w:author="${TOOL_AUTHOR}"`));
  const again = await DocxProjection.load(data);
  assert.equal(again.md, a0.md, '载入时撤掉本工具的修订，投影回到原文');
  const a2 = await DocxProjection.load(data);
  a2.applyText(bp.md, { from: bp });
  const data2 = (await a2.save({ track: true, type: 'uint8array' })).data;
  assert.equal((await allText(data2)).split(`w:author="${TOOL_AUTHOR}"`).length, xml.split(`w:author="${TOOL_AUTHOR}"`).length);
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
});

const REAL = new URL('../../3、C3443 阀门和旋塞制造-杰林石化-2026江苏省先进级智能工厂申报书.docx', import.meta.url);
test('样例申报书：标题按样式识别，替换后可再解析', { skip: !existsSync(REAL) }, async () => {
  const buf = readFileSync(REAL);
  const d = await DocxProjection.load(buf);
  const heads = chapters(d.md).filter((h) => /^#/.test(d.md.slice(h.s, h.s + 1)));
  assert.ok(heads.length >= 40, `识别到 ${heads.length} 个样式标题`);
  const items = scan(d.md, { ...st([]), cats: {} });
  const sel = selection([...items.values()], () => true);
  d.replace(sel.map((x) => ({ ...x, text: '**' })));
  const { data, stats, warnings } = await d.save({ type: 'uint8array' });
  console.log('样例：', sel.length, '处', stats, warnings);
  const d2 = await DocxProjection.load(data);
  assert.ok(d2.md.length > 1000);
});
