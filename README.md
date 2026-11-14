# 数据集打码工作流（docx 脱敏）

离线运行的 Word 脱敏工具：把 `.docx` 投影成 Markdown，在上面按黑名单 / 规则 / 离线 NER 圈出敏感文字，再把改动写回 Word（直接替换，或写成审阅修订）。

```
cd app && npm i
npm run dev        # 浏览器里用（需 Chromium：File System Access API）
npm run build      # 打成单个 HTML，可断网使用
npm test           # round-trip 测试
```

地址后加 `#editor` 关掉脱敏功能，当纯文本投影编辑器用：直接改文字，Ctrl+S 存为 `新版/原名.docx`。Ctrl+Enter 审阅面板按 `新版/原名.docx` → `脱敏/原名.docx` → 手选 的顺序找对比文件，把它与原文的差异写成修订存回该文件（原文不动）。

## 三层投影

```
 .docx  ──DocxProjection.load──▶  md（Markdown 投影）  ──plainView──▶  txt（纯文本）
   ▲                                │  ▲                                  │
   └── save() ◀── splice ◀── diffMd(新 md)       └── mapPlain(fn) ◀── fn(txt) => txt'
```

| 层 | 文件 | 内容 |
|---|---|---|
| docx ⇄ md | `src/docx.js` | 段落、标题、编号、链接、图片说明、表格、文本框、页眉页脚、脚注、图表文字；不依赖脱敏逻辑 |
| md ⇄ txt | `src/plain.js` | 去掉 `\` 转义和 `*` 粗斜体；`张**三**` 在 txt 里就是 `张三` |
| 结构约定 | `src/markup.js` | 所有结构标记 / 魔法字符串都只在这里定义 |
| 识别 | `src/detect.js` + `src/rules.js` | 在 txt 上匹配，命中区间换回 md 坐标 |

改动只落在**可编辑字符**上：Markdown 标记、表格竖线、编号前缀这些结构字符删不掉，也插不进正文；碰到它们的改动记进 `skipped`。

## 投影 API

```js
import { DocxProjection } from './src/docx.js';
import { lineRange } from './src/plain.js';

const d = await DocxProjection.load(buffer);   // ArrayBuffer / Uint8Array / Blob
d.md          // Markdown 投影
d.text        // 纯文本投影
d.warnings    // 没处理的内容（多行文字）
d.unhandled   // 投影不了的元素 [{ tag: 'm:oMath', count: 2 }, …]

d.locate(i)   // md 第 i 个字属于哪个段落 / 文本单元，可不可编辑
await d.media() // word/media/* 的图片 [{ path, bytes, mime, replaceable, uses: [{ s, e }] }]

// 三种改法，按输入的层次分；最后都归到 splice，返回本次的 { del, ins, skipped, fmt }
d.splice([{ s, e, text }])  // md 区间：区间里的可编辑字全删，text 插在第一个被删的字处、用它的格式；s === e 表示纯插入
                            // text 按纯文本写：\n 分段，\u2028 软换行；区间里删掉段落之间的换行 = 合并段落
d.diffMd(newMd, { from })   // 整份 md：和原投影做 diff，转成 splice；只改了 ** / * 的地方记成格式修改（fmt）
                            // from 是另一个 DocxProjection 时只取它的可编辑字
d.mapPlain(fn, { range })   // 纯文本：fn(txt) => txt'，diff 转成 splice；range = [s, e]（md 坐标）时只把这一段交给 fn

d.replaceImages([{ path, bytes }]) // 换图（path 如 word/media/image1.png）；审阅模式下每处都是「原图删除 + 新图插入」
d.blankImages(paths)        // 换成 1×1 占位图（即 replaceImages 占位图）
d.wipeMetadata()            // 抹掉作者、批注、修订记录、rsid 等不可见指纹

const { blob, report, warnings } = await d.save();                         // 浏览器
const { data } = await d.save({ type: 'uint8array', revisions: true });    // Node；revisions 表示写成审阅修订
```

- `save()` 只能调用一次（会改写 DOM）。`report` 是按实际写入修正过的累计 `{ del, ins, skipped, fmt }`。
- `revisions: true` 时所有改动都能在 Word 审阅面板里看到：文字插入 / 删除、拆段 / 新段 / 合并（段落标记的修订）、粗斜体（格式修订）；
  链接地址是「原链接删除 + 新链接插入」，图片说明是「原图删除 + 改了说明的副本插入」；图表 / SmartArt 文字和域代码里的链接地址没有修订语法，直接改并在所在段落加一条批注写明原文。
- 坐标都按**载入时**的投影算。多次调用 `splice` / `diffMd` / `mapPlain` 的改动会叠加，但后一次看不到前一次的结果。
- 载入时会先「拒绝」掉本工具以前写进去的审阅修订，所以反复写修订不会越叠越多。

### 例 1：把全文的「张三」换成「某先生」

```js
d.mapPlain((s) => s.replaceAll('张三', '某先生'));
```

`张**三**` 这种半个字加粗的情况也能换掉；替换后的字沿用原来那个字的格式。链接文字、图片说明、文本框、页眉、图表标题会一起换。

### 例 2：只隐去第 11~22 行的手机号（选区）

```js
d.mapPlain((s) => s.replace(/1[3-9]\d{9}/g, (m) => m.slice(0, 3) + '****' + m.slice(7)),
           { range: lineRange(d.md, 11, 22) });
```

`range` 也可以直接传编辑器里的选区 `[selectionStart, selectionEnd]`。

### 例 3：插入、删除

```js
d.mapPlain((s) => s.replace('联系人：', '联系人（已核实）：'));   // 纯插入
d.mapPlain((s) => s.replace('，电话 13912345678', ''));             // 纯删除
```

### 例 4：识别时注入 NER 函数

`detect` 的 NER 是一个函数参数 `(lines) => Promise<[{ text, type, score }]>`。浏览器里传 `ner.recognize`，测试或 Node 里随便塞一个：

```js
import { detect, selection } from './src/detect.js';

const myNer = async (lines) => lines.flatMap((l) => (l.includes('王小明') ? [{ text: '王小明', type: 'PERSON', score: 1 }] : []));
const items = await detect(d.md, { blacklist: ['张三'] }, { ner: myNer });
const sel = selection([...items.values()], (it) => it.cat === 'black' || it.cat === 'ner');
d.replace(sel.map((x) => ({ s: x.s, e: x.e, text: '**' })));
```

加上 `range` 后只把这一段交给 NER（规则仍然全文扫描）：`detect(md, st, { ner, range: lineRange(md, 30, 80) })`。

### 例 5：拿外部编辑器改完再写回

```sh
npm run di -- md  报告.docx > 报告.md        # 导出投影
code 报告.md                                 # 随便改，结构标记不用管
npm run di -- apply 报告.docx 报告.md        # → 报告_新版.docx，只动有差异的字
```

### 例 6：命令行一句话改 / 脱敏

```sh
# 用 JS 函数改纯文本，只改第 3 行
npm run di -- edit 报告.docx -L 3 -e "s => s.replaceAll('张三', '某')"

# 黑名单 + 电话 + 公司名，用默认替换规则（「.*公司 => 某公司」，其余打 **）
npm run di -- redact 报告.docx -w 张三 -w 李四 -c contact,company

# 写成审阅修订交给同事核对，-r 自定义替换规则（同类按出现顺序编号）
npm run di -- redact 报告.docx -w 张三 -r "::black:: .* => 人员{nth()}" --track -o 待审.docx
```

`npm run di -- -h` 可以查看全部参数。

## fs API（浏览器控制台）

`src/fs.js` 封装了 File System Access API 和 IndexedDB。打开页面后，在 F12 控制台里通过 `__di.fs` 调用。先在界面里打开文件夹，这样 `__di.doc.dir` 已经有读写权限：

```js
const { fs, doc, DocxProjection } = __di;

// 当前文件夹里有哪些 docx
(await fs.list(doc.dir)).map((x) => x.name).filter((n) => n.endsWith('.docx'));

// 批量处理：每份文件的手机号只留前三位，另存为 _隐号.docx
for (const { name, h } of await fs.list(doc.dir)) {
  if (!/\.docx$/.test(name) || /_隐号/.test(name)) continue;
  const d = await DocxProjection.load(await fs.read(h));
  d.mapPlain((s) => s.replace(/(1[3-9]\d)\d{8}/g, '$1********'));
  const { blob } = await d.save();
  await fs.write(await fs.child(doc.dir, name.replace(/\.docx$/, '_隐号.docx'), true), blob);
}

// 记住任意值（含文件句柄），下次打开页面还在
await fs.remember('上次的规则', '.*公司 => 某公司');
await fs.recall('上次的规则');
```

`fs.pickDir()` / `fs.pickFile()` 只能在用户点击时调用：`addEventListener('click', async () => console.log(await __di.fs.pickDir()), { once: true })`，然后点一下页面。

`fs.allowed(h, write, ask)` 检查权限：已有权限时直接返回 `true`；`ask` 为真时会在点击里弹出授权框，选「每次访问都允许」后，下次就不用再点了。

## 文本格式：读写放在一起

每种文本格式由一个模块负责读和写，调用方不自己拼分隔符或标记。

**投影的结构（`src/markup.js`）**：`docx.js` 用 `makeMarks()` 写结构，其他代码用 `mdLines(md)` 读。`mdLines` 逐行标出 `kind`（`open` / `close` / `cell` / `table` / `tableSep` / `rule` / `text`），并标出这一行是否属于结构（`structural`）、是否已进入「其他文本」区（`others`）。识别表格里的姓名、整理章节、准备 NER 输入都通过它来读结构。

| 结构 | 写法 |
|---|---|
| 特殊表格（合并单元格等，1 格 1 行） | 以 `<details>:::特殊表格（1格1行）` 开头，`:::</details>` 收尾 |
| 其他文本区 | `---` + `## 其他文本…`，每节是 `### 页眉 header1` 这样的标题 |

`partnerOf(md, s, e)` 用来实现开头行和收尾行之间的跳转。在界面里点一下特殊表格的开头行，页面会平滑滚到收尾行；点收尾行会滚回开头行。正文里如果有以这些标记开头的段落，投影时会在前面加 `\` 转义。

**左栏输入框（`src/boxes.js`）**：

| 框 | 接口 | 格式 |
|---|---|---|
| 黑名单 | `BlackBox.load / dump / add / remove` | `===` 以上手填，以下随右栏同步 |
| 行号规则 | `lineRule.parse / dump` | `#L11~L22` / `#L11` |
| 替换规则 | `parseRules` | `::cname:: .* => {s[0]}姓人员` |
| 自订规则 | `parseCustom` | `工号 :: EMP\d{6}` |

如果要把 `docx.js` 用在别的项目，可以覆盖 `PROJECTION` 里的默认值：

```js
import { makeMarks } from './src/markup.js';
await DocxProjection.load(buf, {
  marks: makeMarks({ open: '[[表', openNote: '', close: '表]]' }),
  author: '我的工具',  // save({ revisions: true }) 写入的修订作者；载入时先拒绝掉这个作者以前写的修订
  i18n: {              // 界面文字，按组覆盖，默认值见 docx.js 的 I18N
    parts: { header: 'Header', footer: 'Footer', textbox: 'Text box' },
    warn: { noMain: (path) => `main part not found: ${path}` },
  },
});
```

## 没有投影的内容

公式（`m:oMath`）、注音（`w:ruby`）、嵌入的外部文档（`w:altChunk`）这类不常见的元素不会做投影。载入时会按标签名列进清单，打开文档时弹窗提示：

```
以下元素没有投影，其中的文字没有处理：
    <m:oMath> ×3（公式）
    <w:ruby> ×1（注音）
    <w:altChunk> ×1（嵌入的外部文档）
```

内嵌的 Excel / Visio、svg / emf 图片同样只做提示，内容原样保留。
