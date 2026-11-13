// 全部识别规则集中在此（正则 / 标签词 / 清单分组 / NER 类型映射）。detect.js 只负责执行。
// 字段：name 显示名；en 受哪个勾选项控制；cls 高亮配色；warn 需手动复核；re 正则（全局）；check 二次校验；hint 示例

// 微软官方说明：检查文档时可删除的隐藏数据和个人信息
export const FINGERPRINT_URL = 'https://support.microsoft.com/zh-cn/office/collab-files/remove-hidden-data-and-personal-information-by-inspecting-documents-presentations-or-workbooks';
// 简称判定：实体（自动识别 / 手动添加）文字长度 <= 该值
export const ABBR_MAX = 3;
// 称谓代词（XX经理）：红色 badge
export const TITLE_RE = /^[\u4e00-\u9fa5]{1,2}(?:经理|总监|总工|主任|工程师|局长|处长|科长|厂长|书记|院长|教授|老师|博士|主管|会计|总)$/;
// 右侧面板归属：tag=简称/电话/专利号；ent=实体清单；其余（日期、证件号、网址…）进「自动识别」，置信度 0
export const TAG_CATS = new Set(['abbr', 'patent', 'patentWord', 'patentTitle', 'mobile', 'landline', 'email', 'intl', 'tblname']);
export const ENT_CATS = new Set(['company', 'addr']);

const cn = '\\u4e00-\\u9fa5';
const luhn = (s) => {
  const d = s.replace(/\D/g, '');
  if (d.length < 13 || d.length > 19) return false;
  let sum = 0;
  for (let i = 0; i < d.length; i++) { let n = +d[d.length - 1 - i]; if (i % 2) { n *= 2; if (n > 9) n -= 9; } sum += n; }
  return sum % 10 === 0;
};

// 公司名左边界：正则没有分词，会把名称前面的动词短语一起吃进来（"全面采用上海…有限公司"、"部署我公司选用华为…有限公司"）。
// 返回要从左边截掉的字数：截到最后一个「明确的动词」之后；动词不在开头、截完仍 >= 4 字才截（避免伤到 "联合电子" 这类名称）。
const LEAD_VERB = /(?:选用|采用|采购|购买|选择|委托|聘请|引进|依托|配置|部署|接入|对接|使用|签约|通过)/g; // 不含「联合」：它常是名称的一部分（联合电子），截错会漏掉名称前半
function companyLead(text) {
  let k = 0;
  for (const m of text.matchAll(LEAD_VERB)) {
    const end = m.index + m[0].length;
    if (m.index >= 1 && text.length - end >= 4) k = end;
  }
  return k;
}

export const RULES = {
  patent: { name: 'ZL专利号', cls: 'warn', warn: true, re: /ZL\s?\d{4}\s?[\dXx]{6,9}(?:\.[\dXx])?/gi, hint: 'ZL202110123456.7' },
  date: {
    name: '日期和百分比', cls: 'tag', hint: '2022-03-03、2026年5月、30%',
    re: /\d{4}\s?年\s?\d{1,2}\s?月(?:\s?\d{1,2}\s?日)?|\d{4}-\d{1,2}(?:-\d{1,2})?(?!\d)|\d{4}\/\d{1,2}(?:\/\d{1,2})?(?!\d)|\d{4}\.\d{1,2}\.\d{1,2}(?!\d)|\d+(?:\.\d+)?\s?%|\d{4}\s?年度?/g,
  },
  // ZL专利号选项还包括：「发明专利」「专利号」字样，以及「一种……」开头的专利名称（一种 + 5 个以上非空白字符，以空白结束）
  patentWord: { name: '专利字样', en: 'patent', cls: 'warn', warn: true, re: /发明专利|专利号/g, hint: '发明专利、专利号' },
  patentTitle: { name: '专利名称', en: 'patent', cls: 'warn', warn: true, re: /一种[^\s|*]{5,}(?=\s)/g, hint: '一种基于…的方法 ' },
  mobile: { name: '手机号', en: 'contact', cls: 'phone', re: /(?<!\d)1[3-9]\d{9}(?!\d)/g, hint: '18912636641' },
  landline: { name: '座机', en: 'contact', cls: 'phone', re: /(?<!\d)\(?0\d{2,3}\)?-\d{7,8}(?:-\d{1,6})?(?!\d)/g, hint: '025-83123456' },
  email: { name: '电子邮件', en: 'contact', cls: 'mail', re: /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, hint: 'someone@example.com' },
  credit: { name: '统一社会信用代码', cls: 'warn', re: /(?<![0-9A-Za-z])[0-9A-HJ-NP-RTUWXY]{2}\d{6}[0-9A-HJ-NP-RTUWXY]{10}(?![0-9A-Za-z])/g, hint: '91320900MA7HAK6C0X' },
  idcard: { name: '身份证号', cls: 'warn', re: /(?<!\d)[1-9]\d{5}(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx](?!\d)/g, hint: '320902199001011234' },
  cname: { name: '姓名（标签后）', cls: 'warn', re: new RegExp(`(?<=(?:姓名|联系人|负责人|法定代表人|法人代表|填报人|申报人|项目负责人|经办人|代表人)[：:]\\s?)[${cn}·]{2,4}`, 'g'), hint: '联系人：徐天祥' },
  company: { name: '公司／组织名称', cls: 'tag', re: new RegExp(`(?<=[：:，,。；;、“"（(\\s|*]|^)[${cn}A-Za-z0-9（）()]{2,24}?(?:集团股份有限公司|集团有限公司|股份有限公司|有限责任公司|有限公司|分公司|事务所|研究院|研究所)`, 'gm'), lead: companyLead, hint: '超维微电子（盐城）有限公司' },
  addr: { name: '地址', cls: 'tag', re: new RegExp(`[${cn}]{2,12}(?:省|市)[${cn}]{1,10}(?:区|县)[${cn}\\d\\-]{1,25}?(?:号|栋|幢|楼|室)`, 'g'), hint: '盐城市盐都区盐渎路999号' },
  intl: { name: '国际电话', cls: 'phone', re: /(?<![\w+])\+\d{1,3}[\s-]?\(?\d{1,4}\)?[\s-]?\d{3,4}[\s-]?\d{3,4}(?!\d)/g, hint: '+86 138 0013 8000' },
  ssn: { name: '美国社会安全码 (SSN)', cls: 'warn', re: /(?<!\d)\d{3}-\d{2}-\d{4}(?!\d)/g, hint: '123-45-6789' },
  card: { name: '信用卡号', cls: 'warn', re: /(?<!\d)(?:\d[ -]?){12,18}\d(?!\d)/g, check: luhn, hint: '4242 4242 4242 4242' },
  iban: { name: 'IBAN 国际银行账号', cls: 'warn', re: /\b[A-Z]{2}\d{2}(?: ?[A-Z0-9]{4}){2,7}(?: ?[A-Z0-9]{1,3})?\b/g, hint: 'GB82 WEST 1234 5698 7654 32' },
  ip4: { name: 'IPv4 地址', cls: 'tag', re: /(?<![\d.])(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)(?![\d.])/g, hint: '192.168.1.10' },
  ip6: { name: 'IPv6 地址', cls: 'tag', re: /(?<![\w:])(?:[A-Fa-f0-9]{1,4}:){1,7}:?(?:[A-Fa-f0-9]{1,4})?(?![\w:])/g, check: (s) => s.includes('::') || (s.match(/:/g) || []).length === 7, hint: '2001:db8::1' },
  url: { name: '网址', cls: 'tag', re: /https?:\/\/[^\s)）"”'’<>|*]+/g, hint: 'https://example.com' },
  birth: { name: '出生日期', cls: 'warn', re: /(?<=(?:出生日期|出生年月|生日)[：:]\s?)\d{4}[-年/.]\d{1,2}(?:[-月/.]\d{1,2}日?)?/g, hint: '出生日期：1985-03-14' },
};
for (const [id, r] of Object.entries(RULES)) r.en ||= id;

// 勾选清单：与 di-app.png 一致的顺序与默认状态；special 项不是正则
export const MAIN = [
  { id: 'img', name: '默认不保留图片' },
  { id: 'meta', name: '.docx 不可见指纹抹除', info: FINGERPRINT_URL },
  { id: 'black', name: '黑名单词汇' },
  { id: 'abbr', name: '简称词汇打码', def: false },
  { id: 'patent', name: 'ZL专利号打码', note: '请手动复核', grp: 'open' },
  { id: 'tblname', name: '含姓名的表格', note: '请手动复核' },
  { id: 'date', name: '日期和百分比', def: false },
  { id: 'contact', name: '电话和关联实体', grp: 'open' },
];

// 收进 <details> 的其余规则（来自 cmp1~3 的规则页）
export const EXTRA = [
  { id: 'credit', name: '统一社会信用代码' }, { id: 'idcard', name: '身份证号' }, { id: 'cname', name: '姓名（仅头衔）', grp: 'open' },
  { id: 'company', name: '公司／组织名称', grp: 'open' }, { id: 'addr', name: '地址' }, { id: 'intl', name: '国际电话' },
  { id: 'ssn', name: '美国社会安全码 (SSN)' }, { id: 'card', name: '信用卡号' }, { id: 'iban', name: 'IBAN 国际银行账号' },
  { id: 'ip4', name: 'IPv4 地址' }, { id: 'ip6', name: 'IPv6 地址' }, { id: 'url', name: '网址' }, { id: 'birth', name: '出生日期' },
  { id: 'ner_cand', name: '机构 / 产品（模型备选）' },
  { id: 'ner_verify', name: 'NER 校验公司／地址正则', nn: true }, { id: 'ner_alias', name: 'NER 人名放大（张总 / 小张）', nn: true },
  { id: 'ner_num', name: 'NER 日期 / 金额 / 百分比补漏', nn: true },
].map((x) => ({ ...x, def: false }));

// 脱敏工具（勾选项与 di-app.png 一致）：scope=tools 为行为开关，scope=cats 为 NER 高亮
export const TOOLS = [
  { id: 'drag', scope: 'tools', name: '拖选即删', def: false },
  { id: 'chapter', scope: 'tools', name: '章节右键删(T键开关浮框)', def: false },
  { id: 'ner_person', scope: 'cats', name: '离线NLP人名高亮', def: true },
  { id: 'ner_place', scope: 'cats', name: '地名省市名高亮', def: false },
];

// 左栏该项关闭时，对应条目在右侧列表里也隐藏（量大 / 属于「高亮」类）；其余规则关闭时仍列出，只是默认不选中，可点击切换
export const HIDE_WHEN_OFF = new Set(['date', 'ner_person', 'ner_place', 'ner_num']);

// 含姓名的表格：这些标签单元格（或特殊表格中的这些行）的下一格视为姓名
export const NAME_LABELS = /^(?:姓名|联系人|联系人姓名|负责人|项目负责人|法定代表人|法人代表|法人|填报人|申报人|经办人|代表人|总监|技术总监|填表人)$/;
export const NAME_VALUE = /^[一-龥·]{2,4}$/;

// 简称：以下简称“XX”
export const ABBR_RE = /简称[“"「『]([^”"」』\n*|]{2,12})[”"」』]/g;

// NER 类型 -> 分组（决定受哪个勾选项控制）
export const NER_GROUPS = { PERSON: 'person', GPE: 'place', LOC: 'place', ORG: 'cand', FAC: 'cand', PRODUCT: 'cand', NORP: 'cand' };
export const TYPE_NAME = { PERSON: '人名', GPE: '地名', LOC: '地名', DATE: '日期', TIME: '时间', MONEY: '金额', PERCENT: '百分比', QUANTITY: '数量', ORG: '机构', FAC: '设施', PRODUCT: '产品', NORP: '群体' };

// 章节识别（无 # 标题时的回退规则）
export const CHAPTER_FALLBACK = [
  [/^第[一二三四五六七八九十百零\d]+[章节篇部分]/, () => 1],
  [/^[一二三四五六七八九十]+[、.．]/, () => 1],
  [/^[（(][一二三四五六七八九十]+[)）]/, () => 2],
  [/^\d+(?:\.\d+){0,3}(?=[\s、.．]\s*[\u4e00-\u9fa5A-Za-z（(])/, (m) => m[0].split('.').length],
];

// 替换规则的默认文本
export const DEFAULT_REPLACE_RULES = '.*公司 => 某公司\n. => *';

// 姓名值的排除词：标签格右边若是这些字段名，说明不是姓名
export const NAME_STOP = /^(?:电话|手机|手机号|邮箱|邮件|电子邮箱|电子邮件|联系电话|职务|职称|部门|传真|地址|性别|学历|年龄|备注|姓名|联系人|负责人|无|暂无|略)$/;
