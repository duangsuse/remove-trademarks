// 本地文件访问（File System Access API，需 Chromium）：选择、读写，句柄存 IndexedDB 供下次打开时找回。
// 权限：新选的句柄只有本次会话的权限；对找回的句柄在点击 / 按键里调用 allowed(h, write, true)，
// Chromium 会弹出带「每次访问都允许」的对话框，选了之后下次 allowed(h) 直接为真，无需再点

// ---------- 记住：IndexedDB 'di' / 'h'，存句柄或任意可结构化克隆的值 ----------
let dbp = null;
const db = () => (dbp ||= new Promise((res, rej) => {
  const r = indexedDB.open('di', 1);
  r.onupgradeneeded = () => r.result.createObjectStore('h');
  r.onsuccess = () => res(r.result);
  r.onerror = () => { dbp = null; rej(r.error); };
}));
export async function remember(key, v) { try { (await db()).transaction('h', 'readwrite').objectStore('h').put(v, key); } catch { /* 无存储 */ } }
export async function recall(key) {
  try {
    const store = (await db()).transaction('h').objectStore('h');
    return await new Promise((res) => { const q = store.get(key); q.onsuccess = () => res(q.result ?? null); q.onerror = () => res(null); });
  } catch { return null; }
}

// ---------- 权限 ----------
// 已有权限返回 true；ask 为真时没有就请求（须在用户手势内）
export async function allowed(h, write = false, ask = false) {
  if (!h?.queryPermission) return !!h; // 拖入的 File 等，无权限概念
  const mode = write ? 'readwrite' : 'read';
  if ((await h.queryPermission({ mode })) === 'granted') return true;
  return ask && (await h.requestPermission({ mode })) === 'granted';
}

// ---------- 选择 ----------
export const pickDir = () => showDirectoryPicker({ mode: 'readwrite' });
export const pickFile = async (types) => (await showOpenFilePicker({ types }))[0];
export const pickSave = (suggestedName, types) => showSaveFilePicker({ suggestedName, types });

// ---------- 读写 ----------
export const file = (h) => (h.getFile ? h.getFile() : h); // FileSystemFileHandle 或 File 都行
export const read = async (h) => (await file(h)).arrayBuffer();
export async function write(h, data) { const w = await h.createWritable(); await w.write(data); await w.close(); }
// 文件夹里的文件：不存在时返回 null；create 为真则新建（失败照常抛出）
export const child = (dir, name, create = false) => dir.getFileHandle(name, { create }).catch((e) => { if (create) throw e; return null; });
export async function list(dir) {
  const out = [];
  for await (const [name, h] of dir.entries()) if (h.kind === 'file') out.push({ name, h });
  return out;
}
