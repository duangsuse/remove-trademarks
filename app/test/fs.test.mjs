// fs.js 的权限 / 文件夹逻辑：用假句柄测，不依赖浏览器
import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as fs from '../src/fs.js';

const handle = (perm, onRequest = () => perm) => {
  const calls = [];
  return { calls, queryPermission: async ({ mode }) => (calls.push('q:' + mode), perm), requestPermission: async ({ mode }) => (calls.push('r:' + mode), onRequest()) };
};

test('allowed：已有权限不请求；ask 才请求；写权限用 readwrite', async () => {
  const ok = handle('granted');
  assert.equal(await fs.allowed(ok, true, true), true);
  assert.deepEqual(ok.calls, ['q:readwrite']);
  const p = handle('prompt', () => 'granted');
  assert.equal(await fs.allowed(p), false);
  assert.equal(await fs.allowed(p, false, true), true);
  assert.deepEqual(p.calls, ['q:read', 'q:read', 'r:read']);
  assert.equal(await fs.allowed(handle('prompt', () => 'denied'), true, true), false);
  assert.equal(await fs.allowed({ name: 'a.onnx' }), true, '拖入的 File 无权限概念');
  assert.equal(await fs.allowed(null), false);
});

test('child / list：不存在返回 null，create 失败照常抛出；list 只列文件', async () => {
  const dir = {
    getFileHandle: async (name, { create }) => { if (name === 'a.docx') return { name }; throw new Error(create ? '只读' : 'NotFound'); },
    async *entries() { yield ['a.docx', { kind: 'file' }]; yield ['sub', { kind: 'directory' }]; },
  };
  assert.deepEqual(await fs.child(dir, 'a.docx'), { name: 'a.docx' });
  assert.equal(await fs.child(dir, 'b.docx'), null);
  await assert.rejects(fs.child(dir, 'b.docx', true), /只读/);
  assert.deepEqual((await fs.list(dir)).map((x) => x.name), ['a.docx']);
});
