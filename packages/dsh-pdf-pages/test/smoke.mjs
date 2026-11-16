/**
 * Standalone smoke test for `dsh-pdf-pages`.
 *
 * Runs the real plugin module against the real durable attachment store and the
 * real PDFium rasterizer, with a minimal stand-in for the Cordis context, so a
 * rendering or admission regression is visible without booting DSH.
 *
 * Run: node test/smoke.mjs <absolute-pdf-path>
 */

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Context } from '@deepseek-ai/cordis';
import { LocalAttachmentStore } from '@deepseek-ai/dsh-attachment-local';
import { apply, Config } from '../lib/index.js';

const pdfPath = process.argv[2];
assert.ok(typeof pdfPath === 'string' && pdfPath.toLowerCase().endsWith('.pdf'), 'usage: node test/smoke.mjs <absolute-pdf-path>');

const dshHome = await mkdtemp(join(tmpdir(), 'dsh-pdf-pages-home-'));
const attachmentCtx = new Context();
const store = new LocalAttachmentStore(attachmentCtx, { dshHome });

const pdfBytes = new Uint8Array(await readFile(pdfPath));
const pdfRef = await store.saveFile({ data: pdfBytes, name: 'sample.pdf' });
const hostPath = store.fileHostPath(pdfRef);
assert.equal(typeof hostPath, 'string', 'the local store must expose a host path for the PDF');
console.log('stored PDF', pdfRef.name, pdfRef.bytes, 'bytes ->', hostPath);

/* --------------------------------------------------------- fake plugin ctx */

const disposers = [];
const handlers = new Map();
const tools = new Map();
const ctx = {
  attachments: store,
  agents: { list: () => [] },
  tools: { register(tool) { tools.set(tool.name, tool); return () => tools.delete(tool.name); } },
  effect(factory, _label) {
    const dispose = factory();
    if (typeof dispose === 'function') disposers.push(dispose);
    return () => {};
  },
  on(event, handler) {
    const list = handlers.get(event) ?? [];
    list.push(handler);
    handlers.set(event, list);
    return () => {};
  },
  inject(_services, callback) { callback(ctx); },
  get(name) { return name === 'attachments' ? store : undefined; },
};

const config = Config({});
console.log('resolved config', JSON.stringify(config));
apply(ctx, config);

assert.ok(tools.has('pdf_pages'), 'apply() must register the pdf_pages tool');
const tool = tools.get('pdf_pages');

/* ------------------------------------------------------ fake live agent it */

const agentHandlers = new Map();
const agent = {
  id: 'session-smoke',
  ctx: {
    on(event, handler) { agentHandlers.set(event, handler); return () => agentHandlers.delete(event); },
    effect(factory) { const dispose = factory(); if (typeof dispose === 'function') disposers.push(dispose); return () => {}; },
  },
};
for (const listener of handlers.get('agent/created') ?? []) await listener({ agent });
assert.ok(agentHandlers.has('agent/pre-step'), 'the plugin must hook agent/pre-step for the live agent');

/* ------------------------------------------------ automatic conversion path */

const userMessage = Object.freeze({
  id: 'msg-1',
  role: 'user',
  source: { kind: 'user' },
  content: Object.freeze([
    { type: 'text', text: '看看这份 PDF' },
    { type: 'file', attachment: pdfRef },
  ]),
});

const startedAt = Date.now();
const decision = await agentHandlers.get('agent/pre-step')(
  { messages: [userMessage], signal: undefined },
  async () => ({ kind: 'enter', messages: [userMessage] }),
);
console.log('pre-step took', Date.now() - startedAt, 'ms');

assert.equal(decision.kind, 'enter');
assert.equal(decision.messages.length, 2, 'the hook must append exactly one page-images message');
const injected = decision.messages[1];
assert.equal(injected.role, 'user');
assert.equal(injected.source.kind, 'pdf-pages');
assert.equal(injected.content[0].type, 'text');
const imageBlocks = injected.content.filter((block) => block.type === 'image');
assert.ok(imageBlocks.length > 0, 'the injected message must carry page images');
console.log('injected', imageBlocks.length, 'page image(s):', injected.content[0].text.split('\n')[0]);

for (const block of imageBlocks) {
  assert.equal(block.attachment.mediaType, 'image/png');
  assert.ok(block.attachment.width > 0 && block.attachment.height > 0);
  assert.ok(block.attachment.bytes > 0);
  const stored = store.imageHostPath(block.attachment);
  assert.equal(typeof stored, 'string', 'each page must resolve to a stored normalized image');
  const bytes = await readFile(stored);
  assert.deepEqual([...bytes.subarray(0, 8)], [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 'every page must be a PNG');
}

/* ---------------------------------------------------------------- tool path */

const exec = { agent: { id: 'session-smoke', session: {} }, signal: undefined };
const value = await tool.execute({}, exec);
assert.equal(typeof value.pdf, 'string');
assert.ok(value.pageCount >= imageBlocks.length);
assert.equal(value.images.length, Math.min(value.pageCount, config.maxPagesPerCall));
for (const image of value.images) {
  assert.equal(image.mediaType, 'image/png');
  assert.equal(typeof image.attachmentId, 'string');
}
const renderedBlocks = tool.output.render({}, value);
assert.equal(renderedBlocks[0].type, 'text');
assert.equal(renderedBlocks.length, value.images.length + 1);
console.log('tool rendered', value.images.length, 'of', value.pageCount, 'page(s) for', value.pdf);

const single = await tool.execute({ source: '', pages: '1', dpi: 96 }, exec);
assert.equal(single.images.length, 1, 'an explicit page selection must render exactly those pages');
assert.ok(single.images[0].width < value.images[0].width, 'a lower dpi must produce a smaller raster');

let rejected = false;
try {
  await tool.execute({ source: '', pages: '99999' }, exec);
} catch (error) {
  rejected = true;
  console.log('out-of-range page rejected as expected:', String(error.message).slice(0, 120));
}
assert.ok(rejected, 'a page outside the document must fail');

/* --------------------------------------------------------------- teardown */

for (const dispose of disposers.reverse()) await dispose?.();
console.log('smoke test passed');
