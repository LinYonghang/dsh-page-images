/**
 * Standalone smoke test for `dsh-slide-pages`.
 *
 * Runs the real plugin module against the real durable attachment store, the
 * real LibreOffice exporter and the real PDFium rasterizer, with a minimal
 * stand-in for the Cordis context.
 *
 * Run: node test/smoke.mjs <absolute-pptx-path>
 */

import assert from 'node:assert/strict';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Context } from '@deepseek-ai/cordis';
import { LocalAttachmentStore } from '@deepseek-ai/dsh-attachment-local';
import { apply, Config } from '../lib/index.js';

const deckPath = process.argv[2];
assert.ok(typeof deckPath === 'string' && /\.pptx$/iu.test(deckPath), 'usage: node test/smoke.mjs <absolute-pptx-path>');

const dshHome = await mkdtemp(join(tmpdir(), 'dsh-slide-pages-home-'));
const store = new LocalAttachmentStore(new Context(), { dshHome });

const deckBytes = new Uint8Array(await readFile(deckPath));
const deckRef = await store.saveFile({ data: deckBytes, name: 'sample.pptx' });
const hostPath = store.fileHostPath(deckRef);
assert.equal(typeof hostPath, 'string', 'the local store must expose a host path for the deck');
console.log('stored deck', deckRef.name, deckRef.bytes, 'bytes ->', hostPath);

/* --------------------------------------------------------- fake plugin ctx */

const disposers = [];
const handlers = new Map();
const tools = new Map();
const ctx = {
  attachments: store,
  agents: { list: () => [] },
  tools: { register(tool) { tools.set(tool.name, tool); return () => tools.delete(tool.name); } },
  effect(factory) {
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
  get(service) { return service === 'attachments' ? store : undefined; },
};

const config = Config({});
console.log('resolved config', JSON.stringify(config));
apply(ctx, config);
assert.ok(tools.has('slide_pages'), 'apply() must register the slide_pages tool');
const tool = tools.get('slide_pages');

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

const deckMessage = {
  id: 'msg-1',
  role: 'user',
  source: { kind: 'user' },
  content: [
    { type: 'text', text: '看看这份课件' },
    { type: 'file', attachment: deckRef },
  ],
};

const startedAt = Date.now();
const decision = await agentHandlers.get('agent/pre-step')(
  { messages: [deckMessage], signal: undefined },
  async () => ({ kind: 'enter', messages: [deckMessage] }),
);
console.log('pre-step took', Date.now() - startedAt, 'ms');

assert.equal(decision.kind, 'enter');
assert.equal(decision.messages.length, 2, 'the hook must append exactly one slide-images message');
const injected = decision.messages[1];
assert.equal(injected.role, 'user');
assert.equal(injected.source.kind, 'slide-pages');
const textBlocks = injected.content.filter((block) => block.type === 'text');
const imageBlocks = injected.content.filter((block) => block.type === 'image');
assert.equal(textBlocks.length, 2, 'the message must carry a summary and a deck outline');
assert.ok(imageBlocks.length > 0, 'the injected message must carry slide images');
assert.ok(imageBlocks.length <= config.autoMaxSlides, 'the automatic path must respect autoMaxSlides');
assert.match(textBlocks[0].text, /已转换为图片/u);
assert.match(textBlocks[1].text, /文字大纲/u);
console.log('summary:', textBlocks[0].text.split('\n')[0]);
console.log('outline head:', textBlocks[1].text.split('\n').slice(0, 3).join(' | ').slice(0, 160));

for (const block of imageBlocks) {
  assert.equal(block.attachment.mediaType, 'image/png');
  assert.ok(block.attachment.width > 0 && block.attachment.height > 0);
  const stored = store.imageHostPath(block.attachment);
  assert.equal(typeof stored, 'string', 'each slide must resolve to a stored normalized image');
  const bytes = await readFile(stored);
  assert.deepEqual([...bytes.subarray(0, 8)], [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 'every slide must be a PNG');
}

/* ---------------------------------------------------------------- tool path */

const exec = { agent: { id: 'session-smoke', session: {} }, signal: undefined };
const value = await tool.execute({}, exec);
assert.equal(typeof value.deck, 'string');
assert.ok(value.slideCount >= imageBlocks.length);
assert.equal(value.slides.length, Math.min(value.slideCount, config.maxSlidesPerCall));
for (const slide of value.slides) {
  assert.equal(slide.mediaType, 'image/png');
  assert.equal(typeof slide.attachmentId, 'string');
}
const renderedBlocks = tool.output.render({}, value);
assert.equal(renderedBlocks[0].type, 'text');
assert.match(renderedBlocks[0].text, /<deck>/u);
assert.equal(renderedBlocks.length, value.slides.length + 1);
console.log('tool rendered', value.slides.length, 'of', value.slideCount, 'slide(s) for', value.deck);

const pair = await tool.execute({ source: '', slides: '2-3', dpi: 96 }, exec);
assert.equal(pair.slides.length, 2, 'an explicit slide selection must render exactly those slides');
assert.equal(pair.slides[0].index, 2);
assert.ok(pair.slides[0].width < value.slides[0].width, 'a lower dpi must produce a smaller raster');
const outlined = pair.slides.filter((slide) => slide.title !== undefined);
assert.ok(outlined.length > 0, 'the tool must attach per-slide text from the deck outline');
console.log('slide 2 title:', outlined[0].title);

let rejected = false;
try {
  await tool.execute({ source: '', slides: '99999' }, exec);
} catch (error) {
  rejected = true;
  console.log('out-of-range slide rejected as expected:', String(error.message).slice(0, 120));
}
assert.ok(rejected, 'a slide outside the deck must fail');

/* --------------------------------------------------------------- teardown */

for (const dispose of disposers.reverse()) await dispose?.();
console.log('smoke test passed');
