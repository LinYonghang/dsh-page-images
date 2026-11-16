/**
 * Edge cases for `dsh-pdf-pages`: disabled automation, unreadable PDFs,
 * attachments without a host path, and filesystem sources.
 *
 * Run: node test/edges.mjs <absolute-pdf-path> <path-to-write-a-corrupt.pdf>
 */

import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Context } from '@deepseek-ai/cordis';
import { LocalAttachmentStore } from '@deepseek-ai/dsh-attachment-local';
import { apply, Config } from '../lib/index.js';

const goodPdf = process.argv[2];
assert.ok(typeof goodPdf === 'string', 'usage: node test/edges.mjs <absolute-pdf-path>');

const work = await mkdtemp(join(tmpdir(), 'dsh-pdf-pages-edges-'));
const corruptPdf = join(work, 'corrupt.pdf');
await writeFile(corruptPdf, 'this is plain text pretending to be a pdf\n');

const dshHome = await mkdtemp(join(tmpdir(), 'dsh-pdf-pages-home-'));
const store = new LocalAttachmentStore(new Context(), { dshHome });

/** Build a stand-in plugin context around one attachment facade. */
function harness({ attachments, config }) {
  const handlers = new Map();
  const tools = new Map();
  const disposers = [];
  const ctx = {
    attachments,
    agents: { list: () => [] },
    tools: { register(tool) { tools.set(tool.name, tool); return () => tools.delete(tool.name); } },
    effect(factory) { const dispose = factory(); if (typeof dispose === 'function') disposers.push(dispose); return () => {}; },
    on(event, handler) {
      const list = handlers.get(event) ?? [];
      list.push(handler);
      handlers.set(event, list);
      return () => {};
    },
    inject(_services, callback) { callback(ctx); },
    get(name) { return name === 'attachments' ? attachments : undefined; },
  };
  apply(ctx, Config(config));
  const agentHandlers = new Map();
  const agent = {
    id: 'session-edges',
    ctx: {
      on(event, handler) { agentHandlers.set(event, handler); return () => agentHandlers.delete(event); },
      effect(factory) { const dispose = factory(); if (typeof dispose === 'function') disposers.push(dispose); return () => {}; },
    },
  };
  for (const listener of handlers.get('agent/created') ?? []) listener({ agent });
  return { ctx, tools, disposers, preStep: agentHandlers.get('agent/pre-step'), agent };
}

/** Run one pre-step over a message carrying a file attachment. */
async function runPreStep(preStep, attachment) {
  const message = {
    id: 'msg-edge',
    role: 'user',
    source: { kind: 'user' },
    content: [{ type: 'text', text: 'pdf please' }, { type: 'file', attachment }],
  };
  return preStep({ messages: [message], signal: undefined }, async () => ({ kind: 'enter', messages: [message] }));
}

/* 1. autoConvert: false leaves the step untouched. */
{
  const goodBytes = new Uint8Array(await readFile(goodPdf));
  const pdfRef = await store.saveFile({ data: goodBytes, name: 'manual.pdf' });
  const harnessed = harness({ attachments: store, config: { autoConvert: false } });
  const decision = await runPreStep(harnessed.preStep, pdfRef);
  assert.equal(decision.messages.length, 1, 'disabled automation must not append messages');
  const value = await harnessed.tools.get('pdf_pages').execute(
    { source: '', pages: '1' },
    { agent: harnessed.agent, signal: undefined },
  );
  assert.ok(value.pageCount >= 1);
  assert.equal(value.images.length, 1);
  assert.equal(value.pdf, 'manual.pdf');
  console.log('ok: autoConvert=false keeps the step and the tool still works on demand');
  for (const dispose of harnessed.disposers.reverse()) await dispose?.();
}

/* 2. A PDF that cannot be parsed reports a failure instead of breaking the turn. */
{
  const brokenRef = await store.saveFile({ data: new Uint8Array([1, 2, 3, 4]), name: 'broken.pdf' });
  const harnessed = harness({ attachments: store, config: {} });
  const decision = await runPreStep(harnessed.preStep, brokenRef);
  assert.equal(decision.messages.length, 2, 'a failed conversion must still append one message');
  const injected = decision.messages[1];
  assert.equal(injected.content.length, 1);
  assert.equal(injected.content[0].type, 'text');
  assert.match(injected.content[0].text, /失败/u);
  console.log('ok: corrupt PDF ->', injected.content[0].text.slice(0, 90));
  for (const dispose of harnessed.disposers.reverse()) await dispose?.();
}

/* 3. A deployment without host file paths explains itself instead of throwing. */
{
  const portable = {
    imageLimits: store.imageLimits,
    fileHostPath: () => undefined,
    saveImages: store.saveImages.bind(store),
  };
  const harnessed = harness({ attachments: portable, config: {} });
  const ref = { attachmentId: 'portable-id', name: 'remote.pdf', bytes: 10 };
  const decision = await runPreStep(harnessed.preStep, ref);
  assert.equal(decision.messages.length, 2);
  assert.match(decision.messages[1].content[0].text, /宿主文件路径/u);
  console.log('ok: missing host path ->', decision.messages[1].content[0].text.slice(0, 90));
  for (const dispose of harnessed.disposers.reverse()) await dispose?.();
}

/* 4. The tool accepts a plain filesystem path as its source. */
{
  const harnessed = harness({ attachments: store, config: {} });
  const tool = harnessed.tools.get('pdf_pages');
  const value = await tool.execute({ source: goodPdf, pages: '1' }, { agent: harnessed.agent, signal: undefined });
  assert.equal(value.pdf, goodPdf.split(/[\\/]/u).at(-1));
  assert.equal(value.images.length, 1);
  assert.equal(value.images[0].name, `${goodPdf.split(/[\\/]/u).at(-1).replace(/\.pdf$/iu, '')}-p0001.png`);
  console.log('ok: tool rendered a filesystem PDF source');
  for (const dispose of harnessed.disposers.reverse()) await dispose?.();
}

/* 5. Unknown sources and unsupported selections fail with actionable text. */
{
  const harnessed = harness({ attachments: store, config: {} });
  const tool = harnessed.tools.get('pdf_pages');
  await assert.rejects(
    () => tool.execute({ source: 'C:/nope/missing.pdf' }, { agent: { ...harnessed.agent, id: 'session-empty' }, signal: undefined }),
    /PDF not found/u,
  );
  await assert.rejects(
    () => tool.execute({ source: goodPdf, pages: '1-ab' }, { agent: harnessed.agent, signal: undefined }),
    /invalid page/u,
  );
  console.log('ok: unknown path and malformed selection rejected with actionable errors');
  for (const dispose of harnessed.disposers.reverse()) await dispose?.();
}

console.log('edge cases passed');
