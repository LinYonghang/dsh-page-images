/**
 * Edge cases for `dsh-slide-pages`: disabled automation, unreadable decks,
 * attachments without a host path, filesystem sources, and macro-enabled decks
 * that the engine only accepts under a known extension.
 *
 * Run: node test/edges.mjs <absolute-pptx-path> [<smaller-pptx-for-speed>]
 */

import assert from 'node:assert/strict';
import { copyFile, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Context } from '@deepseek-ai/cordis';
import { LocalAttachmentStore } from '@deepseek-ai/dsh-attachment-local';
import { apply, Config } from '../lib/index.js';

const deckPath = process.argv[2];
const smallDeck = process.argv[3] ?? deckPath;
assert.ok(typeof deckPath === 'string', 'usage: node test/edges.mjs <absolute-pptx-path> [<smaller-pptx>]');

const work = await mkdtemp(join(tmpdir(), 'dsh-slide-pages-edges-'));
const corruptDeck = join(work, 'corrupt.pptx');
await writeFile(corruptDeck, 'this is plain text pretending to be a presentation\n');
const macroDeck = join(work, 'macro.pptm');
await copyFile(smallDeck, macroDeck);

const dshHome = await mkdtemp(join(tmpdir(), 'dsh-slide-pages-home-'));
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
    get(service) { return service === 'attachments' ? attachments : undefined; },
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
    content: [{ type: 'text', text: 'deck please' }, { type: 'file', attachment }],
  };
  return preStep({ messages: [message], signal: undefined }, async () => ({ kind: 'enter', messages: [message] }));
}

/* 1. autoConvert: false leaves the step untouched, and the tool still works. */
{
  const bytes = new Uint8Array(await readFile(smallDeck));
  const deckRef = await store.saveFile({ data: bytes, name: 'manual.pptx' });
  const harnessed = harness({ attachments: store, config: { autoConvert: false, outline: false } });
  const decision = await runPreStep(harnessed.preStep, deckRef);
  assert.equal(decision.messages.length, 1, 'disabled automation must not append messages');
  const value = await harnessed.tools.get('slide_pages').execute(
    { source: '', slides: '1' },
    { agent: harnessed.agent, signal: undefined },
  );
  assert.equal(value.deck, 'manual.pptx');
  assert.equal(value.slides.length, 1);
  assert.equal(value.slides[0].title, undefined, 'outline: false must not attach slide text');
  console.log('ok: autoConvert=false keeps the step; the tool works on demand and honours outline=false');
  for (const dispose of harnessed.disposers.reverse()) await dispose?.();
}

/* 2. A deck that cannot be parsed reports a failure instead of breaking the turn. */
{
  const brokenRef = await store.saveFile({ data: new Uint8Array([1, 2, 3, 4]), name: 'broken.pptx' });
  const harnessed = harness({ attachments: store, config: {} });
  const decision = await runPreStep(harnessed.preStep, brokenRef);
  assert.equal(decision.messages.length, 2, 'a failed conversion must still append one message');
  const injected = decision.messages[1];
  assert.equal(injected.content.length, 1);
  assert.equal(injected.content[0].type, 'text');
  assert.match(injected.content[0].text, /失败/u);
  console.log('ok: corrupt deck ->', injected.content[0].text.slice(0, 110));
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
  const ref = { attachmentId: 'portable-id', name: 'remote.pptx', bytes: 10 };
  const decision = await runPreStep(harnessed.preStep, ref);
  assert.equal(decision.messages.length, 2);
  assert.match(decision.messages[1].content[0].text, /宿主文件路径/u);
  console.log('ok: missing host path ->', decision.messages[1].content[0].text.slice(0, 90));
  for (const dispose of harnessed.disposers.reverse()) await dispose?.();
}

/* 4. The tool accepts a plain filesystem path as its source. */
{
  const harnessed = harness({ attachments: store, config: {} });
  const tool = harnessed.tools.get('slide_pages');
  const value = await tool.execute({ source: smallDeck, slides: '1' }, { agent: harnessed.agent, signal: undefined });
  assert.equal(value.deck, smallDeck.split(/[\\/]/u).at(-1));
  assert.equal(value.slides.length, 1);
  assert.equal(value.slides[0].name, `${smallDeck.split(/[\\/]/u).at(-1).replace(/\.pptx$/iu, '')}-slide001.png`);
  console.log('ok: tool rendered a filesystem deck source');
  for (const dispose of harnessed.disposers.reverse()) await dispose?.();
}

/* 5. Unknown sources and malformed selections fail with actionable text. */
{
  const harnessed = harness({ attachments: store, config: {} });
  const tool = harnessed.tools.get('slide_pages');
  await assert.rejects(
    () => tool.execute({ source: 'C:/nope/missing.pptx' }, { agent: { ...harnessed.agent, id: 'session-empty' }, signal: undefined }),
    /deck not found/u,
  );
  await assert.rejects(
    () => tool.execute({ source: smallDeck, slides: '1-ab' }, { agent: harnessed.agent, signal: undefined }),
    /invalid slide/u,
  );
  console.log('ok: unknown path and malformed selection rejected with actionable errors');
  for (const dispose of harnessed.disposers.reverse()) await dispose?.();
}

/* 6. A macro-enabled deck is staged under an extension the engine accepts. */
{
  const harnessed = harness({ attachments: store, config: {} });
  const value = await harnessed.tools.get('slide_pages').execute(
    { source: macroDeck, slides: '1' },
    { agent: harnessed.agent, signal: undefined },
  );
  assert.equal(value.slides.length, 1);
  console.log('ok: .pptm deck staged and rendered as slide images');
  for (const dispose of harnessed.disposers.reverse()) await dispose?.();
}

console.log('edge cases passed');
