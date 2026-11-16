/**
 * `dsh-slide-pages` — turn a presentation attached to a message into one image
 * per slide and post those slides back into the conversation.
 *
 * Two entry points share one pipeline:
 *
 * - `agent/pre-step` detects `.pptx`/`.ppt`/`.odp` file blocks in the messages
 *   about to enter a step, converts the deck and appends a user-role message
 *   carrying the leading slide images plus the deck's text outline. The model
 *   sees both in the same request and the Web UI shows a normal message.
 * - the `slide_pages` tool renders an explicit selection on demand, which also
 *   covers decks longer than the automatic cap.
 *
 * Rendering is two-stage (`convert()` exports the deck to PDF, PDFium rasterizes
 * the pages) because LibreOffice's direct presentation raster clips content;
 * see `./render.js`. Every slide is stored through `ctx.attachments`, so the
 * images are durable and visible to the model on later turns.
 * @module dsh-slide-pages
 */

import { copyFile, mkdir, rm, stat } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import { basename, extname, isAbsolute, join, resolve as resolvePath } from 'node:path';
import { tmpdir } from 'node:os';
import z from '@deepseek-ai/schemastery';
import { defineTool } from '@deepseek-ai/dsh-tools';
import { createUserMessage } from '@deepseek-ai/dsh-llm';
import { createSlideRenderer, parseSlideSelection, slideFailureText } from './render.js';
import { formatOutline, readDeckOutline } from './outline.js';

/** Cordis plugin name used by loader diagnostics and patch rows. */
export const name = 'slide-pages';

/** Services this plugin cannot work without. */
export const inject = ['tools', 'attachments'];

/** Message-source tag for everything this plugin appends. */
const SOURCE = Object.freeze({ kind: 'slide-pages' });

/** Attachment display names this plugin treats as presentations. */
const DECK_NAME = /\.(pptx|pptm|ppt|odp)$/iu;

/** Extensions the shipped engine accepts; anything else is staged under one. */
const ENGINE_EXTENSIONS = new Set(['.ppt', '.pptx', '.odp']);

/** Characters a stored slide name must not carry. */
const UNSAFE_NAME = /[\\/:*?"<>|\u0000-\u001f]+/gu;

/** Bounds that keep long-lived in-memory bookkeeping small. */
const MAX_TRACKED_AGENTS = 12;
const MAX_TRACKED_REFS = 256;
const MAX_CACHED_RENDERS = 4;

/** One-line note about the limits of engine rendering, appended once per deck. */
const RENDER_NOTE = '版面由 LibreOffice 渲染：字体替换、动画、SmartArt、视频与 3D 只取静态帧，可能与 PowerPoint 显示有差异。';

/** Config for the plugin, editable from the profile's `cordis.patch.yml`. */
export const Config = z.object({
  autoConvert: z.boolean().default(true),
  autoMaxSlides: z.number().default(8),
  autoMaxAttachments: z.number().default(2),
  dpi: z.number().default(110),
  maxSlidesPerCall: z.number().default(24),
  outline: z.boolean().default(true),
  outlineBodyChars: z.number().default(300),
  outlineNotesChars: z.number().default(160),
  outlineTotalChars: z.number().default(8000),
  timeoutMs: z.number().default(300_000),
});

/**
 * Whether a file block names a presentation.
 * @param ref - a `FileAttachmentRef` from a message.
 * @returns whether the attachment is a deck.
 */
function isDeckRef(ref) {
  return typeof ref?.name === 'string' && DECK_NAME.test(ref.name);
}

/**
 * Collect deck file references in message order.
 * @param messages - messages about to enter a step.
 * @returns the deck references found, oldest first.
 */
function collectDeckRefs(messages) {
  const found = [];
  for (const message of messages ?? []) {
    const content = Array.isArray(message?.content) ? message.content : [];
    for (const block of content) {
      if (block?.type === 'file' && isDeckRef(block.attachment)) found.push(block.attachment);
    }
  }
  return found;
}

/**
 * Project a stored image reference into the `ImageAttachmentRef` an image block
 * carries, dropping fields the block does not own.
 * @param ref - the durable reference returned by `ctx.attachments`.
 * @returns the image block attachment payload.
 */
function imageRef(ref) {
  return {
    attachmentId: ref.attachmentId,
    mediaType: ref.mediaType,
    bytes: ref.bytes,
    width: ref.width,
    height: ref.height,
    ...(typeof ref.name === 'string' ? { name: ref.name } : {}),
    ...(ref.originalDimensions === undefined ? {} : { originalDimensions: { ...ref.originalDimensions } }),
  };
}

/**
 * Display name for one rendered slide image.
 * @param label - the deck's display name.
 * @param slide - the one-based slide number.
 * @returns a filesystem-safe PNG name that keeps the slide number sortable.
 */
function slideName(label, slide) {
  const stem = label.replace(DECK_NAME, '').replace(UNSAFE_NAME, '_').slice(-80);
  return `${stem}-slide${String(slide).padStart(3, '0')}.png`;
}

/**
 * Clamp a user-supplied number into an inclusive range.
 * @param value - the candidate value.
 * @param fallback - the value used when the candidate is not a finite number.
 * @param min - lower bound.
 * @param max - upper bound.
 * @returns the clamped integer.
 */
function clampInteger(value, fallback, min, max) {
  if (!Number.isFinite(value)) return fallback;
  return Math.min(max, Math.max(min, Math.round(value)));
}

/**
 * Register the slide pages plugin.
 * @param ctx - the plugin context.
 * @param config - the resolved `Config` value.
 */
export function apply(ctx, config) {
  const settings = {
    autoConvert: config.autoConvert !== false,
    autoMaxSlides: clampInteger(config.autoMaxSlides, 8, 1, 60),
    autoMaxAttachments: clampInteger(config.autoMaxAttachments, 2, 1, 8),
    dpi: clampInteger(config.dpi, 110, 48, 300),
    maxSlidesPerCall: clampInteger(config.maxSlidesPerCall, 24, 1, 60),
    outline: config.outline !== false,
    outlineBodyChars: clampInteger(config.outlineBodyChars, 300, 0, 4000),
    outlineNotesChars: clampInteger(config.outlineNotesChars, 160, 0, 4000),
    outlineTotalChars: clampInteger(config.outlineTotalChars, 8000, 0, 40_000),
  };

  const renderer = createSlideRenderer({ timeoutMs: clampInteger(config.timeoutMs, 300_000, 10_000, 900_000) });
  ctx.effect(() => () => renderer.dispose(), 'slide-pages: renderer');

  /** Latest deck attachment per conversation, for the tool's default source. */
  const latestByAgent = new Map();
  /** Durable references by attachment id, so a tool call can name one. */
  const refsById = new Map();
  /** Completed renders, keyed by deck identity and selection. */
  const renderedCache = new Map();
  /** Which `agent.id` each installed pre-step listener belongs to. */
  const agentListeners = new Map();

  /**
   * Remember where a deck came from so a later tool call can find it again.
   * @param agentId - owning conversation id.
   * @param ref - the attachment reference.
   */
  function remember(agentId, ref) {
    latestByAgent.delete(agentId);
    latestByAgent.set(agentId, ref);
    refsById.delete(ref.attachmentId);
    refsById.set(ref.attachmentId, ref);
    while (latestByAgent.size > MAX_TRACKED_AGENTS) latestByAgent.delete(latestByAgent.keys().next().value);
    while (refsById.size > MAX_TRACKED_REFS) refsById.delete(refsById.keys().next().value);
  }

  /**
   * Give the engine a path it accepts, staging a copy when the extension is not
   * one it recognises (macro-enabled decks share the OOXML container).
   * @param path - the source path.
   * @returns the path to hand to the engine plus a cleanup for a staged copy.
   */
  async function stageDeck(path) {
    if (ENGINE_EXTENSIONS.has(extname(path).toLowerCase())) return { path, cleanup: async () => {} };
    const dir = join(tmpdir(), `dsh-slide-pages-stage-${randomUUID()}`);
    await mkdir(dir, { recursive: true });
    const staged = join(dir, 'deck.pptx');
    await copyFile(path, staged);
    return { path: staged, cleanup: async () => { await rm(dir, { recursive: true, force: true }).catch(() => {}); } };
  }

  /**
   * Resolve a deck source into an absolute path the engine can read.
   * @param source - an explicit path, an attachment id, or empty for the newest attachment.
   * @param agentId - the calling conversation.
   * @returns the ref (when stored), the display name and the absolute path.
   * @throws when the source cannot be resolved to a readable deck.
   */
  async function resolveSource(source, agentId) {
    const text = typeof source === 'string' ? source.trim() : '';
    if (text !== '' && (text.includes('/') || text.includes('\\') || DECK_NAME.test(text))) {
      const candidate = isAbsolute(text) ? text : resolvePath(process.cwd(), text);
      try {
        const info = await stat(candidate);
        if (info.isFile()) {
          return { path: candidate, label: basename(candidate), ref: undefined, key: `${candidate}|${info.size}|${info.mtimeMs}` };
        }
      } catch {
        // Fall through to attachment lookup, then report below.
      }
      if (text.includes('/') || text.includes('\\')) {
        throw new Error(`deck not found at "${text}"; pass an absolute path to an existing presentation, an attachment id, or omit the argument to use the deck attached to this conversation`);
      }
    }
    const byId = text === '' ? undefined : refsById.get(text);
    const ref = byId ?? (text === '' ? latestByAgent.get(agentId) : undefined)
      ?? (text === '' ? [...refsById.values()].at(-1) : undefined);
    if (ref === undefined) {
      throw new Error(text === ''
        ? 'no deck attachment is known for this conversation yet; attach a presentation to a message or pass an absolute path'
        : `no attached deck matches "${text}"; pass the attachment id, an absolute path, or omit the argument to use the newest deck in this conversation`);
    }
    const hostPath = ctx.attachments.fileHostPath?.(ref);
    if (typeof hostPath !== 'string' || hostPath === '') {
      throw new Error(`the attached deck "${ref.name}" is not backed by a host file path, so this deployment cannot rasterize it`);
    }
    return { path: hostPath, label: ref.name, ref, key: `attachment:${ref.attachmentId}` };
  }

  /**
   * Read the deck outline when it is enabled and the deck carries one.
   * @param path - absolute deck path.
   * @returns the per-slide outline, or undefined.
   */
  async function outlineFor(path) {
    if (!settings.outline) return undefined;
    try {
      return await readDeckOutline(path, { bodyChars: settings.outlineBodyChars, notesChars: settings.outlineNotesChars });
    } catch {
      return undefined;
    }
  }

  /**
   * Render the requested slides and admit each one as a durable image attachment.
   * @param request - what to render.
   * @param request.path - absolute deck path.
   * @param request.key - deck identity, for the PDF cache and the render cache.
   * @param request.label - display name for the rendered slides.
   * @param request.slides - explicit slide numbers, or undefined for the leading `maxSlides`.
   * @param request.dpi - raster density.
   * @param request.maxSlides - slide-count cap for the "every slide" selection.
   * @param request.signal - caller cancellation.
   * @returns the stored slides, the deck's slide count and font diagnostics.
   */
  async function convertAndStore({ path, key, label, slides, dpi, maxSlides, signal }) {
    const cacheKey = `${key}|${dpi}|${slides === undefined ? 'all' : slides.join(',')}|${maxSlides}`;
    const cached = renderedCache.get(cacheKey);
    if (cached !== undefined) return cached;

    const staged = await stageDeck(path);
    let result;
    try {
      result = await renderer.render({
        inputPath: staged.path,
        key,
        pages: slides,
        dpi,
        maxPages: slides === undefined ? maxSlides : slides.length,
        signal,
      });
    } finally {
      await staged.cleanup();
    }
    signal?.throwIfAborted();

    const batchLimit = Math.max(1, Math.min(ctx.attachments.imageLimits?.maxImagesPerMessage ?? 20, 20));
    const entries = [];
    for (let offset = 0; offset < result.images.length; offset += batchLimit) {
      const batch = result.images.slice(offset, offset + batchLimit);
      const stored = await ctx.attachments.saveImages(batch.map((image) => ({
        data: image.data,
        mediaType: 'image/png',
        name: slideName(label, image.page),
      })));
      for (const [index, ref] of stored.entries()) {
        const image = batch[index];
        entries.push({
          slide: image.page,
          width: ref.width,
          height: ref.height,
          bytes: ref.bytes,
          mediaType: 'image/png',
          attachmentId: String(ref.attachmentId),
          name: slideName(label, image.page),
          ref: imageRef(ref),
        });
      }
    }

    const outcome = {
      label,
      key,
      path,
      slideCount: result.slideCount,
      missingFonts: result.missingFonts ?? [],
      entries,
      omittedSlides: Math.max(0, result.slideCount - entries.length),
    };
    renderedCache.set(cacheKey, outcome);
    while (renderedCache.size > MAX_CACHED_RENDERS) renderedCache.delete(renderedCache.keys().next().value);
    return outcome;
  }

  /**
   * Build one conversation message for a converted deck.
   * @param outcome - the conversion outcome.
   * @param dpi - raster density used.
   * @param outline - the deck outline, when available.
   * @returns content blocks: a summary, the deck outline, then the slide images.
   */
  function deckContent(outcome, dpi, outline) {
    const first = outcome.entries[0]?.slide ?? 1;
    const last = outcome.entries.at(-1)?.slide ?? first;
    const range = first === last ? `第 ${first} 张` : `第 ${first}–${last} 张`;
    const total = outcome.slideCount;
    const head = [
      `演示文稿《${outcome.label}》${range}已转换为图片（共 ${total} 张，${dpi} dpi，每张都是可持久化的附件）。`,
      outcome.omittedSlides > 0
        ? `其余 ${outcome.omittedSlides} 张未转图，可用 slide_pages 指定页码继续（例如 slides: "${last + 1}-${Math.min(total, last + 8)}"）。`
        : '全 deck 已全部转换为图片。',
      outcome.missingFonts.length > 0 ? `提示：本机缺少字体 ${outcome.missingFonts.slice(0, 6).join('、')}，部分版式可能与 PowerPoint 不同。` : '',
      RENDER_NOTE,
    ].filter((line) => line !== '');

    const outlineText = outline === undefined ? '' : formatOutline(outline, { totalChars: settings.outlineTotalChars });
    const blocks = [{ type: 'text', text: head.join('\n') }];
    if (outlineText !== '') {
      blocks.push({ type: 'text', text: `全 deck 文字大纲（标题 / 正文摘要 / 演讲备注）：\n${outlineText}` });
    }
    for (const entry of outcome.entries) blocks.push({ type: 'image', attachment: entry.ref });
    return blocks;
  }

  /* ------------------------------------------------------------------ tool */

  ctx.tools.register(defineTool({
    name: 'slide_pages',
    description: 'Render the slides of a presentation (.pptx/.ppt/.odp) into images and post them into this conversation. Use it for a deck attached to a message (omit source to use the newest deck), a deck already on disk, or when specific slides or a higher resolution are needed. Slides arrive as image blocks, so read the layout, figures and text directly from them.',
    parameters: {
      source: {
        type: 'string',
        description: 'Optional deck source: an attachment id from an earlier message, an absolute path, or a file name. Omit to use the newest deck attached in this conversation.',
      },
      slides: {
        type: 'string',
        description: 'Optional slide selection such as "1-3,7"; omit or pass "all" for the leading slides up to maxSlides.',
      },
      dpi: {
        type: 'number',
        description: 'Optional raster density from 48 to 300; default 110, which keeps a slide legible without inflating image tokens.',
      },
      maxSlides: {
        type: 'number',
        description: 'Optional cap on rendered slides. Raise it to read a long deck in full.',
      },
      outline: {
        type: 'boolean',
        description: 'Optional: include each rendered slide title and body text beside the images. Default true.',
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          deck: { type: 'string', required: true },
          dpi: { type: 'number', required: true },
          slideCount: { type: 'number', required: true },
          omittedSlides: { type: 'number', required: true },
          missingFonts: { type: 'array', required: true, items: { type: 'string' } },
          slides: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                index: { type: 'number', required: true },
                attachmentId: { type: 'string', required: true },
                mediaType: { type: 'string', required: true },
                bytes: { type: 'number', required: true },
                width: { type: 'number', required: true },
                height: { type: 'number', required: true },
                name: { type: 'string' },
                title: { type: 'string' },
                body: { type: 'string' },
                notes: { type: 'string' },
              },
            },
          },
        },
      },
      render: (_args, value) => {
        const first = value.slides[0]?.index ?? 1;
        const last = value.slides.at(-1)?.index ?? first;
        const head = [
          `<deck>${value.deck}</deck>`,
          '<type>slide-pages</type>',
          '<content>',
          `${value.slides.length} slide image(s) at ${value.dpi} dpi, slides ${first}-${last} of ${value.slideCount}${value.omittedSlides > 0 ? ` (${value.omittedSlides} slide(s) not rendered)` : ''}`,
          ...(value.missingFonts.length === 0 ? [] : [`missing fonts may change the layout: ${value.missingFonts.slice(0, 6).join(', ')}`]),
          '</content>',
        ];
        const outlined = value.slides.filter((slide) => slide.title !== undefined || slide.body !== undefined || slide.notes !== undefined);
        if (outlined.length > 0) {
          head.push('<outline>');
          for (const slide of outlined) {
            const parts = [`第 ${slide.index} 张`];
            if (slide.title !== undefined && slide.title !== '') parts.push(`《${slide.title}》`);
            if (slide.body !== undefined && slide.body !== '') parts.push(`— ${slide.body}`);
            if (slide.notes !== undefined && slide.notes !== '') parts.push(`（备注：${slide.notes}）`);
            head.push(parts.join(' '));
          }
          head.push('</outline>');
        }
        return [
          { type: 'text', text: head.join('\n') },
          ...value.slides.map((slide) => ({
            type: 'image',
            attachment: {
              attachmentId: slide.attachmentId,
              mediaType: slide.mediaType,
              bytes: slide.bytes,
              width: slide.width,
              height: slide.height,
              ...(slide.name === undefined ? {} : { name: slide.name }),
            },
          })),
        ];
      },
      presentationMeta: (args) => (typeof args.source === 'string' ? { source: args.source } : {}),
    },
    async execute(args, exec) {
      const agentId = String(exec.agent?.id ?? '');
      const source = await resolveSource(args.source, agentId);
      if (source.ref !== undefined) remember(agentId, source.ref);
      const slides = parseSlideSelection(args.slides);
      if (slides !== undefined && slides.length > settings.maxSlidesPerCall) {
        throw new Error(`selecting ${slides.length} slides exceeds this plugin's per-call limit of ${settings.maxSlidesPerCall}; narrow the selection`);
      }
      const dpi = clampInteger(args.dpi, settings.dpi, 48, 300);
      const maxSlides = clampInteger(args.maxSlides, settings.maxSlidesPerCall, 1, 60);
      const outline = args.outline === false ? undefined : await outlineFor(source.path);
      const outcome = await convertAndStore({
        path: source.path,
        key: source.key,
        label: source.label,
        slides,
        dpi,
        maxSlides: slides === undefined ? maxSlides : slides.length,
        signal: exec.signal,
      });
      const byIndex = new Map((outline ?? []).map((entry) => [entry.index, entry]));
      return {
        deck: source.label,
        dpi,
        slideCount: outcome.slideCount,
        omittedSlides: outcome.omittedSlides,
        missingFonts: outcome.missingFonts,
        slides: outcome.entries.map(({ ref: _ref, ...rest }) => {
          const text = byIndex.get(rest.slide);
          return {
            index: rest.slide,
            attachmentId: rest.attachmentId,
            mediaType: rest.mediaType,
            bytes: rest.bytes,
            width: rest.width,
            height: rest.height,
            ...(rest.name === undefined ? {} : { name: rest.name }),
            ...(text?.title === undefined || text.title === '' ? {} : { title: text.title }),
            ...(text?.body === undefined ? {} : { body: text.body }),
            ...(text?.notes === undefined ? {} : { notes: text.notes }),
          };
        }),
      };
    },
  }));

  /* --------------------------------------------------- automatic conversion */

  /**
   * Install the step-boundary hook for one agent.
   * @param agent - the live agent handle.
   */
  function watchAgent(agent) {
    const agentId = String(agent.id ?? '');
    if (agentId === '' || agentListeners.has(agentId)) return;
    const off = agent.ctx.on('agent/pre-step', async (payload, next) => {
      const decision = await next();
      if (decision.kind !== 'enter') return decision;
      const entering = collectDeckRefs(decision.messages ?? payload.messages);
      const refs = [...new Map(entering.map((ref) => [String(ref.attachmentId), ref])).values()];
      if (refs.length === 0) return decision;
      for (const ref of refs) remember(agentId, ref);
      if (!settings.autoConvert) return decision;

      const additions = [];
      for (const ref of refs.slice(-settings.autoMaxAttachments)) {
        if (payload.signal?.aborted === true) return decision;
        try {
          const path = ctx.attachments.fileHostPath?.(ref);
          if (typeof path !== 'string' || path === '') {
            additions.push(createUserMessage({
              content: [{ type: 'text', text: `演示文稿附件 ${ref.name} 没有可用的宿主文件路径，无法在本机转换为图片。` }],
              source: SOURCE,
            }));
            continue;
          }
          const outline = await outlineFor(path);
          const outcome = await convertAndStore({
            path,
            key: `attachment:${ref.attachmentId}`,
            label: ref.name,
            slides: undefined,
            dpi: settings.dpi,
            maxSlides: settings.autoMaxSlides,
            signal: payload.signal,
          });
          additions.push(createUserMessage({ content: deckContent(outcome, settings.dpi, outline), source: SOURCE }));
        } catch (error) {
          if (payload.signal?.aborted === true) return decision;
          additions.push(createUserMessage({
            content: [{ type: 'text', text: `演示文稿 ${ref.name} 转换为图片失败：${slideFailureText(error)}` }],
            source: SOURCE,
          }));
        }
      }
      if (additions.length === 0) return decision;
      return { ...decision, messages: [...decision.messages, ...additions] };
    });
    agentListeners.set(agentId, off);
    agent.ctx.effect(() => () => {
      off();
      agentListeners.delete(agentId);
      latestByAgent.delete(agentId);
    }, `slide-pages: agent ${agentId}`);
  }

  ctx.on('agent/created', ({ agent }) => {
    watchAgent(agent);
  });

  // Agents that were already live when this plugin loaded never emit
  // `agent/created` again, so adopt the current roster once; `watchAgent` is
  // idempotent per agent id.
  ctx.inject(['agents'], (agentsCtx) => {
    for (const agent of agentsCtx.agents.list()) watchAgent(agent);
  });

  ctx.effect(() => () => {
    for (const [agentId, off] of agentListeners) {
      off();
      agentListeners.delete(agentId);
    }
  }, 'slide-pages: agent hooks');
}
