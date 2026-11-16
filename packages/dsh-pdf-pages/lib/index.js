/**
 * `dsh-pdf-pages` — turn the PDF attached to a message into one image per page
 * and post those images back into the conversation.
 *
 * Two entry points share one conversion pipeline:
 *
 * - `agent/pre-step` detects PDF file blocks in the messages about to enter a
 *   step, renders them, and appends a user-role message carrying the page
 *   images. The model sees the pages in the same request and the Web UI shows
 *   the images as a normal conversation message.
 * - the `pdf_pages` tool renders an explicit selection (pages, density) on
 *   demand, which also covers documents larger than the automatic cap.
 *
 * Rendering goes through the PDFium rasterizer DSH ships
 * (`@deepseek-ai/libreoffice-kit`), and every page is stored through
 * `ctx.attachments` so the images are durable, addressable, and visible to the
 * model on later turns.
 * @module dsh-pdf-pages
 */

import { stat } from 'node:fs/promises';
import { basename } from 'node:path';
import z from '@deepseek-ai/schemastery';
import { defineTool } from '@deepseek-ai/dsh-tools';
import { createUserMessage } from '@deepseek-ai/dsh-llm';
import { createPdfRenderer, parsePageSelection, absoluteInputPath } from './render.js';

/** Cordis plugin name used by loader diagnostics and patch rows. */
export const name = 'pdf-pages';

/** Services this plugin cannot work without. */
export const inject = ['tools', 'attachments'];

/** Message-source tag for everything this plugin appends. */
const SOURCE = Object.freeze({ kind: 'pdf-pages' });

/** Attachment display names this plugin treats as PDFs. */
const PDF_NAME = /\.pdf$/iu;

/** Bounds that keep long-lived in-memory bookkeeping small. */
const MAX_TRACKED_AGENTS = 12;
const MAX_TRACKED_REFS = 256;
const MAX_CACHED_RENDERS = 6;
/** Characters a stored page name must not carry. */
const UNSAFE_NAME = /[\\/:*?"<>|\u0000-\u001f]+/gu;

/** Config for the plugin, editable from the profile's `cordis.patch.yml`. */
export const Config = z.object({
  autoConvert: z.boolean().default(true),
  autoMaxPages: z.number().default(10),
  autoMaxAttachments: z.number().default(2),
  dpi: z.number().default(110),
  maxPagesPerCall: z.number().default(24),
  timeoutMs: z.number().default(180_000),
});

/**
 * Whether a file block names a PDF.
 * @param ref - a `FileAttachmentRef` from a message.
 * @returns whether the attachment is a PDF.
 */
function isPdfRef(ref) {
  return typeof ref?.name === 'string' && PDF_NAME.test(ref.name);
}

/**
 * Collect PDF file references in message order.
 * @param messages - messages about to enter a step.
 * @returns the PDF references found, oldest first.
 */
function collectPdfRefs(messages) {
  const found = [];
  for (const message of messages ?? []) {
    const content = Array.isArray(message?.content) ? message.content : [];
    for (const block of content) {
      if (block?.type === 'file' && isPdfRef(block.attachment)) found.push(block.attachment);
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
 * Human-readable reason for a failed conversion, without stack noise.
 * @param error - the caught failure.
 * @returns a one-line message suitable for a conversation message or tool error.
 */
function failureText(error) {
  const code = typeof error?.code === 'string' ? `${error.code}: ` : '';
  const message = error instanceof Error ? error.message : String(error);
  return `${code}${message}`.slice(0, 600);
}

/**
 * Display name for one rendered page image.
 * @param label - the document's display name.
 * @param page - the one-based page number.
 * @returns a filesystem-safe PNG name that keeps the page number sortable.
 */
function pageName(label, page) {
  const stem = label.replace(/\.pdf$/iu, '').replace(UNSAFE_NAME, '_').slice(-80);
  return `${stem}-p${String(page).padStart(4, '0')}.png`;
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
 * Register the PDF pages plugin.
 * @param ctx - the plugin context.
 * @param config - the resolved `Config` value.
 */
export function apply(ctx, config) {
  const settings = {
    autoConvert: config.autoConvert !== false,
    autoMaxPages: clampInteger(config.autoMaxPages, 10, 1, 60),
    autoMaxAttachments: clampInteger(config.autoMaxAttachments, 2, 1, 8),
    dpi: clampInteger(config.dpi, 110, 48, 300),
    maxPagesPerCall: clampInteger(config.maxPagesPerCall, 24, 1, 60),
  };

  const renderer = createPdfRenderer({ timeoutMs: clampInteger(config.timeoutMs, 180_000, 10_000, 900_000) });
  ctx.effect(() => () => renderer.dispose(), 'pdf-pages: renderer');

  /** Latest PDF attachment per conversation, for the tool's default source. */
  const latestByAgent = new Map();
  /** Durable references by attachment id, so a tool call can name one. */
  const refsById = new Map();
  /** Completed renders, keyed by source bytes and selection, to skip re-rendering. */
  const renderedCache = new Map();
  /** Which `agent.id` each installed pre-step listener belongs to. */
  const agentListeners = new Map();

  /**
   * Remember where a PDF came from so a later tool call can find it again.
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
   * Resolve a PDF source into an absolute path the rasterizer can read.
   * @param source - an explicit path, an attachment id, or empty for the newest attachment.
   * @param agentId - the calling conversation.
   * @returns the ref (when the source is a stored attachment) and the file path.
   * @throws when the source cannot be resolved to a readable PDF.
   */
  async function resolveSource(source, agentId) {
    const text = typeof source === 'string' ? source.trim() : '';
    if (text !== '' && (text.includes('/') || text.includes('\\') || PDF_NAME.test(text))) {
      const candidate = absoluteInputPath(text);
      try {
        const info = await stat(candidate);
        if (info.isFile()) return { path: candidate, label: basename(candidate), ref: undefined };
      } catch {
        // Not a path that exists; fall through to attachment lookup and report below.
      }
      if (text.includes('/') || text.includes('\\')) {
        throw new Error(`PDF not found at "${text}"; pass an absolute path to an existing .pdf file, an attachment id, or omit the argument to use the PDF attached to this conversation`);
      }
    }
    const byId = text === '' ? undefined : refsById.get(text);
    const ref = byId ?? (text === '' ? latestByAgent.get(agentId) : undefined)
      ?? (text === '' ? [...refsById.values()].at(-1) : undefined);
    if (ref === undefined) {
      throw new Error(text === ''
        ? 'no PDF attachment is known for this conversation yet; attach a PDF to a message or pass an absolute .pdf path'
        : `no attached PDF matches "${text}"; pass the attachment id, an absolute .pdf path, or omit the argument to use the newest PDF in this conversation`);
    }
    const hostPath = ctx.attachments.fileHostPath?.(ref);
    if (typeof hostPath !== 'string' || hostPath === '') {
      throw new Error(`the attached PDF "${ref.name}" is not backed by a host file path, so this deployment cannot rasterize it`);
    }
    return { path: hostPath, label: ref.name, ref };
  }

  /**
   * Render the requested pages and admit each one as a durable image attachment.
   * @param request - what to render.
   * @param request.path - absolute PDF path.
   * @param request.label - display name for the rendered pages.
   * @param request.pages - explicit page numbers, or undefined for every page.
   * @param request.dpi - raster density.
   * @param request.maxPages - page-count cap applied to the "every page" selection.
   * @param request.signal - caller cancellation.
   * @returns the stored pages plus the document's page count.
   */
  async function convertAndStore({ path, label, pages, dpi, maxPages, signal }) {
    const cacheKey = `${path}|${dpi}|${pages === undefined ? 'all' : pages.join(',')}|${maxPages}`;
    const cached = renderedCache.get(cacheKey);
    if (cached !== undefined) return cached;

    const render = pages === undefined
      ? renderer.renderAll({ inputPath: path, dpi, maxPages, signal })
      : renderer.renderAll({ inputPath: path, pages, dpi, maxPages: pages.length, signal });

    const result = await render;
    signal?.throwIfAborted();

    const batchLimit = Math.max(1, Math.min(ctx.attachments.imageLimits?.maxImagesPerMessage ?? 20, 20));
    const entries = [];
    for (let offset = 0; offset < result.images.length; offset += batchLimit) {
      const batch = result.images.slice(offset, offset + batchLimit);
      const stored = await ctx.attachments.saveImages(batch.map((image) => ({
        data: image.data,
        mediaType: 'image/png',
        name: pageName(label, image.page),
      })));
      for (const [index, ref] of stored.entries()) {
        const image = batch[index];
        entries.push({
          page: image.page,
          width: ref.width,
          height: ref.height,
          bytes: ref.bytes,
          mediaType: 'image/png',
          attachmentId: String(ref.attachmentId),
          name: pageName(label, image.page),
          ref: imageRef(ref),
        });
      }
    }

    const outcome = {
      label,
      pageCount: result.pageCount,
      entries,
      omittedPages: Math.max(0, result.pageCount - entries.length),
    };
    renderedCache.set(cacheKey, outcome);
    while (renderedCache.size > MAX_CACHED_RENDERS) renderedCache.delete(renderedCache.keys().next().value);
    return outcome;
  }

  /**
   * Build the conversation content for a batch of rendered pages.
   * @param outcome - the conversion outcome.
   * @param dpi - raster density used.
   * @returns content blocks: a summary line followed by the page images.
   */
  function pageContent(outcome, dpi) {
    const name = outcome.label;
    const first = outcome.entries[0]?.page ?? 1;
    const last = outcome.entries.at(-1)?.page ?? first;
    const range = first === last ? `第 ${first} 页` : `第 ${first}–${last} 页`;
    const missing = outcome.omittedPages > 0
      ? `\n文档共 ${outcome.pageCount} 页，其余 ${outcome.omittedPages} 页未转换；可用 pdf_pages 工具指定页码继续转换。`
      : `\n文档共 ${outcome.pageCount} 页，已全部转换为图片。`;
    return [
      {
        type: 'text',
        text: `PDF 附件已逐页转换为图片：${name} 的${range}（${dpi} dpi，每张图片都是可持久化的附件）。${missing}`,
      },
      ...outcome.entries.map((entry) => ({ type: 'image', attachment: entry.ref })),
    ];
  }

  /* ------------------------------------------------------------------ tool */

  ctx.tools.register(defineTool({
    name: 'pdf_pages',
    description: 'Render pages of a PDF into images and post them into this conversation. Use it for a PDF attached to a message (omit source to use the newest attached PDF), a PDF already on disk, or when specific pages or a higher resolution are needed. Pages arrive as image blocks, so read the figures and text directly from them.',
    parameters: {
      source: {
        type: 'string',
        description: 'Optional PDF source: an attachment id from an earlier message, an absolute .pdf path, or a file name. Omit to use the newest PDF attached in this conversation.',
      },
      pages: {
        type: 'string',
        description: 'Optional page selection such as "1-3,7"; omit or pass "all" for every page (capped by maxPages).',
      },
      dpi: {
        type: 'number',
        description: 'Optional raster density from 48 to 300; default 110, which keeps a page legible without inflating image tokens.',
      },
      maxPages: {
        type: 'number',
        description: 'Optional cap on rendered pages. Raise it to read a long document in full.',
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          pdf: { type: 'string', required: true },
          dpi: { type: 'number', required: true },
          pageCount: { type: 'number', required: true },
          omittedPages: { type: 'number', required: true },
          images: {
            type: 'array',
            required: true,
            items: {
              type: 'object',
              additionalProperties: false,
              properties: {
                page: { type: 'number', required: true },
                attachmentId: { type: 'string', required: true },
                mediaType: { type: 'string', required: true },
                bytes: { type: 'number', required: true },
                width: { type: 'number', required: true },
                height: { type: 'number', required: true },
                name: { type: 'string' },
              },
            },
          },
        },
      },
      render: (_args, value) => {
        const head = value.images[0]?.page ?? 1;
        const tail = value.images.at(-1)?.page ?? head;
        const summary = `<pdf>${value.pdf}</pdf>\n<type>pdf-pages</type>\n<content>\n${value.images.length} page image(s) at ${value.dpi} dpi, pages ${head}-${tail} of ${value.pageCount}${value.omittedPages > 0 ? ` (${value.omittedPages} page(s) not rendered)` : ''}\n</content>`;
        return [
          { type: 'text', text: summary },
          ...value.images.map((image) => ({
            type: 'image',
            attachment: {
              attachmentId: image.attachmentId,
              mediaType: image.mediaType,
              bytes: image.bytes,
              width: image.width,
              height: image.height,
              ...(image.name === undefined ? {} : { name: image.name }),
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
      const pages = parsePageSelection(args.pages);
      if (pages !== undefined && pages.length > settings.maxPagesPerCall) {
        throw new Error(`selecting ${pages.length} pages exceeds this plugin's per-call limit of ${settings.maxPagesPerCall}; narrow the selection`);
      }
      const dpi = clampInteger(args.dpi, settings.dpi, 48, 300);
      const maxPages = clampInteger(args.maxPages, settings.maxPagesPerCall, 1, 60);
      const outcome = await convertAndStore({
        path: source.path,
        label: source.label,
        pages,
        dpi,
        maxPages: pages === undefined ? maxPages : pages.length,
        signal: exec.signal,
      });
      return {
        pdf: source.label,
        dpi,
        pageCount: outcome.pageCount,
        omittedPages: outcome.omittedPages,
        images: outcome.entries.map(({ ref: _ref, ...rest }) => rest),
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
      const entering = collectPdfRefs(decision.messages ?? payload.messages);
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
              content: [{ type: 'text', text: `PDF 附件 ${ref.name} 没有可用的宿主文件路径，无法在本机转换为图片。` }],
              source: SOURCE,
            }));
            continue;
          }
          const outcome = await convertAndStore({
            path,
            label: ref.name,
            pages: undefined,
            dpi: settings.dpi,
            maxPages: settings.autoMaxPages,
            signal: payload.signal,
          });
          additions.push(createUserMessage({ content: pageContent(outcome, settings.dpi), source: SOURCE }));
        } catch (error) {
          if (payload.signal?.aborted === true) return decision;
          additions.push(createUserMessage({
            content: [{ type: 'text', text: `PDF 附件 ${ref.name} 转换为图片失败：${failureText(error)}` }],
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
    }, `pdf-pages: agent ${agentId}`);
  }

  ctx.on('agent/created', ({ agent }) => {
    watchAgent(agent);
  });

  // Agents that were already live when this plugin loaded (a restored session,
  // a hot reload in a running app) never emit `agent/created` again, so adopt
  // the current roster once. `watchAgent` is idempotent per agent id.
  ctx.inject(['agents'], (agentsCtx) => {
    for (const agent of agentsCtx.agents.list()) watchAgent(agent);
  });

  ctx.effect(() => () => {
    for (const [agentId, off] of agentListeners) {
      off();
      agentListeners.delete(agentId);
    }
  }, 'pdf-pages: agent hooks');
}
