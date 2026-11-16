/**
 * PDF page rasterization for `dsh-pdf-pages`.
 *
 * The rasterizer is the one DSH already ships: `@deepseek-ai/libreoffice-kit`
 * imports a PDF through PDFium and writes one PNG per page. The package lives
 * in the DSH installation rather than in the profile, so the import goes
 * through the profile package resolver; the fallbacks below only cover the
 * case where that resolver is unavailable (plain `node` runs, tests).
 * @module dsh-pdf-pages/render
 */

import { readFile, rm } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import { tmpdir } from 'node:os';
import { dirname, isAbsolute, join, resolve as resolvePath } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

/** Shipped package that owns the native PDFium rasterizer. */
const KIT_SPECIFIER = '@deepseek-ai/libreoffice-kit';

/** Package whose resolved location tells us where the DSH installation lives. */
const ANCHOR_SPECIFIERS = ['@deepseek-ai/dsh-tools', '@deepseek-ai/dsh-llm'];

/** Page-count probe used when a document is larger than the caller's cap. */
const PROBE_PAGE = 1;

let kitModule;

/**
 * Candidate module URLs for the rasterizer, most specific first.
 * @returns absolute file URLs to try with `import()`.
 */
function kitCandidates() {
  const candidates = [];
  for (const anchor of ANCHOR_SPECIFIERS) {
    try {
      const resolved = import.meta.resolve(anchor);
      if (!resolved.startsWith('file:')) continue;
      // <root>/node_modules/@deepseek-ai/<anchor>/lib/index.js -> <root>/node_modules
      const nodeModules = dirname(dirname(dirname(fileURLToPath(resolved))));
      candidates.push(join(nodeModules, KIT_SPECIFIER, 'lib/index.js'));
    } catch {
      // The anchor is not resolvable from here; try the next strategy.
    }
  }
  const roots = [];
  if (typeof process.resourcesPath === 'string' && process.resourcesPath.length > 0) {
    roots.push(join(process.resourcesPath, 'app', 'node_modules'));
  }
  if (typeof process.execPath === 'string' && process.execPath.length > 0) {
    roots.push(join(dirname(process.execPath), 'resources', 'app', 'node_modules'));
  }
  for (const root of roots) candidates.push(join(root, KIT_SPECIFIER, 'lib/index.js'));
  return candidates;
}

/**
 * Resolve the shipped rasterizer once per process.
 * @returns the `@deepseek-ai/libreoffice-kit` module namespace.
 * @throws when neither the resolver nor any fallback anchor can reach it.
 */
async function loadKit() {
  kitModule ??= (async () => {
    try {
      return await import(KIT_SPECIFIER);
    } catch (directError) {
      for (const candidate of kitCandidates()) {
        try {
          const url = candidate.startsWith('file:') ? candidate : pathToFileURL(candidate).href;
          return await import(url);
        } catch {
          // Try the next candidate.
        }
      }
      throw new Error(
        `dsh-pdf-pages: cannot load ${KIT_SPECIFIER}; the PDF rasterizer ships with the DSH installation, so this plugin must run inside a DSH profile`,
        { cause: directError },
      );
    }
  })();
  return kitModule;
}

/** Error codes `ConversionError` uses for the failures a caller can act on. */
const PAGE_COUNT_EXCEEDED = 'output-too-large';

/** The kit reports an out-of-range page as a bare `TypeError`. */
const PAGE_OUTSIDE_DOCUMENT = /outside the (\d+)-page document/u;

/**
 * Rewrite a rasterizer failure into the wording a caller can act on.
 * @param error - the failure raised by the kit.
 * @returns the failure to surface.
 */
function describeRenderFailure(error) {
  const message = error instanceof Error ? error.message : String(error);
  const outside = PAGE_OUTSIDE_DOCUMENT.exec(message);
  if (outside !== null) return new Error(`the requested page is outside this ${outside[1]}-page document; select pages within 1-${outside[1]}`);
  return error;
}

/**
 * Create the single long-lived PDF renderer for one plugin instance.
 * @param options - renderer limits.
 * @param options.timeoutMs - per-render deadline enforced by the kit.
 * @param options.maxPixels - per-page pixel cap.
 * @param options.maxDimension - per-page edge cap.
 * @returns the renderer handle.
 */
export function createPdfRenderer({ timeoutMs = 180_000, maxPixels = 16_777_216, maxDimension = 8192 } = {}) {
  let converterPromise;
  let disposed = false;

  const converter = () => {
    if (disposed) return Promise.reject(new Error('dsh-pdf-pages: the PDF renderer has been disposed'));
    converterPromise ??= loadKit().then(({ createConverter }) => createConverter({ timeoutMs }));
    return converterPromise;
  };

  /**
   * Render selected pages of one PDF into PNG bytes.
   * @param request - the render request.
   * @param request.inputPath - absolute path of an authorized PDF file.
   * @param request.pages - explicit one-based page numbers, or `'all'`.
   * @param request.dpi - raster density.
   * @param request.maxPages - page-count limit for this request.
   * @param request.signal - caller cancellation.
   * @returns the rasterized pages plus the document's total page count.
   */
  async function render({ inputPath, pages = 'all', dpi = 120, maxPages = 20, signal }) {
    const active = await converter();
    signal?.throwIfAborted();
    const outputDir = join(tmpdir(), `dsh-pdf-pages-${randomUUID()}`);
    try {
      const result = await active.renderImages(
        { inputPath, outputDir, pages, dpi, maxPages, maxPixels, maxDimension },
        signal,
      );
      const images = [];
      for (const image of result.images) {
        images.push({
          page: image.page,
          width: image.width,
          height: image.height,
          data: new Uint8Array(await readFile(image.path)),
        });
      }
      return {
        pageCount: result.pageCount,
        backend: result.backend,
        rasterEngine: result.rasterEngine,
        images,
      };
    } catch (error) {
      throw describeRenderFailure(error);
    } finally {
      await rm(outputDir, { recursive: true, force: true }).catch(() => {});
    }
  }

  /**
   * Render every page a caller asked for, capping an oversized document
   * instead of failing: `pages: 'all'` beyond the cap probes page 1 for the
   * document length and then renders the leading `maxPages` pages.
   * @param request - the render request, with `pages` omitted for "all".
   * @returns the rasterized pages plus the document's total page count.
   */
  async function renderAll({ inputPath, pages, dpi, maxPages, signal }) {
    if (pages !== undefined) return render({ inputPath, pages, dpi, maxPages, signal });
    try {
      return await render({ inputPath, pages: 'all', dpi, maxPages, signal });
    } catch (error) {
      if (error?.code !== PAGE_COUNT_EXCEEDED) throw error;
      const probe = await render({ inputPath, pages: [PROBE_PAGE], dpi, maxPages: 1, signal });
      const wanted = [];
      for (let page = 1; page <= Math.min(probe.pageCount, maxPages); page += 1) wanted.push(page);
      if (wanted.length === 1) return probe;
      return render({ inputPath, pages: wanted, dpi, maxPages, signal });
    }
  }

  async function dispose() {
    if (disposed) return;
    disposed = true;
    const pending = converterPromise;
    converterPromise = undefined;
    if (pending === undefined) return;
    try {
      const active = await pending;
      await active.dispose?.();
    } catch {
      // A converter that never finished starting owns nothing to release.
    }
  }

  return { render, renderAll, dispose };
}

/**
 * Parse a page selection such as `1-3,7` or `all` into explicit page numbers.
 * @param text - the raw user or model supplied selection.
 * @returns one-based page numbers, or undefined for "every page".
 * @throws when the selection is not a list of positive page numbers or ranges.
 */
export function parsePageSelection(text) {
  if (text === undefined || text === null) return undefined;
  const trimmed = String(text).trim();
  if (trimmed === '' || trimmed.toLowerCase() === 'all') return undefined;
  const pages = new Set();
  for (const part of trimmed.split(',')) {
    const piece = part.trim();
    if (piece === '') continue;
    const range = /^(\d+)\s*-\s*(\d+)$/u.exec(piece);
    if (range !== null) {
      const start = Number(range[1]);
      const end = Number(range[2]);
      if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 1 || end < start) {
        throw new Error(`invalid page range "${piece}"; use e.g. "1-3,7" with positive page numbers`);
      }
      if (end - start > 500) throw new Error(`page range "${piece}" is too large; select at most 500 pages at a time`);
      for (let page = start; page <= end; page += 1) pages.add(page);
      continue;
    }
    if (!/^\d+$/u.test(piece) || Number(piece) < 1) {
      throw new Error(`invalid page "${piece}"; use e.g. "1-3,7" with positive page numbers`);
    }
    pages.add(Number(piece));
  }
  return pages.size === 0 ? undefined : [...pages].sort((left, right) => left - right);
}

/**
 * Resolve a file path the same way the rasterizer needs it: absolute, existing,
 * and a regular file. Relative paths resolve against the caller's workspace.
 * @param path - the supplied path.
 * @returns the absolute path.
 */
export function absoluteInputPath(path) {
  return isAbsolute(path) ? path : resolvePath(process.cwd(), path);
}
