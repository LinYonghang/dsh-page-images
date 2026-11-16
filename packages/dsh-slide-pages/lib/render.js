/**
 * Slide rasterization for `dsh-slide-pages`.
 *
 * LibreOffice's direct presentation raster is not faithful (it draws slide
 * content at roughly 1.5x and clips the right edge), so every deck goes through
 * two stages: `convert()` exports the deck to PDF — one PDF page per slide,
 * laid out exactly as PowerPoint exports it — and PDFium then rasterizes the
 * requested pages. The converted PDF is cached per deck so repeat requests and
 * "show me slides 12-15" calls do not pay the export cost again.
 *
 * Both stages come from `@deepseek-ai/libreoffice-kit`, which ships with the DSH
 * installation, so this plugin declares no install-time dependency on it.
 * @module dsh-slide-pages/render
 */

import { createHash, randomUUID } from 'node:crypto';
import { mkdir, readFile, readdir, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

/** Shipped package that owns the LibreOffice engine and the PDFium rasterizer. */
const KIT_SPECIFIER = '@deepseek-ai/libreoffice-kit';

/** Packages whose resolved location reveals where the DSH installation lives. */
const ANCHOR_SPECIFIERS = ['@deepseek-ai/dsh-tools', '@deepseek-ai/dsh-llm'];

/** Both stages share one deadline; this bounds the per-stage slice. */
const PROBE_PAGE = 1;

/** The kit reports an out-of-range page as a bare `TypeError`. */
const PAGE_OUTSIDE_DOCUMENT = /outside the (\d+)-page document/u;

/** Failure classes a caller can act on, in the wording a user needs. */
const FAILURE_TEXT = new Map([
  ['unavailable', '本机缺少随 DSH 发行的 LibreOffice 渲染引擎，无法把幻灯片转成图片'],
  ['invalid-document', '文件损坏、加密或不是可解析的演示文稿'],
  ['input-too-large', '文件超过引擎的 64 MiB 输入上限，请拆分后再试'],
  ['output-too-large', '生成的 PDF 或图片超过引擎的输出上限'],
  ['timeout', '转换超时，可提高 timeoutMs 或只转换部分页面'],
  ['unsupported-format', '引擎不支持这种演示文稿格式'],
]);

let kitModule;

/**
 * Candidate module URLs for the engine, most specific first.
 * @returns absolute paths to try with `import()`.
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
      // Anchor not resolvable here; try the next strategy.
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
 * Resolve the shipped engine once per process.
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
          return await import(pathToFileURL(candidate).href);
        } catch {
          // Try the next candidate.
        }
      }
      throw new Error(
        `dsh-slide-pages: cannot load ${KIT_SPECIFIER}; the slide renderer ships with the DSH installation, so this plugin must run inside a DSH profile`,
        { cause: directError },
      );
    }
  })();
  return kitModule;
}

/**
 * Rewrite an engine failure into the wording a caller can act on.
 * @param error - the failure raised by the kit.
 * @returns the failure to surface.
 */
function describeFailure(error) {
  const message = error instanceof Error ? error.message : String(error);
  const outside = PAGE_OUTSIDE_DOCUMENT.exec(message);
  if (outside !== null) {
    return new Error(`requested slide is outside this ${outside[1]}-slide deck; pick slides within 1-${outside[1]}`);
  }
  const code = typeof error?.code === 'string' ? error.code : undefined;
  const text = code === undefined ? undefined : FAILURE_TEXT.get(code);
  if (text === undefined) return error;
  return Object.assign(new Error(`${text}（${code}）`), { code });
}

/**
 * Create the single long-lived slide renderer for one plugin instance.
 * @param options - renderer limits and cache policy.
 * @param options.timeoutMs - per-stage deadline enforced by the kit.
 * @param options.cacheEntries - converted PDFs to keep for reuse.
 * @param options.maxPixels - per-slide pixel cap.
 * @param options.maxDimension - per-slide edge cap.
 * @returns the renderer handle.
 */
export function createSlideRenderer({ timeoutMs = 300_000, cacheEntries = 3, maxPixels = 16_777_216, maxDimension = 8192 } = {}) {
  let converterPromise;
  let disposed = false;
  let cacheDir;
  /** Converted PDF per deck key: `{ path, bytes }`, oldest first. */
  const pdfs = new Map();
  /** In-flight conversions per deck key, so concurrent requests share one export. */
  const converting = new Map();

  const converter = () => {
    if (disposed) return Promise.reject(new Error('dsh-slide-pages: the slide renderer has been disposed'));
    converterPromise ??= loadKit().then(({ createConverter }) => createConverter({ timeoutMs }));
    return converterPromise;
  };

  const ensureCacheDir = async () => {
    cacheDir ??= join(tmpdir(), `dsh-slide-pages-${randomUUID()}`);
    await mkdir(cacheDir, { recursive: true });
    return cacheDir;
  };

  /**
   * Delete cached PDFs that no live entry claims, which is what a crashed or
   * evicted conversion leaves behind.
   */
  async function sweepOrphans() {
    const dir = await ensureCacheDir();
    const live = new Set([...pdfs.values()].map((entry) => entry.path));
    for (const name of await readdir(dir).catch(() => [])) {
      const path = join(dir, name);
      if (live.has(path)) continue;
      await rm(path, { force: true }).catch(() => {});
    }
  }

  /**
   * Export one deck to PDF, reusing a cached conversion when available.
   * @param request - the export request.
   * @param request.inputPath - absolute path of an authorized deck.
   * @param request.key - content identity of the deck.
   * @param request.signal - caller cancellation.
   * @returns the PDF path plus the conversion's font diagnostics.
   */
  async function ensurePdf({ inputPath, key, signal }) {
    const cached = pdfs.get(key);
    if (cached !== undefined) return cached;
    const inflight = converting.get(key);
    if (inflight !== undefined) return inflight;

    const task = (async () => {
      const active = await converter();
      signal?.throwIfAborted();
      const dir = await ensureCacheDir();
      const outputPath = join(dir, `${createHash('sha256').update(key).digest('hex').slice(0, 32)}.pdf`);
      const existing = await stat(outputPath).catch(() => undefined);
      let missingFonts = [];
      if (existing === undefined || existing.size === 0) {
        await rm(outputPath, { force: true }).catch(() => {});
        const converted = await active.convert({ inputPath, outputPath }, signal);
        missingFonts = Array.isArray(converted?.missingFonts) ? [...converted.missingFonts] : [];
      }
      const info = await stat(outputPath);
      const entry = { path: outputPath, bytes: info.size, missingFonts };
      pdfs.set(key, entry);
      while (pdfs.size > cacheEntries) {
        const [oldestKey] = pdfs.keys();
        const oldest = pdfs.get(oldestKey);
        pdfs.delete(oldestKey);
        if (oldest !== undefined) await rm(oldest.path, { force: true }).catch(() => {});
      }
      await sweepOrphans();
      return entry;
    })();

    converting.set(key, task);
    try {
      return await task;
    } finally {
      converting.delete(key);
    }
  }

  /**
   * Rasterize selected pages of an already converted PDF.
   * @param request - the raster request.
   * @returns page images plus the document's page count.
   */
  async function rasterize({ pdfPath, pages, dpi, maxPages, signal }) {
    const active = await converter();
    signal?.throwIfAborted();
    const outputDir = join(tmpdir(), `dsh-slide-pages-png-${randomUUID()}`);
    try {
      const result = await active.renderImages(
        { inputPath: pdfPath, outputDir, pages, dpi, maxPages, maxPixels, maxDimension },
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
      return { slideCount: result.pageCount, backend: result.backend, rasterEngine: result.rasterEngine, images };
    } finally {
      await rm(outputDir, { recursive: true, force: true }).catch(() => {});
    }
  }

  /**
   * Convert a deck when needed and rasterize the requested slides.
   * @param request - the slide request.
   * @param request.inputPath - absolute path of an authorized deck.
   * @param request.key - content identity of the deck, used as the PDF cache key.
   * @param request.pages - explicit one-based slide numbers, or undefined for the leading `maxPages`.
   * @param request.dpi - raster density.
   * @param request.maxPages - slide-count cap for the "every slide" selection.
   * @param request.signal - caller cancellation.
   * @returns the rasterized slides plus the deck's slide count and font diagnostics.
   */
  async function render({ inputPath, key, pages, dpi, maxPages, signal }) {
    try {
      const pdf = await ensurePdf({ inputPath, key, signal });
      if (pages !== undefined) {
        const raster = await rasterize({ pdfPath: pdf.path, pages, dpi, maxPages: pages.length, signal });
        return { ...raster, missingFonts: pdf.missingFonts, pdfBytes: pdf.bytes };
      }
      // "Every slide" cannot know the deck length up front: the engine refuses a
      // selection longer than `maxPages`, so probe slide 1 first and keep it.
      const probe = await rasterize({ pdfPath: pdf.path, pages: [PROBE_PAGE], dpi, maxPages: 1, signal });
      const rest = [];
      for (let page = PROBE_PAGE + 1; page <= Math.min(probe.slideCount, maxPages); page += 1) rest.push(page);
      if (rest.length === 0) return { ...probe, missingFonts: pdf.missingFonts, pdfBytes: pdf.bytes };
      const tail = await rasterize({ pdfPath: pdf.path, pages: rest, dpi, maxPages: rest.length, signal });
      return {
        slideCount: probe.slideCount,
        backend: tail.backend,
        rasterEngine: tail.rasterEngine,
        missingFonts: pdf.missingFonts,
        pdfBytes: pdf.bytes,
        images: [...probe.images, ...tail.images],
      };
    } catch (error) {
      throw describeFailure(error);
    }
  }

  async function dispose() {
    if (disposed) return;
    disposed = true;
    const pending = converterPromise;
    converterPromise = undefined;
    pdfs.clear();
    if (pending !== undefined) {
      try {
        const active = await pending;
        await active.dispose?.();
      } catch {
        // A converter that never finished starting owns nothing to release.
      }
    }
    if (cacheDir !== undefined) await rm(cacheDir, { recursive: true, force: true }).catch(() => {});
  }

  return { render, dispose };
}

/**
 * Parse a slide selection such as `1-3,7` or `all` into explicit slide numbers.
 * @param text - the raw user or model supplied selection.
 * @returns one-based slide numbers, or undefined for "every slide".
 * @throws when the selection is not a list of positive slide numbers or ranges.
 */
export function parseSlideSelection(text) {
  if (text === undefined || text === null) return undefined;
  const trimmed = String(text).trim();
  if (trimmed === '' || trimmed.toLowerCase() === 'all') return undefined;
  const slides = new Set();
  for (const part of trimmed.split(',')) {
    const piece = part.trim();
    if (piece === '') continue;
    const range = /^(\d+)\s*-\s*(\d+)$/u.exec(piece);
    if (range !== null) {
      const start = Number(range[1]);
      const end = Number(range[2]);
      if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 1 || end < start) {
        throw new Error(`invalid slide range "${piece}"; use e.g. "1-3,7" with positive slide numbers`);
      }
      if (end - start > 500) throw new Error(`slide range "${piece}" is too large; select at most 500 slides at a time`);
      for (let slide = start; slide <= end; slide += 1) slides.add(slide);
      continue;
    }
    if (!/^\d+$/u.test(piece) || Number(piece) < 1) {
      throw new Error(`invalid slide "${piece}"; use e.g. "1-3,7" with positive slide numbers`);
    }
    slides.add(Number(piece));
  }
  return slides.size === 0 ? undefined : [...slides].sort((left, right) => left - right);
}

/**
 * Render the renderer's own failure wording for a caught error.
 * @param error - the caught failure.
 * @returns a one-line message suitable for a conversation message or tool error.
 */
export function slideFailureText(error) {
  const described = describeFailure(error);
  const message = described instanceof Error ? described.message : String(described);
  const code = typeof described?.code === 'string' ? `${described.code}: ` : '';
  return `${code}${message}`.slice(0, 600);
}
