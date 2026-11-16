/**
 * Timing probe: how long does one PDF page cost at a given density?
 * Run: node test/timing.mjs <absolute-pdf-path>
 */
import { readFile } from 'node:fs/promises';
import { createPdfRenderer } from '../lib/render.js';

const pdfPath = process.argv[2];
const bytes = new Uint8Array(await readFile(pdfPath));
console.log('input', pdfPath, bytes.byteLength, 'bytes');

const renderer = createPdfRenderer({ timeoutMs: 180_000 });

/** Render one selection and report its elapsed time. */
async function timed(label, request) {
  const started = Date.now();
  const result = await renderer.render(request);
  const ms = Date.now() - started;
  const kb = result.images.reduce((sum, image) => sum + image.data.byteLength, 0) / 1024;
  console.log(`${label}: ${ms} ms, ${result.images.length} page(s), ${Math.round(kb)} KiB total, doc pages ${result.pageCount}`);
  return { ms, result };
}

const probe = await timed('probe page 1 @110dpi', { inputPath: pdfPath, pages: [1], dpi: 110, maxPages: 1 });
const ten = await timed('pages 1-10 @110dpi', { inputPath: pdfPath, pages: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], dpi: 110, maxPages: 10 });
const ten96 = await timed('pages 1-10 @96dpi', { inputPath: pdfPath, pages: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], dpi: 96, maxPages: 10 });
await timed('pages 1-24 @110dpi', { inputPath: pdfPath, pages: Array.from({ length: Math.min(24, probe.result.pageCount) }, (_, i) => i + 1), dpi: 110, maxPages: 24 });
console.log('marginal per page @110:', Math.round((ten.ms - probe.ms) / 9), 'ms');
console.log('marginal per page @96 :', Math.round((ten96.ms - probe.ms) / 9), 'ms');
await renderer.dispose();
