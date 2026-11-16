/**
 * Timing probe for the slide renderer: conversion cost versus raster cost.
 * Run: node test/timing.mjs <absolute-pptx-path>
 */
import { randomUUID } from 'node:crypto';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import z from '@deepseek-ai/schemastery';
import { Config } from '../lib/index.js';
import { createSlideRenderer } from '../lib/render.js';

const deck = process.argv[2];
const config = Config({});
const renderer = createSlideRenderer({ timeoutMs: 300_000 });

/** Render one selection and report its elapsed time. */
async function timed(label, request) {
  const started = Date.now();
  const result = await renderer.render({ inputPath: deck, key: `timing:${deck}`, dpi: config.dpi, ...request });
  const ms = Date.now() - started;
  const kb = result.images.reduce((sum, image) => sum + image.data.byteLength, 0) / 1024;
  console.log(`${label}: ${ms} ms (${result.images.length} slide(s), ${Math.round(kb)} KiB), deck slides ${result.slideCount}`);
  return ms;
}

const first = await timed('slide 1 (cold: convert + raster)', { pages: [1], maxPages: 1 });
const ten = await timed('slides 1-10 (warm: raster only)', { pages: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], maxPages: 10 });
const all = await timed('every slide up to 24', { pages: undefined, maxPages: 24 });
console.log('marginal per slide:', Math.round((ten - first) / 9), 'ms');
console.log('probe-first "every slide" total:', all, 'ms');

await renderer.dispose();
const scratch = await mkdtemp(join(tmpdir(), 'slide-timing-'));
await rm(scratch, { recursive: true, force: true });
console.log('timing probe done');
