/**
 * Deck text outline for `dsh-slide-pages`.
 *
 * A deck is a ZIP of XML parts, so the per-slide text — what the author wrote:
 * titles, body runs, speaker notes — is readable with `fflate` alone, no engine
 * and no new dependency. The outline lets the model know what the slides it did
 * *not* see as images still say, and it is what makes an 8-slide automatic
 * conversion useful on a 40-slide deck.
 *
 * Only OOXML decks (`.pptx`/`.pptm`) carry this structure; a legacy binary
 * `.ppt` yields no outline and the caller simply omits it.
 * @module dsh-slide-pages/outline
 */

import { readFile } from 'node:fs/promises';

/** Named and numeric XML entities that appear in presentation text. */
const ENTITIES = new Map([
  ['amp', '&'],
  ['lt', '<'],
  ['gt', '>'],
  ['quot', '"'],
  ['apos', "'"],
  ['nbsp', ' '],
]);

/**
 * Decode the XML entities a text run can carry.
 * @param text - raw text between tags.
 * @returns the decoded text.
 */
function decodeEntities(text) {
  return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/giu, (whole, entity) => {
    if (entity.startsWith('#x') || entity.startsWith('#X')) return String.fromCodePoint(Number.parseInt(entity.slice(2), 16));
    if (entity.startsWith('#')) return String.fromCodePoint(Number.parseInt(entity.slice(1), 10));
    return ENTITIES.get(entity.toLowerCase()) ?? whole;
  });
}

/** Paragraphs that carry only decoration: bullets, page markers such as `0 / 10`. */
const DECORATION_ONLY = /^[^\p{L}\p{N}]+$/u;
const PAGE_MARKER = /^\d{1,3}\s*\/\s*\d{1,3}$/u;

/**
 * Extract every drawing text paragraph from one XML part.
 *
 * Text runs are concatenated inside their paragraph, because a single sentence
 * is routinely split across runs by formatting; paragraphs then stay separate,
 * which is what makes table-heavy slides readable instead of "0 / 10 / ● / …".
 * Pure decoration (bullet glyphs, `0 / 10` page markers) is dropped.
 * @param xml - the part contents.
 * @returns trimmed, non-empty paragraphs in document order.
 */
function textParagraphs(xml) {
  const paragraphs = [];
  const paragraphPattern = /<a:p(?:\s[^>]*)?>([\s\S]*?)<\/a:p>/gu;
  for (const paragraph of xml.matchAll(paragraphPattern)) {
    const runs = [...paragraph[1].matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/gu)]
      .map((run) => decodeEntities(run[1]));
    const text = runs.join('').replace(/\s+/gu, ' ').trim();
    if (text === '' || DECORATION_ONLY.test(text) || PAGE_MARKER.test(text)) continue;
    paragraphs.push(text);
  }
  if (paragraphs.length > 0) return paragraphs;

  // A part without `a:p` wrappers still carries runs; fall back to flat runs.
  const flat = [];
  for (const match of xml.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/gu)) {
    const text = decodeEntities(match[1]).replace(/\s+/gu, ' ').trim();
    if (text === '' || DECORATION_ONLY.test(text)) continue;
    flat.push(text);
  }
  return flat;
}

/**
 * Read the slide part paths of a deck in presentation order.
 * @param files - the deck's ZIP entries, keyed by part name.
 * @returns part names such as `slides/slide3.xml`, in slide order.
 */
function orderedSlideParts(files) {
  const presentation = files.get('ppt/presentation.xml');
  const relationships = files.get('ppt/_rels/presentation.xml.rels');
  if (presentation !== undefined && relationships !== undefined) {
    const targets = new Map();
    const relPattern = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*>/gu;
    for (const match of relationships.matchAll(relPattern)) targets.set(match[1], match[2]);
    const ordered = [];
    const idPattern = /<p:sldId\b[^>]*\br:id="([^"]+)"[^>]*>/gu;
    for (const match of presentation.matchAll(idPattern)) {
      const target = targets.get(match[1]);
      if (target === undefined) continue;
      const normalised = target.replace(/^\.\.?\//u, '').replace(/^\/?ppt\//u, '');
      if (!normalised.startsWith('slides/')) continue;
      ordered.push(normalised);
    }
    if (ordered.length > 0) return ordered;
  }
  return [...files.keys()]
    .filter((name) => /^slides\/slide\d+\.xml$/u.test(name))
    .sort((left, right) => Number(/(\d+)/u.exec(left)[1]) - Number(/(\d+)/u.exec(right)[1]));
}

/**
 * Resolve the notes part that belongs to one slide.
 * @param files - the deck's ZIP entries.
 * @param slidePart - the slide part name, e.g. `slides/slide7.xml`.
 * @returns text runs of the notes slide, or an empty list.
 */
function notesFor(files, slidePart) {
  const base = slidePart.replace(/^slides\//u, '');
  const rels = files.get(`ppt/slides/_rels/${base}.rels`);
  if (rels === undefined) return [];
  const match = /<Relationship\b[^>]*\bType="[^"]*\/notesSlide"[^>]*\bTarget="([^"]+)"[^>]*>/u.exec(rels)
    ?? /<Relationship\b[^>]*\bTarget="([^"]*notesSlide[^"]*)"[^>]*>/u.exec(rels);
  if (match === null) return [];
  const target = match[1].replace(/^\.\.?\//u, '').replace(/^\/?ppt\//u, '');
  const part = files.get(`ppt/${target}`);
  if (part === undefined) return [];
  // The notes slide repeats the slide image and its page number; keep prose only.
  return textParagraphs(part).filter((run) => !/^\d{1,4}$/u.test(run));
}

/**
 * Truncate one field without splitting a surrogate pair.
 * @param text - the field text.
 * @param limit - maximum characters.
 * @returns the text, ellipsised when it had to be cut.
 */
function clamp(text, limit) {
  if (text.length <= limit) return text;
  return `${text.slice(0, Math.max(0, limit - 1))}…`;
}

/**
 * Read the per-slide text outline of a deck.
 * @param inputPath - absolute path of the deck.
 * @param options - outline limits.
 * @param options.bodyChars - maximum body characters kept per slide.
 * @param options.notesChars - maximum speaker-note characters kept per slide.
 * @returns one entry per slide, or undefined for a non-OOXML deck or a password-protected one.
 */
export async function readDeckOutline(inputPath, { bodyChars = 300, notesChars = 160 } = {}) {
  let bytes;
  try {
    bytes = new Uint8Array(await readFile(inputPath));
  } catch {
    return undefined;
  }
  // A ZIP local-file header; anything else is the legacy binary format.
  if (bytes.byteLength < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) return undefined;

  let unzipSync;
  let strFromU8;
  try {
    ({ unzipSync, strFromU8 } = await import('fflate'));
  } catch {
    return undefined;
  }

  let entries;
  try {
    entries = unzipSync(bytes);
  } catch {
    return undefined;
  }
  const files = new Map();
  for (const [name, content] of Object.entries(entries)) {
    if (name.startsWith('ppt/slides/') || name.startsWith('ppt/notesSlides/') || name.startsWith('ppt/_rels/') || name.startsWith('ppt/slides/_rels/') || name === 'ppt/presentation.xml') {
      files.set(name, strFromU8(content));
    }
  }

  const parts = orderedSlideParts(files);
  if (parts.length === 0) return undefined;

  const outline = [];
  for (const [offset, part] of parts.entries()) {
    const xml = files.get(`ppt/${part}`);
    if (xml === undefined) continue;
    const paragraphs = textParagraphs(xml);
    const title = paragraphs[0] ?? '';
    const body = clamp(paragraphs.slice(1).join(' / '), bodyChars);
    const notes = clamp(notesFor(files, part).join(' / '), notesChars);
    outline.push({
      index: offset + 1,
      title,
      ...(body === '' ? {} : { body }),
      ...(notes === '' ? {} : { notes }),
    });
  }
  return outline.length === 0 ? undefined : outline;
}

/**
 * Render an outline as the compact text block a conversation message carries.
 * @param outline - the per-slide entries.
 * @param options - rendering limits.
 * @param options.totalChars - maximum characters across the whole block.
 * @param options.from - first slide number to include (one-based).
 * @param options.to - last slide number to include (one-based, inclusive).
 * @returns the text block, without a trailing newline.
 */
export function formatOutline(outline, { totalChars = 8000, from, to } = {}) {
  const entries = outline.filter((entry) => (from === undefined || entry.index >= from) && (to === undefined || entry.index <= to));
  const lines = [];
  let used = 0;
  let omitted = 0;
  for (const entry of entries) {
    const parts = [`第 ${entry.index} 张`];
    if (entry.title !== '') parts.push(`《${entry.title}》`);
    if (entry.body !== undefined) parts.push(`— ${entry.body}`);
    if (entry.notes !== undefined) parts.push(`（备注：${entry.notes}）`);
    const line = parts.join(' ');
    if (used + line.length > totalChars) {
      omitted += 1;
      continue;
    }
    lines.push(line);
    used += line.length + 1;
  }
  if (omitted > 0) lines.push(`…另有 ${omitted} 张未列出`);
  return lines.join('\n');
}
