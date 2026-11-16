/**
 * Unit test for the deck outline parser: runs on a synthetic OOXML deck so the
 * speaker-notes path is covered without depending on a deck that happens to
 * carry notes.
 *
 * Run: node test/outline.mjs [<absolute-pptx-path> ...]
 */
import assert from 'node:assert/strict';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { zipSync, strToU8 } from 'fflate';
import { formatOutline, readDeckOutline } from '../lib/outline.js';

/** Assemble a two-slide deck whose second slide carries speaker notes. */
function syntheticDeck() {
  const parts = {
    'ppt/presentation.xml': `<?xml version="1.0"?><p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': `<?xml version="1.0"?><Relationships><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>`,
    'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld><p:cSld><p:spTree><p:sp><p:txBody>
      <a:p><a:r><a:t>封面标题</a:t></a:r></a:p>
      <a:p><a:r><a:t>副标题</a:t></a:r><a:r><a:t>折行续写</a:t></a:r></a:p>
      <a:p><a:r><a:t>0 / 12</a:t></a:r></a:p>
      <a:p><a:r><a:t>●</a:t></a:r></a:p>
      </p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
    'ppt/slides/slide2.xml': `<?xml version="1.0"?><p:sld><p:cSld><p:spTree><p:sp><p:txBody>
      <a:p><a:r><a:t>第二页 &amp; 结论</a:t></a:r></a:p>
      <a:p><a:r><a:t>要点一</a:t></a:r></a:p>
      <a:p><a:r><a:t>要点二</a:t></a:r></a:p>
      </p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
    'ppt/slides/_rels/slide2.xml.rels': `<?xml version="1.0"?><Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide2.xml"/></Relationships>`,
    'ppt/notesSlides/notesSlide2.xml': `<?xml version="1.0"?><p:notes><p:cSld><p:spTree><p:sp><p:txBody>
      <a:p><a:r><a:t>2</a:t></a:r></a:p>
      <a:p><a:r><a:t>讲这一页时先回顾上一页</a:t></a:r></a:p>
      <a:p><a:r><a:t>再给出结论</a:t></a:r></a:p>
      </p:txBody></p:sp></p:spTree></p:cSld></p:notes>`,
  };
  const entries = {};
  for (const [name, text] of Object.entries(parts)) entries[name] = strToU8(text);
  return zipSync(entries);
}

const dir = await mkdtemp(join(tmpdir(), 'slide-outline-'));
const synthetic = join(dir, 'synthetic.pptx');
await writeFile(synthetic, syntheticDeck());

const outline = await readDeckOutline(synthetic, { bodyChars: 300, notesChars: 160 });
assert.ok(outline !== undefined, 'a synthetic OOXML deck must produce an outline');
assert.equal(outline.length, 2, 'both slides must be listed in presentation order');

assert.equal(outline[0].title, '封面标题');
assert.equal(outline[0].body, '副标题折行续写', 'runs of one paragraph must be joined, not separated');
assert.equal(outline[0].notes, undefined);
assert.ok(!outline[0].body.includes('0 / 12'), 'page markers must be dropped');
console.log('ok: title/body grouping', JSON.stringify(outline[0]));

assert.equal(outline[1].title, '第二页 & 结论', 'XML entities must be decoded');
assert.equal(outline[1].body, '要点一 / 要点二', 'paragraphs must stay separate');
assert.equal(outline[1].notes, '讲这一页时先回顾上一页 / 再给出结论', 'notes must drop the page-number placeholder');
console.log('ok: notes', JSON.stringify(outline[1].notes));

const text = formatOutline(outline, { totalChars: 2000 });
assert.match(text, /第 1 张 《封面标题》/u);
assert.match(text, /备注：讲这一页时/u);
console.log('ok: formatting\n' + text);

const truncated = formatOutline(outline, { totalChars: 20 });
assert.match(truncated, /未列出/u, 'an exhausted budget must say how many slides were dropped');
console.log('ok: budget', JSON.stringify(truncated));

const legacyText = join(dir, 'legacy.ppt');
await writeFile(legacyText, 'not a zip at all');
assert.equal(await readDeckOutline(legacyText), undefined, 'a binary/legacy deck must yield no outline');
console.log('ok: legacy deck has no outline');

for (const deck of process.argv.slice(2)) {
  const real = await readDeckOutline(deck, {});
  assert.ok(real !== undefined, `${deck} must produce an outline`);
  const withNotes = real.filter((entry) => entry.notes !== undefined).length;
  console.log(`ok: ${deck.split(/[\\/]/u).at(-1)} -> ${real.length} slides, ${withNotes} with speaker notes`);
}

console.log('outline test passed');
