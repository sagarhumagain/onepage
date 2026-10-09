/**
 * tables.test.js — merged cells and tables inside tables: how they are
 * written in the source, and that the page and the Word file draw the same
 * grid from it.
 *
 * Reading them out of pasted HTML needs a DOM and is covered by
 * test/word-paste.js.
 *
 * Run: npm test
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { inflateRawSync } from 'node:zlib';

import { parse } from '../src/parse.js';
import { render } from '../src/render.js';
import { applyMarks, collectUnits } from '../src/marks.js';
import { buildDocx } from '../src/to-docx.js';

const tableOf = (src) => parse(`Doc\n\n${src}`).find((b) => b.type === 'table');

/* --- merged cells -------------------------------------------------------- */

test('<< joins a cell to the one on its left', () => {
  const t = tableOf('| A | B | C |\n| --- | --- | --- |\n| wide | << | x |');
  assert.deepEqual(t.spans, { 'rows.0.0': { rows: 1, cols: 2 } });
  assert.deepEqual(t.covered, { 'rows.0.1': 'rows.0.0' });
  const html = render([t]);
  assert.ok(html.includes('<tr><td colspan="2">wide</td><td>x</td></tr>'), html);
});

test('^^ joins a cell to the one above', () => {
  const t = tableOf('| A | B |\n| --- | --- |\n| tall | 1 |\n| ^^ | 2 |\n| ^^ | 3 |');
  assert.deepEqual(t.spans, { 'rows.0.0': { rows: 3, cols: 1 } });
  const html = render([t]);
  assert.ok(html.includes('<tr><td rowspan="3">tall</td><td>1</td></tr><tr><td>2</td></tr><tr><td>3</td></tr>'), html);
});

test('a block of cells merges as one rectangle', () => {
  const t = tableOf('| A | B | C |\n| --- | --- | --- |\n| big | << | x |\n| ^^ | ^^ | y |\n| p | q | r |');
  assert.deepEqual(t.spans, { 'rows.0.0': { rows: 2, cols: 2 } });
  assert.deepEqual(Object.keys(t.covered).sort(), ['rows.0.1', 'rows.1.0', 'rows.1.1']);
  assert.ok(render([t]).includes('<td colspan="2" rowspan="2">big</td>'));
});

test('a header cell spans columns but never runs down into the body', () => {
  const t = tableOf('| Title | << |\n| --- | --- |\n| ^^ | b |');
  assert.deepEqual(t.spans, { 'head.0': { rows: 1, cols: 2 } });
  const html = render([t]);
  assert.ok(html.includes('<thead><tr><th colspan="2">Title</th></tr></thead>'), html);
  assert.ok(html.includes('<td>^^</td>'), 'a marker nothing reaches stays the text it is');
});

test('a marker in the first column, or under a cell it cannot square with, is text', () => {
  const t = tableOf('| A | B |\n| --- | --- |\n| << | x |');
  assert.equal(t.spans, undefined);
  assert.ok(render([t]).includes('<td>&lt;&lt;</td>'));
});

test('the cells a merge covers are not units, so nothing can mark them', () => {
  const units = collectUnits(parse('Doc\n\n| A | B |\n| --- | --- |\n| wide | << |\n| c | d |'));
  assert.deepEqual(units.map((u) => u.plain), ['Doc', 'A', 'B', 'wide', 'c', 'd']);
});

/* --- tables inside tables ---------------------------------------------- */

const NESTED = [
  '| Region | Detail |',
  '| --- | --- |',
  '| East | Cases by month:<br><br>\\| Jan \\| Feb \\|<br>\\| --- \\| --- \\|<br>\\| 12 \\| one<br\\>two \\|<br>{: widths="30 70"} |',
].join('\n');

test('a cell can hold a table of its own', () => {
  const t = tableOf(NESTED);
  const inner = t.cells['rows.0.1'];
  assert.deepEqual(inner.map((b) => b.type), ['p', 'table']);
  const nested = inner[1];
  assert.deepEqual(nested.head, ['Jan', 'Feb']);
  assert.deepEqual(nested.widths, [30, 70]);
  assert.deepEqual(nested.cells['rows.0.1'].map((b) => b.text), ['one\ntwo'], "the inner cell's own line break stays in it");
});

test('a table inside a cell renders inside the cell, and its cells are units', () => {
  const blocks = parse(`Doc\n\n${NESTED}`);
  applyMarks(blocks, []);
  const html = render(blocks, { editor: true });
  assert.match(html, /<td class="doc-cell"><p[^>]*>Cases by month:<\/p>\n<table data-widths="30 70">/);
  const units = collectUnits(blocks);
  assert.deepEqual(units.slice(-3).map((u) => u.plain), ['Feb', '12', 'one\ntwo']);
});

/* --- the Word file ---------------------------------------------------------- */

/** The text of one file in a .docx (a zip), read without a zip library. */
function zipEntry(buf, name) {
  for (let at = 0; at + 30 < buf.length; ) {
    if (buf.readUInt32LE(at) !== 0x04034b50) break;
    const method = buf.readUInt16LE(at + 8);
    const size = buf.readUInt32LE(at + 18);
    const nameLen = buf.readUInt16LE(at + 26);
    const extra = buf.readUInt16LE(at + 28);
    const entry = buf.toString('utf8', at + 30, at + 30 + nameLen);
    const data = buf.subarray(at + 30 + nameLen + extra, at + 30 + nameLen + extra + size);
    if (entry === name) return (method === 8 ? inflateRawSync(data) : data).toString('utf8');
    at += 30 + nameLen + extra + size;
  }
  return null;
}

async function documentXml(src) {
  const blocks = parse(src);
  applyMarks(blocks, []);
  const blob = await buildDocx({
    blocks,
    layout: { fontPt: 11, columns: 1, lineHeight: 1.38, marginMm: 15 },
    face: { docx: 'Calibri', label: 'Calibri' },
  });
  return zipEntry(Buffer.from(await blob.arrayBuffer()), 'word/document.xml');
}

test('Word draws the same merges: a grid span, and a vertical merge', async () => {
  const xml = await documentXml('Doc\n\n| A | B | C |\n| --- | --- | --- |\n| big | << | x |\n| ^^ | ^^ | y |');
  assert.match(xml, /<w:gridSpan w:val="2"\/>/);
  assert.match(xml, /<w:vMerge w:val="restart"\/>/);
  assert.match(xml, /<w:vMerge w:val="continue"\/>/);
  const rows = xml.match(/<w:tr>[\s\S]*?<\/w:tr>/g) || xml.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g);
  assert.equal(rows.length, 3);
  // Each row covers the three columns: the merged cell, or its continuation, and one more.
  for (const r of rows.slice(1)) assert.equal((r.match(/<w:tc>/g) || []).length, 2, r);
  assert.ok(!xml.includes('&lt;&lt;') && !xml.includes('^^'), 'the markers never reach the file');
});

test('Word nests a table inside a cell, and the cell still ends with a paragraph', async () => {
  const xml = await documentXml(`Doc\n\n${NESTED}`);
  assert.match(xml, /<w:tc>(?:(?!<\/w:tc>)[\s\S])*<w:tbl>[\s\S]*<\/w:tbl><w:p[ >/]/);
  assert.ok(xml.includes('Jan') && xml.includes('Feb'));
});
