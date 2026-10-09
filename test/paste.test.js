/**
 * paste.test.js — what a pasted Word document needs from the model: nested
 * and ticked lists, cells that hold paragraphs, tables without a header row,
 * hard line breaks, links, whole-unit formatting, and the pictures that only
 * Word's RTF flavour carries.
 *
 * The DOM half of a Word paste (VML text boxes, Word's list paragraphs) needs
 * a browser and is covered by test/word-paste.js.
 *
 * Run: npm test
 */

import test from 'node:test';
import assert from 'node:assert/strict';

import { parse } from '../src/parse.js';
import { render, padOf } from '../src/render.js';
import { tokenizeInline, plainText } from '../src/inline.js';
import { applyMarks, collectUnits } from '../src/marks.js';
import { extractRtfPictures } from '../src/rtf.js';
import { declarations, mapSymbols, toHex, toPt } from '../src/word.js';
import { marksFromRecords } from '../src/clipboard.js';

const TICK = String.fromCharCode(0x2713);

/* --- lists ---------------------------------------------------------------- */

test('indentation nests a list item under the one above it', () => {
  const [, list] = parse('Doc\n\n1. Protect yourself:\n  - Wear sleeves\n  - Use nets\n2. Prevent breeding:\n  - Remove habitats');
  assert.equal(list.type, 'ol');
  assert.deepEqual(list.items, ['Protect yourself:', 'Wear sleeves', 'Use nets', 'Prevent breeding:', 'Remove habitats']);
  assert.deepEqual(list.levels, [0, 1, 1, 0, 1]);
  assert.deepEqual(list.kinds, ['ol', 'ul', 'ul', 'ol', 'ul']);
});

test('a plain list keeps exactly the shape it always had', () => {
  const [, list] = parse('Doc\n\n- one\n- two');
  assert.deepEqual(list, { type: 'ul', items: ['one', 'two'] });
});

test('a tick is a bullet that is drawn as written', () => {
  const [, list] = parse(`Doc\n\n- Categories:\n  ${TICK} without warning signs\n  ${TICK} with warning signs`);
  assert.deepEqual(list.glyphs, [null, TICK, TICK]);
  const html = render([list]);
  assert.ok(html.includes(`<li data-glyph="${TICK}">without warning signs</li>`), html);
});

test('nested items render inside the item above them, numbering intact', () => {
  const html = render(parse('Doc\n\n1. Protect:\n  - Wear\n2. Prevent:\n  - Remove').slice(1));
  assert.equal(html, '<ol><li>Protect:<ul><li>Wear</li></ul></li><li>Prevent:<ul><li>Remove</li></ul></li></ol>');
});

test('every item is still one unit, however deep it sits', () => {
  const units = collectUnits(parse('Doc\n\n- a\n  - b\n    - c'));
  assert.deepEqual(units.map((u) => u.sub), ['text', 'items.0', 'items.1', 'items.2']);
});

/* --- hard line breaks ------------------------------------------------------ */

test('a trailing backslash breaks the line without ending the paragraph', () => {
  const blocks = parse('Government of Nepal\\\nMinistry of Health\\\nCurative Service Division\n\nNext.');
  assert.deepEqual(blocks.map((b) => b.type), ['p', 'p']);
  assert.equal(blocks[0].text, 'Government of Nepal\nMinistry of Health\nCurative Service Division');
});

test('a hard break is a real newline on the page, so offsets do not drift', () => {
  const html = render(parse('Doc\n\nfirst\\\nsecond'));
  assert.ok(html.includes('first<span class="doc-br">\n</span>second'), html);
});

/* --- tables ----------------------------------------------------------------- */

const TABLE = [
  'Doc',
  '',
  '|  |  |',
  '| --- | --- |',
  '| **Transmission** | - Bites of infected mosquitoes<br> |',
  '| **Clinical** | **Incubation:** 4-10 days.<br>Three phases:<br><br>1. Febrile<br>2. Critical |',
  '{: widths="14 86" border="#000000"}',
].join('\n');

test('an empty header row is a table without one', () => {
  const t = parse(TABLE).find((b) => b.type === 'table');
  assert.equal(t.headless, true);
  const html = render([t]);
  assert.ok(!html.includes('<thead>'));
  assert.ok(html.includes('data-head="0"'));
});

test('widths and a rule colour ride on the line after the table', () => {
  const t = parse(TABLE).find((b) => b.type === 'table');
  assert.deepEqual(t.widths, [14, 86]);
  assert.equal(t.border, '#000000');
  const html = render([t]);
  assert.ok(html.includes('<col style="width:14%"><col style="width:86%">'));
  assert.ok(html.includes('--table-rule:#000000'));
});

test('an attribute line that is not about a known key changes nothing', () => {
  const t = parse('Doc\n\n| A | B |\n|---|---|\n| 1 | 2 |\n{: border="url(evil)" widths="a b"}').find((b) => b.type === 'table');
  assert.equal(t.border, undefined);
  assert.equal(t.widths, undefined);
});

test('a cell with line breaks holds paragraphs and lists; a dash alone is just a dash', () => {
  const t = parse(TABLE).find((b) => b.type === 'table');
  assert.deepEqual(t.cells['rows.0.1'].map((b) => b.type), ['ul']);
  assert.deepEqual(t.cells['rows.1.1'].map((b) => b.type), ['p', 'ol']);
  assert.equal(t.cells['rows.1.1'][0].text, '**Incubation:** 4-10 days.\nThree phases:');
  assert.equal(t.cells['rows.0.0'], undefined, 'a one-line cell stays a plain cell');
  const dash = parse('Doc\n\n| A | B |\n|---|---|\n| - 5 | x |').find((b) => b.type === 'table');
  assert.equal(dash.cells, undefined, '"- 5" in a cell is a value, not a list');
});

test('initials at the start of a wrapped line do not start a list', () => {
  const blocks = parse('Doc\n\nIt was written by\nJ. K. Rowling in 1997.');
  assert.deepEqual(blocks.map((b) => b.type), ['h1', 'p']);
});

test('a hard break holds the next line in its paragraph, whatever it looks like', () => {
  const [, p, list] = parse('Doc\n\nHead Office\\\n- Floor 2\n\n- a\\\n  b\\\n  c');
  assert.equal(p.text, 'Head Office\n- Floor 2');
  assert.deepEqual(list.items, ['a\nb\nc']);
});

test('the paragraphs inside a cell are units of their own, at the table size', () => {
  const units = collectUnits(parse(TABLE));
  const inner = units.filter((u) => u.inTable && u.kind !== 'cell');
  assert.deepEqual(inner.map((u) => u.plain), ['Bites of infected mosquitoes', 'Incubation: 4-10 days.\nThree phases:', 'Febrile', 'Critical']);
  assert.ok(!units.some((u) => u.sub === 'head.0'), 'an empty header row has no units');
});

test('an escaped pipe stays inside its cell', () => {
  const t = parse('Doc\n\n| a \\| b | c |\n|---|---|\n| 1 | 2 |').find((b) => b.type === 'table');
  assert.deepEqual(t.head, ['a | b', 'c']);
});

/* --- links and whole-unit formatting ---------------------------------------- */

test('a link can have words that are not its address', () => {
  const t = tokenizeInline('Dengue [[Internet](https://who.int/x?utm)]. Geneva');
  assert.deepEqual(t.find((x) => x.href), { text: 'Internet', href: 'https://who.int/x?utm' });
  assert.equal(plainText('Dengue [[Internet](https://who.int/x?utm)]. Geneva'), 'Dengue [Internet]. Geneva');
});

const mark = (sig, start, end, fmt) => ({ sig, nth: 0, text: sig.slice(start, end), start, end, fmt });

test('alignment and vertical alignment belong to the cell, not its words', () => {
  const blocks = parse('Doc\n\n|  |  |\n| --- | --- |\n| Transmission | x |');
  applyMarks(blocks, [mark('Transmission', 0, 12, { band: '#9bbb59', align: 'center', valign: 'middle' })]);
  const html = render(blocks);
  assert.ok(html.includes('<td class="doc-band" style="background:#9bbb59;text-align:center;vertical-align:middle">Transmission</td>'), html);
});

test('a list item set small in every word has a small number too', () => {
  const blocks = parse('Doc\n\n1. World Health Organization.\n2. Pan American Health Organization.');
  applyMarks(blocks, [mark('World Health Organization.', 0, 26, { size: 0.78 })]);
  const html = render(blocks);
  assert.ok(html.includes('<li style="font-size:0.78em">World Health Organization.</li>'), html);
  assert.ok(!/<span[^>]*font-size/.test(html), 'and the words do not carry the size twice');
});

test('a heading set in the regular weight stays regular', () => {
  const blocks = parse('Dengue | Clinical Management\n\nBody.');
  applyMarks(blocks, [mark('Dengue | Clinical Management', 0, 28, { size: 1.64, bold: false })]);
  const html = render(blocks);
  assert.ok(html.includes('<span style="font-size:1.64em;font-weight:400">Dengue | Clinical Management</span>'), html);
});

test('a band shared by every paragraph of a cell shades the cell', () => {
  const blocks = parse('Doc\n\n|  |  |\n| --- | --- |\n| a | one<br>two |');
  applyMarks(blocks, [mark('one\ntwo', 0, 7, { band: '#9bbb59' })]);
  const html = render(blocks);
  assert.ok(html.includes('<td class="doc-cell" style="background:#9bbb59">'), html);
  assert.ok(!html.includes('<p class="doc-band"'), 'the paragraph does not draw a second box inside it');
});

/* --- records from a Word paste become marks ---------------------------------- */

test('a pasted run keeps its size relative to the body, whatever element it lands in', () => {
  const units = collectUnits(parse('Dengue | Clinical Management\n\nBody text.'));
  const records = [
    {
      plain: 'Dengue | Clinical Management',
      runs: [{ start: 0, end: 28, pt: 28, bold: false, italic: false }],
      unit: { band: '#538135', align: 'center', dark: true },
    },
    { plain: 'Body text.', runs: [{ start: 0, end: 10, pt: 9 }], unit: { align: 'left' } },
  ];
  const marks = marksFromRecords(units, records, 0, 9);
  const title = marks.find((m) => m.sig === 'Dengue | Clinical Management');
  // 28pt over a 9pt body, in a heading drawn at 1.9: 1.64 of the heading.
  assert.equal(title.fmt.size, 1.64);
  assert.equal(title.fmt.bold, false);
  assert.equal(title.fmt.fg, '#ffffff', 'unstated text on a dark fill is white, as Word draws it');
  assert.equal(title.fmt.band, '#538135');
  assert.equal(title.fmt.align, 'center');
  const body = marks.find((m) => m.sig === 'Body text.');
  assert.deepEqual(body.fmt, { align: 'left' }, 'body text at body size needs no size mark');
});

test('a record finds its own words, not an earlier paragraph that repeats them', () => {
  const units = collectUnits(parse('Doc\n\nSame words.\n\nSame words.'));
  const marks = marksFromRecords(units, [{ plain: 'Same words.', runs: [], unit: { band: '#ff0000' } }], 2, 9);
  assert.equal(marks.length, 1);
  assert.equal(marks[0].nth, 1);
});

/* --- RTF pictures --------------------------------------------------------------- */

// A 1x1 PNG.
const PNG =
  '89504e470d0a1a0a0000000d4948445200000001000000010806000000' +
  '1f15c4890000000d49444154789c6360000002000154a24f5d0000000049454e44ae426082';

const pict = (extra = '') =>
  `{\\pict\\picscalex50\\picscaley50\\picw100\\pich100\\picwgoal400\\pichgoal200\\pngblip\\bliptag255{\\*\\blipuid aa11}${extra}${PNG}}`;

test('pictures come out of RTF as PNG, at the size Word showed them', () => {
  const [pic] = extractRtfPictures(`{\\rtf1 {\\shp{\\sp{\\sn pib}{\\sv ${pict()}}}}}`);
  assert.equal(pic.mime, 'image/png');
  assert.equal(pic.widthPt, 10); // 400 twips at 50%
  assert.equal(pic.heightPt, 5);
  assert.equal(pic.pxWidth, 1);
  assert.equal(Buffer.from(pic.base64, 'base64').toString('hex'), PNG);
});

test('the legacy copy of a picture and a repeat of the same picture are skipped', () => {
  const wmf = `{\\nonshppict{\\pict\\wmetafile8\\picw1\\pich1 0102}}`;
  const pics = extractRtfPictures(`{\\rtf1 ${pict()} ${wmf} ${pict()}}`);
  assert.equal(pics.length, 1);
});

/* --- Word's CSS -------------------------------------------------------------------- */

test("Word's lengths, colours and symbol fonts read as what they draw", () => {
  assert.equal(toPt('9.0pt'), 9);
  assert.equal(toPt('12px'), 9);
  assert.equal(Math.round(toPt('1cm') * 100) / 100, 28.35);
  assert.equal(toHex('#9BBB59'), '#9bbb59');
  assert.equal(toHex('red'), '#ff0000');
  assert.equal(toHex('auto'), null);
  assert.equal(mapSymbols('ü', 'Wingdings'), TICK);
  assert.equal(mapSymbols('', 'Calibri'), '→');
  assert.equal(declarations("font-size:9.0pt;\n  color:black").get('color'), 'black');
});

/* --- highlight padding ------------------------------------------------------- */

test('a highlight carries its own padding, mostly sideways', () => {
  const blocks = parse('Doc\n\nBody text.');
  applyMarks(blocks, [mark('Body text.', 0, 4, { bg: '#ffe08a', pad: 6 })]);
  const html = render(blocks);
  // Vertical padding paints over the neighbouring lines rather than moving
  // them, so it is kept a fraction of the horizontal.
  assert.ok(html.includes('<mark style="background:#ffe08a;padding:0.21em 0.6em">Body</mark>'), html);
});

test('padding without a highlight, or out of range, never reaches the page', () => {
  const blocks = parse('Doc\n\nBody text.');
  applyMarks(blocks, [mark('Body text.', 0, 4, { pad: 6 })]);
  assert.ok(!render(blocks).includes('padding'));
  assert.deepEqual(padOf({ pad: 99 }), { h: 2, v: 0.7 });
  assert.equal(padOf({}), null);
});
