/**
 * unit.test.js — the pure logic: parsing, inline tokenising, and the
 * invariant that document.css and the DOCX exporter still agree.
 *
 * Run: npm test
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

import { parse, normalize, wordCount } from '../src/parse.js';
import { render, escapeHtml } from '../src/render.js';
import { tokenizeInline, tokenizeMarked, plainText } from '../src/inline.js';
import { applyMarks, collectUnits, formatRange, clearRange, allHave, fmtAt, stepSize } from '../src/marks.js';
import { SCALE, COLUMN_GAP_MM, TYPEFACES, resolveTypeface, mmToTwip, ptToHalfPoint, ptToTwip, A4_TWIP } from '../src/scale.js';

const SRC = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'src');
const types = (blocks) => blocks.map((b) => b.type);

/* --- parsing ----------------------------------------------------------- */

test('first line becomes the title', () => {
  const b = parse('Quarterly Report\n\nSome body text follows here.');
  assert.equal(b[0].type, 'h1');
  assert.equal(b[0].text, 'Quarterly Report');
});

test('ALL CAPS lines are headings', () => {
  const b = parse('Title Here\n\nEXECUTIVE SUMMARY\nBody text.');
  assert.deepEqual(types(b), ['h1', 'h2', 'p']);
});

test('a lone numbered line followed by prose is a heading, not a list', () => {
  const b = parse('Doc\n\n1. Methodology\nWe used a survey.\n\n2. Results\nCoverage rose.');
  assert.deepEqual(types(b), ['h1', 'h2', 'p', 'h2', 'p']);
  assert.equal(b[1].text, '1. Methodology');
});

test('a real numbered sequence is a list', () => {
  const b = parse('Doc\n\nSteps:\n1. First\n2. Second\n3. Third');
  assert.deepEqual(types(b), ['h1', 'p', 'ol']);
  assert.deepEqual(b[2].items, ['First', 'Second', 'Third']);
});

test('bullet continuation lines join the previous item', () => {
  const b = parse('Doc\n\n- one item\n  wrapped onward\n- two');
  const ul = b.find((x) => x.type === 'ul');
  assert.deepEqual(ul.items, ['one item wrapped onward', 'two']);
});

test('Word-style bullet glyphs are recognised', () => {
  const bullet = String.fromCharCode(0x2022);
  const b = parse(`Doc\n\n${bullet} alpha\n${bullet} beta`);
  const ul = b.find((x) => x.type === 'ul');
  assert.deepEqual(ul.items, ['alpha', 'beta']);
});

test('pipe tables parse into head and rows', () => {
  const b = parse('Doc\n\n| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |');
  const t = b.find((x) => x.type === 'table');
  assert.deepEqual(t.head, ['A', 'B']);
  assert.deepEqual(t.rows, [['1', '2'], ['3', '4']]);
});

test('blockquotes and rules and fenced code', () => {
  const b = parse('Doc\n\n> quoted line\n\n---\n\n```\ncode here\n```');
  assert.deepEqual(types(b), ['h1', 'quote', 'hr', 'code']);
});

test('normalize collapses NBSP, zero-width and CRLF', () => {
  const nbsp = String.fromCharCode(0x00a0);
  const zwsp = String.fromCharCode(0x200b);
  // NBSP becomes a real space; a zero-width character is removed outright
  // rather than becoming a space, which would split a word in two.
  assert.equal(normalize(`a${nbsp}b${zwsp}c\r\nd`), 'a bc\nd');
  assert.equal(normalize('a\r\n\r\n\r\n\r\nb'), 'a\n\nb');
});

test('word count ignores markup noise', () => {
  assert.equal(wordCount('one two   three\n\nfour'), 4);
});

test('empty input yields no blocks', () => {
  assert.deepEqual(parse(''), []);
  assert.deepEqual(parse('   \n  \n'), []);
});

/* --- inline + rendering ------------------------------------------------- */

test('inline tokeniser marks bold, italic, code and links', () => {
  const t = tokenizeInline('a **bold** and *it* and `c` and https://x.test/p');
  assert.ok(t.some((x) => x.bold && x.text === 'bold'));
  assert.ok(t.some((x) => x.italic && x.text === 'it'));
  assert.ok(t.some((x) => x.code && x.text === 'c'));
  assert.ok(t.some((x) => x.href === 'https://x.test/p'));
  assert.equal(plainText('a **bold** end'), 'a bold end');
});

test('markup in user text is escaped, never rendered', () => {
  const html = render(parse('Doc\n\n<script>alert(1)</script> and <b>x</b>'));
  assert.ok(!html.includes('<script>'));
  assert.ok(html.includes('&lt;script&gt;'));
  assert.equal(escapeHtml(`<&">'`), '&lt;&amp;&quot;&gt;&#39;');
});

test('only http and mailto links are emitted', () => {
  const t = tokenizeInline('mail me at a@b.test');
  assert.ok(t.some((x) => x.href === 'mailto:a@b.test'));
  // A javascript: string is not a link pattern at all, so it stays plain text.
  const html = render(parse('Doc\n\njavascript:alert(1)'));
  assert.ok(!html.includes('href="javascript'));
});

test('ordered list start attribute survives rendering', () => {
  const html = render(parse('Doc\n\n5. five\n6. six'));
  assert.ok(html.includes('<ol start="5">'));
});

/* --- images ------------------------------------------------------------- */

test('an image reference is a block of its own, with text either side', () => {
  const b = parse('Report\n\nBefore the picture.\n\n[image:ab12]\n\nAfter the picture.');
  assert.deepEqual(types(b), ['h1', 'p', 'image', 'p']);
  assert.equal(b[2].id, 'ab12');
});

test('an image reference breaks a paragraph and a list rather than joining them', () => {
  const para = parse('Doc\n\nOne line of prose.\n[image:zz99]\nAnother line of prose.');
  assert.deepEqual(types(para), ['h1', 'p', 'image', 'p']);

  const list = parse('Doc\n\n- alpha\n[image:zz99]\n- beta');
  assert.deepEqual(types(list), ['h1', 'ul', 'image', 'ul']);
});

test('text that merely looks like a reference stays text', () => {
  assert.deepEqual(types(parse('Doc\n\n[image:has spaces]')), ['h1', 'p']);
  assert.deepEqual(types(parse('Doc\n\nsee [image:ab12] here')), ['h1', 'p']);
});

test('an image renders as a sized figure, not a full-width band', () => {
  const html = render([{ type: 'image', id: 'ab12', src: 'data:image/png;base64,AAA', width: 45, align: 'left', wrap: true }]);
  assert.ok(html.includes('class="doc-image"'));
  assert.ok(html.includes('--img-w:45'));
  assert.ok(html.includes('data-align="left"'));
  assert.ok(html.includes('data-wrap="1"'));
});

test('image widths are clamped and sources are restricted to what we produced', () => {
  const wide = render([{ type: 'image', src: 'data:image/png;base64,AAA', width: 900 }]);
  assert.ok(wide.includes('--img-w:100'));
  const narrow = render([{ type: 'image', src: 'data:image/png;base64,AAA', width: 1 }]);
  assert.ok(narrow.includes('--img-w:10'));
  // A javascript: or file: src is not rendered at all.
  assert.equal(render([{ type: 'image', src: 'javascript:alert(1)' }]), '');
});

/* --- formatting a run of words ------------------------------------------ */

const mark = (sig, text, start, end, fmt) => ({ sig, nth: 0, text, start, end, fmt });

test('units are every markable run of text, in reading order', () => {
  const blocks = parse('Title\n\nBody text.\n\n- one\n- two\n\n| A | B |\n|---|---|\n| 1 | 2 |');
  const units = collectUnits(blocks);
  assert.deepEqual(units.map((u) => u.plain), ['Title', 'Body text.', 'one', 'two', 'A', 'B', '1', '2']);
  assert.deepEqual(units.map((u) => u.sub), ['text', 'text', 'items.0', 'items.1', 'head.0', 'head.1', 'rows.0.0', 'rows.0.1']);
});

test('a mark is offsets into the text the reader sees, not into the source', () => {
  const blocks = parse('Doc\n\nThe **bold** word.');
  // "The bold word." — colour "bold", which sits behind ** markers.
  applyMarks(blocks, [mark('The bold word.', 'bold', 4, 8, { bg: '#ffe08a' })]);
  assert.ok(render(blocks).includes('<mark style="background:#ffe08a"><strong>bold</strong></mark>'));
});

test('every property a run can carry reaches the page', () => {
  const blocks = parse('Doc\n\nBody text.');
  applyMarks(blocks, [mark('Body text.', 'Body', 0, 4, { bg: '#ffe08a', fg: '#b42318', size: 1.35, bold: true, italic: true })]);
  const html = render(blocks);
  assert.ok(html.includes('background:#ffe08a'));
  assert.ok(html.includes('color:#b42318'));
  assert.ok(html.includes('font-size:1.35em'), 'size must be relative, so the fitter still owns absolute size');
  assert.ok(html.includes('font-weight:700'));
  assert.ok(html.includes('font-style:italic'));
});

test('a run with colour but no background is a span, not a highlight', () => {
  const blocks = parse('Doc\n\nBody text.');
  applyMarks(blocks, [mark('Body text.', 'Body', 0, 4, { fg: '#b42318' })]);
  const html = render(blocks);
  assert.ok(html.includes('<span style="color:#b42318">Body</span>'), html);
  assert.ok(!html.includes('<mark'));
});

test('a mark follows its words when the text around them is edited', () => {
  const m = mark('Second line.', 'Second', 0, 6, { bg: '#ffe08a' });
  const blocks = parse('Doc\n\nA new first line.\n\nSecond line, now longer.');
  applyMarks(blocks, [m]);
  assert.ok(render(blocks).includes('<mark style="background:#ffe08a">Second</mark>'));

  // And a mark whose words are gone is dropped, never moved onto others.
  const gone = parse('Doc\n\nNothing of the sort.');
  applyMarks(gone, [m]);
  assert.ok(!render(gone).includes('<mark'));
});

test('formatting merges property by property, and null takes one off', () => {
  let m = formatRange([], 2, 9, { bg: '#ffe08a' }, 20);
  m = formatRange(m, 5, 12, { bold: true }, 20);
  assert.deepEqual(m, [
    { start: 2, end: 5, fmt: { bg: '#ffe08a' } },
    { start: 5, end: 9, fmt: { bg: '#ffe08a', bold: true } },
    { start: 9, end: 12, fmt: { bold: true } },
  ]);

  assert.ok(allHave(m, 5, 9, 'bold', true));
  assert.ok(!allHave(m, 2, 9, 'bold', true));
  assert.equal(fmtAt(m, 6).bg, '#ffe08a');

  const unbolded = formatRange(m, 0, 20, { bold: null }, 20);
  assert.deepEqual(unbolded, [{ start: 2, end: 9, fmt: { bg: '#ffe08a' } }]);
  assert.deepEqual(clearRange(m, 0, 20, 20), []);
});

test('a later colour wins where two marks overlap', () => {
  const m = formatRange(formatRange([], 0, 6, { bg: '#ffe08a' }, 10), 3, 10, { bg: '#a5d8ff' }, 10);
  assert.deepEqual(m, [
    { start: 0, end: 3, fmt: { bg: '#ffe08a' } },
    { start: 3, end: 10, fmt: { bg: '#a5d8ff' } },
  ]);
});

test('size steps walk a ladder rather than compounding', () => {
  assert.equal(stepSize(1, 1), 1.15);
  assert.equal(stepSize(1.15, 1), 1.35);
  assert.equal(stepSize(1, -1), 0.9);
  assert.equal(stepSize(2.25, 1), 2.25, 'the ladder has a top');
  assert.equal(stepSize(0.6, -1), 0.6, 'and a bottom');
});

test('marked tokens are cut at the boundary, keeping their emphasis', () => {
  const t = tokenizeMarked('a **bold** end', [{ start: 4, end: 8, fmt: { bg: '#fff000' } }]);
  assert.deepEqual(
    t.map((x) => [x.text, x.fmt ? x.fmt.bg : null, Boolean(x.bold)]),
    [['a ', null, false], ['bo', null, true], ['ld', '#fff000', true], [' e', '#fff000', false], ['nd', null, false]]
  );
});

test('a band is the line\'s own background, not a highlight around the words', () => {
  const blocks = parse('Doc\n\nEXECUTIVE SUMMARY\n\nBody text.');
  applyMarks(blocks, [mark('EXECUTIVE SUMMARY', 'EXECUTIVE SUMMARY', 0, 17, { band: '#ffe08a' })]);
  const html = render(blocks);
  assert.ok(html.includes('<h2 class="doc-band" style="background:#ffe08a">EXECUTIVE SUMMARY</h2>'), html);
  assert.ok(!html.includes('<mark'), 'a band must not also wrap the text');
});

test('a band and a highlight can sit on the same line', () => {
  const blocks = parse('Doc\n\nBody text here.');
  applyMarks(blocks, [
    mark('Body text here.', 'Body text here.', 0, 15, { band: '#eef2f7' }),
    { sig: 'Body text here.', nth: 0, text: 'text', start: 5, end: 9, fmt: { bg: '#ffe08a' } },
  ]);
  const html = render(blocks);
  assert.ok(html.includes('class="doc-band" style="background:#eef2f7"'));
  assert.ok(html.includes('<mark style="background:#ffe08a">text</mark>'), html);
});

test('two marks on one paragraph both survive re-anchoring', () => {
  // Each range is stored as its own mark against the same unit, so anchoring
  // must never treat a unit as claimed by the first mark that lands on it.
  const blocks = parse('Doc\n\nAlpha beta gamma.');
  applyMarks(blocks, [
    { sig: 'Alpha beta gamma.', nth: 0, text: 'Alpha', start: 0, end: 5, fmt: { bg: '#ffe08a' } },
    { sig: 'Alpha beta gamma.', nth: 0, text: 'gamma', start: 11, end: 16, fmt: { bg: '#a5d8ff' } },
  ]);
  const html = render(blocks);
  assert.ok(html.includes('<mark style="background:#ffe08a">Alpha</mark>'), html);
  assert.ok(html.includes('<mark style="background:#a5d8ff">gamma</mark>'), html);
});

test('a colour that is not a plain hex never reaches the style attribute', () => {
  const blocks = parse('Doc\n\nBody text.');
  applyMarks(blocks, [mark('Body text.', 'Body', 0, 4, { bg: 'url(evil)', fg: 'expression(x)' })]);
  const html = render(blocks);
  assert.ok(!html.includes('evil'));
  assert.ok(!html.includes('expression'));
  assert.ok(!html.includes('<mark'));
});

/* --- unit conversions --------------------------------------------------- */

test('A4 uses the twip literals Word itself writes', () => {
  // docx 9.7.1's convertMillimetersToTwip truncates to 11905 x 16837, which is
  // a twip short of Word's own A4 and can flip a borderline layout to 2 pages.
  assert.equal(A4_TWIP.width, 11906);
  assert.equal(A4_TWIP.height, 16838);
});

test('mm and pt conversions round rather than truncate', () => {
  assert.equal(mmToTwip(15), 850); // 15/25.4*1440 = 850.39
  assert.equal(mmToTwip(25), 1417); // 1417.32
  assert.equal(ptToTwip(9.5), 190);
  assert.equal(ptToHalfPoint(9.5), 19);
});

test('typeface fallback resolves to an installed family', () => {
  const none = () => false;
  const all = () => true;
  assert.equal(resolveTypeface('sans', all).docx, 'Calibri');
  // Calibri missing -> Arial, which has no further fallback.
  assert.equal(resolveTypeface('sans', none).docx, 'Arial');
  assert.equal(resolveTypeface('serif', none).docx, 'Times New Roman');
});

/* --- the invariant that keeps Word matching the preview ------------------ */

const css = readFileSync(path.join(SRC, 'document.css'), 'utf8');

/** Pull one rule block out of document.css. */
function rule(selector) {
  const m = css.match(new RegExp(`${selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*\\{([^}]*)\\}`));
  assert.ok(m, `document.css is missing a rule for ${selector}`);
  return m[1];
}

const prop = (body, name) => {
  const m = body.match(new RegExp(`(?:^|;|\\n)\\s*${name}\\s*:\\s*([^;]+)`));
  return m ? m[1].trim() : null;
};

/** The bottom value of a `margin: a b c` / `margin: a b` / `margin: a` shorthand. */
function marginParts(body) {
  const v = prop(body, 'margin');
  if (!v) return null;
  const p = v.split(/\s+/);
  if (p.length === 1) return { top: p[0], bottom: p[0] };
  if (p.length === 2) return { top: p[0], bottom: p[0] };
  return { top: p[0], bottom: p[2] };
}

const em = (v) => (v === '0' ? 0 : Number(String(v).replace('em', '')));

test('document.css font sizes match the DOCX scale', () => {
  for (const [key, selector] of [
    ['h1', '.page-body h1'],
    ['h2', '.page-body h2'],
    ['h3', '.page-body h3'],
  ]) {
    assert.equal(em(prop(rule(selector), 'font-size')), SCALE[key].size, `${key} font-size`);
    assert.equal(Number(prop(rule(selector), 'line-height')), SCALE[key].lineHeight, `${key} line-height`);
  }
  assert.equal(em(prop(rule('.page-body pre'), 'font-size')), SCALE.code.size);
  assert.equal(Number(prop(rule('.page-body pre'), 'line-height')), SCALE.code.lineHeight);
  assert.equal(em(prop(rule('.page-body table'), 'font-size')), SCALE.table.size);
});

test('document.css margins match the DOCX scale', () => {
  for (const [key, selector] of [
    ['h1', '.page-body h1'],
    ['h2', '.page-body h2'],
    ['h3', '.page-body h3'],
    ['p', '.page-body p'],
    ['li', '.page-body li'],
    ['quote', '.page-body blockquote'],
  ]) {
    const m = marginParts(rule(selector));
    assert.equal(em(m.bottom), SCALE[key].after, `${key} margin-bottom`);
    assert.equal(em(m.top), SCALE[key].before, `${key} margin-top`);
  }
});

test('a band keeps banded and unbanded lines on the same left edge', () => {
  // The negative margins have to cancel the padding exactly, or every banded
  // line would sit indented from the ones above and below it.
  const band = rule('.page-body .doc-band');
  const inset = prop(band, 'padding').replace(/^\S+\s+/, '');
  assert.equal(prop(band, 'margin-left'), `calc(-1 * ${inset})`);
  assert.equal(prop(band, 'margin-right'), `calc(-1 * ${inset})`);
  // Measured from the document size, so every band bleeds the same distance.
  assert.ok(prop(band, '--band-inset').includes('var(--doc-font-pt)'));
});

test('the image scale the fitter turns down reaches both the box and its cap', () => {
  const figure = rule('.page-body .doc-image');
  assert.ok(prop(figure, 'width').includes('var(--doc-image-scale)'), 'image width must follow the scale');
  const img = rule('.page-body .doc-image img');
  assert.ok(prop(img, 'max-height').includes('var(--doc-image-scale)'), 'image height cap must follow the scale');
  assert.equal(prop(rule('.page'), '--doc-image-scale'), '1');
});

test('column gap matches the DOCX column spacing', () => {
  assert.equal(prop(rule('.page-body'), 'column-gap'), `${COLUMN_GAP_MM}mm`);
});

test('the page box is exactly A4 and owns its own margins', () => {
  const page = rule('.page');
  assert.equal(prop(page, 'width'), '210mm');
  assert.equal(prop(page, 'height'), '297mm');
  assert.equal(prop(page, 'box-sizing'), 'border-box');
  assert.equal(prop(page, 'padding'), 'calc(var(--doc-margin-mm) * 1mm)');
  assert.ok(/@page\s*\{[^}]*size:\s*A4 portrait/.test(css), '@page must declare A4');
  assert.ok(/@page\s*\{[^}]*margin:\s*0/.test(css), '@page margin must be 0');
});

test('every typeface names a family the DOCX exporter can use', () => {
  for (const [key, face] of Object.entries(TYPEFACES)) {
    assert.ok(face.docx && face.css && face.probe, `${key} is incomplete`);
    assert.ok(face.css.includes(face.docx), `${key}: preview stack must lead with ${face.docx}`);
    if (face.fallback) assert.ok(TYPEFACES[face.fallback], `${key}: unknown fallback ${face.fallback}`);
  }
});
