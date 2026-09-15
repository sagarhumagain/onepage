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
import { tokenizeInline, plainText } from '../src/inline.js';
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
