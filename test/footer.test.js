/**
 * footer.test.js — the footer's two slots: what is drawn, what is saved, and
 * that the stylesheet reserves the same strip footer.js describes.
 *
 * Run: npm test
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

import { FOOTER_GAP_MM, FOOTER_MM, footerHtml, hasFooter, restoreFooter, serializeFooter } from '../src/footer.js';

const PNG = 'data:image/png;base64,AAAA';
const DEFAULTS = { left: { src: PNG, alt: 'Left', preset: true }, right: { src: PNG, alt: 'Right', preset: true } };

test('both slots are always drawn, so a lone right logo stays on the right', () => {
  const html = footerHtml({ left: null, right: { src: PNG, alt: 'Logo' } });
  assert.ok(html.startsWith('<figure class="footer-slot" data-slot="left" data-empty="1"></figure>'), html);
  assert.ok(html.includes('<figure class="footer-slot" data-slot="right"><img src="data:image/png;base64,AAAA" alt="Logo"></figure>'));
});

test('only a picture the app read in is drawn', () => {
  const html = footerHtml({ left: { src: 'javascript:alert(1)' }, right: { src: 'https://x.test/a.png' } });
  assert.ok(!html.includes('<img'), html);
  assert.equal(hasFooter({ left: null, right: null }), false);
});

test('a first run gets the defaults; a cleared slot stays cleared', () => {
  assert.deepEqual(restoreFooter(undefined, DEFAULTS), DEFAULTS);
  const cleared = restoreFooter({ left: null, right: 'default' }, DEFAULTS);
  assert.equal(cleared.left, null);
  assert.equal(cleared.right, DEFAULTS.right);
});

test('a default is saved as a word, a chosen picture as itself', () => {
  const saved = serializeFooter({ left: DEFAULTS.left, right: { src: PNG, alt: 'Mine', w: 4, h: 2 } });
  assert.deepEqual(saved, { left: 'default', right: { src: PNG, alt: 'Mine', w: 4, h: 2 } });
  assert.deepEqual(restoreFooter(saved, DEFAULTS).right, { src: PNG, alt: 'Mine', w: 4, h: 2 });
});

test('document.css reserves the strip footer.js describes', () => {
  const css = readFileSync(path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'src', 'document.css'), 'utf8');
  assert.ok(new RegExp(`--footer-mm:\\s*${FOOTER_MM};`).test(css));
  assert.ok(new RegExp(`--footer-gap-mm:\\s*${FOOTER_GAP_MM};`).test(css));
  assert.ok(/\.page\[data-footer='1'\] \.page-body\s*\{[^}]*height:\s*calc\(100% - \(var\(--footer-mm\) \+ var\(--footer-gap-mm\)\) \* 1mm\)/.test(css));
});
