/**
 * footer.js — the two logos at the foot of the page, in the real app.
 *
 * A fresh document carries the bundled logos; either can be replaced or
 * removed; an empty slot is an editor-only "+ Image" target; the choice
 * survives a reload and undoes in one step; and the printed sheet is still one
 * A4 page with the footer on it.
 *
 * Usage: ONEPAGE_DEV_URL=http://localhost:5183 npx electron test/footer.js
 */

import { app, BrowserWindow } from 'electron';
import { mkdirSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const DEV_URL = process.env.ONEPAGE_DEV_URL || 'http://localhost:5183';
const OUT = path.join(os.tmpdir(), 'onepage-footer');

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok: Boolean(ok), detail });
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? `\n          ${detail}` : ''}`);
};

const ready = (win) =>
  win.webContents.executeJavaScript(
    `new Promise(r => { const t = setInterval(() => {
       if (document.documentElement.dataset.ready === 'true') { clearInterval(t); r(true); }
     }, 40); })`
  );

const run = (win, code) => win.webContents.executeJavaScript(`(async () => { ${code} })()`);

const TEXT = ['Field Note', '', 'A paragraph of text above the footer.'].join('\n');

const PROBE = `
  const doc = document.getElementById('sheet').contentDocument;
  const page = doc.getElementById('page');
  const body = doc.getElementById('page-body');
  const foot = doc.getElementById('page-footer');
  const slots = [...foot.querySelectorAll('.footer-slot')].map((f) => ({
    slot: f.dataset.slot, empty: f.dataset.empty === '1', img: Boolean(f.querySelector('img')),
  }));
`;

async function main() {
  mkdirSync(OUT, { recursive: true });
  const win = new BrowserWindow({ show: false, width: 1440, height: 900, webPreferences: { sandbox: true, contextIsolation: true } });
  await win.loadURL(DEV_URL);
  await ready(win);
  await win.webContents.executeJavaScript('localStorage.clear(); true');
  await win.reload();
  await ready(win);

  const fresh = await run(
    win,
    `await window.__onepage.setText(${JSON.stringify(TEXT)});
     ${PROBE}
     const f = foot.getBoundingClientRect(), b = body.getBoundingClientRect();
     return { slots, flag: page.getAttribute('data-footer'), footTop: f.top, bodyBottom: b.bottom, html: await window.__onepage.sheetHtml() };`
  );
  check('a fresh document carries both default logos', fresh.slots.every((s) => s.img) && fresh.flag === '1', JSON.stringify(fresh.slots));
  check('the text area stops above the footer', fresh.bodyBottom <= fresh.footTop, `${fresh.bodyBottom} <= ${fresh.footTop}`);
  check('and the printed sheet carries them', (fresh.html.match(/class="footer-slot"/g) || []).length === 2 && /data-footer="1"/.test(fresh.html));

  const replaced = await run(
    win,
    `const c = document.createElement('canvas'); c.width = 300; c.height = 100;
     const g = c.getContext('2d'); g.fillStyle = '#c0392b'; g.fillRect(0, 0, 300, 100);
     await window.__onepage.setFooter('left', c.toDataURL('image/png'), 300, 100);
     await window.__onepage.setFooter('right', null);
     ${PROBE}
     return { slots, footer: window.__onepage.footer(), html: await window.__onepage.sheetHtml() };`
  );
  check('the left logo can be replaced', replaced.footer.left && replaced.footer.left.w === 300 && !replaced.footer.left.preset);
  check('the right one removed, leaving an empty slot in the editor', replaced.footer.right === null && replaced.slots[1].empty);
  check(
    'an empty slot never reaches the printed sheet as anything but an empty box',
    !/\+ Image/.test(replaced.html) && !/data-selected/.test(replaced.html) && (replaced.html.match(/<img /g) || []).length === 1
  );

  const selected = await run(
    win,
    `window.__onepage.selectFooter('left');
     await new Promise(r => setTimeout(r, 50));
     return { tools: !document.getElementById('foot-tools').hidden, html: await window.__onepage.sheetHtml() };`
  );
  check('clicking a logo offers Replace and Remove', selected.tools);
  check('and the selection ring stays out of the export', !/data-selected/.test(selected.html));

  await win.reload();
  await ready(win);
  const reloaded = await run(win, `return window.__onepage.footer();`);
  check('the choice survives a reload, the removal included', reloaded.left && reloaded.left.w === 300 && reloaded.right === null);

  const undone = await run(
    win,
    `await window.__onepage.setFooter('left', null);
     const both = window.__onepage.footer();
     window.__onepage.undo();
     await new Promise(r => setTimeout(r, 300));
     return { removed: both.left === null, back: window.__onepage.footer().left };`
  );
  check('removing a logo is one undo step', undone.removed && undone.back && undone.back.w === 300);

  const cleared = await run(
    win,
    `await window.__onepage.setFooter('left', null);
     ${PROBE}
     const f = foot.getBoundingClientRect(), p = page.getBoundingClientRect();
     return { flag: page.getAttribute('data-footer'), html: await window.__onepage.sheetHtml(), slots,
       inMargin: f.top > p.bottom - 15 * 96 / 25.4 };`
  );
  check('with both slots empty the page is the page it was', cleared.flag === null && !/<div class="page-footer"/.test(cleared.html) && !/ data-footer="1"/.test(cleared.html));
  check('and the empty slots wait in the bottom margin', cleared.slots.every((s) => s.empty) && cleared.inMargin);

  // Back to the defaults for the print check: a full page of text above them.
  await win.webContents.executeJavaScript('localStorage.clear(); true');
  await win.reload();
  await ready(win);
  const long = Array.from({ length: 60 }, (_, i) => `Paragraph ${i + 1} of a long field report, written to fill the page well past one sheet.`).join('\n\n');
  const full = await run(win, `await window.__onepage.setText(${JSON.stringify(`Report\n\n${long}`)}); return { html: await window.__onepage.sheetHtml(), layout: window.__onepage.layout() };`);
  const pdf = await renderPdf(full.html);
  writeFileSync(path.join(OUT, 'footer.pdf'), pdf);
  const pages = pdfPageCount(pdf);
  check('a full page of text plus the footer is still exactly one A4 page', pages === 1 && !full.layout.overflow, `${pages} page(s), ${full.layout.fontPt}pt`);

  const failed = results.filter((r) => !r.ok);
  console.log(`\n  ${results.length - failed.length}/${results.length} passed\n  output: ${OUT}\n`);
  app.exit(failed.length ? 1 : 0);
}

async function renderPdf(html) {
  const file = path.join(OUT, 'sheet.html');
  writeFileSync(file, html);
  const win = new BrowserWindow({ show: false, webPreferences: { sandbox: true, contextIsolation: true } });
  try {
    await win.loadFile(file);
    return await win.webContents.printToPDF({ pageSize: 'A4', preferCSSPageSize: true, printBackground: true, margins: { marginType: 'none' } });
  } finally {
    win.destroy();
  }
}

/** Page count without pdfinfo: the page tree's own /Count. */
function pdfPageCount(buffer) {
  const text = buffer.toString('latin1');
  const counts = [...text.matchAll(/\/Type\s*\/Pages\b[^>]*?\/Count\s+(\d+)/g)].map((m) => Number(m[1]));
  if (counts.length) return Math.max(...counts);
  return (text.match(/\/Type\s*\/Page\b(?!s)/g) || []).length;
}

app.whenReady().then(() =>
  main().catch((err) => {
    console.error(err);
    app.exit(1);
  })
);
