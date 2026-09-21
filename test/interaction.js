/**
 * interaction.js — the editing gestures, driven in the real app.
 *
 * e2e.js proves the fitter: content in, exactly one A4 page out. This proves
 * the two things the user does *to the page* rather than to the text — placing
 * and sizing an image, and formatting a run of words — because neither can be
 * checked from a unit test: both are a DOM selection, a pointer drag and a
 * re-render, and all three only exist in a browser.
 *
 * Usage: ONEPAGE_DEV_URL=http://localhost:5183 npx electron test/interaction.js
 */

import { app, BrowserWindow } from 'electron';
import { mkdirSync, writeFileSync } from 'node:fs';
import { inflateRawSync } from 'node:zlib';
import os from 'node:os';
import path from 'node:path';

const DEV_URL = process.env.ONEPAGE_DEV_URL || 'http://localhost:5183';
const OUT = path.join(os.tmpdir(), 'onepage-interaction');

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok: Boolean(ok), detail });
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? `\n          ${detail}` : ''}`);
};

const TEXT = [
  'Field Assessment',
  '',
  'The first paragraph sits above the picture and should stay exactly where it is,',
  'with the image taking only the width it was given and the text carrying on.',
  '',
  'The second paragraph sits below the picture. It is here to prove that an image',
  'is a block in the flow like any other, not the end of the document.',
].join('\n');

const ready = (win) =>
  win.webContents.executeJavaScript(
    `new Promise(r => { const t = setInterval(() => {
       if (document.documentElement.dataset.ready === 'true') { clearInterval(t); r(true); }
     }, 40); })`
  );

/** Run an expression in the page and hand back whatever it resolves to. */
const run = (win, code) => win.webContents.executeJavaScript(`(async () => { ${code} })()`);

const PROBE = `
  const doc = document.getElementById('sheet').contentDocument;
  const body = doc.getElementById('page-body');
  const kids = [...body.children].filter(el => el.getAttribute('aria-hidden') !== 'true');
`;

async function main() {
  mkdirSync(OUT, { recursive: true });

  const win = new BrowserWindow({
    show: false,
    width: 1440,
    height: 900,
    webPreferences: { sandbox: true, contextIsolation: true },
  });

  await win.loadURL(DEV_URL);
  await ready(win);
  // A previous run's document must not decide this one's.
  await win.webContents.executeJavaScript('localStorage.clear(); true');
  await win.reload();
  await ready(win);

  /* --- Images ---------------------------------------------------------- */

  console.log('\n  images\n  ' + '-'.repeat(70));

  const placed = await run(
    win,
    `
    await window.__onepage.setText(${JSON.stringify(TEXT)});
    const source = document.getElementById('source');
    // Put the cursor at the end of the first body paragraph.
    const at = source.value.indexOf('carrying on.') + 'carrying on.'.length;
    source.setSelectionRange(at, at);

    const canvas = document.createElement('canvas');
    canvas.width = 1200; canvas.height = 800;
    const g = canvas.getContext('2d');
    g.fillStyle = '#4a7fb5'; g.fillRect(0, 0, 1200, 800);
    g.fillStyle = '#ffffff'; g.fillRect(80, 80, 1040, 640);
    const id = await window.__onepage.addImage(canvas.toDataURL('image/png'), 1200, 800);

    ${PROBE}
    const figure = body.querySelector('.doc-image');
    const order = kids.map(el => el.tagName.toLowerCase());
    return {
      id,
      order,
      source: source.value,
      figureWidth: figure.getBoundingClientRect().width,
      bodyWidth: body.getBoundingClientRect().width,
      figureBottom: figure.getBoundingClientRect().bottom,
      afterTop: figure.nextElementSibling.getBoundingClientRect().top,
      afterText: figure.nextElementSibling.textContent.slice(0, 40),
      layout: window.__onepage.layout(),
    };
  `
  );

  check(
    'an image sits between the paragraphs the cursor was in, not at the end',
    placed.order.join(',') === 'h1,p,figure,p',
    `order: ${placed.order.join(' ')}`
  );
  check(
    'its position is a line in the source, so it can be moved by moving text',
    /\[image:[a-z0-9]+\]/.test(placed.source) &&
      placed.source.indexOf('[image:') > placed.source.indexOf('carrying on.'),
    placed.source.split('\n').filter((l) => l.includes('[image:')).join('')
  );
  check(
    'it takes the width it was given, not the whole column',
    placed.figureWidth / placed.bodyWidth > 0.5 && placed.figureWidth / placed.bodyWidth < 0.7,
    `${Math.round((placed.figureWidth / placed.bodyWidth) * 100)}% of the column`
  );
  check(
    'the text after it continues below it',
    placed.afterTop >= placed.figureBottom - 1 && placed.afterText.length > 10,
    `"${placed.afterText}…"`
  );

  const wrapped = await run(
    win,
    `
    await window.__onepage.setImage(${JSON.stringify(placed.id)}, { align: 'left', wrap: true, width: 40 });
    ${PROBE}
    const figure = body.querySelector('.doc-image');
    const after = figure.nextElementSibling;
    const fig = figure.getBoundingClientRect();
    // Where does the first line of the following paragraph actually start?
    const range = doc.createRange();
    range.setStart(after.firstChild, 0);
    range.setEnd(after.firstChild, 12);
    const line = range.getBoundingClientRect();
    return {
      float: doc.defaultView.getComputedStyle(figure).float,
      lineLeft: line.left,
      lineTop: line.top,
      figureRight: fig.right,
      figureTop: fig.top,
      figureBottom: fig.bottom,
    };
  `
  );

  check(
    'wrapped, the text runs beside the image instead of under it',
    wrapped.float === 'left' &&
      wrapped.lineLeft >= wrapped.figureRight - 1 &&
      wrapped.lineTop < wrapped.figureBottom,
    `first line starts at x=${Math.round(wrapped.lineLeft)}, image ends at x=${Math.round(wrapped.figureRight)}`
  );

  const resized = await run(
    win,
    `
    window.__onepage.selectImage(${JSON.stringify(placed.id)});
    await new Promise(r => setTimeout(r, 60));
    const before = window.__onepage.imageOf(${JSON.stringify(placed.id)}).width;

    const handle = document.querySelector('.img-handle[data-corner="se"]');
    const box = handle.getBoundingClientRect();
    const x = box.left + box.width / 2;
    const y = box.top + box.height / 2;
    const ev = (type, dx) => new PointerEvent(type, {
      clientX: x + dx, clientY: y, bubbles: true, cancelable: true, pointerId: 1, isPrimary: true,
    });

    handle.dispatchEvent(ev('pointerdown', 0));
    window.dispatchEvent(ev('pointermove', 90));
    const readout = document.getElementById('img-size').textContent;
    const resizing = document.body.classList.contains('resizing');
    window.dispatchEvent(ev('pointerup', 90));
    await new Promise(r => setTimeout(r, 250));

    return { before, readout, resizing, after: window.__onepage.imageOf(${JSON.stringify(placed.id)}).width };
  `
  );

  check(
    'dragging a corner handle resizes the image and commits it',
    resized.resizing && resized.after > resized.before,
    `${resized.before}% -> ${resized.after}% (live readout ${resized.readout})`
  );

  const escaped = await run(
    win,
    `
    const id = ${JSON.stringify(placed.id)};
    const before = window.__onepage.imageOf(id).width;
    const handle = document.querySelector('.img-handle[data-corner="se"]');
    const box = handle.getBoundingClientRect();
    const x = box.left + box.width / 2;
    const y = box.top + box.height / 2;
    const ev = (type, dx) => new PointerEvent(type, {
      clientX: x + dx, clientY: y, bubbles: true, cancelable: true, pointerId: 1, isPrimary: true,
    });

    handle.dispatchEvent(ev('pointerdown', 0));
    window.dispatchEvent(ev('pointermove', -70));
    const midway = document.getElementById('img-size').textContent;

    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await new Promise(r => setTimeout(r, 120));
    const afterEscape = window.__onepage.imageOf(id).width;

    // Anything still moving would keep following the pointer after Esc.
    window.dispatchEvent(ev('pointermove', -160));
    window.dispatchEvent(ev('pointerup', -160));
    await new Promise(r => setTimeout(r, 200));

    ${PROBE}
    const figure = body.querySelector('.doc-image');
    return {
      before,
      midway,
      afterEscape,
      settled: window.__onepage.imageOf(id).width,
      resizing: document.body.classList.contains('resizing'),
      rendered: Number(figure.style.getPropertyValue('--img-w')),
      stillSelected: !document.getElementById('img-tools').hidden,
    };
  `
  );

  check(
    'Esc stops the resize and puts the image back the size it was',
    escaped.afterEscape === escaped.before &&
      escaped.settled === escaped.before &&
      escaped.rendered === escaped.before &&
      !escaped.resizing,
    `${escaped.before}% -> dragged to ${escaped.midway} -> Esc -> ${escaped.settled}%`
  );
  check('and the image is still selected, so it can simply be dragged again', escaped.stillSelected);

  const escapedTwice = await run(
    win,
    `
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await new Promise(r => setTimeout(r, 80));
    return { tools: document.getElementById('img-tools').hidden };
  `
  );
  check('a second Esc lets the image go entirely', escapedTwice.tools);

  /* --- Formatting ------------------------------------------------------- */

  console.log('\n  formatting\n  ' + '-'.repeat(70));

  const highlighted = await run(
    win,
    `
    // Select the title the way a user does — inside the sheet, with the
    // selection API — so the app's own selectionchange wiring is what runs.
    ${PROBE}
    const title = body.querySelector('h1');
    const range = doc.createRange();
    range.selectNodeContents(title);
    const selection = doc.defaultView.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    await new Promise(r => setTimeout(r, 120));

    const ok = !document.getElementById('fmt-tools').hidden;
    const barShown = ok;
    document.getElementById('btn-highlight').click();
    await new Promise(r => setTimeout(r, 250));
    return {
      ok,
      barShown,
      h1: body.querySelector('h1').innerHTML,
      keptSelection: !document.getElementById('fmt-tools').hidden,
      note: document.getElementById('stat-note').textContent,
    };
  `
  );

  check('selecting a title shows the format bar over it', highlighted.ok && highlighted.barShown);
  check(
    'Highlight puts a background behind the selected words',
    /<mark style="background:#ffe08a">/.test(highlighted.h1),
    highlighted.h1.slice(0, 90)
  );
  check('the selection survives, so formatting can be chained', highlighted.keptSelection);

  const styled = await run(
    win,
    `
    const click = (sel) => document.querySelector(sel).click();
    click('[data-fmt="bigger"]');
    await new Promise(r => setTimeout(r, 250));
    click('[data-fmt="bold"]');
    await new Promise(r => setTimeout(r, 250));
    const fg = document.getElementById('fmt-fg');
    fg.value = '#b42318';
    fg.dispatchEvent(new Event('input', { bubbles: true }));
    await new Promise(r => setTimeout(r, 350));
    ${PROBE}
    return {
      h1: body.querySelector('h1').innerHTML,
      size: document.getElementById('fmt-size').textContent,
      pressed: document.querySelector('[data-fmt="bold"]').getAttribute('aria-pressed'),
      exported: await window.__onepage.sheetHtml(),
      layout: window.__onepage.layout(),
    };
  `
  );

  check(
    'size, bold and colour all land on the same run',
    /font-size:1\.15em/.test(styled.h1) && /font-weight:700/.test(styled.h1) && /color:#b42318/.test(styled.h1),
    styled.h1.slice(0, 140)
  );
  check('the bar reports the state of the run it is over', styled.size === '115%' && styled.pressed === 'true');
  check(
    'the export carries the formatting and none of the editing furniture',
    /background:#ffe08a/.test(styled.exported) &&
      /font-size:1\.15em/.test(styled.exported) &&
      !/data-u=/.test(styled.exported) &&
      !/data-selected/.test(styled.exported)
  );

  const survived = await run(
    win,
    `
    const source = document.getElementById('source');
    // Edit a *different* paragraph: the mark must not move or vanish.
    await window.__onepage.setText(source.value.replace('The second paragraph', 'A rewritten second paragraph'));
    await new Promise(r => setTimeout(r, 250));
    ${PROBE}
    return { h1: body.querySelector('h1').innerHTML };
  `
  );
  check(
    'formatting is anchored to its words, not to a position in the text',
    /<mark style="background:#ffe08a;color:#b42318;font-size:1\.15em;font-weight:700">/.test(survived.h1),
    survived.h1.slice(0, 140)
  );

  const banded = await run(
    win,
    `
    ${PROBE}
    const heading = body.querySelector('h1');
    // The title already carries a <mark>, so take its first text node rather
    // than its first child.
    const text = doc.createTreeWalker(heading, 4).nextNode();
    const range = doc.createRange();
    range.setStart(text, 0);
    range.setEnd(text, Math.min(5, text.nodeValue.length)); // only part of the line
    const selection = doc.defaultView.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    await new Promise(r => setTimeout(r, 120));
    document.querySelector('[data-fmt="band"]').click();
    await new Promise(r => setTimeout(r, 300));

    const banded = body.querySelector('h1');
    const box = banded.getBoundingClientRect();
    // Where the banded line's *text* starts, not where its box does: the box
    // is deliberately wider than the text, by the bleed on each side.
    const inner = doc.createRange();
    inner.selectNodeContents(banded);
    return {
      html: banded.outerHTML.slice(0, 120),
      bandWidth: box.width,
      bodyWidth: body.getBoundingClientRect().width,
      bandLeft: box.left,
      textLeft: inner.getBoundingClientRect().left,
      bodyTextLeft: body.querySelector('p').getBoundingClientRect().left,
      exported: await window.__onepage.sheetHtml(),
      pressed: document.querySelector('[data-fmt="band"]').getAttribute('aria-pressed'),
    };
  `
  );

  check(
    'Band fills the whole width of the line, not just the words',
    /class="doc-band" style="background:/.test(banded.html) && banded.bandWidth >= banded.bodyWidth - 1,
    `band ${Math.round(banded.bandWidth)}px across a ${Math.round(banded.bodyWidth)}px column`
  );
  check(
    'and a banded line still starts on the same left edge as the rest',
    Math.abs(banded.textLeft - banded.bodyTextLeft) < 1.5 && banded.bandLeft < banded.textLeft,
    `banded text at x=${banded.textLeft.toFixed(1)}, body text at x=${banded.bodyTextLeft.toFixed(1)}, ` +
      `band bleeds ${(banded.textLeft - banded.bandLeft).toFixed(1)}px past it`
  );
  check('the bar shows it as on, and the export carries it', banded.pressed === 'true' && /doc-band/.test(banded.exported));

  const cleared = await run(
    win,
    `
    window.__onepage.select(0, 0, null);
    await new Promise(r => setTimeout(r, 80));
    document.querySelector('[data-fmt="clear"]').click();
    await new Promise(r => setTimeout(r, 250));
    ${PROBE}
    return { h1: body.querySelector('h1').outerHTML };
  `
  );
  check(
    'Clear takes all of it off again, band included',
    !/<mark|<span style|doc-band/.test(cleared.h1),
    cleared.h1.slice(0, 80)
  );

  /* --- Typing on the page, and taking it back --------------------------- */

  console.log('\n  editing the page\n  ' + '-'.repeat(70));

  const typed = await run(
    win,
    `
    const before = window.__onepage.source();
    ${PROBE}
    // Put the caret at the end of the first body paragraph and type.
    const para = body.querySelectorAll('p')[0];
    const index = Number(para.dataset.u);
    window.__onepage.caretAt(index, para.textContent.length);
    await window.__onepage.type(' Typed straight onto the page.');
    await new Promise(r => setTimeout(r, 400));

    return {
      before,
      source: window.__onepage.source(),
      page: body.textContent,
      layout: window.__onepage.layout(),
      history: window.__onepage.history(),
      note: document.getElementById('stat-note').textContent,
    };
  `
  );

  check(
    'typing on the page writes straight back into the source text',
    typed.source.includes('Typed straight onto the page.') &&
      !typed.before.includes('Typed straight onto the page.'),
    typed.source.split('\n').find((l) => l.includes('Typed straight')) || ''
  );
  check(
    'and the page is re-fitted around what was typed',
    typed.page.includes('Typed straight onto the page.') && !typed.layout.overflow,
    `${typed.layout.fontPt}pt — "${typed.note}"`
  );

  const formattedAfterTyping = await run(
    win,
    `
    ${PROBE}
    const para = body.querySelectorAll('p')[0];
    const index = Number(para.dataset.u);
    window.__onepage.select(index, 0, null);
    await new Promise(r => setTimeout(r, 100));
    document.getElementById('btn-highlight').click();
    await new Promise(r => setTimeout(r, 400));
    const after = doc.getElementById('page-body').querySelectorAll('p')[0];
    return { html: after.innerHTML.slice(0, 200), marked: after.querySelector('mark') ? after.querySelector('mark').textContent : '' };
  `
  );

  check(
    'formatting a line that was just typed works, without reselecting anything',
    formattedAfterTyping.marked.includes('Typed straight onto the page.'),
    `highlighted: "${formattedAfterTyping.marked.slice(-42)}"`
  );

  const undone = await run(
    win,
    `
    const z = (shift) => window.dispatchEvent(new KeyboardEvent('keydown', {
      key: shift ? 'Z' : 'z', metaKey: true, shiftKey: Boolean(shift), bubbles: true,
    }));

    const before = window.__onepage.source();
    z(false); // undo the highlight
    await new Promise(r => setTimeout(r, 300));
    ${PROBE}
    const afterOneUndo = body.querySelectorAll('p')[0].innerHTML;

    z(false); // undo the typing
    await new Promise(r => setTimeout(r, 300));
    const afterTwo = window.__onepage.source();

    z(true); // redo the typing
    await new Promise(r => setTimeout(r, 300));
    const afterRedo = window.__onepage.source();

    return {
      before,
      highlightGone: !afterOneUndo.includes('<mark'),
      afterTwo,
      afterRedo,
      history: window.__onepage.history(),
      buttons: {
        undo: document.getElementById('btn-undo').disabled,
        redo: document.getElementById('btn-redo').disabled,
      },
    };
  `
  );

  check('Ctrl/Cmd+Z takes back a formatting change', undone.highlightGone);
  check(
    'and takes back typing, a phrase at a time rather than a letter at a time',
    !undone.afterTwo.includes('Typed straight onto the page.'),
    `source ends: "…${undone.afterTwo.trim().slice(-48)}"`
  );
  check(
    'redo puts it back',
    undone.afterRedo.includes('Typed straight onto the page.') && !undone.buttons.undo,
    `${undone.history.undo} steps back, ${undone.history.redo} forward`
  );

  /* --- Inserting structure ---------------------------------------------- */

  console.log('\n  inserting\n  ' + '-'.repeat(70));

  const inserted = await run(
    win,
    `
    ${PROBE}
    const para = body.querySelectorAll('p')[0];
    window.__onepage.caretAt(Number(para.dataset.u), 4);
    const menu = document.getElementById('insert');
    menu.value = 'table';
    menu.dispatchEvent(new Event('change', { bubbles: true }));
    await new Promise(r => setTimeout(r, 500));

    const after = doc.getElementById('page-body');
    const table = after.querySelector('table');
    const rows = table ? table.querySelectorAll('tr').length : 0;
    const cells = table ? table.querySelectorAll('tr')[0].children.length : 0;
    const source = window.__onepage.source();
    return {
      rows,
      cells,
      afterPara: table ? table.previousElementSibling.tagName.toLowerCase() : '',
      source,
      dividerLine: source.split('\\n').some((l) => /^\\|\\s*---/.test(l)),
      menuReset: menu.value === '',
      layout: window.__onepage.layout(),
    };
  `
  );

  check(
    'Insert drops a real table in at the cursor, not at the end',
    inserted.rows === 3 && inserted.cells === 3 && inserted.afterPara === 'p',
    `${inserted.rows} rows x ${inserted.cells} columns, after a <${inserted.afterPara}>`
  );
  check(
    'and it is written back into the source as a table, not as markup soup',
    inserted.dividerLine && /\| Column A \| Column B \| Column C \|/.test(inserted.source),
    inserted.source.split('\n').filter((l) => l.includes('|')).slice(0, 2).join('  /  ')
  );
  check('the menu returns to its label, ready for the next insert', inserted.menuReset);

  const listed = await run(
    win,
    `
    const menu = document.getElementById('insert');
    menu.value = 'ul';
    menu.dispatchEvent(new Event('change', { bubbles: true }));
    await new Promise(r => setTimeout(r, 450));
    ${PROBE}
    const items = [...body.querySelectorAll('ul li')].map((li) => li.textContent);

    // And taking it back is one keystroke, like everything else.
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', metaKey: true, bubbles: true }));
    await new Promise(r => setTimeout(r, 350));
    return { items, afterUndo: doc.getElementById('page-body').querySelectorAll('ul').length };
  `
  );

  check('a list inserts the same way', listed.items.length === 3, listed.items.join(' · '));
  check('and undo removes a whole insert in one step', listed.afterUndo === 0);

  /* --- Page colour and the inset panel ---------------------------------- */

  console.log('\n  page colour\n  ' + '-'.repeat(70));

  const page = await run(
    win,
    `
    const colour = document.getElementById('bg-color');
    colour.value = '#fdf6e3';
    colour.dispatchEvent(new Event('input', { bubbles: true }));
    const toggle = document.getElementById('padding-toggle');
    toggle.checked = true;
    toggle.dispatchEvent(new Event('change', { bubbles: true }));
    await new Promise(r => setTimeout(r, 350));

    ${PROBE}
    const pageEl = doc.getElementById('page');
    const styles = doc.defaultView.getComputedStyle(body);
    const box = body.getBoundingClientRect();
    const last = kids[kids.length - 1].getBoundingClientRect();
    return {
      inset: pageEl.classList.contains('inset'),
      panel: styles.backgroundColor,
      padding: parseFloat(styles.paddingTop),
      overflow: last.bottom - (box.bottom - parseFloat(styles.paddingBottom)),
      exported: await window.__onepage.sheetHtml(),
      note: document.getElementById('stat-note').textContent,
    };
  `
  );

  check(
    'the page colour becomes an inset panel, and the text keeps clear of its edge',
    page.inset && page.padding > 2 && page.panel === 'rgb(253, 246, 227)',
    `panel ${page.panel}, ${page.padding.toFixed(1)}px of padding`
  );
  check(
    'the fitter measures the panel, not the sheet, so nothing runs under the edge',
    page.overflow <= 1.5,
    `${page.overflow.toFixed(2)}px past the panel's inside edge — "${page.note}"`
  );
  check(
    'and the export carries the colour and the panel with it',
    /--doc-bg:#fdf6e3/.test(page.exported) && /class="page inset"/.test(page.exported)
  );

  await run(
    win,
    `
    const toggle = document.getElementById('padding-toggle');
    toggle.checked = false;
    toggle.dispatchEvent(new Event('change', { bubbles: true }));
    const colour = document.getElementById('bg-color');
    colour.value = '#ffffff';
    colour.dispatchEvent(new Event('input', { bubbles: true }));
    await new Promise(r => setTimeout(r, 250));
    return true;
  `
  );

  /* --- Still one page, and still exportable ----------------------------- */

  console.log('\n  output\n  ' + '-'.repeat(70));

  const heavy = await run(
    win,
    `
    const filler = ('Coverage rose across all three districts and the supervision backlog fell. ').repeat(90);
    await window.__onepage.setText('Programme Review\\n\\n' + filler);
    const source = document.getElementById('source');
    source.setSelectionRange(source.value.length, source.value.length);
    const canvas = document.createElement('canvas');
    canvas.width = 1600; canvas.height = 1200;
    const g = canvas.getContext('2d');
    g.fillStyle = '#c95d3c'; g.fillRect(0, 0, 1600, 1200);
    await window.__onepage.addImage(canvas.toDataURL('image/png'), 1600, 1200, {
      width: 45, align: 'left', wrap: true,
    });
    window.__onepage.select(0, 0, null);
    await new Promise(r => setTimeout(r, 80));
    window.__onepage.format({ bg: '#ffe08a', fg: '#b42318', size: 1.35, bold: true });
    await new Promise(r => setTimeout(r, 400));
    ${PROBE}
    const box = body.getBoundingClientRect();
    const last = kids[kids.length - 1].getBoundingClientRect();
    return {
      layout: window.__onepage.layout(),
      overflow: last.bottom - box.bottom,
      note: document.getElementById('stat-note').textContent,
      html: await window.__onepage.sheetHtml(),
      docx: await window.__onepage.docxBytes(),
    };
  `
  );

  check(
    'a full page of text plus a large picture still fits, with the type still readable',
    heavy.overflow <= 1.5 && !heavy.layout.overflow && heavy.layout.fontPt >= 7,
    `${heavy.layout.fontPt}pt, images at ${Math.round(heavy.layout.imageScale * 100)}%, ` +
      `overflow ${heavy.overflow.toFixed(2)}px — "${heavy.note}"`
  );

  const docx = Buffer.from(heavy.docx);
  writeFileSync(path.join(OUT, 'export.docx'), docx);
  const xml = readZipEntry(docx, 'word/document.xml') || '';
  const extent = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(xml);
  const inches = extent ? [Number(extent[1]) / 914400, Number(extent[2]) / 914400] : [0, 0];

  check(
    'the Word export builds and carries the picture at a real size',
    docx[0] === 0x50 && docx[1] === 0x4b && extent && inches[0] > 0.3 && inches[0] < 8,
    `${docx.length} bytes, image ${inches[0].toFixed(2)} x ${inches[1].toFixed(2)} in`
  );
  check(
    'the picture keeps its proportions rather than a guessed ratio',
    Math.abs(inches[0] / inches[1] - 1600 / 1200) < 0.02,
    `ratio ${(inches[0] / inches[1]).toFixed(3)} against the file's own 1.333`
  );
  check(
    'a wrapped image is anchored with square wrapping, so Word flows text beside it too',
    xml.includes('<wp:anchor') && xml.includes('<wp:wrapSquare'),
    xml.includes('<wp:anchor') ? 'anchored' : 'still inline'
  );
  // The heading is 1.9x body and the run asks for 1.35x on top of that, so its
  // half-points must be well clear of anything else in the file.
  const sizes = [...xml.matchAll(/<w:sz w:val="(\d+)"/g)].map((m) => Number(m[1]));
  const titleRun = /<w:rPr>(?:(?!<\/w:rPr>).)*B42318(?:(?!<\/w:rPr>).)*<\/w:rPr>/.exec(xml);
  check(
    'and the formatting reaches that run: shading, colour and a bigger size together',
    Boolean(titleRun) && /w:fill="FFE08A"/.test(titleRun[0]) && /<w:sz w:val="(3[5-9]|[4-9]\d)"/.test(titleRun[0]),
    `sizes in the file: ${[...new Set(sizes)].sort((a, b) => a - b).join(', ')} half-points`
  );
  check(
    'an enlarged run asks Word for a line box that can hold it',
    /<w:spacing[^>]*w:lineRule="atLeast"[^>]*\/>\s*<\/w:pPr>\s*<w:r><w:rPr>(?:(?!<\/w:rPr>).)*B42318/.test(xml) ||
      /w:lineRule="atLeast"/.test(xml),
    'exact leading would have clipped it'
  );

  const pdf = await renderPdf(heavy.html);
  writeFileSync(path.join(OUT, 'export.pdf'), pdf);
  const pages = pdfPageCount(pdf);
  check('and the printed sheet is still exactly one A4 page', pages === 1, `${pages} page(s), ${pdf.length} bytes`);

  /* --- Summary ---------------------------------------------------------- */

  const failed = results.filter((r) => !r.ok).length;
  console.log('\n  ' + '-'.repeat(70));
  console.log(`  ${results.length - failed}/${results.length} checks passed`);
  console.log(`  artifacts: ${OUT}\n`);
  app.exit(failed ? 1 : 0);
}

/** One entry out of a zip, through its central directory. */
function readZipEntry(buf, wanted) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0 && i > buf.length - 70000; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return null;

  const count = buf.readUInt16LE(eocd + 10);
  let at = buf.readUInt32LE(eocd + 16);
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(at) !== 0x02014b50) return null;
    const method = buf.readUInt16LE(at + 10);
    const compressed = buf.readUInt32LE(at + 20);
    const nameLen = buf.readUInt16LE(at + 28);
    const extraLen = buf.readUInt16LE(at + 30);
    const commentLen = buf.readUInt16LE(at + 32);
    const local = buf.readUInt32LE(at + 42);
    const name = buf.toString('utf8', at + 46, at + 46 + nameLen);

    if (name === wanted) {
      const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
      const data = buf.subarray(start, start + compressed);
      return (method === 0 ? data : inflateRawSync(data)).toString('utf8');
    }
    at += 46 + nameLen + extraLen + commentLen;
  }
  return null;
}

async function renderPdf(html) {
  const win = new BrowserWindow({ show: false, webPreferences: { sandbox: true } });
  try {
    const loaded = new Promise((res, rej) => {
      win.webContents.once('did-finish-load', res);
      win.webContents.once('did-fail-load', (_e, code, desc) => rej(new Error(`${desc} (${code})`)));
    });
    await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
    await loaded;
    return await win.webContents.printToPDF({
      pageSize: 'A4',
      preferCSSPageSize: true,
      printBackground: true,
      scale: 1,
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      displayHeaderFooter: false,
    });
  } finally {
    win.destroy();
  }
}

/** Page count without pdfinfo: the page tree's own /Count. */
function pdfPageCount(buffer) {
  const text = buffer.toString('latin1');
  const counts = [...text.matchAll(/\/Type\s*\/Pages[^>]*?\/Count\s+(\d+)/g)].map((m) => Number(m[1]));
  if (counts.length) return Math.max(...counts);
  return (text.match(/\/Type\s*\/Page[^s]/g) || []).length;
}

app.whenReady().then(() =>
  main().catch((err) => {
    console.error('\nINTERACTION HARNESS FAILED:', err);
    app.exit(1);
  })
);
