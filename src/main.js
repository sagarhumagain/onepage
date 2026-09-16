/**
 * main.js — wiring.
 *
 * Pipeline: source text -> parse -> render -> inject into the sheet -> fit.
 * Print, PDF and Word all read the same fitted layout, so the three outputs
 * and the preview always agree.
 */

import { parse, wordCount } from './parse.js';
import { render } from './render.js';
import { createFitter, A4 } from './fit.js';
import { mountSheet, standaloneHtml } from './sheet.js';
import { textFromPaste } from './clipboard.js';
import { TYPEFACES, resolveTypeface } from './scale.js';
import { exportDocx, exportPdf, printSheet } from './exporters.js';

const $ = (id) => document.getElementById(id);

const els = {
  source: $('source'),
  sheet: $('sheet'),
  scaler: $('scaler'),
  preview: $('preview'),
  margin: $('margin'),
  typeface: $('typeface'),
  columns: $('columns'),
  fill: $('fill'),
  clear: $('btn-clear'),
  print: $('btn-print'),
  pdf: $('btn-pdf'),
  docx: $('btn-docx'),
  statWords: $('stat-words'),
  statSize: $('stat-size'),
  statLayout: $('stat-layout'),
  statNote: $('stat-note'),
  fontUp: $('font-up'),
  fontDown: $('font-down'),
  fontSizeVal: $('font-size-val'),
  bgColor: $('bg-color'),
  paddingToggle: $('padding-toggle'),
  btnImage: $('btn-image'),
  imageFile: $('image-file'),
};

const STORE_KEY = 'onepage.state.v2';

let layout = { fontPt: 11, columns: 1, lineHeight: 1.38, marginMm: 15, overflow: false, belowFloor: false };
let sheetRefs = null;
let fitter = null;
let docTitle = 'Document';
let running = false;
let queued = false;

/* --- Formatting state -------------------------------------------------- */

let fontScale = 1.0;
let bgColor = '#ffffff';
let bgPadding = false;
let imageBlocks = [];

/* --- Typeface availability --------------------------------------------- */

function fontAvailable(family) {
  try {
    const doc = sheetRefs ? sheetRefs.doc : document;
    if (!doc.fonts || !doc.fonts.check) return true;
    return doc.fonts.check(`12pt "${family}"`);
  } catch {
    return true;
  }
}

const currentFace = () => resolveTypeface(els.typeface.value, fontAvailable);

/* --- Settings ---------------------------------------------------------- */

const readSettings = () => ({
  marginMm: Number(els.margin.value),
  columns: els.columns.value === 'auto' ? 'auto' : Number(els.columns.value),
  fillPage: els.fill.checked,
});

function saveState() {
  try {
    localStorage.setItem(
      STORE_KEY,
      JSON.stringify({
        ...readSettings(),
        typeface: els.typeface.value,
        text: els.source.value,
        fontScale,
        bgColor,
        bgPadding,
        imageBlocks,
      })
    );
  } catch {
    /* private mode or blocked storage */
  }
}

function restoreState() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (!raw) return;
    const s = JSON.parse(raw);
    if (typeof s.text === 'string') els.source.value = s.text;
    if (s.marginMm) els.margin.value = String(s.marginMm);
    if (s.typeface && TYPEFACES[s.typeface]) els.typeface.value = s.typeface;
    if (s.columns != null) els.columns.value = String(s.columns);
    if (typeof s.fillPage === 'boolean') els.fill.checked = s.fillPage;
    if (typeof s.fontScale === 'number') fontScale = s.fontScale;
    if (typeof s.bgColor === 'string') bgColor = s.bgColor;
    if (typeof s.bgPadding === 'boolean') bgPadding = s.bgPadding;
    if (Array.isArray(s.imageBlocks)) imageBlocks = s.imageBlocks;

    els.bgColor.value = bgColor;
    els.paddingToggle.checked = bgPadding;
    updateFontSizeUI();
  } catch {
    /* ignore corrupt or unavailable storage */
  }
}

/* --- Font size controls ------------------------------------------------- */

function updateFontSizeUI() {
  els.fontSizeVal.textContent = Math.round(fontScale * 11);
  els.fontDown.disabled = fontScale <= 0.3;
  els.fontUp.disabled = fontScale >= 3.0;
}

function adjustFontScale(delta) {
  fontScale = Math.max(0.3, Math.min(3.0, fontScale + delta));
  updateFontSizeUI();
  applyFormatting();
  update();
  saveState();
}

/* --- Formatting application -------------------------------------------- */

function applyFormatting() {
  if (!sheetRefs) return;
  sheetRefs.page.style.setProperty('--doc-bg', bgColor);
}

/* --- Preview scaling ---------------------------------------------------- */

function rescalePreview() {
  const pad = 48;
  const byWidth = (els.preview.clientWidth - pad) / A4.widthPx;
  const byHeight = (els.preview.clientHeight - pad) / A4.heightPx;
  const scale = Math.max(0.2, Math.min(1, byWidth, byHeight));
  els.scaler.style.transform = `scale(${scale})`;
  els.scaler.style.height = `${A4.heightPx * scale}px`;
}

/* --- The pipeline ------------------------------------------------------- */

function currentDoc() {
  const blocks = parse(els.source.value);

  // Merge in image blocks at their stored positions
  if (imageBlocks.length) {
    const merged = [];
    let imgIdx = 0;
    for (let i = 0; i < blocks.length; i++) {
      // Insert images before the block at their stored position
      while (imgIdx < imageBlocks.length && imageBlocks[imgIdx].beforeIndex <= i) {
        merged.push(imageBlocks[imgIdx]);
        imgIdx++;
      }
      merged.push(blocks[i]);
    }
    while (imgIdx < imageBlocks.length) {
      merged.push(imageBlocks[imgIdx]);
      imgIdx++;
    }
    const first = merged.find((b) => b.type === 'h1' || b.type === 'h2' || b.type === 'p');
    docTitle = first && first.text ? first.text.slice(0, 80) : 'Document';
    return { blocks: merged, html: render(merged) };
  }

  const first = blocks.find((b) => b.type === 'h1' || b.type === 'h2' || b.type === 'p');
  docTitle = first && first.text ? first.text.slice(0, 80) : 'Document';
  return { blocks, html: render(blocks) };
}

async function update() {
  if (running) {
    queued = true;
    return;
  }
  running = true;
  try {
    const settings = readSettings();
    const { html } = currentDoc();
    const face = currentFace();

    sheetRefs.body.innerHTML = html;
    sheetRefs.page.style.setProperty('--doc-font-family', face.css);

    // Apply bg padding to page body
    if (bgPadding) {
      sheetRefs.body.classList.add('bg-padded');
    } else {
      sheetRefs.body.classList.remove('bg-padded');
    }

    applyFormatting();

    const words = wordCount(els.source.value);

    if (!html.trim() && !imageBlocks.length) {
      layout = {
        fontPt: 11, columns: 1, lineHeight: 1.38, marginMm: settings.marginMm,
        overflow: false, belowFloor: false, marginsReduced: false,
      };
      fitter.applyState(11 * fontScale, 1, 1.38, settings.marginMm);
      setStatus(0, null, face);
      setExportsEnabled(false);
      return;
    }

    layout = await fitter.fit({ ...settings, fontScale });
    setStatus(words, layout, face);
    setExportsEnabled(true);
  } finally {
    running = false;
    if (queued) {
      queued = false;
      update();
    }
  }
}

function setExportsEnabled(on) {
  for (const b of [els.print, els.pdf, els.docx]) b.disabled = !on;
}

function setStatus(words, fitted, face) {
  els.statWords.textContent = `${words.toLocaleString()} ${words === 1 ? 'word' : 'words'}`;

  if (!fitted) {
    els.statSize.textContent = '—';
    els.statLayout.textContent = '—';
    els.statNote.textContent = '';
    els.statNote.removeAttribute('data-level');
    return;
  }

  els.statSize.textContent = `${fitted.fontPt.toFixed(1)} pt`;
  els.statLayout.textContent =
    `${fitted.columns === 1 ? 'single column' : `${fitted.columns} columns`} · ${fitted.marginMm}mm`;

  const chosen = TYPEFACES[els.typeface.value];
  if (chosen && face.label !== chosen.label) {
    els.statNote.textContent = `${chosen.label} is not installed — using ${face.label}`;
    els.statNote.dataset.level = 'warn';
  } else if (fitted.overflow) {
    els.statNote.textContent = 'Too long for one page even at minimum size — trim the text';
    els.statNote.dataset.level = 'danger';
  } else if (fitted.belowFloor) {
    els.statNote.textContent = `Very small type (${fitted.fontPt.toFixed(1)}pt) — consider trimming`;
    els.statNote.dataset.level = 'warn';
  } else if (fitted.marginsReduced) {
    els.statNote.textContent = `Margins reduced to ${fitted.marginMm}mm to fit one page`;
    els.statNote.dataset.level = 'warn';
  } else {
    els.statNote.textContent = 'Fits one A4 page';
    els.statNote.dataset.level = 'ok';
  }
}

/* --- Image handling ----------------------------------------------------- */

function readFileAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function getImageDimensions(src) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = () => resolve({ width: 400, height: 300 });
    img.src = src;
  });
}

async function addImageFromFile(file) {
  if (!file || !file.type.startsWith('image/')) return;
  const src = await readFileAsDataURL(file);
  const dims = await getImageDimensions(src);
  const blockCount = parse(els.source.value).length;
  imageBlocks.push({
    type: 'image',
    src,
    alt: file.name || '',
    align: 'center',
    beforeIndex: blockCount,
  });
  update();
  saveState();
}

function handleImagePaste(event) {
  const dt = event.clipboardData;
  if (!dt) return false;

  const items = dt.items;
  for (const item of items) {
    if (item.type.startsWith('image/')) {
      event.preventDefault();
      const file = item.getAsFile();
      if (file) addImageFromFile(file);
      return true;
    }
  }
  return false;
}

/* --- Exports ------------------------------------------------------------ */

async function sheetHtml() {
  const { html } = currentDoc();
  return standaloneHtml(html, {
    fontPt: layout.fontPt,
    columns: layout.columns,
    lineHeight: layout.lineHeight,
    marginMm: layout.marginMm,
    fontFamily: currentFace().css,
    title: docTitle,
    bgColor,
  });
}

/* --- Events ------------------------------------------------------------- */

function debounce(fn, ms) {
  let t = 0;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

const scheduleUpdate = debounce(() => {
  update();
  saveState();
}, 140);

function insertAtCursor(textarea, text) {
  const { selectionStart: s, selectionEnd: e, value } = textarea;
  textarea.value = value.slice(0, s) + text + value.slice(e);
  const pos = s + text.length;
  textarea.setSelectionRange(pos, pos);
}

function onPaste(event) {
  if (handleImagePaste(event)) return;
  const text = textFromPaste(event);
  if (!text) return;
  event.preventDefault();
  insertAtCursor(els.source, text);
  update();
  saveState();
}

async function withBusy(button, fn) {
  const label = button.textContent;
  button.disabled = true;
  button.textContent = 'Working…';
  try {
    await fn();
  } catch (err) {
    console.error(err);
    els.statNote.textContent = `Export failed: ${err && err.message ? err.message : err}`;
    els.statNote.dataset.level = 'danger';
  } finally {
    button.textContent = label;
    button.disabled = false;
  }
}

function wire() {
  els.source.addEventListener('input', scheduleUpdate);
  els.source.addEventListener('paste', onPaste);

  for (const el of [els.margin, els.typeface, els.columns, els.fill]) {
    el.addEventListener('change', () => {
      update();
      saveState();
    });
  }

  els.clear.addEventListener('click', () => {
    els.source.value = '';
    imageBlocks = [];
    els.source.focus();
    update();
    saveState();
  });

  // Font size controls
  els.fontUp.addEventListener('click', () => adjustFontScale(0.1));
  els.fontDown.addEventListener('click', () => adjustFontScale(-0.1));

  // Background color
  els.bgColor.addEventListener('input', () => {
    bgColor = els.bgColor.value;
    applyFormatting();
    saveState();
  });

  // Padding toggle
  els.paddingToggle.addEventListener('change', () => {
    bgPadding = els.paddingToggle.checked;
    update();
    saveState();
  });

  // Image upload
  els.btnImage.addEventListener('click', () => els.imageFile.click());
  els.imageFile.addEventListener('change', () => {
    for (const file of els.imageFile.files) {
      addImageFromFile(file);
    }
    els.imageFile.value = '';
  });

  // Drag and drop images onto the source area
  els.source.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
  });
  els.source.addEventListener('drop', (e) => {
    e.preventDefault();
    for (const file of e.dataTransfer.files) {
      if (file.type.startsWith('image/')) addImageFromFile(file);
    }
  });

  els.print.addEventListener('click', () =>
    withBusy(els.print, async () => {
      await printSheet({ html: await sheetHtml(), iframe: els.sheet });
    })
  );

  els.pdf.addEventListener('click', () =>
    withBusy(els.pdf, async () => {
      const r = await exportPdf({ html: await sheetHtml(), title: docTitle });
      if (r && r.saved) flash(`Saved ${r.filePath}`);
    })
  );

  els.docx.addEventListener('click', () =>
    withBusy(els.docx, async () => {
      const { blocks } = currentDoc();
      const r = await exportDocx({ blocks, layout, face: currentFace(), title: docTitle });
      if (r && r.saved && r.filePath) flash(`Saved ${r.filePath}`);
    })
  );

  window.addEventListener('resize', debounce(rescalePreview, 60));

  window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
      e.preventDefault();
      els.print.click();
    }
  });
}

let flashTimer = 0;
function flash(message) {
  els.statNote.textContent = message;
  els.statNote.dataset.level = 'ok';
  clearTimeout(flashTimer);
  flashTimer = setTimeout(() => setStatus(wordCount(els.source.value), layout, currentFace()), 4000);
}

/* --- Boot --------------------------------------------------------------- */

function exposeForTests() {
  if (!import.meta.env.DEV) return;
  window.__onepage = {
    setText: async (text) => {
      els.source.value = text;
      await update();
    },
    layout: () => ({ ...layout }),
    face: () => currentFace(),
    sheetHtml,
    docxBytes: async () => {
      const { blocks } = currentDoc();
      const { buildDocx } = await import('./to-docx.js');
      const blob = await buildDocx({ blocks, layout, face: currentFace(), title: docTitle });
      return Array.from(new Uint8Array(await blob.arrayBuffer()));
    },
  };
}

async function boot() {
  restoreState();
  sheetRefs = await mountSheet(els.sheet);
  fitter = createFitter({ page: sheetRefs.page, body: sheetRefs.body });
  wire();
  rescalePreview();
  exposeForTests();
  updateFontSizeUI();
  applyFormatting();
  await update();
  els.source.focus();
  document.documentElement.dataset.ready = 'true';
}

boot();
