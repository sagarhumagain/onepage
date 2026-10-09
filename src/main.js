/**
 * main.js — wiring.
 *
 * Pipeline: source text -> parse -> render -> inject into the sheet -> fit.
 * Print, PDF and Word all read the same fitted layout, so the three outputs
 * and the preview always agree.
 *
 * The source textarea is the single source of truth for the document. Images
 * and highlights are tokens inside that text (see markup.js) rather than a
 * parallel list of decorations, which is what makes an image move when you
 * type above it and what lets both survive a reload.
 */

import { parse, wordCount } from './parse.js';
import { render } from './render.js';
import { createFitter, A4 } from './fit.js';
import { mountSheet, standaloneHtml } from './sheet.js';
import { textFromPaste } from './clipboard.js';
import { TYPEFACES, resolveTypeface } from './scale.js';
import { exportDocx, exportPdf, printSheet } from './exporters.js';
import * as images from './images.js';
import {
  DEFAULT_ALIGN,
  DEFAULT_WIDTH_PCT,
  highlight,
  imageToken,
  insertImageAt,
  referencedIds,
  removeImage,
  stripHighlights,
  updateImage,
} from './markup.js';

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
  hlColor: $('hl-color'),
  hlPad: $('hl-pad'),
  hlApply: $('hl-apply'),
  hlClear: $('hl-clear'),
  bgColor: $('bg-color'),
  btnImage: $('btn-image'),
  imageFile: $('image-file'),
  overlay: $('img-overlay'),
  frame: $('img-frame'),
  handle: $('img-handle'),
  tools: $('img-tools'),
  widthRange: $('img-width'),
  widthVal: $('img-width-val'),
  removeImage: $('img-remove'),
};

const STORE_KEY = 'onepage.state.v3';

let layout = { fontPt: 11, columns: 1, lineHeight: 1.38, marginMm: 15, overflow: false, belowFloor: false };
let sheetRefs = null;
let fitter = null;
let docTitle = 'Document';
let running = false;
let queued = false;

/* --- Formatting state -------------------------------------------------- */

let fontScale = 1.0;
let pageBg = '#ffffff';
let hlColor = '#ffe066';
let hlPad = 4;

/** The image currently showing resize chrome, by its reference id. */
let selectedRef = null;
/** Set while a resize drag is in flight, so the fit does not fight the drag. */
let resizing = false;
let previewScale = 1;

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
  // Images the text no longer mentions are dead weight in a quota-limited
  // store, and the text is the only thing that decides what is still in use.
  images.keepOnly(referencedIds(els.source.value));

  try {
    localStorage.setItem(
      STORE_KEY,
      JSON.stringify({
        ...readSettings(),
        typeface: els.typeface.value,
        text: els.source.value,
        fontScale,
        pageBg,
        hlColor,
        hlPad,
      })
    );
  } catch {
    /* private mode or blocked storage */
  }

  if (!images.save() && referencedIds(els.source.value).length) {
    note('Images are too large to remember between sessions — the document still works', 'warn');
  }
}

function restoreState() {
  images.restore();
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
    if (typeof s.pageBg === 'string') pageBg = s.pageBg;
    if (typeof s.hlColor === 'string') hlColor = s.hlColor;
    if (typeof s.hlPad === 'number') hlPad = s.hlPad;
  } catch {
    /* ignore corrupt or unavailable storage */
  }

  els.bgColor.value = pageBg;
  els.hlColor.value = hlColor;
  els.hlPad.value = String(hlPad);
  updateFontSizeUI();
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
  update();
  saveState();
}

/* --- Preview scaling ---------------------------------------------------- */

function rescalePreview() {
  const pad = 48;
  const byWidth = (els.preview.clientWidth - pad) / A4.widthPx;
  const byHeight = (els.preview.clientHeight - pad) / A4.heightPx;
  previewScale = Math.max(0.2, Math.min(1, byWidth, byHeight));
  els.scaler.style.transform = `scale(${previewScale})`;
  els.scaler.style.height = `${A4.heightPx * previewScale}px`;
  // The overlay lives inside the scaler, so its controls must undo the zoom
  // or they would shrink along with the page they are meant to sit above.
  els.overlay.style.setProperty('--inv-scale', String(1 / previewScale));
  positionOverlay();
}

/* --- The pipeline ------------------------------------------------------- */

/**
 * Parse the source and resolve every image reference to its bytes.
 *
 * The parser only ever sees text, so it emits a reference; resolving it here
 * keeps the store out of the parser and means the preview and both exporters
 * are handed identical blocks.
 */
function currentDoc() {
  const blocks = parse(els.source.value).map((b) => {
    if (b.type !== 'image') return b;
    const stored = images.get(b.ref);
    if (!stored) return b;
    return { ...b, src: stored.src, naturalWidth: stored.width, naturalHeight: stored.height };
  });

  const first = blocks.find((b) => b.type === 'h1' || b.type === 'h2' || b.type === 'p');
  docTitle = first && first.text ? first.text.slice(0, 80) : 'Document';
  return { blocks, html: render(blocks) };
}

/**
 * An image with no layout box yet measures as zero height, so fitting before
 * it decodes produces a size that is far too large and the picture is then
 * pushed off the bottom of the page, where `overflow: hidden` silently eats
 * it. Waiting is the difference between the image appearing and not.
 */
async function imagesReady(body) {
  const pending = [...body.querySelectorAll('img')]
    .filter((img) => !img.complete || !img.naturalWidth)
    .map((img) => img.decode().catch(() => {}));
  if (pending.length) await Promise.all(pending);
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
    sheetRefs.page.style.setProperty('--doc-bg', pageBg);

    await imagesReady(sheetRefs.body);

    const words = wordCount(els.source.value);

    if (!html.trim()) {
      layout = {
        fontPt: 11, columns: 1, lineHeight: 1.38, marginMm: settings.marginMm,
        overflow: false, belowFloor: false, marginsReduced: false,
      };
      fitter.applyState(11 * fontScale, 1, 1.38, settings.marginMm);
      setStatus(0, null, face);
      setExportsEnabled(false);
      select(null);
      return;
    }

    layout = await fitter.fit({ ...settings, fontScale });
    setStatus(words, layout, face);
    setExportsEnabled(true);
    reselect();
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
    note(`${chosen.label} is not installed — using ${face.label}`, 'warn');
  } else if (fitted.overflow) {
    note('Too long for one page even at minimum size — trim the text', 'danger');
  } else if (fitted.belowFloor) {
    note(`Very small type (${fitted.fontPt.toFixed(1)}pt) — consider trimming`, 'warn');
  } else if (fitted.marginsReduced) {
    note(`Margins reduced to ${fitted.marginMm}mm to fit one page`, 'warn');
  } else {
    note('Fits one A4 page', 'ok');
  }
}

function note(message, level) {
  els.statNote.textContent = message;
  els.statNote.dataset.level = level;
}

/* --- Images ------------------------------------------------------------- */

/**
 * Import files and drop a reference for each at the caret, so a picture lands
 * where the user was typing and everything after it flows down.
 */
async function insertImages(files) {
  const list = [...files].filter((f) => f && String(f.type || '').startsWith('image/'));
  if (!list.length) return;

  for (const file of list) {
    try {
      const { id } = await images.importFile(file);
      const area = els.source;
      const { text, caret } = insertImageAt(
        area.value,
        area.selectionStart,
        area.selectionEnd,
        imageToken({ id, widthPct: DEFAULT_WIDTH_PCT, align: DEFAULT_ALIGN })
      );
      area.value = text;
      area.setSelectionRange(caret, caret);
      selectedRef = id;
    } catch (err) {
      note(`Could not add ${file.name || 'that image'}: ${err.message}`, 'danger');
      return;
    }
  }

  await update();
  saveState();
}

/** Find the rendered figure for a reference, if it is on the page. */
const figureFor = (ref) =>
  ref && sheetRefs ? sheetRefs.body.querySelector(`.doc-image[data-ref="${CSS.escape(ref)}"]`) : null;

function select(ref) {
  selectedRef = ref;
  if (sheetRefs) {
    for (const fig of sheetRefs.body.querySelectorAll('.doc-image.is-selected')) {
      fig.classList.remove('is-selected');
    }
  }
  const fig = figureFor(ref);
  if (!fig) {
    els.overlay.hidden = true;
    return;
  }
  fig.classList.add('is-selected');
  els.overlay.hidden = false;

  const block = parse(els.source.value).find((b) => b.type === 'image' && b.ref === ref);
  const pct = block ? block.widthPct : DEFAULT_WIDTH_PCT;
  els.widthRange.value = String(Math.round(pct));
  els.widthVal.textContent = `${Math.round(pct)}%`;
  for (const btn of els.tools.querySelectorAll('[data-align]')) {
    btn.setAttribute('aria-pressed', String(btn.dataset.align === (block ? block.align : DEFAULT_ALIGN)));
  }
  positionOverlay();
}

/** Re-attach the chrome after a re-render replaced the DOM it was tracking. */
function reselect() {
  if (!selectedRef) return;
  if (!figureFor(selectedRef)) {
    selectedRef = null;
    els.overlay.hidden = true;
    return;
  }
  select(selectedRef);
}

/**
 * The overlay is a sibling of the iframe inside the scaler, so it shares the
 * sheet's unscaled coordinate system and the image's own rect can be used
 * directly — no zoom arithmetic, and nothing to drift when the window resizes.
 */
function positionOverlay() {
  if (els.overlay.hidden || !sheetRefs) return;
  const fig = figureFor(selectedRef);
  if (!fig) return;
  const img = fig.querySelector('img');
  if (!img) return;

  const box = img.getBoundingClientRect();
  const root = sheetRefs.doc.documentElement.getBoundingClientRect();
  const left = box.left - root.left;
  const top = box.top - root.top;

  els.frame.style.left = `${left}px`;
  els.frame.style.top = `${top}px`;
  els.frame.style.width = `${box.width}px`;
  els.frame.style.height = `${box.height}px`;

  els.tools.style.left = `${left + box.width / 2}px`;
  // Above the image when there is room for the bar, below it otherwise.
  const barHeight = 34 / previewScale;
  els.tools.style.top = top > barHeight ? `${top - barHeight}px` : `${top + box.height + 8 / previewScale}px`;
}

function setWidth(pct, { commit }) {
  if (!selectedRef) return;
  const clamped = Math.max(10, Math.min(100, Math.round(pct)));
  els.widthVal.textContent = `${clamped}%`;
  els.widthRange.value = String(clamped);

  // While dragging, write straight to the figure: re-running the fitter on
  // every pointer move would make the drag lag behind the cursor and the page
  // jump under it as the type size changed.
  const fig = figureFor(selectedRef);
  if (fig) {
    fig.style.setProperty('--img-width', `${clamped}%`);
    positionOverlay();
  }

  if (!commit) return;
  els.source.value = updateImage(els.source.value, selectedRef, { widthPct: clamped });
  update().then(saveState);
}

function setAlign(align) {
  if (!selectedRef) return;
  els.source.value = updateImage(els.source.value, selectedRef, { align });
  update().then(saveState);
}

function dropSelectedImage() {
  if (!selectedRef) return;
  const ref = selectedRef;
  select(null);
  els.source.value = removeImage(els.source.value, ref);
  images.remove(ref);
  update().then(saveState);
}

function wireResize() {
  let startX = 0;
  let startWidth = 0;
  let containerWidth = 1;

  const onMove = (e) => {
    if (!resizing) return;
    // Pointer movement is in app pixels; the sheet is drawn scaled, so the
    // delta has to be converted back into the sheet's own pixels.
    const delta = (e.clientX - startX) / previewScale;
    const next = ((startWidth + delta) / containerWidth) * 100;
    setWidth(next, { commit: false });
  };

  const onUp = () => {
    if (!resizing) return;
    resizing = false;
    window.removeEventListener('pointermove', onMove);
    window.removeEventListener('pointerup', onUp);
    setWidth(Number(els.widthRange.value), { commit: true });
  };

  els.handle.addEventListener('pointerdown', (e) => {
    const fig = figureFor(selectedRef);
    const img = fig && fig.querySelector('img');
    if (!img) return;
    e.preventDefault();
    resizing = true;
    startX = e.clientX;
    startWidth = img.getBoundingClientRect().width;
    containerWidth = fig.getBoundingClientRect().width || 1;
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  });
}

/* --- Highlighting -------------------------------------------------------- */

/**
 * Highlighting rewrites the selected run of source text in place, so the
 * colour is part of the document rather than a decoration layered over it.
 */
function applyHighlight(clearInstead) {
  const area = els.source;
  const { selectionStart: s, selectionEnd: e, value } = area;
  if (s === e) {
    note('Select some text in the source pane first, then apply a highlight', 'warn');
    return;
  }

  const chosen = value.slice(s, e);
  const replacement = clearInstead ? stripHighlights(chosen) : highlight(chosen, hlColor, hlPad);
  area.value = value.slice(0, s) + replacement + value.slice(e);
  area.setSelectionRange(s, s + replacement.length);
  area.focus();
  update();
  saveState();
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
    bgColor: pageBg,
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

/**
 * Text wins over pixels.
 *
 * Word, Excel and most browsers put a bitmap of the selection on the
 * clipboard *alongside* the text. Checking for an image first therefore turns
 * every pasted document into a screenshot and drops the text on the floor, so
 * the image path is only taken when there is no text to be had.
 */
function onPaste(event) {
  const text = textFromPaste(event);
  if (text && text.trim()) {
    event.preventDefault();
    insertAtCursor(els.source, text);
    update();
    saveState();
    return;
  }

  const dt = event.clipboardData;
  if (!dt) return;
  const files = [...(dt.files || [])].filter((f) => String(f.type || '').startsWith('image/'));
  const fromItems = [...(dt.items || [])]
    .filter((i) => i.kind === 'file' && String(i.type || '').startsWith('image/'))
    .map((i) => i.getAsFile())
    .filter(Boolean);
  const all = files.length ? files : fromItems;
  if (!all.length) return;

  event.preventDefault();
  insertImages(all);
}

async function withBusy(button, fn) {
  const label = button.textContent;
  button.disabled = true;
  button.textContent = 'Working…';
  try {
    await fn();
  } catch (err) {
    console.error(err);
    note(`Export failed: ${err && err.message ? err.message : err}`, 'danger');
  } finally {
    button.textContent = label;
    button.disabled = false;
  }
}

/** Dropping an image anywhere in the app adds it; dropping a file elsewhere must not navigate. */
function wireDrop() {
  const hasImage = (e) =>
    e.dataTransfer && [...(e.dataTransfer.items || [])].some((i) => String(i.type || '').startsWith('image/'));

  for (const target of [els.source, els.preview]) {
    target.addEventListener('dragover', (e) => {
      if (!hasImage(e)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
      target.classList.add('is-drop-target');
    });
    target.addEventListener('dragleave', () => target.classList.remove('is-drop-target'));
    target.addEventListener('drop', (e) => {
      target.classList.remove('is-drop-target');
      if (!e.dataTransfer) return;
      e.preventDefault();
      insertImages(e.dataTransfer.files);
    });
  }

  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());
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
    images.clear();
    select(null);
    els.source.focus();
    update();
    saveState();
  });

  els.fontUp.addEventListener('click', () => adjustFontScale(0.1));
  els.fontDown.addEventListener('click', () => adjustFontScale(-0.1));

  els.hlColor.addEventListener('input', () => {
    hlColor = els.hlColor.value;
    saveState();
  });
  els.hlPad.addEventListener('change', () => {
    hlPad = Math.max(0, Math.min(20, Number(els.hlPad.value) || 0));
    els.hlPad.value = String(hlPad);
    saveState();
  });
  els.hlApply.addEventListener('click', () => applyHighlight(false));
  els.hlClear.addEventListener('click', () => applyHighlight(true));

  els.bgColor.addEventListener('input', () => {
    pageBg = els.bgColor.value;
    if (sheetRefs) sheetRefs.page.style.setProperty('--doc-bg', pageBg);
    saveState();
  });

  els.btnImage.addEventListener('click', () => els.imageFile.click());
  els.imageFile.addEventListener('change', () => {
    insertImages(els.imageFile.files);
    els.imageFile.value = '';
  });

  wireDrop();
  wireResize();

  // Selecting an image: a click inside the sheet, which is a separate document.
  sheetRefs.doc.addEventListener('mousedown', (e) => {
    const fig = e.target.closest ? e.target.closest('.doc-image') : null;
    select(fig ? fig.dataset.ref : null);
  });

  els.widthRange.addEventListener('input', () => setWidth(Number(els.widthRange.value), { commit: false }));
  els.widthRange.addEventListener('change', () => setWidth(Number(els.widthRange.value), { commit: true }));
  els.removeImage.addEventListener('click', dropSelectedImage);
  for (const btn of els.tools.querySelectorAll('[data-align]')) {
    btn.addEventListener('click', () => setAlign(btn.dataset.align));
  }

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
  els.preview.addEventListener('scroll', positionOverlay, { passive: true });

  window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
      e.preventDefault();
      els.print.click();
      return;
    }
    if (e.key === 'Escape' && selectedRef) select(null);
    if ((e.key === 'Delete' || e.key === 'Backspace') && selectedRef && document.activeElement !== els.source) {
      e.preventDefault();
      dropSelectedImage();
    }
  });
}

let flashTimer = 0;
function flash(message) {
  note(message, 'ok');
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
  await update();
  els.source.focus();
  document.documentElement.dataset.ready = 'true';
}

boot();
