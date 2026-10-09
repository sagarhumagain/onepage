/**
 * main.js — wiring.
 *
 * Pipeline: source text -> parse -> render -> inject into the sheet -> fit.
 * Print, PDF and Word all read the same fitted layout, so the three outputs
 * and the preview always agree.
 *
 * The page is also the editor. Typing into it, sizing an image, colouring a
 * run of words — each is an edit to the document model, and the text in the
 * left pane stays the document of record: an edited page is turned straight
 * back into source text by the same walker that reads pasted Word HTML. That
 * is what keeps one representation, one undo history, and one thing to export.
 */

import { parse, wordCount, IMAGE_REF } from './parse.js';
import { render } from './render.js';
import { createFitter, A4 } from './fit.js';
import { mountSheet, standaloneHtml } from './sheet.js';
import { textFromPaste, htmlToText, richFromPaste, marksFromRecords } from './clipboard.js';
import { TYPEFACES, resolveTypeface } from './scale.js';
import { exportDocx, exportPdf, printSheet } from './exporters.js';
import { applyMarks, collectUnits, storedMark, formatRange, clearRange, allHave, fmtAt, stepSize } from './marks.js';
import { footerHtml, hasFooter, restoreFooter, serializeFooter } from './footer.js';
import defaultLeft from './assets/footer-left.png?inline';
import defaultRight from './assets/footer-right.png?inline';

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
  insetToggle: $('padding-toggle'),
  highlight: $('btn-highlight'),
  insert: $('insert'),
  imageFile: $('image-file'),
  undo: $('btn-undo'),
  redo: $('btn-redo'),
  imgTools: $('img-tools'),
  imgFrame: $('img-frame'),
  imgBar: $('img-bar'),
  imgSize: $('img-size'),
  imgWidth: $('img-width'),
  fmtTools: $('fmt-tools'),
  fmtBar: $('fmt-bar'),
  fmtSize: $('fmt-size'),
  fmtFg: $('fmt-fg'),
  fmtBg: $('fmt-bg'),
  fmtPad: $('fmt-pad'),
  footTools: $('foot-tools'),
  footFrame: $('foot-frame'),
  footBar: $('foot-bar'),
  footerFile: $('footer-file'),
};

const STORE_KEY = 'onepage.state.v3';
const LEGACY_KEY = 'onepage.state.v2';

let layout = {
  fontPt: 11, columns: 1, lineHeight: 1.38, marginMm: 15,
  imageScale: 1, overflow: false, belowFloor: false,
};
let sheetRefs = null;
let fitter = null;
let docTitle = 'Document';
let running = false;
let queued = null;

/* --- Document state ----------------------------------------------------- */

let fontScale = 1.0;
let bgColor = '#ffffff';
let inset = false;
let hlColor = '#ffe08a';
/** Padding around a highlight, in tenths of an em; 1 is the look it always had. */
let hlPad = 1;
/** The default padding is the stylesheet's, so it is stored as no padding at all. */
const padMark = (pad) => (pad === 1 ? null : pad);
let fgColor = '#b42318';

/** id -> {src, alt, w, h, width, align, wrap}. Positions live in the text. */
let images = {};
/**
 * The two logos at the foot of every page; see footer.js. Until something
 * else is chosen they are the bundled defaults, which are saved as the word
 * "default" rather than as their bytes.
 */
const DEFAULT_FOOTER = {
  left: { src: defaultLeft, alt: 'Government of Nepal', w: 210, h: 212, preset: true },
  right: { src: defaultRight, alt: 'World Health Organization Nepal', w: 417, h: 214, preset: true },
};
let footer = { ...DEFAULT_FOOTER };
/** The footer slot being edited, 'left' or 'right', or null. */
let selectedSlot = null;
/** The slot a picked file goes into. */
let footerTarget = null;
/** Content-anchored run formatting; see marks.js. */
let marks = [];

/** Rebuilt on every render: the markable runs of text, in reading order. */
let units = [];
/** The current selection inside the sheet, as ranges over those units. */
let pendingSelection = null;
/**
 * Where the selection was, counted in characters from the start of the page,
 * to be put back after the next re-render. Characters rather than unit ranges:
 * a re-render can change which units exist — typing a new line makes one — and
 * the caret still has to land where it was.
 */
let keepOffsets = null;
/** True when the page has been typed into and the model has not caught up. */
let pageDirty = false;

let selectedImage = null;
let drag = null;
let previewScale = 1;

const clamp = (n, lo, hi) => Math.max(lo, Math.min(hi, n));

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

/** The images the text actually refers to. Anything else is not saved. */
function usedImageIds() {
  const ids = new Set();
  for (const line of els.source.value.split('\n')) {
    const m = line.match(IMAGE_REF);
    if (m) ids.add(m[1]);
  }
  return ids;
}

let storageWarned = false;

function saveState() {
  const kept = {};
  for (const id of usedImageIds()) if (images[id]) kept[id] = images[id];

  try {
    localStorage.setItem(
      STORE_KEY,
      JSON.stringify({
        ...readSettings(),
        typeface: els.typeface.value,
        text: els.source.value,
        fontScale,
        bgColor,
        inset,
        hlColor,
        hlPad,
        fgColor,
        images: kept,
        marks,
        footer: serializeFooter(footer),
      })
    );
  } catch {
    // Pictures are data URLs and localStorage is a few megabytes. Say so once
    // rather than letting the document look saved when it is not.
    if (!storageWarned && Object.keys(kept).length) {
      storageWarned = true;
      flash('Too much image data to remember between sessions — export before closing');
      els.statNote.dataset.level = 'warn';
    }
  }
}

function restoreState() {
  let s = null;
  try {
    s = JSON.parse(localStorage.getItem(STORE_KEY) || 'null');
  } catch {
    /* corrupt or unavailable storage */
  }
  if (!s) s = migrateLegacy();
  if (!s) return;

  if (typeof s.text === 'string') els.source.value = s.text;
  if (s.marginMm) els.margin.value = String(s.marginMm);
  if (s.typeface && TYPEFACES[s.typeface]) els.typeface.value = s.typeface;
  if (s.columns != null) els.columns.value = String(s.columns);
  if (typeof s.fillPage === 'boolean') els.fill.checked = s.fillPage;
  if (typeof s.fontScale === 'number') fontScale = s.fontScale;
  if (typeof s.bgColor === 'string') bgColor = s.bgColor;
  if (typeof s.inset === 'boolean') inset = s.inset;
  if (typeof s.hlColor === 'string') hlColor = s.hlColor;
  if (typeof s.hlPad === 'number') hlPad = s.hlPad;
  if (typeof s.fgColor === 'string') fgColor = s.fgColor;
  if (s.images && typeof s.images === 'object') images = s.images;
  if (Array.isArray(s.marks)) marks = s.marks;
  footer = restoreFooter(s.footer, DEFAULT_FOOTER);

  els.bgColor.value = bgColor;
  els.fmtBg.value = hlColor;
  els.fmtPad.value = String(hlPad);
  els.fmtFg.value = fgColor;
  els.insetToggle.checked = inset;
}

/**
 * The previous version stored images as blocks pinned to a block index, which
 * drifted the moment the text above them was edited. Convert them into the
 * text-anchored form by appending a reference for each, in order.
 */
function migrateLegacy() {
  let old = null;
  try {
    old = JSON.parse(localStorage.getItem(LEGACY_KEY) || 'null');
  } catch {
    return null;
  }
  if (!old) return null;

  const refs = [];
  for (const b of Array.isArray(old.imageBlocks) ? old.imageBlocks : []) {
    if (!b || !b.src) continue;
    const id = newImageId();
    images[id] = {
      src: b.src,
      alt: b.alt || '',
      w: 0,
      h: 0,
      width: 60,
      align: b.align === 'left' || b.align === 'right' ? b.align : 'center',
      wrap: false,
    };
    refs.push(`[image:${id}]`);
  }

  const text = [old.text || '', ...refs].filter(Boolean).join('\n\n');
  try {
    localStorage.removeItem(LEGACY_KEY);
  } catch {
    /* ignore */
  }
  return { ...old, text, images, marks: [] };
}

/* --- Document font size ------------------------------------------------- */

function updateFontSizeUI() {
  els.fontSizeVal.textContent = Math.round(fontScale * 11);
  els.fontDown.disabled = fontScale <= 0.3;
  els.fontUp.disabled = fontScale >= 3.0;
}

function adjustFontScale(delta) {
  fontScale = clamp(fontScale + delta, 0.3, 3.0);
  updateFontSizeUI();
  commit('setting');
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
  syncTools();
}

/* --- The pipeline ------------------------------------------------------- */

/** An image reference in the text, resolved against the picture library. */
function imageBlock(id) {
  const rec = images[id];
  // A reference with no picture behind it is shown as the text it is, rather
  // than disappearing and taking the user's placeholder with it.
  if (!rec) return { type: 'p', text: `[image:${id}]` };
  return {
    type: 'image',
    id,
    src: rec.src,
    alt: rec.alt || '',
    caption: rec.caption || '',
    width: rec.width,
    align: rec.align,
    wrap: Boolean(rec.wrap),
    w: rec.w,
    h: rec.h,
  };
}

function currentDoc(opts = {}) {
  const blocks = parse(els.source.value).map((b) => (b.type === 'image' ? imageBlock(b.id) : b));
  const docUnits = applyMarks(blocks, marks);

  const first = blocks.find((b) => b.type === 'h1' || b.type === 'h2' || b.type === 'p');
  docTitle = first && first.text ? first.text.slice(0, 80) : 'Document';

  return { blocks, units: docUnits, html: render(blocks, opts) };
}

/**
 * Re-render the page from the model, then fit it.
 *
 * @param {{refit?: boolean, keepDom?: boolean}} [opts]
 *   `refit: false` — a colour changes no geometry, so re-render without paying
 *   for a fresh binary search.
 *   `keepDom: true` — the page is being typed into: fit what is on the screen
 *   and leave the DOM, and the caret sitting in it, alone.
 */
async function update(opts = {}) {
  if (running) {
    queued = opts;
    return;
  }
  running = true;
  try {
    const settings = readSettings();
    const face = currentFace();

    if (!opts.keepDom) {
      const doc = currentDoc({ editor: true });
      units = doc.units;
      pendingSelection = null; // offsets belong to the DOM about to be replaced
      sheetRefs.body.innerHTML = doc.html;
      pageDirty = false;
    }

    sheetRefs.page.style.setProperty('--doc-font-family', face.css);
    sheetRefs.page.style.setProperty('--doc-bg', bgColor);
    sheetRefs.page.classList.toggle('inset', inset);
    renderFooter();

    const words = wordCount(els.source.value);
    const hasImages = Boolean(sheetRefs.body.querySelector('.doc-image'));

    if (!hasImages && !sheetRefs.body.textContent.trim()) {
      layout = {
        fontPt: 11, columns: 1, lineHeight: 1.38, marginMm: settings.marginMm,
        imageScale: 1, overflow: false, belowFloor: false, marginsReduced: false,
      };
      fitter.applyState(11 * fontScale, 1, 1.38, settings.marginMm, 1);
      setStatus(0, null, face);
      setExportsEnabled(false);
      return;
    }

    if (opts.refit === false) {
      fitter.applyState(layout.fontPt, layout.columns, layout.lineHeight, layout.marginMm, layout.imageScale);
    } else {
      layout = await fitter.fit({ ...settings, fontScale, hasImages });
    }

    setStatus(words, layout, face);
    setExportsEnabled(true);
  } finally {
    running = false;
    restoreOffsets();
    restoreImageSelection();
    syncFooterTools();
    const next = queued;
    queued = null;
    if (next) update(next);
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
  } else if (fitted.imagesShrunk) {
    els.statNote.textContent = `Images scaled to ${Math.round(fitted.imageScale * 100)}% to fit one page`;
    els.statNote.dataset.level = 'warn';
  } else if (fitted.marginsReduced) {
    els.statNote.textContent = `Margins reduced to ${fitted.marginMm}mm to fit one page`;
    els.statNote.dataset.level = 'warn';
  } else {
    els.statNote.textContent = 'Fits one A4 page';
    els.statNote.dataset.level = 'ok';
  }
}

/* --- Images ------------------------------------------------------------- */

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
    img.onerror = () => resolve({ width: 0, height: 0 });
    img.src = src;
  });
}

function newImageId() {
  let id;
  do {
    id = Math.random().toString(36).slice(2, 8);
  } while (images[id]);
  return id;
}

/**
 * Put the reference on a line of its own where the cursor is.
 *
 * Its own line is what makes it a block of the document like any other, so the
 * text above it and the text below it keep flowing around it — an image is
 * never the end of the page.
 */
function insertImageRef(id) {
  const ta = els.source;
  const token = `[image:${id}]`;
  const at = ta.selectionStart == null ? ta.value.length : ta.selectionStart;
  const to = ta.selectionEnd == null ? at : ta.selectionEnd;

  const before = ta.value.slice(0, at).replace(/\s+$/, '');
  const after = ta.value.slice(to).replace(/^\s+/, '');
  ta.value = [before, token, after].filter((p) => p !== '').join('\n\n');

  const pos = (before ? before.length + 2 : 0) + token.length;
  ta.setSelectionRange(pos, pos);
}

/** Longest edge kept for an imported picture: well past 300dpi on A4. */
const MAX_IMAGE_EDGE = 1600;

/**
 * A phone photo is several megapixels and nothing on an A4 page can use more
 * than a fraction of that; kept whole it fills the browser's storage and slows
 * every fit. A picture with transparency stays a PNG; anything else becomes a
 * JPEG.
 */
async function downscaled(src, type, width, height) {
  const scale = MAX_IMAGE_EDGE / Math.max(width, height);
  if (!(scale < 1)) return { src, width, height };
  const img = new Image();
  await new Promise((resolve, reject) => {
    img.onload = resolve;
    img.onerror = reject;
    img.src = src;
  });
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(width * scale);
  canvas.height = Math.round(height * scale);
  canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
  const out = type === 'image/png' ? canvas.toDataURL('image/png') : canvas.toDataURL('image/jpeg', 0.86);
  return { src: out, width: canvas.width, height: canvas.height };
}

async function addImageFromFile(file) {
  if (!file || !file.type.startsWith('image/')) return;
  const read = await readFileAsDataURL(file);
  const natural = await getImageDimensions(read);
  const { src, width, height } = await downscaled(read, file.type, natural.width, natural.height);

  const id = newImageId();
  images[id] = {
    src,
    alt: String(file.name || '').replace(/\.[a-z0-9]+$/i, ''),
    w: width,
    h: height,
    // A portrait photograph at 60% of the column is most of the page; start it
    // smaller so there is still a document around it.
    width: height > width * 1.2 ? 40 : 60,
    align: 'center',
    wrap: false,
  };

  insertImageRef(id);
  commit('image');
  await update();
  selectImage(id);
  saveState();
}

function handleImagePaste(event) {
  const dt = event.clipboardData;
  if (!dt || !dt.items) return false;
  for (const item of dt.items) {
    if (item.type.startsWith('image/')) {
      event.preventDefault();
      const file = item.getAsFile();
      if (file) addImageFromFile(file);
      return true;
    }
  }
  return false;
}

function removeImage(id) {
  els.source.value = els.source.value
    .split('\n')
    .filter((line) => {
      const m = line.match(IMAGE_REF);
      return !m || m[1] !== id;
    })
    .join('\n')
    .replace(/\n{3,}/g, '\n\n');

  delete images[id];
  selectImage(null);
  commit('image');
  update();
  saveState();
}

/* --- The footer ---------------------------------------------------------- */

let footerShown = null;

/**
 * Draw the footer into the page. Its pictures are only replaced when they
 * change: a re-fit on every keystroke must not make the logos flash.
 */
function renderFooter() {
  const html = footerHtml(footer);
  if (html !== footerShown) {
    sheetRefs.footer.innerHTML = html;
    footerShown = html;
  }
  if (hasFooter(footer)) sheetRefs.page.setAttribute('data-footer', '1');
  else sheetRefs.page.removeAttribute('data-footer');
  const slot = selectedSlot && sheetRefs.footer.querySelector(`[data-slot="${selectedSlot}"]`);
  for (const el of sheetRefs.footer.querySelectorAll('.footer-slot')) el.toggleAttribute('data-selected', el === slot);
  if (slot) slot.setAttribute('data-selected', '1');
}

function selectFooterSlot(slot) {
  selectedSlot = slot && footer[slot] ? slot : null;
  if (selectedSlot) {
    selectImage(null);
    const sel = sheetRefs.win.getSelection();
    if (sel) sel.removeAllRanges();
    pendingSelection = null;
  }
  renderFooter();
  syncTools();
}

function syncFooterTools() {
  const el = selectedSlot && sheetRefs ? sheetRefs.footer.querySelector(`[data-slot="${selectedSlot}"] img`) : null;
  if (!el) {
    els.footTools.hidden = true;
    return;
  }
  els.footTools.hidden = false;
  const box = paneBox(el.getBoundingClientRect());
  els.footFrame.style.left = `${box.left}px`;
  els.footFrame.style.top = `${box.top}px`;
  els.footFrame.style.width = `${box.width}px`;
  els.footFrame.style.height = `${box.height}px`;
  placeBar(els.footBar, box);
}

/** One change to the footer: one undo step, and a re-fit, since it moves the text's foot. */
function setFooterSlot(slot, image) {
  footer = { ...footer, [slot]: image };
  if (!image && selectedSlot === slot) selectedSlot = null;
  commit('footer');
  update();
  saveState();
}

function pickFooterImage(slot) {
  footerTarget = slot;
  els.footerFile.click();
}

async function footerImageFromFile(file) {
  const slot = footerTarget;
  footerTarget = null;
  if (!slot || !file || !file.type.startsWith('image/')) return;
  const src = await readFileAsDataURL(file);
  const { width, height } = await getImageDimensions(src);
  setFooterSlot(slot, { src, alt: String(file.name || '').replace(/\.[a-z0-9]+$/i, ''), w: width, h: height });
}

function onFooterAction(event) {
  const act = event.target && event.target.dataset ? event.target.dataset.act : null;
  if (!act || !selectedSlot) return;
  if (act === 'replace') pickFooterImage(selectedSlot);
  else if (act === 'remove') setFooterSlot(selectedSlot, null);
}

/* --- Image selection and resizing --------------------------------------- */

const figureFor = (id) => (id ? sheetRefs.body.querySelector(`.doc-image[data-img="${id}"]`) : null);

function selectImage(id) {
  const previous = figureFor(selectedImage);
  if (previous) previous.removeAttribute('data-selected');
  selectedImage = id && images[id] ? id : null;
  const next = figureFor(selectedImage);
  if (next && selectedSlot) {
    selectedSlot = null;
    renderFooter();
  }
  if (next) {
    next.setAttribute('data-selected', '1');
    // A picture and a run of words are never both "the thing being edited".
    const sel = sheetRefs.win.getSelection();
    if (sel) sel.removeAllRanges();
    pendingSelection = null;
  }
  syncTools();
}

/** After a re-render the figure is a new element, so re-mark and re-measure. */
function restoreImageSelection() {
  if (selectedImage && !figureFor(selectedImage)) selectedImage = null;
  const fig = figureFor(selectedImage);
  if (fig) fig.setAttribute('data-selected', '1');
  syncTools();
}

/**
 * A rectangle inside the sheet, in the preview pane's own coordinates.
 *
 * The sheet is an iframe scaled by a CSS transform, so a rectangle read inside
 * it is in unscaled page pixels; the sheet's own rect is already transformed.
 * Scrolling is added back because the tools are absolutely positioned inside
 * the scrolling pane and move with its content.
 */
function paneBox(rect) {
  const frame = els.sheet.getBoundingClientRect();
  const pane = els.preview.getBoundingClientRect();
  return {
    left: frame.left + rect.left * previewScale - pane.left + els.preview.scrollLeft,
    top: frame.top + rect.top * previewScale - pane.top + els.preview.scrollTop,
    width: rect.width * previewScale,
    height: rect.height * previewScale,
  };
}

/** Place a floating bar above a box, or below it when there is no room. */
function placeBar(bar, box) {
  const height = bar.offsetHeight || 30;
  const above = box.top - height - 10;
  bar.style.left = `${Math.max(4, box.left)}px`;
  bar.style.top = `${above > 4 ? above : box.top + box.height + 10}px`;
}

function syncTools() {
  syncImageTools();
  syncTextTools();
  syncFooterTools();
}

function setFrame(box) {
  els.imgFrame.style.left = `${box.left}px`;
  els.imgFrame.style.top = `${box.top}px`;
  els.imgFrame.style.width = `${box.width}px`;
  els.imgFrame.style.height = `${box.height}px`;
}

function syncImageTools() {
  const fig = figureFor(selectedImage);
  if (!fig) {
    els.imgTools.hidden = true;
    return;
  }
  const rec = images[selectedImage];
  els.imgTools.hidden = false;

  const box = paneBox(fig.getBoundingClientRect());
  setFrame(box);
  placeBar(els.imgBar, box);

  els.imgSize.textContent = `${Math.round(rec.width)}%`;
  els.imgWidth.value = String(Math.round(rec.width));
  for (const btn of els.imgBar.querySelectorAll('.img-btn')) {
    const act = btn.dataset.act;
    if (act === 'left' || act === 'center' || act === 'right') {
      btn.setAttribute('aria-pressed', String(rec.align === act));
    } else if (act === 'wrap') {
      btn.setAttribute('aria-pressed', String(Boolean(rec.wrap)));
      btn.disabled = rec.align === 'center';
    }
  }
}

function startResize(event) {
  const corner = event.target && event.target.dataset ? event.target.dataset.corner : null;
  if (!corner) return;
  const fig = figureFor(selectedImage);
  if (!fig) return;

  const rec = images[selectedImage];
  const scale = layout.imageScale || 1;
  const width = fig.getBoundingClientRect().width;

  event.preventDefault();
  drag = {
    corner,
    startX: event.clientX,
    startWidth: rec.width,
    figWidth: width,
    // The figure's width *is* its percentage of the column, so the column's
    // width falls straight out of it — no assumption about the layout needed.
    columnWidth: width / ((rec.width * scale) / 100),
    scale,
    fig,
  };
  try {
    event.target.setPointerCapture(event.pointerId);
  } catch {
    // A pointer that is no longer active cannot be captured; the window-level
    // move and up listeners still carry the drag.
  }
  document.body.classList.add('resizing');
}

function moveResize(event) {
  if (!drag) return;
  const dx = (event.clientX - drag.startX) / previewScale;
  const direction = drag.corner.endsWith('e') ? 1 : -1;
  const pct = clamp(((drag.figWidth + direction * dx) / drag.columnWidth) * (100 / drag.scale), 10, 100);
  drag.fig.style.setProperty('--img-w', String(Math.round(pct)));
  els.imgSize.textContent = `${Math.round(pct)}%`;
  setFrame(paneBox(drag.fig.getBoundingClientRect()));
}

function endResize() {
  if (!drag) return;
  const applied = Number(drag.fig.style.getPropertyValue('--img-w')) || drag.startWidth;
  const changed = Math.round(applied) !== Math.round(drag.startWidth);
  drag = null;
  document.body.classList.remove('resizing');
  if (!selectedImage || !changed) return;
  images[selectedImage].width = clamp(Math.round(applied), 10, 100);
  commit('image');
  update();
  saveState();
}

/** Esc: put the picture back the size it was and stop resizing. */
function cancelResize() {
  if (!drag) return;
  const { fig, startWidth } = drag;
  drag = null;
  document.body.classList.remove('resizing');
  fig.style.setProperty('--img-w', String(startWidth));
  if (selectedImage) images[selectedImage].width = startWidth;
  syncImageTools();
}

function onImageAction(event) {
  const act = event.target && event.target.dataset ? event.target.dataset.act : null;
  if (!act || !selectedImage) return;
  const rec = images[selectedImage];

  if (act === 'remove') {
    removeImage(selectedImage);
    return;
  }
  if (act === 'wrap') {
    rec.wrap = !rec.wrap;
    // Text cannot run beside something that is centred in the column.
    if (rec.wrap && rec.align === 'center') rec.align = 'left';
  } else {
    rec.align = act;
    if (act === 'center') rec.wrap = false;
  }
  commit('image');
  update();
  saveState();
}

/* --- Selection ---------------------------------------------------------- */

/**
 * The current selection inside the sheet, as ranges over markable units.
 *
 * Offsets are measured in the text the reader sees, which is also the space
 * marks.js stores them in, so nothing has to be mapped back through the
 * inline markers in the source.
 */
function readSelection() {
  if (!sheetRefs) return null;
  const sel = sheetRefs.win.getSelection();
  if (!sel || !sel.rangeCount || sel.isCollapsed) return null;

  const body = sheetRefs.body;
  const range = sel.getRangeAt(0);
  if (!body.contains(range.commonAncestorContainer) && range.commonAncestorContainer !== body) return null;

  const absolute = (node, offset) => {
    const probe = sheetRefs.doc.createRange();
    probe.selectNodeContents(body);
    try {
      probe.setEnd(node, offset);
    } catch {
      return 0;
    }
    return probe.toString().length;
  };

  const start = absolute(range.startContainer, range.startOffset);
  const end = start + range.toString().length;
  if (end <= start) return null;

  const out = [];
  // One walk for every unit's starting offset. Asking a Range for each one
  // instead is quadratic, and this runs on every selectionchange — which is
  // every mouse move while dragging a selection out.
  const walker = sheetRefs.doc.createTreeWalker(body, 1 /* elements */ | 4 /* text */);
  let base = 0;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node.nodeType === 3) {
      base += node.nodeValue.length;
      continue;
    }
    if (!node.dataset || node.dataset.u == null) continue;
    const index = Number(node.dataset.u);
    if (!units[index]) continue;
    const length = ownLength(node);
    const from = Math.max(start, base) - base;
    const to = Math.min(end, base + length) - base;
    if (to > from) out.push({ index, start: from, end: to });
  }
  return out.length ? out : null;
}

/**
 * A unit's own length on the page. A list item's element also holds the list
 * nested under it, whose items are units of their own; counting their text
 * here would stretch the parent over its children.
 */
function ownLength(el) {
  let length = el.textContent.length;
  for (const child of el.children) {
    const tag = child.tagName.toUpperCase();
    if (tag === 'UL' || tag === 'OL') length -= child.textContent.length;
  }
  return length;
}

/** The text node and offset that a unit's character offset lands on. */
function locate(el, offset) {
  const walker = sheetRefs.doc.createTreeWalker(el, 4 /* text */);
  let seen = 0;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const length = node.nodeValue.length;
    if (seen + length >= offset) return { node, offset: offset - seen };
    seen += length;
  }
  return { node: el, offset: el.childNodes.length };
}

/** The selection as characters from the start of the page, caret included. */
function readSelectionOffsets() {
  if (!sheetRefs) return null;
  const sel = sheetRefs.win.getSelection();
  if (!sel || !sel.rangeCount) return null;

  const body = sheetRefs.body;
  const range = sel.getRangeAt(0);
  if (!body.contains(range.commonAncestorContainer) && range.commonAncestorContainer !== body) return null;

  const probe = sheetRefs.doc.createRange();
  probe.selectNodeContents(body);
  try {
    probe.setEnd(range.startContainer, range.startOffset);
  } catch {
    return null;
  }
  const start = probe.toString().length;
  return { start, end: start + range.toString().length };
}

/**
 * Put the selection back after a re-render.
 *
 * Re-rendering replaces every node in the page, so a caret or a selection has
 * to be re-found by counting characters. Keeping it is what lets someone type
 * a line and immediately colour it, or highlight a phrase and then enlarge it,
 * without reaching for the mouse again in between.
 */
function restoreOffsets() {
  const wanted = keepOffsets;
  keepOffsets = null;
  if (!wanted) return;

  const body = sheetRefs.body;
  try {
    const a = locate(body, wanted.start);
    const b = locate(body, wanted.end);
    const range = sheetRefs.doc.createRange();
    range.setStart(a.node, a.offset);
    range.setEnd(b.node, b.offset);
    const sel = sheetRefs.win.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    pendingSelection = readSelection();
  } catch {
    pendingSelection = null;
  }
  syncTextTools();
}

/* --- Typing on the page -------------------------------------------------- */

/**
 * The page is contenteditable, and what comes out of it is text.
 *
 * Nothing tries to patch the model from a DOM diff. The edited page is walked
 * back into source text — by `htmlToText`, the same walker that reads pasted
 * Word HTML, so there is only one place where markup becomes text — and that
 * text is the document from then on. The page is only re-rendered from it when
 * something actually needs the model, which is what stops the structure
 * rearranging itself under the caret mid-sentence.
 */
function enablePageEditing() {
  const { doc, body } = sheetRefs;
  body.setAttribute('contenteditable', 'true');
  body.setAttribute('spellcheck', 'false');
  // Enter should split the paragraph, not wrap the new line in a bare <div>.
  try {
    doc.execCommand('defaultParagraphSeparator', false, 'p');
  } catch {
    /* older engines keep their own default */
  }

  body.addEventListener('input', onPageInput);
  body.addEventListener('paste', onPagePaste);
  body.addEventListener('blur', () => settlePage());
}

const sourceFromPage = () => htmlToText(sheetRefs.body.innerHTML);

const refitPage = debounce(() => update({ keepDom: true }), 160);

function onPageInput(kind = 'type') {
  pageDirty = true;
  // Written on every keystroke, not on a timer: every other path in the app
  // reads the source, and none of them should ever see a stale one.
  els.source.value = sourceFromPage();
  commit(kind);
  refitPage();
  saveState();
}

/** Re-render from the text, putting the caret back where the typing left it. */
async function settlePage() {
  if (!pageDirty) return false;
  keepOffsets = readSelectionOffsets();
  await update();
  return true;
}

/* --- Inserting structure ------------------------------------------------- */

/**
 * What the Insert menu drops in, written in the same conventions the source
 * pane uses. Nothing here is HTML: each template is run through the app's own
 * parse and render, so an inserted table is the same markup as a pasted one
 * and there is no second idea of what a table looks like.
 */
const INSERTS = {
  table: '| Column A | Column B | Column C |\n| --- | --- | --- |\n|  |  |  |\n|  |  |  |',
  ul: '- First item\n- Second item\n- Third item',
  ol: '1. First item\n2. Second item\n3. Third item',
  h2: '## Section heading',
  quote: '> Quoted text',
  hr: '---',
};

/** The top-level block of the page the caret is sitting in. */
function blockAtCaret() {
  const sel = sheetRefs.win.getSelection();
  if (!sel || !sel.rangeCount) return null;
  let node = sel.getRangeAt(0).startContainer;
  if (!sheetRefs.body.contains(node)) return null;
  while (node && node.parentNode !== sheetRefs.body) node = node.parentNode;
  return node && node.parentNode === sheetRefs.body ? node : null;
}

async function insertStructure(kind) {
  if (kind === 'image') {
    els.imageFile.click();
    return;
  }
  const template = INSERTS[kind];
  if (!template) return;

  // Typing in the source pane? Then the insert belongs at the source cursor,
  // where the user is actually looking.
  if (document.activeElement === els.source) {
    insertAtCursor(els.source, `\n\n${template}\n\n`);
    commit('insert');
    update();
    saveState();
    return;
  }

  const holder = sheetRefs.doc.createElement('div');
  holder.innerHTML = render(parse(template));
  const nodes = [...holder.childNodes].filter((n) => n.nodeType === 1);
  if (!nodes.length) return;

  const anchor = blockAtCaret();
  const before = anchor ? anchor.nextSibling : null;
  for (const node of nodes) sheetRefs.body.insertBefore(node, before);

  // Leave the first words of what was inserted selected, so typing replaces
  // the placeholder instead of having to delete it first.
  const first = sheetRefs.doc.createTreeWalker(nodes[0], 4 /* text */).nextNode();
  if (first) {
    const range = sheetRefs.doc.createRange();
    range.setStart(first, 0);
    range.setEnd(first, first.nodeValue.length);
    const sel = sheetRefs.win.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }

  onPageInput('insert');
  await settlePage();
}

/* --- Undo and redo ------------------------------------------------------- */

/**
 * One history for everything.
 *
 * The browser has its own undo stack for a contenteditable, and it is the
 * wrong one: it would put DOM back without putting the model back, and it
 * knows nothing about resizing an image or colouring a word. So every change
 * commits a snapshot of the whole document — text, pictures, marks, settings —
 * and Ctrl/Cmd+Z is taken over everywhere, including inside the page.
 *
 * Runs of typing coalesce, so undo steps back a phrase at a time rather than a
 * letter at a time.
 */
const HISTORY_LIMIT = 120;
const COALESCE_MS = 700;
/** Changes that arrive in a stream: a run of typing, a colour wheel dragged. */
const COALESCING = new Set(['type', 'colour']);

let committed = null;
let undoStack = [];
let redoStack = [];
let lastCommitAt = 0;
let lastCommitKind = '';

const cloneImages = () => {
  const out = {};
  // The data URL is a string, so this shares the picture and copies only the
  // handful of numbers around it.
  for (const [id, rec] of Object.entries(images)) out[id] = { ...rec };
  return out;
};

const snapshot = () => ({
  text: els.source.value,
  images: cloneImages(),
  marks: marks.map((m) => ({ ...m, fmt: { ...m.fmt } })),
  fontScale,
  bgColor,
  inset,
  hlColor,
  hlPad,
  fgColor,
  margin: els.margin.value,
  typeface: els.typeface.value,
  columns: els.columns.value,
  fill: els.fill.checked,
  footer: { ...footer },
});

function commit(kind) {
  const now = Date.now();
  const coalesce = COALESCING.has(kind) && kind === lastCommitKind && now - lastCommitAt < COALESCE_MS;

  if (committed && !coalesce) {
    undoStack.push(committed);
    if (undoStack.length > HISTORY_LIMIT) undoStack.shift();
  }
  committed = snapshot();
  lastCommitAt = now;
  lastCommitKind = kind;
  redoStack = [];
  syncHistoryButtons();
}

function applySnapshot(state) {
  els.source.value = state.text;
  images = state.images;
  marks = state.marks;
  fontScale = state.fontScale;
  bgColor = state.bgColor;
  inset = state.inset;
  hlColor = state.hlColor;
  hlPad = state.hlPad == null ? 1 : state.hlPad;
  fgColor = state.fgColor;
  els.margin.value = state.margin;
  els.typeface.value = state.typeface;
  els.columns.value = state.columns;
  els.fill.checked = state.fill;
  footer = { ...state.footer };
  selectedSlot = null;
  els.bgColor.value = bgColor;
  els.fmtBg.value = hlColor;
  els.fmtPad.value = String(hlPad);
  els.fmtFg.value = fgColor;
  els.insetToggle.checked = inset;

  pageDirty = false;
  selectImage(null);
  updateFontSizeUI();
  update();
  saveState();
  syncHistoryButtons();
}

function undo() {
  if (!undoStack.length) return;
  redoStack.push(committed);
  committed = undoStack.pop();
  lastCommitKind = '';
  applySnapshot(committed);
}

function redo() {
  if (!redoStack.length) return;
  undoStack.push(committed);
  committed = redoStack.pop();
  lastCommitKind = '';
  applySnapshot(committed);
}

function syncHistoryButtons() {
  els.undo.disabled = undoStack.length === 0;
  els.redo.disabled = redoStack.length === 0;
}

/* --- Formatting a run of words ------------------------------------------ */

/** Replace every stored mark that belongs to `unit` with `ranges`. */
function setUnitMarks(unit, ranges) {
  marks = marks.filter((m) => !unit.owners.includes(m));
  for (const r of ranges) marks.push(storedMark(unit, r));
}

/**
 * The selection, with the model guaranteed to describe the page it is on.
 *
 * If the page has been typed into, the model is a step behind — a line typed a
 * moment ago is not a unit yet — so it is brought up to date first, and the
 * selection is carried across the re-render by its character offsets.
 */
async function needSelection() {
  await settlePage();
  const selection = pendingSelection || readSelection();
  if (selection) return selection;
  flash('Select some words in the page on the right first');
  els.statNote.dataset.level = 'warn';
  return null;
}

/**
 * @param {object} patch — properties to set; a `null` value removes one
 * @param {{selection?: object[], refit?: boolean}} [opts]
 */
async function applyFormat(patch, opts = {}) {
  const selection = opts.selection || (await needSelection());
  if (!selection) return;

  for (const s of selection) {
    const unit = units[s.index];
    if (!unit) continue;
    setUnitMarks(unit, formatRange(unit.marks, s.start, s.end, patch, unit.plain.length));
  }

  keepOffsets = readSelectionOffsets();
  commit('format');
  update({ refit: opts.refit });
  saveState();
}

/**
 * A band is the background of a whole line, so a partial selection still means
 * the whole unit — the alternative is a header bar that stops mid-sentence.
 */
async function toggleBand() {
  const selection = await needSelection();
  if (!selection) return;

  const whole = selection
    .filter((s) => units[s.index])
    .map((s) => ({ index: s.index, start: 0, end: units[s.index].plain.length }));
  const on = whole.every((s) => allHave(units[s.index].marks, s.start, s.end, 'band', hlColor));

  for (const s of whole) {
    const unit = units[s.index];
    setUnitMarks(unit, formatRange(unit.marks, s.start, s.end, { band: on ? null : hlColor }, unit.plain.length));
  }

  keepOffsets = readSelectionOffsets();
  commit('format');
  update();
  saveState();
}

/**
 * Alignment belongs to the paragraph, so like a band it applies to every unit
 * the selection touches, whole. Pressing the alignment that is already on
 * takes it off, back to the element's own.
 */
async function toggleAlign(value) {
  const selection = await needSelection();
  if (!selection) return;
  const whole = selection.filter((s) => units[s.index]).map((s) => units[s.index]);
  const on = whole.every((u) => allHave(u.marks, 0, u.plain.length, 'align', value));
  for (const unit of whole) {
    setUnitMarks(unit, formatRange(unit.marks, 0, unit.plain.length, { align: on ? null : value }, unit.plain.length));
  }
  keepOffsets = readSelectionOffsets();
  commit('format');
  update();
  saveState();
}

/** Press once to apply, press again to take it off — as a toolbar should. */
async function toggleFormat(key, value, opts = {}) {
  const selection = await needSelection();
  if (!selection) return;
  const on = selection.every((s) => units[s.index] && allHave(units[s.index].marks, s.start, s.end, key, value));
  await applyFormat({ [key]: on ? null : value }, { ...opts, selection });
}

/**
 * A highlight is a colour and its padding, put on and taken off together, so
 * a run that loses its colour does not keep a padding nothing can see.
 */
async function toggleHighlight() {
  const selection = await needSelection();
  if (!selection) return;
  const on = selection.every((s) => units[s.index] && allHave(units[s.index].marks, s.start, s.end, 'bg', hlColor));
  await applyFormat(on ? { bg: null, pad: null } : { bg: hlColor, pad: padMark(hlPad) }, { selection, refit: false });
}

/**
 * The padding control: on a highlighted selection it changes only the
 * padding; on words with no highlight yet it highlights them at that padding.
 */
async function setHighlightPad(pad) {
  hlPad = pad;
  const selection = pendingSelection || readSelection();
  if (!selection) {
    saveState();
    return;
  }
  const lit = selection.every((s) => units[s.index] && allHave(units[s.index].marks, s.start, s.end, 'bg'));
  await applyFormat(lit ? { pad: padMark(pad) } : { bg: hlColor, pad: padMark(pad) }, { selection, refit: false });
}

async function clearFormat() {
  const selection = await needSelection();
  if (!selection) return;
  for (const s of selection) {
    const unit = units[s.index];
    if (!unit) continue;
    setUnitMarks(unit, clearRange(unit.marks, s.start, s.end, unit.plain.length));
  }
  keepOffsets = readSelectionOffsets();
  commit('format');
  update();
  saveState();
}

/**
 * Size steps apply to the selection. They are multipliers of whatever size the
 * fitter chose, so an emphasised line grows and shrinks with the rest of the
 * page and the one-page guarantee still holds.
 */
async function stepSelectionSize(direction) {
  const selection = await needSelection();
  if (!selection) return;
  const first = selection[0];
  const unit = units[first.index];
  if (!unit) return;
  const current = fmtAt(unit.marks, first.start).size || 1;
  await applyFormat({ size: stepSize(current, direction) }, { selection });
}

function syncTextTools() {
  const selection = !drag && !selectedImage ? pendingSelection : null;
  if (!selection || !sheetRefs) {
    els.fmtTools.hidden = true;
    return;
  }

  const sel = sheetRefs.win.getSelection();
  if (!sel || !sel.rangeCount || sel.isCollapsed) {
    els.fmtTools.hidden = true;
    return;
  }

  const rect = sel.getRangeAt(0).getBoundingClientRect();
  if (!rect || (!rect.width && !rect.height)) {
    els.fmtTools.hidden = true;
    return;
  }

  els.fmtTools.hidden = false;
  placeBar(els.fmtBar, paneBox(rect));

  const every = (key, value) =>
    selection.every((s) => units[s.index] && allHave(units[s.index].marks, s.start, s.end, key, value));
  for (const btn of els.fmtBar.querySelectorAll('.fmt-btn[data-fmt]')) {
    const act = btn.dataset.fmt;
    if (act === 'bold' || act === 'italic') btn.setAttribute('aria-pressed', String(every(act, true)));
    else if (act === 'band') {
      btn.setAttribute(
        'aria-pressed',
        String(selection.every((s) => units[s.index] && allHave(units[s.index].marks, 0, units[s.index].plain.length, 'band')))
      );
    } else if (act === 'align') {
      const value = btn.dataset.value;
      btn.setAttribute(
        'aria-pressed',
        String(selection.every((s) => units[s.index] && allHave(units[s.index].marks, 0, units[s.index].plain.length, 'align', value)))
      );
    }
  }

  const first = selection[0];
  const fmt = units[first.index] ? fmtAt(units[first.index].marks, first.start) : {};
  els.fmtSize.textContent = `${Math.round((fmt.size || 1) * 100)}%`;
  if (fmt.fg) els.fmtFg.value = fmt.fg;
  if (fmt.bg) els.fmtBg.value = fmt.bg;
  els.fmtPad.value = String(fmt.bg && fmt.pad != null ? fmt.pad : hlPad);
}

function onFormatAction(event) {
  const act = event.target && event.target.dataset ? event.target.dataset.fmt : null;
  if (!act) return;
  if (act === 'bold' || act === 'italic') toggleFormat(act, true);
  else if (act === 'band') toggleBand();
  else if (act === 'align') toggleAlign(event.target.dataset.value);
  else if (act === 'bigger') stepSelectionSize(1);
  else if (act === 'smaller') stepSelectionSize(-1);
  else if (act === 'clear') clearFormat();
}

/* --- Exports ------------------------------------------------------------ */

async function sheetHtml() {
  const { html } = currentDoc();
  return standaloneHtml(html, {
    fontPt: layout.fontPt,
    columns: layout.columns,
    lineHeight: layout.lineHeight,
    marginMm: layout.marginMm,
    imageScale: layout.imageScale,
    fontFamily: currentFace().css,
    title: docTitle,
    bgColor,
    inset,
    footerHtml: hasFooter(footer) ? footerHtml(footer) : '',
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
  commit('type');
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
  // Word first: its clipboard can also carry a picture of the selection,
  // which is not what was copied.
  const rich = richFromPaste(event, { newImageId });
  if (rich) {
    event.preventDefault();
    pasteRichIntoSource(rich);
    return;
  }
  if (handleImagePaste(event)) return;
  const text = textFromPaste(event);
  if (!text) return;
  event.preventDefault();
  insertAtCursor(els.source, text);
  commit('paste');
  update();
  saveState();
}

/* --- Pasting from Word ---------------------------------------------------- */

/**
 * A document pasted into an empty page brings its own typeface and its single
 * column with it: that is how it looked where it was copied from. Pasted into
 * an existing document, it takes that document's settings instead.
 */
function adoptPastedSettings(rich) {
  const face = rich.typeface && TYPEFACES[rich.typeface] ? rich.typeface : null;
  if (face) els.typeface.value = face;
  els.columns.value = '1';
  // A document that brings its own floating logos has its own letterhead:
  // the default footer steps aside rather than draw them twice. The empty
  // slots stay there to be filled again.
  if (Object.values(rich.images).some((img) => img.wrap)) footer = { left: null, right: null };
}

/** A Word paste's pictures and formatting, once its text is in the source. */
function applyRichPaste(rich, unitsBefore) {
  Object.assign(images, rich.images);
  const all = collectUnits(parse(els.source.value));
  marks.push(...marksFromRecords(all, rich.records, unitsBefore, rich.basePt));
}

/** Paste into the source pane: a block of its own at the cursor. */
function pasteRichIntoSource(rich) {
  const ta = els.source;
  const wasEmpty = !ta.value.trim();
  const at = ta.selectionStart == null ? ta.value.length : ta.selectionStart;
  const to = ta.selectionEnd == null ? at : ta.selectionEnd;
  const before = ta.value.slice(0, at).replace(/\s+$/, '');
  const after = ta.value.slice(to).replace(/^\s+/, '');
  ta.value = [before, rich.text, after].filter((p) => p !== '').join('\n\n');
  const caret = (before ? before.length + 2 : 0) + rich.text.length;
  ta.setSelectionRange(caret, caret);

  applyRichPaste(rich, before ? collectUnits(parse(before)).length : 0);
  if (wasEmpty) adoptPastedSettings(rich);
  commit('paste');
  update();
  saveState();
}

/**
 * Paste into the page: the pasted document is drawn after the block the caret
 * is in, and the page is then read back into the source like any other edit.
 * Left to itself the browser would paste Word's HTML as it is — every style
 * Word wrote, inline.
 */
async function onPagePaste(event) {
  const rich = richFromPaste(event, { newImageId });
  if (!rich) return;
  event.preventDefault();

  const wasEmpty = !els.source.value.trim();
  Object.assign(images, rich.images);
  const holder = sheetRefs.doc.createElement('div');
  holder.innerHTML = render(parse(rich.text).map((b) => (b.type === 'image' ? imageBlock(b.id) : b)));
  const nodes = [...holder.childNodes].filter((n) => n.nodeType === 1);
  if (!nodes.length) return;

  const anchor = blockAtCaret();
  const before = anchor ? anchor.nextSibling : null;
  for (const node of nodes) sheetRefs.body.insertBefore(node, before);

  // Units above the paste are the ones already numbered on the page.
  let unitsBefore = 0;
  for (const el of sheetRefs.body.querySelectorAll('[data-u]')) {
    if (nodes[0].compareDocumentPosition(el) & 2 /* preceding */) unitsBefore++;
  }

  // One history step for the whole paste: the text, its pictures and its marks.
  pageDirty = true;
  els.source.value = sourceFromPage();
  applyRichPaste(rich, unitsBefore);
  if (wasEmpty) adoptPastedSettings(rich);
  commit('paste');
  await settlePage();
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

/** Esc, wherever focus happens to be: stop resizing, then let go. */
function onEscape() {
  if (drag) {
    cancelResize();
    return true;
  }
  if (selectedImage) {
    selectImage(null);
    return true;
  }
  if (selectedSlot) {
    selectFooterSlot(null);
    return true;
  }
  if (pendingSelection) {
    const sel = sheetRefs.win.getSelection();
    if (sel) sel.removeAllRanges();
    pendingSelection = null;
    syncTextTools();
    return true;
  }
  return false;
}

function onKeyDown(event) {
  if (event.key === 'Escape') {
    if (onEscape()) event.preventDefault();
    return;
  }

  const accel = event.ctrlKey || event.metaKey;
  const key = event.key.toLowerCase();

  // The browser's own undo is taken over on purpose: inside the page it would
  // put DOM back without putting the model back, and it knows nothing about a
  // resized image or a coloured word.
  if (accel && (key === 'z' || key === 'y')) {
    event.preventDefault();
    if (key === 'y' || event.shiftKey) redo();
    else undo();
    return;
  }

  if (accel && event.shiftKey && key === 'h') {
    event.preventDefault();
    toggleHighlight();
    return;
  }
  // Only ever hijack B and I for the page, never for the text being typed.
  if (accel && !event.shiftKey && (key === 'b' || key === 'i') && event.target !== els.source) {
    if (!pendingSelection) return;
    event.preventDefault();
    toggleFormat(key === 'b' ? 'bold' : 'italic', true);
    return;
  }
  if (accel && !event.shiftKey && key === 'p') {
    event.preventDefault();
    els.print.click();
    return;
  }

  if ((event.key === 'Delete' || event.key === 'Backspace') && selectedImage && event.target !== els.source) {
    event.preventDefault();
    removeImage(selectedImage);
  } else if ((event.key === 'Delete' || event.key === 'Backspace') && selectedSlot && event.target !== els.source) {
    event.preventDefault();
    setFooterSlot(selectedSlot, null);
  }
}

function wireSheet() {
  const { doc } = sheetRefs;

  doc.addEventListener('mousedown', (event) => {
    const slot = event.target.closest ? event.target.closest('.footer-slot') : null;
    if (slot) {
      event.preventDefault();
      // An empty slot is a "+ Image" target; a full one is selected.
      if (slot.dataset.empty === '1') pickFooterImage(slot.dataset.slot);
      else selectFooterSlot(slot.dataset.slot);
      return;
    }
    if (selectedSlot) selectFooterSlot(null);
    const fig = event.target.closest ? event.target.closest('.doc-image') : null;
    selectImage(fig ? fig.dataset.img : null);
  });

  // The selection is cached as unit ranges the moment it is made: clicking a
  // toolbar button moves focus out of the sheet, and a cached range is also
  // what survives the re-render that applying a format causes.
  doc.addEventListener('selectionchange', () => {
    const found = readSelection();
    if (found && selectedImage) selectImage(null);
    pendingSelection = found;
    syncTextTools();
  });

  doc.addEventListener('keydown', onKeyDown);
}

function wire() {
  els.source.addEventListener('input', scheduleUpdate);
  els.source.addEventListener('paste', onPaste);

  for (const el of [els.margin, els.typeface, els.columns, els.fill]) {
    el.addEventListener('change', () => {
      commit('setting');
      update();
      saveState();
    });
  }

  els.clear.addEventListener('click', () => {
    els.source.value = '';
    images = {};
    marks = [];
    selectImage(null);
    els.source.focus();
    commit('clear');
    update();
    saveState();
  });

  els.fontUp.addEventListener('click', () => adjustFontScale(0.1));
  els.fontDown.addEventListener('click', () => adjustFontScale(-0.1));

  els.bgColor.addEventListener('input', () => {
    bgColor = els.bgColor.value;
    sheetRefs.page.style.setProperty('--doc-bg', bgColor);
    commit('colour');
    saveState();
  });

  els.insetToggle.addEventListener('change', () => {
    inset = els.insetToggle.checked;
    commit('setting');
    update();
    saveState();
  });

  els.highlight.addEventListener('click', () => toggleHighlight());

  // The colour pickers fire continuously while the wheel is dragged, and each
  // one is a re-render; a colour changes no geometry, so no re-fit either.
  const pickFg = debounce(() => applyFormat({ fg: fgColor }, { refit: false }), 120);
  els.fmtFg.addEventListener('input', () => {
    fgColor = els.fmtFg.value;
    pickFg();
  });

  const pickBg = debounce(() => applyFormat({ bg: hlColor, pad: padMark(hlPad) }, { refit: false }), 120);
  els.fmtPad.addEventListener('change', () => {
    const pad = Math.max(0, Math.min(20, Math.round(Number(els.fmtPad.value)) || 0));
    els.fmtPad.value = String(pad);
    setHighlightPad(pad);
  });

  els.fmtBg.addEventListener('input', () => {
    hlColor = els.fmtBg.value; // also becomes what the toolbar button applies
    pickBg();
  });

  els.fmtBar.addEventListener('click', onFormatAction);
  // Pressing a button must not take the selection away from the sheet.
  els.fmtBar.addEventListener('mousedown', (e) => {
    if (e.target.tagName !== 'INPUT') e.preventDefault();
  });

  els.insert.addEventListener('change', async () => {
    const kind = els.insert.value;
    els.insert.value = '';
    await insertStructure(kind);
  });

  els.imageFile.addEventListener('change', async () => {
    for (const file of els.imageFile.files) await addImageFromFile(file);
    els.imageFile.value = '';
  });

  els.source.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
  });
  els.source.addEventListener('drop', async (e) => {
    e.preventDefault();
    for (const file of e.dataTransfer.files) {
      if (file.type.startsWith('image/')) await addImageFromFile(file);
    }
  });

  // A picture dropped anywhere else in the window lands at the source caret,
  // and a stray file dropped on the window never navigates the app away.
  const dropImages = async (e) => {
    if (e.defaultPrevented) return;
    e.preventDefault();
    for (const file of e.dataTransfer ? e.dataTransfer.files : []) {
      if (file.type.startsWith('image/')) await addImageFromFile(file);
    }
  };
  const allowDrop = (e) => {
    e.preventDefault();
    if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
  };
  window.addEventListener('dragover', allowDrop);
  window.addEventListener('drop', dropImages);
  sheetRefs.doc.addEventListener('dragover', allowDrop);
  sheetRefs.doc.addEventListener('drop', dropImages);

  // The width slider previews as it moves and commits when it is let go.
  els.imgWidth.addEventListener('input', () => {
    const fig = figureFor(selectedImage);
    if (!fig) return;
    fig.style.setProperty('--img-w', els.imgWidth.value);
    els.imgSize.textContent = `${els.imgWidth.value}%`;
    setFrame(paneBox(fig.getBoundingClientRect()));
  });
  els.imgWidth.addEventListener('change', () => {
    if (!selectedImage) return;
    images[selectedImage].width = clamp(Number(els.imgWidth.value), 10, 100);
    commit('image');
    update();
    saveState();
  });

  els.undo.addEventListener('click', undo);
  els.redo.addEventListener('click', redo);

  els.imgFrame.addEventListener('pointerdown', startResize);
  window.addEventListener('pointermove', moveResize);
  window.addEventListener('pointerup', endResize);
  window.addEventListener('pointercancel', cancelResize);
  els.imgBar.addEventListener('click', onImageAction);
  els.footBar.addEventListener('click', onFooterAction);
  els.footerFile.addEventListener('change', async () => {
    await footerImageFromFile(els.footerFile.files[0]);
    els.footerFile.value = '';
  });

  els.preview.addEventListener('scroll', syncTools, { passive: true });
  els.preview.addEventListener('mousedown', (event) => {
    // A click on the canvas around the sheet lets the picture go.
    if (event.target === els.preview || event.target === els.scaler) {
      selectImage(null);
      selectFooterSlot(null);
    }
  });

  els.print.addEventListener('click', () =>
    withBusy(els.print, async () => {
      await settlePage();
      await printSheet({ html: await sheetHtml(), iframe: els.sheet });
    })
  );

  els.pdf.addEventListener('click', () =>
    withBusy(els.pdf, async () => {
      await settlePage();
      const r = await exportPdf({ html: await sheetHtml(), title: docTitle });
      if (r && r.saved) flash(`Saved ${r.filePath}`);
    })
  );

  els.docx.addEventListener('click', () =>
    withBusy(els.docx, async () => {
      await settlePage();
      const { blocks } = currentDoc();
      const r = await exportDocx({ blocks, layout, face: currentFace(), title: docTitle, footer });
      if (r && r.saved && r.filePath) flash(`Saved ${r.filePath}`);
    })
  );

  window.addEventListener('resize', debounce(rescalePreview, 60));
  window.addEventListener('keydown', onKeyDown);
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
    units: () => units.map((u) => ({ index: u.index, plain: u.plain, marks: u.marks })),
    select: (index, start, end) => {
      const el = sheetRefs.body.querySelector(`[data-u="${index}"]`);
      if (!el) return false;
      const a = locate(el, start);
      const b = locate(el, end == null ? el.textContent.length : end);
      const range = sheetRefs.doc.createRange();
      range.setStart(a.node, a.offset);
      range.setEnd(b.node, b.offset);
      const sel = sheetRefs.win.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      pendingSelection = readSelection();
      syncTextTools();
      return Boolean(pendingSelection);
    },
    type: async (text) => {
      // What typing into the page does, without a keyboard: change the DOM,
      // then let the app's own input path take it from there.
      const sel = sheetRefs.win.getSelection();
      if (!sel || !sel.rangeCount) return false;
      const range = sel.getRangeAt(0);
      range.deleteContents();
      const node = sheetRefs.doc.createTextNode(text);
      range.insertNode(node);
      range.setStart(node, text.length);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
      onPageInput();
      return true;
    },
    caretAt: (index, offset) => {
      const el = sheetRefs.body.querySelector(`[data-u="${index}"]`);
      if (!el) return false;
      const a = locate(el, offset);
      const range = sheetRefs.doc.createRange();
      range.setStart(a.node, a.offset);
      range.collapse(true);
      const sel = sheetRefs.win.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      return true;
    },
    source: () => els.source.value,
    settle: settlePage,
    undo,
    redo,
    history: () => ({ undo: undoStack.length, redo: redoStack.length }),
    format: (patch) => applyFormat(patch),
    toggle: (key, value) => toggleFormat(key, value),
    addImage: async (src, w, h, extra = {}) => {
      const id = newImageId();
      images[id] = { src, alt: '', w, h, width: 60, align: 'center', wrap: false, ...extra };
      insertImageRef(id);
      await update();
      return id;
    },
    setImage: async (id, patch) => {
      Object.assign(images[id], patch);
      await update();
    },
    imageOf: (id) => ({ ...images[id] }),
    footer: () => ({ left: footer.left ? { ...footer.left } : null, right: footer.right ? { ...footer.right } : null }),
    setFooter: async (slot, src, w, h) => {
      footer = { ...footer, [slot]: src ? { src, alt: '', w, h } : null };
      commit('footer');
      await update();
      saveState();
    },
    selectFooter: (slot) => selectFooterSlot(slot),
    selectImage,
    sheetHtml,
    docxBytes: async () => {
      const { blocks } = currentDoc();
      const { buildDocx } = await import('./to-docx.js');
      const blob = await buildDocx({ blocks, layout, face: currentFace(), title: docTitle, footer });
      return Array.from(new Uint8Array(await blob.arrayBuffer()));
    },
  };
}

async function boot() {
  restoreState();
  sheetRefs = await mountSheet(els.sheet);
  fitter = createFitter({ page: sheetRefs.page, body: sheetRefs.body });
  wire();
  wireSheet();
  enablePageEditing();
  rescalePreview();
  exposeForTests();
  updateFontSizeUI();
  commit('init');
  await update();
  els.source.focus();
  document.documentElement.dataset.ready = 'true';
}

boot();
