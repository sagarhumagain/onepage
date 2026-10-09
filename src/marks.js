/**
 * marks.js — formatting applied to a run of words, anchored to content.
 *
 * A mark is "these characters look like this": a background, an ink colour, a
 * size relative to whatever the fitter chose, bold, italic. The obvious way to
 * store one is "block 4, characters 10-25", and it is wrong — the source text
 * is edited constantly, and every insertion above would silently move every
 * mark below it onto the wrong words.
 *
 * So a mark remembers *what it was on*: the plain text of its unit, which
 * occurrence of that text it was, and the marked substring itself.
 * Re-anchoring then degrades gracefully — the exact unit if it still exists,
 * otherwise any unit still containing the marked words, otherwise the mark is
 * dropped rather than applied to text the user never chose.
 *
 * A *unit* is one independently markable run of text: a paragraph, a heading, a
 * single list item, a single table cell. Code blocks are deliberately not units:
 * their text is rendered verbatim rather than tokenised, and a colour behind
 * part of a listing reads as a rendering fault.
 *
 * Sizes are multipliers, never point sizes. The fitter owns absolute type size
 * — it is the whole product — so "make this bigger" has to mean "1.3 times
 * whatever this element ends up being", or one emphasised line would break the
 * one-page guarantee.
 */

import { plainText } from './inline.js';

/**
 * `bold: false` is meaningful: it is a heading, or a header cell, set in the
 * regular weight its source used. `pad` is a highlight's padding in tenths of
 * an em, and only means anything alongside `bg`.
 * @typedef {{bg?:string, pad?:number, band?:string, fg?:string, size?:number, bold?:boolean,
 *   italic?:boolean, underline?:boolean, align?:string, valign?:string}} Fmt
 */
/** @typedef {{start:number, end:number, fmt:Fmt}} Mark */
/** @typedef {{sig:string, nth:number, text:string, start:number, end:number, fmt:Fmt}} StoredMark */
/**
 * `block` is the top-level block the unit sits in; `node` is the block object
 * that owns it, which is a block inside a table cell when the cell holds more
 * than a line. `kind` is the element it renders as, and `inTable` says it is
 * set at the table's size.
 * @typedef {{index:number, block:number, node:object, sub:string, plain:string, nth:number,
 *   kind:string, inTable:boolean}} Unit
 */

const MARKABLE_TEXT = new Set(['h1', 'h2', 'h3', 'p', 'quote']);

/** The sizes the size buttons step through. 1 is "whatever the fitter chose". */
const SIZE_STEPS = [0.6, 0.7, 0.8, 0.9, 1, 1.15, 1.35, 1.6, 1.9, 2.25];

export function stepSize(current, direction) {
  const at = SIZE_STEPS.reduce(
    (best, v, i) => (Math.abs(v - (current || 1)) < Math.abs(SIZE_STEPS[best] - (current || 1)) ? i : best),
    0
  );
  const next = Math.max(0, Math.min(SIZE_STEPS.length - 1, at + (direction > 0 ? 1 : -1)));
  return SIZE_STEPS[next];
}

/**
 * Every markable run of text in the document, in reading order.
 *
 * `sub` is the unit's address inside its block — 'text', 'items.2',
 * 'head.1', 'rows.3.0' — and is the key both the HTML and the Word renderer
 * use to look the unit's marks back up. A cell that holds paragraphs and
 * lists is not a unit itself: each paragraph and item inside it is.
 *
 * @param {import('./parse.js').Block[]} blocks
 * @returns {Unit[]}
 */
export function collectUnits(blocks) {
  /** @type {Unit[]} */
  const units = [];
  const seen = new Map();

  const add = (block, node, sub, text, kind, inTable) => {
    const plain = plainText(text);
    const nth = seen.get(plain) || 0;
    seen.set(plain, nth + 1);
    units.push({ index: units.length, block, node, sub, plain, nth, kind, inTable });
  };

  const visit = (b, i, inTable) => {
    if (MARKABLE_TEXT.has(b.type)) add(i, b, 'text', b.text, b.type, inTable);
    else if (b.type === 'ul' || b.type === 'ol') (b.items || []).forEach((it, j) => add(i, b, `items.${j}`, it, 'li', inTable));
    else if (b.type === 'table') {
      const cells = b.cells || {};
      const cell = (sub, text) => {
        // A cell a merge runs over is not drawn, so it has nothing to mark.
        if (b.covered && b.covered[sub]) return;
        if (cells[sub]) cells[sub].forEach((inner) => visit(inner, i, true));
        else add(i, b, sub, text, 'cell', true);
      };
      if (!b.headless) (b.head || []).forEach((c, j) => cell(`head.${j}`, c));
      (b.rows || []).forEach((row, r) => row.forEach((c, j) => cell(`rows.${r}.${j}`, c)));
    }
  };
  blocks.forEach((b, i) => visit(b, i, false));

  return units;
}

/**
 * The unit a stored mark belongs to now, with its offsets rebased.
 *
 * An exact hit is never rationed: one unit can carry any number of marks — two
 * highlighted words in the same paragraph are two marks with the same anchor.
 * `taken` only rations the *guesses* below it, so that several marks whose
 * units are gone do not all pile onto whichever unit happens to match first.
 */
function anchor(units, taken, m) {
  const exact = units.find((u) => u.plain === m.sig && u.nth === m.nth);
  if (exact) return { unit: exact, start: m.start, end: m.end, exact: true };

  const sameText = units.find((u) => !taken.has(u.index) && u.plain === m.sig);
  if (sameText) return { unit: sameText, start: m.start, end: m.end };

  // The unit was edited. If the marked words are still somewhere in the
  // document, follow them; a mark that cannot find its words is dropped.
  if (m.text) {
    for (const u of units) {
      if (taken.has(u.index)) continue;
      const at = u.plain.indexOf(m.text);
      if (at >= 0) return { unit: u, start: at, end: at + m.text.length };
    }
  }
  return null;
}

/* --- The format itself --------------------------------------------------- */

/**
 * `bg` is a highlighter: it hugs the words it is on. `band` is the whole
 * line's background — a section header bar — and is therefore only ever
 * applied to a complete unit, never to part of one. The same is true of
 * `align` and `valign`, which belong to the paragraph or the cell rather than
 * to any of its words.
 */
const KEYS = ['bg', 'pad', 'band', 'fg', 'size', 'bold', 'italic', 'underline', 'align', 'valign'];

/** Properties of the whole unit; see above. */
export const UNIT_KEYS = ['band', 'align', 'valign'];

const isEmptyFmt = (fmt) => !fmt || !KEYS.some((k) => fmt[k] != null);

const fmtKey = (fmt) => KEYS.map((k) => (fmt && fmt[k] != null ? `${k}:${fmt[k]}` : '')).join('|');

/**
 * Apply a patch to a format. A `null` value *removes* that property, which is
 * what makes every button a toggle: bold off is `{bold: null}`, not `false`.
 */
function patched(fmt, patch) {
  const out = {};
  for (const k of KEYS) if (fmt && fmt[k] != null) out[k] = fmt[k];
  for (const k of KEYS) {
    if (!(k in patch)) continue;
    if (patch[k] == null) delete out[k];
    else out[k] = patch[k];
  }
  return isEmptyFmt(out) ? null : out;
}

/** Marks -> one format per character. */
function spread(marks, length) {
  const chars = new Array(length).fill(null);
  for (const m of marks || []) {
    for (let i = Math.max(0, m.start); i < Math.min(length, m.end); i++) {
      chars[i] = patched(chars[i], m.fmt || {});
    }
  }
  return chars;
}

/** One format per character -> the shortest disjoint, ordered set of marks. */
function gather(chars) {
  const out = [];
  let i = 0;
  while (i < chars.length) {
    if (!chars[i]) {
      i++;
      continue;
    }
    const key = fmtKey(chars[i]);
    let j = i;
    while (j < chars.length && chars[j] && fmtKey(chars[j]) === key) j++;
    out.push({ start: i, end: j, fmt: chars[i] });
    i = j;
  }
  return out;
}

/** Flatten overlapping marks; later ones win property by property. */
function normaliseMarks(list, length) {
  if (!list || !list.length || length <= 0) return [];
  return gather(spread(list, length));
}

/** Set, change or remove properties across `[start, end)`. */
export function formatRange(marks, start, end, patch, length) {
  const chars = spread(marks, length);
  for (let i = Math.max(0, start); i < Math.min(length, end); i++) chars[i] = patched(chars[i], patch);
  return gather(chars);
}

/** Take every mark off `[start, end)`. */
export function clearRange(marks, start, end, length) {
  const chars = spread(marks, length);
  for (let i = Math.max(0, start); i < Math.min(length, end); i++) chars[i] = null;
  return gather(chars);
}

/**
 * Does every character in the range carry this property (with this value)?
 *
 * This is what decides whether a button is showing "on" and whether pressing
 * it will add the formatting or take it away.
 */
export function allHave(marks, start, end, key, value) {
  if (end <= start) return false;
  const chars = spread(marks, end);
  for (let i = start; i < end; i++) {
    const fmt = chars[i];
    if (!fmt || fmt[key] == null) return false;
    if (value !== undefined && fmt[key] !== value) return false;
  }
  return true;
}

/** The format at the first character of the range — what the bar reports. */
export function fmtAt(marks, at) {
  const chars = spread(marks, at + 1);
  return chars[at] || {};
}

/* --- Resolving stored marks onto a freshly parsed document --------------- */

/**
 * Blocks gain two private fields, both keyed by the unit's `sub` address:
 * `_marks` (the runs to paint) and `_u` (the unit index, which the preview
 * uses to turn a DOM selection back into a mark). They are written to the
 * block that owns the unit — a block inside a cell, for a cell's paragraphs.
 * Blocks are rebuilt on every keystroke, so writing to them is safe.
 *
 * @param {import('./parse.js').Block[]} blocks
 * @param {StoredMark[]} stored
 * @returns {Unit[]} the units, each carrying its resolved `marks` and the
 *   stored objects (`owners`) that produced them
 */
export function applyMarks(blocks, stored) {
  const units = collectUnits(blocks);
  const byUnit = new Map();

  for (const u of units) {
    const b = u.node;
    if (!b._u) b._u = {};
    b._u[u.sub] = u.index;
    u.marks = [];
    u.owners = [];
  }

  const taken = new Set();
  for (const m of stored || []) {
    if (isEmptyFmt(m.fmt)) continue;
    const hit = anchor(units, taken, m);
    if (!hit) continue;
    if (!hit.exact) taken.add(hit.unit.index);
    const list = byUnit.get(hit.unit.index) || [];
    list.push({ start: hit.start, end: hit.end, fmt: m.fmt });
    byUnit.set(hit.unit.index, list);
    hit.unit.owners.push(m);
  }

  for (const [index, list] of byUnit) {
    const u = units[index];
    const marks = normaliseMarks(list, u.plain.length);
    if (!marks.length) continue;
    u.marks = marks;
    const b = u.node;
    if (!b._marks) b._marks = {};
    b._marks[u.sub] = marks;
  }

  return units;
}

/** A storable mark for a range of one unit. */
export function storedMark(unit, mark) {
  return {
    sig: unit.plain,
    nth: unit.nth,
    text: unit.plain.slice(mark.start, mark.end),
    start: mark.start,
    end: mark.end,
    fmt: mark.fmt,
  };
}
