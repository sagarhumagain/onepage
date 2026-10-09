/**
 * markup.js — the two conventions OnePage adds to the source text.
 *
 * Both exist so that a formatting choice lives in the document rather than in
 * a side table: the text you can see, edit, copy and reload is the whole
 * state. That is what makes an image shift when you type above it, and what
 * lets a highlight survive a reload without a second store to keep in sync.
 *
 *   image      ![alt](img:7f3a2b){62%|center}   on a line of its own
 *   highlight  {=#ffe066:6}some words{=}        inline, anywhere
 *
 * The image reference points into images.js; the percentage is of the column
 * width and the last field is the alignment. The highlight carries its colour
 * and its padding in tenths of an em.
 */

export const IMAGE_LINE =
  /^!\[([^\]]*)\]\(img:([A-Za-z0-9_-]+)\)(?:\{(\d{1,3})%\|(left|center|right)\})?$/;

/** Global form, used to find every reference in a document. */
const IMAGE_GLOBAL = new RegExp(IMAGE_LINE.source.slice(1, -1), 'gm');

export const HIGHLIGHT = /\{=(#[0-9a-fA-F]{3,8})(?::(\d{1,2}))?\}([\s\S]*?)\{=\}/;

export const DEFAULT_WIDTH_PCT = 62;
export const DEFAULT_ALIGN = 'center';

/** @param {{id:string, alt?:string, widthPct?:number, align?:string}} spec */
export function imageToken({ id, alt = '', widthPct = DEFAULT_WIDTH_PCT, align = DEFAULT_ALIGN }) {
  const pct = Math.max(5, Math.min(100, Math.round(widthPct)));
  return `![${alt}](img:${id}){${pct}%|${align}}`;
}

/** Every image id referenced by the text, in document order. */
export function referencedIds(text) {
  const ids = [];
  for (const m of String(text || '').matchAll(IMAGE_GLOBAL)) ids.push(m[2]);
  return ids;
}

/**
 * Rewrite one image's token in place.
 *
 * @param {string} text
 * @param {string} id
 * @param {{widthPct?:number, align?:string, alt?:string}} changes
 * @returns {string}
 */
export function updateImage(text, id, changes) {
  return String(text).replace(IMAGE_GLOBAL, (whole, alt, refId, pct, align) => {
    if (refId !== id) return whole;
    return imageToken({
      id,
      alt: changes.alt != null ? changes.alt : alt,
      widthPct: changes.widthPct != null ? changes.widthPct : Number(pct) || DEFAULT_WIDTH_PCT,
      align: changes.align != null ? changes.align : align || DEFAULT_ALIGN,
    });
  });
}

/** Remove an image's token, and the blank line it was sitting on. */
export function removeImage(text, id) {
  const lines = String(text).split('\n');
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].trim().match(IMAGE_LINE);
    if (m && m[2] === id) {
      // Drop one of the blank lines that surrounded it, so removing an image
      // does not leave a widening gap behind.
      if (out.length && !out[out.length - 1].trim() && !String(lines[i + 1] || '').trim()) i++;
      continue;
    }
    out.push(lines[i]);
  }
  return out.join('\n');
}

/**
 * Insert an image token at the caret, always on its own line.
 *
 * @returns {{text:string, caret:number}}
 */
export function insertImageAt(text, caretStart, caretEnd, token) {
  const before = String(text).slice(0, caretStart);
  const after = String(text).slice(caretEnd);
  const lead = before === '' || before.endsWith('\n\n') ? '' : before.endsWith('\n') ? '\n' : '\n\n';
  const trail = after === '' ? '\n' : after.startsWith('\n\n') ? '' : after.startsWith('\n') ? '\n' : '\n\n';
  const insert = `${lead}${token}${trail}`;
  return { text: before + insert + after, caret: before.length + insert.length };
}

const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Strip any highlight markers already inside a run of text. */
export const stripHighlights = (s) =>
  String(s).replace(new RegExp(HIGHLIGHT.source, 'g'), (_m, _c, _p, inner) => inner);

/**
 * Anything at the head of a line that tells the parser what kind of block the
 * line is: indentation, a bullet or number, a quote arrow, a heading's hashes.
 * A highlight must open *after* it, or the marker text ends up inside the
 * highlight and the line stops being a list item at all.
 */
const LINE_PREFIX = /^(\s*(?:[-*+\u2022\u00b7]\s+|\d{1,3}[.)]\s+|[a-zA-Z][.)]\s+|>\s?|#{1,6}\s+)?)/;

/**
 * Wrap a selection in highlight markers.
 *
 * A selection that spans several lines is wrapped line by line rather than as
 * one run: a marker pair straddling a paragraph break would put the opener in
 * one block and the closer in another, and neither would render.
 *
 * @param {string} selection
 * @param {string} color
 * @param {number} padTenths padding in tenths of an em, 0-99
 * @returns {string}
 */
export function highlight(selection, color, padTenths) {
  const pad = Math.max(0, Math.min(99, Math.round(padTenths || 0)));
  const open = `{=${color}${pad ? `:${pad}` : ''}}`;
  return stripHighlights(selection)
    .split('\n')
    .map((line) => {
      if (!line.trim()) return line;
      const prefix = line.match(LINE_PREFIX)[1];
      const rest = line.slice(prefix.length);
      // Trailing whitespace stays outside too, so the coloured box ends at the
      // last word rather than trailing off past it.
      const [, body, trail] = rest.match(/^([\s\S]*?)(\s*)$/);
      return body ? `${prefix}${open}${body}{=}${trail}` : line;
    })
    .join('\n');
}

/** True when the text is entirely wrapped in one highlight already. */
export const isHighlighted = (s) => new RegExp(`^${escapeRe('{=')}`).test(String(s).trim());
