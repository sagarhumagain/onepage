/**
 * parse.js — plain text -> structured document model.
 *
 * The app deliberately discards source formatting on paste and rebuilds the
 * document from the text alone. That keeps output consistent, makes the
 * auto-fit engine the single owner of typography, and lets the DOCX/PDF
 * exporters emit exact values instead of guessing at Word's intent.
 */

/** @typedef {{type:'h1'|'h2'|'h3'|'p'|'quote'|'code', text:string}} TextBlock */
/** @typedef {{type:'ul'|'ol', items:string[], start?:number}} ListBlock */
/** @typedef {{type:'hr'}} RuleBlock */
/** @typedef {{type:'table', head:string[], rows:string[][]}} TableBlock */
/** @typedef {{type:'image', id:string}} ImageRef */
/** @typedef {TextBlock|ListBlock|RuleBlock|TableBlock|ImageRef} Block */

/** Build a string from code points. Keeps this source file pure ASCII. */
const ch = (...codes) => String.fromCharCode(...codes);

const C = {
  NBSP: 0x00a0,
  MIDDOT: 0x00b7,
  BULLET: 0x2022,
  TRIANGLE_BULLET: 0x2023,
  WHITE_BULLET: 0x25e6,
  BLACK_SQUARE: 0x25aa,
  EN_DASH: 0x2013,
  EM_DASH: 0x2014,
  ZWSP: 0x200b,
  ZWJ: 0x200d,
  BOM: 0xfeff,
  LINE_SEP: 0x2028,
  PARA_SEP: 0x2029,
  DEVANAGARI_LO: 0x0900,
  DEVANAGARI_HI: 0x097f,
};

const BULLET_CHARS = ch(C.BULLET, C.MIDDOT, C.TRIANGLE_BULLET, C.BLACK_SQUARE, C.WHITE_BULLET, C.EN_DASH, C.EM_DASH);
const ZERO_WIDTH = new RegExp('[' + ch(C.ZWSP) + '-' + ch(C.ZWJ) + ch(C.BOM) + ']', 'g');
const LINE_SEPS = new RegExp('[' + ch(C.LINE_SEP, C.PARA_SEP) + ']', 'g');
const NBSP_RE = new RegExp(ch(C.NBSP), 'g');
// Latin plus Devanagari, so Nepali/Hindi text is treated as real letters.
const LETTER = '[A-Za-z' + ch(C.DEVANAGARI_LO) + '-' + ch(C.DEVANAGARI_HI) + ']';
const HAS_LETTER = new RegExp(LETTER);
const STARTS_UPPER = new RegExp('^[A-Z' + ch(C.DEVANAGARI_LO) + '-' + ch(C.DEVANAGARI_HI) + ']');

const BULLET_RE = new RegExp('^[ \\t]*[-*+' + BULLET_CHARS + ']\\s+(.*)$');
const ORDERED = /^[ \t]*(\d{1,3})[.)]\s+(.*)$/;
const ORDERED_ALPHA = /^[ \t]*[a-zA-Z][.)]\s+(.*)$/;
const ATX = /^(#{1,6})\s+(.*?)\s*#*$/;
const RULE = /^[ \t]*([-*_=])\1{2,}[ \t]*$/;
const QUOTE = /^[ \t]*>\s?(.*)$/;
const FENCE = /^[ \t]*(```|~~~)/;
const NUMBERED_HEADING = /^\d+(\.\d+)*\.?\s+\S/;
const TABLE_DIVIDER = /^[ \t]*\|?[ \t]*:?-{2,}:?[ \t]*(\|[ \t]*:?-{2,}:?[ \t]*)+\|?[ \t]*$/;
const SETEXT = /^[ \t]*(=|-){3,}[ \t]*$/;
/**
 * An image placed in the document. The picture itself lives outside the text
 * — a data URL is not something anyone wants to edit — but its *position* is
 * a line like any other, so moving an image is moving a line, and text above
 * and below it flows around it exactly as two paragraphs would.
 */
export const IMAGE_REF = /^[ \t]*\[image:([A-Za-z0-9_-]{1,40})\][ \t]*$/;

/** Collapse the many ways real-world text encodes whitespace into something uniform. */
export function normalize(raw) {
  return String(raw == null ? '' : raw)
    .replace(/\r\n?/g, '\n')
    .replace(NBSP_RE, ' ') // Word and the web are full of non-breaking spaces
    .replace(ZERO_WIDTH, '')
    .replace(LINE_SEPS, '\n')
    .replace(/[ \t]+$/gm, '')
    .replace(/\n{3,}/g, '\n\n') // never more than one blank line
    .trim();
}

const isBlank = (l) => !l || !l.trim();

/** A line that reads like a heading: short, no terminal punctuation, has letters. */
function looksLikeHeading(line, next) {
  const t = line.trim();
  if (t.length === 0 || t.length > 90) return false;
  if (!HAS_LETTER.test(t)) return false;
  if (/[.,;:!?]$/.test(t)) return false;
  if (BULLET_RE.test(line) || ORDERED.test(line)) return false;

  const words = t.split(/\s+/);
  if (words.length > 14) return false;

  // ALL CAPS is an unambiguous heading signal.
  const letters = t.replace(/[^A-Za-z]/g, '');
  if (letters.length >= 2 && letters === letters.toUpperCase()) return true;

  // "1.2 Scope" / "3. Methodology"
  if (NUMBERED_HEADING.test(t) && t.length <= 70) return true;

  // "Executive Summary" followed by a blank line.
  if (isBlank(next) && t.length <= 70 && STARTS_UPPER.test(t)) {
    const capitalised = words.filter((w) => /^[A-Z]/.test(w)).length;
    if (capitalised / words.length >= 0.5) return true;
  }
  return false;
}

function headingLevel(line) {
  const t = line.trim();
  const m = t.match(/^(\d+)(\.\d+)*\.?\s/);
  if (m) return m[2] ? 3 : 2;
  const letters = t.replace(/[^A-Za-z]/g, '');
  if (letters.length >= 2 && letters === letters.toUpperCase()) return 2;
  return 3;
}

/**
 * "1. Methodology" is a numbered heading; "1. Buy milk" inside a run of
 * numbered lines is a list item. They are the same shape, so decide by
 * looking for a sibling marker nearby — a real list almost always has one.
 */
function isNumberedHeading(lines, i) {
  const m = lines[i].match(ORDERED);
  if (!m) return false;
  const t = lines[i].trim();
  if (t.length > 70) return false;
  if (/[.,;:!?]$/.test(t)) return false;
  if (t.split(/\s+/).length > 12) return false;
  if (!HAS_LETTER.test(m[2])) return false;

  const n = parseInt(m[1], 10);
  for (let j = i + 1; j < lines.length; j++) {
    if (isBlank(lines[j])) continue;
    if (/^[ \t]/.test(lines[j])) return false; // wrapped continuation -> list item
    const next = lines[j].match(ORDERED);
    if (next && parseInt(next[1], 10) === n + 1) return false; // a real sequence -> list
    return true; // prose follows -> section heading
  }
  return true;
}

const splitRow = (line) =>
  line
    .trim()
    .replace(/^\|/, '')
    .replace(/\|$/, '')
    .split('|')
    .map((c) => c.trim());

/**
 * @param {string} raw
 * @returns {Block[]}
 */
export function parse(raw) {
  const text = normalize(raw);
  if (!text) return [];
  const lines = text.split('\n');
  /** @type {Block[]} */
  const blocks = [];
  let i = 0;
  let sawTitle = false;

  const pushPara = (buf) => {
    const joined = buf.join(' ').replace(/\s+/g, ' ').trim();
    if (joined) blocks.push({ type: 'p', text: joined });
  };

  while (i < lines.length) {
    const line = lines[i];

    if (isBlank(line)) {
      i++;
      continue;
    }

    const imageRef = line.match(IMAGE_REF);
    if (imageRef) {
      blocks.push({ type: 'image', id: imageRef[1] });
      i++;
      continue;
    }

    // ``` fenced code
    if (FENCE.test(line)) {
      const fence = line.trim().slice(0, 3);
      const buf = [];
      i++;
      while (i < lines.length && !lines[i].trim().startsWith(fence)) buf.push(lines[i++]);
      i++; // closing fence
      if (buf.length) blocks.push({ type: 'code', text: buf.join('\n') });
      continue;
    }

    if (RULE.test(line)) {
      blocks.push({ type: 'hr' });
      i++;
      continue;
    }

    // Markdown-ish table: header row followed by a divider row.
    if (line.includes('|') && i + 1 < lines.length && TABLE_DIVIDER.test(lines[i + 1])) {
      const head = splitRow(line);
      i += 2;
      const rows = [];
      while (i < lines.length && lines[i].includes('|') && !isBlank(lines[i])) {
        rows.push(splitRow(lines[i++]));
      }
      blocks.push({ type: 'table', head, rows });
      continue;
    }

    // # ATX heading
    const atx = line.match(ATX);
    if (atx) {
      const level = Math.min(atx[1].length, 3);
      blocks.push({ type: `h${level}`, text: atx[2].trim() });
      sawTitle = sawTitle || level === 1;
      i++;
      continue;
    }

    // Setext heading: text underlined with === or ---
    if (i + 1 < lines.length && SETEXT.test(lines[i + 1])) {
      const level = lines[i + 1].trim()[0] === '=' ? 1 : 2;
      blocks.push({ type: `h${level}`, text: line.trim() });
      sawTitle = sawTitle || level === 1;
      i += 2;
      continue;
    }

    // > blockquote
    if (QUOTE.test(line)) {
      const buf = [];
      while (i < lines.length && QUOTE.test(lines[i])) buf.push(lines[i++].match(QUOTE)[1]);
      blocks.push({ type: 'quote', text: buf.join(' ').replace(/\s+/g, ' ').trim() });
      continue;
    }

    if (isNumberedHeading(lines, i)) {
      blocks.push({ type: `h${headingLevel(line)}`, text: line.trim() });
      i++;
      continue;
    }

    // Lists. Continuation lines without a marker belong to the previous item.
    if (BULLET_RE.test(line) || ORDERED.test(line) || ORDERED_ALPHA.test(line)) {
      const ordered = !BULLET_RE.test(line);
      const startMatch = line.match(ORDERED);
      const items = [];
      while (i < lines.length && !isBlank(lines[i])) {
        const l = lines[i];
        const b = l.match(BULLET_RE);
        const o = l.match(ORDERED);
        const oa = l.match(ORDERED_ALPHA);
        const isMarker = ordered ? !!(o || oa) : !!b;
        if (isMarker) {
          items.push((b ? b[1] : o ? o[2] : oa[1]).trim());
        } else if (
          items.length &&
          !looksLikeHeading(l, lines[i + 1]) &&
          !ATX.test(l) &&
          !QUOTE.test(l) &&
          !IMAGE_REF.test(l)
        ) {
          items[items.length - 1] += ' ' + l.trim();
        } else break;
        i++;
      }
      if (items.length) {
        const blk = { type: ordered ? 'ol' : 'ul', items };
        if (ordered && startMatch) blk.start = parseInt(startMatch[1], 10);
        blocks.push(blk);
      }
      continue;
    }

    // The first non-empty line becomes the document title if it reads like one.
    if (!sawTitle && blocks.length === 0) {
      const t = line.trim();
      if (t.length <= 90 && !/[.,;:]$/.test(t) && t.split(/\s+/).length <= 16) {
        blocks.push({ type: 'h1', text: t });
        sawTitle = true;
        i++;
        continue;
      }
    }

    if (looksLikeHeading(line, lines[i + 1])) {
      blocks.push({ type: `h${headingLevel(line)}`, text: line.trim() });
      i++;
      continue;
    }

    // Otherwise: a paragraph, running until a blank line or a structural line.
    const buf = [];
    while (i < lines.length && !isBlank(lines[i])) {
      const l = lines[i];
      if (
        buf.length &&
        (BULLET_RE.test(l) ||
          ORDERED.test(l) ||
          ATX.test(l) ||
          RULE.test(l) ||
          QUOTE.test(l) ||
          FENCE.test(l) ||
          IMAGE_REF.test(l))
      )
        break;
      if (buf.length && looksLikeHeading(l, lines[i + 1])) break;
      buf.push(l.trim());
      i++;
    }
    pushPara(buf);
  }

  return tidy(blocks);
}

const FULL_BOLD = /^\*\*([\s\S]+)\*\*$/;

/**
 * Remove emphasis that the surrounding element already provides.
 *
 * Word has no heading styles in its clipboard output — a title arrives as an
 * ordinary paragraph wrapped in <b>, which becomes "**Title**". Once that line
 * is recognised as a heading the markers are redundant, and the same is true of
 * bold text in a table's header row.
 */
function tidy(blocks) {
  const unwrap = (text) => {
    const m = String(text).match(FULL_BOLD);
    return m && !m[1].includes('**') ? m[1].trim() : text;
  };
  for (const b of blocks) {
    if (b.type === 'h1' || b.type === 'h2' || b.type === 'h3') b.text = unwrap(b.text);
    else if (b.type === 'table' && b.head) {
      b.head = b.head.map(unwrap);
      // Square the grid off here, once, so every consumer — the HTML
      // renderer, the Word exporter and the highlight index — addresses the
      // same cells.
      const cols = b.head.length;
      if (cols) b.rows = (b.rows || []).map((r) => Array.from({ length: cols }, (_, j) => r[j] || ''));
    }
  }
  return blocks;
}

/** Rough word count, used for the overflow ladder and the UI readout. */
export function wordCount(raw) {
  const t = normalize(raw);
  if (!t) return 0;
  return (t.match(/\S+/g) || []).length;
}
