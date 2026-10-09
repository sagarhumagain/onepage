/**
 * parse.js — plain text -> structured document model.
 *
 * The text is the document of record: structure is written in it, and
 * everything a reader would call *styling* — a colour, a size, a band — lives
 * beside it as marks (see marks.js). That keeps output consistent, makes the
 * auto-fit engine the single owner of typography, and lets the DOCX/PDF
 * exporters emit exact values instead of guessing at Word's intent.
 */

/** @typedef {{type:'h1'|'h2'|'h3'|'p'|'quote'|'code', text:string}} TextBlock */
/**
 * A list. `levels`, `kinds`, `glyphs`, `nums` and `styles` are parallel to
 * `items` and present only when the list is more than a flat run of one kind
 * of marker — so a plain list is exactly the shape it always was, and every
 * item is still addressed as `items.N` however deep it sits.
 * @typedef {{type:'ul'|'ol', items:string[], start?:number, levels?:number[],
 *   kinds?:('ul'|'ol')[], glyphs?:(string|null)[], nums?:(number|null)[],
 *   styles?:(string|null)[]}} ListBlock
 */
/** @typedef {{type:'hr'}} RuleBlock */
/**
 * `cells` holds the parsed content of any cell that is more than one line of
 * text, keyed by its address ('head.1', 'rows.2.0'). `headless` is a table
 * whose header row is empty: the convention for a table that has none.
 *
 * `spans` and `covered` describe merged cells. A merged cell is written as the
 * cell where it starts — its *anchor* — and `<<` or `^^` in each cell it runs
 * over: `<<` joins a cell to the one on its left, `^^` to the one above. The
 * anchor's address maps to how many rows and columns it covers in `spans`,
 * and each cell it runs over maps to the anchor in `covered`; those cells are
 * not drawn.
 * @typedef {{type:'table', head:string[], rows:string[][], headless?:boolean,
 *   widths?:number[], border?:string, cells?:Object<string, Block[]>,
 *   spans?:Object<string, {rows:number, cols:number}>, covered?:Object<string, string>}} TableBlock
 */
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
  CHECK: 0x2713,
  HEAVY_CHECK: 0x2714,
  BALLOT_X: 0x2717,
  HEAVY_BALLOT_X: 0x2718,
  ARROWHEAD: 0x27a2,
  HEAVY_ARROW: 0x27a4,
  POINTER: 0x25ba,
  SMALL_TRIANGLE: 0x25b8,
  DIAMOND_CROSS: 0x2756,
  SHADOWED_SQUARE: 0x2751,
  LARGE_SQUARE: 0x25a0,
  WHITE_SQUARE: 0x25a1,
  BLACK_CIRCLE: 0x25cf,
  WHITE_CIRCLE: 0x25cb,
  BLACK_DIAMOND: 0x25c6,
  WHITE_DIAMOND: 0x25c7,
  ZWSP: 0x200b,
  ZWJ: 0x200d,
  BOM: 0xfeff,
  LINE_SEP: 0x2028,
  PARA_SEP: 0x2029,
  DEVANAGARI_LO: 0x0900,
  DEVANAGARI_HI: 0x097f,
};

/** Markers that mean "a bullet", whichever one: they all render as the default disc. */
const PLAIN_BULLETS = ch(C.BULLET, C.MIDDOT, C.EN_DASH, C.EM_DASH);
/**
 * Markers that are themselves the formatting — a checklist of ticks, Word's
 * arrowheads and squares. These are kept and drawn as written.
 */
export const LIST_GLYPHS = ch(
  C.TRIANGLE_BULLET, C.BLACK_SQUARE, C.WHITE_BULLET, C.CHECK, C.HEAVY_CHECK, C.BALLOT_X, C.HEAVY_BALLOT_X,
  C.ARROWHEAD, C.HEAVY_ARROW, C.POINTER, C.SMALL_TRIANGLE, C.DIAMOND_CROSS, C.SHADOWED_SQUARE,
  C.LARGE_SQUARE, C.WHITE_SQUARE, C.BLACK_CIRCLE, C.WHITE_CIRCLE, C.BLACK_DIAMOND, C.WHITE_DIAMOND
);
const BULLET_CHARS = PLAIN_BULLETS + LIST_GLYPHS;
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
/** Any list item: its indent, then a bullet, a number or a letter. */
const LIST_LINE = new RegExp('^([ \\t]*)(?:([-*+' + BULLET_CHARS + '])|(\\d{1,3})[.)]|([a-zA-Z])[.)])\\s+(.*)$');
/** A line ending in a backslash breaks the line without ending the paragraph. */
const HARD_BREAK = /\\$/;
/** Attributes for the block above, kramdown style: `{: widths="20 80"}`. */
const ATTRS = /^[ \t]*\{:\s*([^}]*)\}[ \t]*$/;
/**
 * Inside a table cell, `<br>` is the line separator — as it is in GitHub's
 * tables. A table inside a cell has line breaks of its own; those are written
 * one backslash deeper, `<br\>`, so they stay inside the inner table's cell.
 */
const CELL_BREAK = /<br\s*\/?>/gi;
const NESTED_BREAK = /<br\\(\\*)>/g;

/** A cell's text as the lines of the small document it holds. */
const cellLines = (text) => text.replace(CELL_BREAK, '\n').replace(NESTED_BREAK, '<br$1>');

/** A cell that is part of the merged cell to its left, or the one above. */
export const MERGE_LEFT = '<<';
export const MERGE_UP = '^^';
const HEX_COLOR = /^#[0-9a-f]{6}$/i;
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
  // A line that ends in a hard break is the first line of a paragraph.
  if (HARD_BREAK.test(t)) return false;
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

/**
 * Cells split on pipes that are not escaped; `\|` is a pipe inside a cell,
 * which is how the clipboard walker writes one.
 */
const splitRow = (line) =>
  line
    .trim()
    .replace(/^\|/, '')
    .replace(/(?<!\\)\|$/, '')
    .split(/(?<!\\)\|/)
    .map((c) => c.trim().replace(/\\\|/g, '|'));

/** Columns of indentation, a tab counting as four. */
const indentOf = (ws) => ws.replace(/\t/g, '    ').length;

/** Join a paragraph's lines: a soft wrap is a space, a hard break stays a line break. */
function joinLines(buf, hard) {
  let out = '';
  for (const line of buf) {
    const l = line.replace(/[ \t]+/g, ' ').trim();
    if (!out) out = l;
    else if (hard || HARD_BREAK.test(out)) out = `${out.replace(HARD_BREAK, '')}\n${l}`;
    else out = `${out} ${l}`;
  }
  // A break with nothing after it breaks nothing.
  return out.replace(HARD_BREAK, '');
}

/** `key="value"` pairs from an attribute line. Only the keys a table understands survive. */
function tableAttrs(line) {
  const out = {};
  const body = line.match(ATTRS)[1];
  for (const m of body.matchAll(/([a-z]+)\s*=\s*"([^"]*)"/gi)) {
    const key = m[1].toLowerCase();
    if (key === 'widths') {
      const widths = m[2].trim().split(/\s+/).map(Number);
      if (widths.length && widths.every((w) => w > 0 && w <= 100)) out.widths = widths;
    } else if (key === 'border' && HEX_COLOR.test(m[2].trim())) {
      out.border = m[2].trim().toLowerCase();
    }
  }
  return out;
}

/**
 * A list item's marker, as the parts the model keeps: which kind of list, the
 * glyph if it is one that should be drawn as written, and the number.
 */
function listMarker(m) {
  if (m[2]) return { kind: 'ul', glyph: LIST_GLYPHS.includes(m[2]) ? m[2] : null, num: null, style: null };
  if (m[3]) return { kind: 'ol', glyph: null, num: parseInt(m[3], 10), style: null };
  const letter = m[4];
  const upper = letter === letter.toUpperCase();
  return {
    kind: 'ol',
    glyph: null,
    num: letter.toLowerCase().charCodeAt(0) - 96,
    style: upper ? 'upper-alpha' : 'lower-alpha',
  };
}

/**
 * @param {string} raw
 * @param {{cell?: boolean}} [opts] — `cell`: the content of one table cell.
 *   A cell has no title and no inferred headings, and each of its lines is a
 *   line: `<br>` in a table cell means a line break, as it does on GitHub.
 * @returns {Block[]}
 */
export function parse(raw, opts = {}) {
  const cell = Boolean(opts.cell);
  const text = normalize(raw);
  if (!text) return [];
  const lines = text.split('\n');
  /** @type {Block[]} */
  const blocks = [];
  let i = 0;
  let sawTitle = cell;
  const headingLike = (line, next) => !cell && looksLikeHeading(line, next);

  const pushPara = (buf) => {
    const joined = joinLines(buf, cell);
    if (joined) blocks.push({ type: 'p', text: joined });
  };

  while (i < lines.length) {
    const line = lines[i];

    if (isBlank(line)) {
      i++;
      continue;
    }

    const imageRef = !cell && line.match(IMAGE_REF);
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

    // Markdown-ish table: header row followed by a divider row. A cell can
    // hold one too: a table inside a table.
    if (line.includes('|') && i + 1 < lines.length && TABLE_DIVIDER.test(lines[i + 1])) {
      const head = splitRow(line);
      i += 2;
      const rows = [];
      while (i < lines.length && lines[i].includes('|') && !isBlank(lines[i]) && !ATTRS.test(lines[i])) {
        rows.push(splitRow(lines[i++]));
      }
      const table = { type: 'table', head, rows };
      if (i < lines.length && ATTRS.test(lines[i])) Object.assign(table, tableAttrs(lines[i++]));
      blocks.push(table);
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
    if (!cell && i + 1 < lines.length && SETEXT.test(lines[i + 1])) {
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
      blocks.push({ type: 'quote', text: joinLines(buf, false) });
      continue;
    }

    if (!cell && isNumberedHeading(lines, i)) {
      blocks.push({ type: `h${headingLevel(line)}`, text: line.trim() });
      i++;
      continue;
    }

    // Lists. Indentation nests an item under the one above it; a line with no
    // marker continues the previous item.
    if (LIST_LINE.test(line)) {
      const items = [];
      const levels = [];
      const markers = [];
      const stack = [];
      while (i < lines.length && !isBlank(lines[i])) {
        const l = lines[i];
        const m = l.match(LIST_LINE);
        if (m) {
          const indent = indentOf(m[1]);
          if (!stack.length || indent > stack[stack.length - 1]) stack.push(indent);
          else while (stack.length > 1 && indent < stack[stack.length - 1]) stack.pop();
          items.push(m[5].trim());
          levels.push(stack.length - 1);
          markers.push(listMarker(m));
        } else if (
          items.length &&
          !headingLike(l, lines[i + 1]) &&
          !ATX.test(l) &&
          !QUOTE.test(l) &&
          !IMAGE_REF.test(l)
        ) {
          // Join onto the item as it stands, so a break written two lines up
          // is still there; a trailing break is trimmed once the list ends.
          const prev = items[items.length - 1];
          items[items.length - 1] = HARD_BREAK.test(prev)
            ? `${prev.replace(HARD_BREAK, '')}\n${l.replace(/[ \t]+/g, ' ').trim()}`
            : `${prev} ${l.replace(/[ \t]+/g, ' ').trim()}`;
        } else break;
        i++;
      }
      if (items.length) blocks.push(listBlock(items.map((it) => it.replace(HARD_BREAK, '')), levels, markers));
      continue;
    }

    // The first non-empty line becomes the document title if it reads like one.
    if (!sawTitle && blocks.length === 0) {
      const t = line.trim();
      if (t.length <= 90 && !/[.,;:\\]$/.test(t) && t.split(/\s+/).length <= 16) {
        blocks.push({ type: 'h1', text: t });
        sawTitle = true;
        i++;
        continue;
      }
    }

    if (headingLike(line, lines[i + 1])) {
      blocks.push({ type: `h${headingLevel(line)}`, text: line.trim() });
      i++;
      continue;
    }

    // Otherwise: a paragraph, running until a blank line or a structural line.
    const buf = [];
    while (i < lines.length && !isBlank(lines[i])) {
      const l = lines[i];
      // A line after a hard break belongs to this paragraph whatever it looks
      // like. Only a bullet or a number starts a list mid-paragraph; a letter
      // would end "written by" at "J. K. Rowling".
      const afterBreak = buf.length && HARD_BREAK.test(buf[buf.length - 1]);
      if (
        buf.length &&
        !afterBreak &&
        (BULLET_RE.test(l) ||
          ORDERED.test(l) ||
          ATX.test(l) ||
          RULE.test(l) ||
          QUOTE.test(l) ||
          FENCE.test(l) ||
          (!cell && IMAGE_REF.test(l)))
      )
        break;
      // After a hard break the next line belongs to this paragraph, however
      // much it looks like a heading on its own.
      if (buf.length && !afterBreak && headingLike(l, lines[i + 1])) break;
      buf.push(l.trim());
      i++;
    }
    pushPara(buf);
  }

  return tidy(blocks);
}

/**
 * A plain run of one kind of marker keeps the list's original, flat shape;
 * anything richer carries the parallel arrays that describe it.
 */
function listBlock(items, levels, markers) {
  const first = markers[0];
  const block = { type: first.kind, items };
  if (first.kind === 'ol' && first.num != null && !first.style) block.start = first.num;

  const flat =
    levels.every((l) => l === 0) &&
    markers.every((m) => m.kind === first.kind && !m.glyph && !m.style);
  if (flat) return block;

  block.levels = levels;
  block.kinds = markers.map((m) => m.kind);
  block.glyphs = markers.map((m) => m.glyph);
  block.nums = markers.map((m) => m.num);
  block.styles = markers.map((m) => m.style);
  return block;
}

const FULL_BOLD = /^\*\*([\s\S]+)\*\*$/;

/**
 * A cell holds more than a line of text when it has line breaks. A one-item
 * list in a cell is written with a trailing `<br>`, so a cell that merely
 * starts with a dash — "- 5", "– n/a" — stays the text it is.
 */
const isRichCell = (text) => /<br\s*\/?>/i.test(text);

/**
 * Merged cells, from the `<<` and `^^` written in the cells they cover.
 *
 * Each cell that is not itself a marker is an anchor. It runs right over the
 * `<<` cells after it, then down over each row below whose cells under it are
 * all `^^` (or `<<`, past the first column) — always a rectangle, the only
 * shape a table cell can have. A header cell does not run down into the body:
 * HTML and Word both keep a table's header rows apart. A marker that no
 * anchor reaches stays the text it is, so nothing typed is lost.
 */
function mergeCells(b) {
  const grid = [];
  if (!b.headless) grid.push(b.head.map((c, j) => ({ sub: `head.${j}`, text: c.trim() })));
  b.rows.forEach((r, ri) => grid.push(r.map((c, j) => ({ sub: `rows.${ri}.${j}`, text: c.trim() }))));

  const covered = {};
  const spans = {};
  const free = (r, c, ...markers) => grid[r] && grid[r][c] && !covered[grid[r][c].sub] && markers.includes(grid[r][c].text);
  grid.forEach((row, r) =>
    row.forEach((at, c) => {
      if (covered[at.sub] || at.text === MERGE_LEFT || at.text === MERGE_UP) return;
      let cols = 1;
      while (free(r, c + cols, MERGE_LEFT)) cols++;
      let rows = 1;
      const isHead = !b.headless && r === 0;
      const rowFits = (rr) =>
        Array.from({ length: cols }, (_, k) => (k ? free(rr, c + k, MERGE_UP, MERGE_LEFT) : free(rr, c, MERGE_UP))).every(Boolean);
      while (!isHead && r + rows < grid.length && rowFits(r + rows)) rows++;
      if (cols === 1 && rows === 1) return;
      spans[at.sub] = { rows, cols };
      for (let i = 0; i < rows; i++)
        for (let k = 0; k < cols; k++) if (i || k) covered[grid[r + i][c + k].sub] = at.sub;
    })
  );
  if (Object.keys(spans).length) {
    b.spans = spans;
    b.covered = covered;
  }
}

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
      if (b.widths && b.widths.length !== cols) delete b.widths;
      if (cols && b.head.every((c) => !c.trim())) b.headless = true;

      mergeCells(b);

      // A cell with more than a line in it is a small document of its own.
      const cells = {};
      const visit = (text, address) => {
        if (isRichCell(text) && !(b.covered && b.covered[address])) cells[address] = parse(cellLines(text), { cell: true });
      };
      if (!b.headless) b.head.forEach((c, j) => visit(c, `head.${j}`));
      b.rows.forEach((r, ri) => r.forEach((c, j) => visit(c, `rows.${ri}.${j}`)));
      if (Object.keys(cells).length) b.cells = cells;
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
