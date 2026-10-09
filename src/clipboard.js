/**
 * clipboard.js — turn whatever the clipboard offers into clean structured text.
 *
 * The app's canonical input is plain text, because that is what the user
 * edits and what every exporter reads. But Word's and a browser's plain-text
 * clipboard flavour throws away list nesting and headings, while their HTML
 * flavour keeps that structure buried in a mound of `mso-` noise.
 *
 * So when HTML is on the clipboard we walk it and emit the lightweight text
 * conventions that parse.js already understands. From a web page that is
 * structure only — a page's CSS is the site's look, not the author's.
 *
 * From Word it is more. A Word document's formatting *is* the author's: the
 * green title bar, the shaded label column, the size of the references. That
 * is read too (see `htmlToRich`), and comes back as marks — the same
 * content-anchored formatting the format bar writes — so a pasted document
 * looks like the one that was copied, while the text stays the text.
 *
 * The same walker runs in the other direction: when the preview is typed into,
 * the edited page is turned back into source text through `htmlToText`. That
 * is deliberate — a second, nearly identical serialiser would be a second
 * place for the two representations to disagree.
 */

import { tokenizeInline } from './inline.js';
import { LIST_GLYPHS, MERGE_LEFT, MERGE_UP } from './parse.js';
import { formatRange, storedMark } from './marks.js';
import { SCALE } from './scale.js';
import { extractRtfPictures } from './rtf.js';
import {
  declarations,
  expandVml,
  isWordHtml,
  luminance,
  mapSymbols,
  readWordStyles,
  stripFallbacks,
  symbolFont,
  toHex,
  toPt,
} from './word.js';

/** Elements whose text content is markup metadata, never document content. */
const SKIP = new Set(['STYLE', 'SCRIPT', 'HEAD', 'META', 'LINK', 'TITLE', 'NOSCRIPT', 'XML', 'O:P']);
const BLOCK = new Set([
  'P', 'DIV', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'MAIN', 'ASIDE',
  'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI', 'BLOCKQUOTE',
  'PRE', 'TABLE', 'TR', 'HR', 'FIGURE', 'FIGCAPTION', 'DL', 'DT', 'DD',
]);

const collapse = (s) => s.replace(/[\s ]+/g, ' ');

/**
 * Word does not emit <ul>/<li> for bulleted lists. It emits a flat run of
 * <p class=MsoListParagraph style='mso-list:l0 level1 lfo1'> elements, each
 * beginning with a literal bullet glyph in a Symbol-font span (and, in Word's
 * older output, the letter "o" or "v" for the second and third levels).
 *
 * Detecting that and re-emitting a real list marker is the difference between
 * a pasted Word document keeping its bullets and silently losing them.
 */
const BULLET_GLYPH = /^[\s ]*([·•●▪◦‣§-]|o|v)[\s ]+/;

const isWordListItem = (el) => {
  const cls = el.getAttribute('class') || '';
  const style = el.getAttribute('style') || '';
  return /MsoList/i.test(cls) || /mso-list\s*:/i.test(style);
};

const stripLeadingBullet = (s) => s.replace(BULLET_GLYPH, '').trim();

/** Emphasis markers around the words, with the spaces either side left outside. */
function wrap(inner, marker) {
  const m = inner.match(/^(\s*)([\s\S]*?)(\s*)$/);
  return m[2] ? `${m[1]}${marker}${m[2]}${marker}${m[3]}` : inner;
}

const LINKABLE = /^(https?:|mailto:)/i;

/** A list item's first line, as parse.js reads one. */
const LIST_START = new RegExp(`^\\s*(?:[-*+\u2022\u00b7\u2013\u2014${LIST_GLYPHS}]|\\d{1,3}[.)]|[a-zA-Z][.)])\\s`);

/**
 * A link written as `[words](href)`. A scheme is lower-cased so the tokeniser
 * reads it; emphasis inside the words goes outside the brackets, because a
 * link's words are not tokenised again. Words the tokeniser would link by
 * themselves — a bare address, an email — need no syntax at all.
 */
function linkMarkdown(words, href) {
  const url = href.replace(/^[a-z]+:/i, (scheme) => scheme.toLowerCase()).replace(/[()\s]/g, encodeURIComponent);
  const bare = words.replace(/\*+/g, '');
  const own = tokenizeInline(bare);
  if (own.length === 1 && own[0].href && own[0].text === bare) return bare;
  const strong = /^\*\*[^*]+\*\*$/.test(words);
  const link = `[${bare}](${url})`;
  return strong ? `**${link}**` : link;
}

function inlineText(node, out) {
  for (const child of node.childNodes) {
    if (child.nodeType === 3) {
      out.push(collapse(child.nodeValue));
      continue;
    }
    if (child.nodeType !== 1) continue;
    const tag = child.tagName.toUpperCase();
    if (SKIP.has(tag)) continue;
    if (tag === 'BR') {
      out.push('\n');
      continue;
    }
    // OnePage's own hard line break: a newline that is part of the paragraph.
    if (child.classList && child.classList.contains('doc-br')) {
      out.push('\\\n');
      continue;
    }
    if (tag === 'STRONG' || tag === 'B') {
      out.push(wrap(textOf(child), '**'));
      continue;
    }
    if (tag === 'EM' || tag === 'I') {
      out.push(wrap(textOf(child), '*'));
      continue;
    }
    if (tag === 'CODE' || tag === 'KBD' || tag === 'SAMP') {
      const inner = textOf(child);
      out.push(inner.trim() ? `\`${inner.trim()}\`` : '');
      continue;
    }
    if (tag === 'A') {
      // A link whose words are its address needs no syntax: the address is
      // recognised as a link wherever it appears.
      const href = child.getAttribute('href') || '';
      const inner = textOf(child);
      const words = inner.trim();
      if (LINKABLE.test(href) && words && words !== href && !/[[\]\n]/.test(words)) {
        const m = inner.match(/^(\s*)[\s\S]*?(\s*)$/);
        out.push(`${m[1]}${linkMarkdown(words, href)}${m[2]}`);
      } else out.push(inner);
      continue;
    }
    inlineText(child, out);
  }
}

function textOf(node) {
  const out = [];
  inlineText(node, out);
  return out.join('').replace(/[ \t]{2,}/g, ' ');
}

const cellText = (td) => textOf(td).replace(/\|/g, '\\|').replace(/\n/g, ' ').trim();

/* --- Lists ---------------------------------------------------------------- */

/** The marker a rendered or pasted list item is written back with. */
function listMarker(list, li, n) {
  if (list.tagName.toUpperCase() !== 'OL') {
    const glyph = li.getAttribute('data-glyph');
    return glyph && glyph.length === 1 && LIST_GLYPHS.includes(glyph) ? glyph : '-';
  }
  const type = list.getAttribute('type');
  if (type === 'a' || type === 'A') {
    const letter = String.fromCharCode(96 + Math.min(26, Math.max(1, n)));
    return `${type === 'A' ? letter.toUpperCase() : letter}.`;
  }
  return `${n}.`;
}

function walkList(list, lines, depth) {
  let n = Number(list.getAttribute('start')) || 1;
  const pad = '  '.repeat(depth);
  for (const li of list.children) {
    if (li.tagName.toUpperCase() !== 'LI') continue;
    // Nested lists are written by the recursive call below, so take only
    // this item's own text here.
    const own = li.ownerDocument.createElement('div');
    for (const kid of li.childNodes) {
      const t = kid.nodeType === 1 ? kid.tagName.toUpperCase() : '';
      if (t === 'UL' || t === 'OL') continue;
      own.appendChild(kid.cloneNode(true));
    }
    const t = stripLeadingBullet(textOf(own)).replace(/\\?\n/g, ' ');
    if (t) lines.push({ kind: 'block', text: `${pad}${listMarker(list, li, n)} ${t}` });
    n++;
    for (const kid of li.children) {
      const kt = kid.tagName.toUpperCase();
      if (kt === 'UL' || kt === 'OL') walkList(kid, lines, depth + 1);
    }
  }
}

/* --- Tables --------------------------------------------------------------- */

const cellsOf = (tr) => [...tr.children].filter((c) => /^T[HD]$/.test(c.tagName.toUpperCase()));

/** Own rows only: a table nested in a cell is that cell's content. */
const rowsOf = (table) => [...table.querySelectorAll('tr')].filter((tr) => tr.closest('table') === table);

/** A colspan or rowspan as a count; 0 is HTML's "to the end", which `limit` is. */
function spanOf(td, name, limit) {
  const raw = td.getAttribute(name);
  if (raw == null || raw.trim() === '') return 1;
  const n = Math.floor(Number(raw));
  if (n === 0 && name === 'rowspan') return limit;
  return Math.max(1, Math.min(limit, Number.isFinite(n) ? n : 1));
}

/**
 * A table's rows as a square grid: each cell where it starts, and `<<` or `^^`
 * in each grid position a merged cell runs over (see parse.js). A cell a
 * row span carries down is placed before the next cell of that row is, so
 * every cell lands in the column it is drawn in. A position nothing fills is
 * null.
 *
 * @returns {{grid: (Element|string|null)[][], width: number}}
 */
function tableGrid(rows) {
  const grid = rows.map(() => []);
  rows.forEach((tr, r) => {
    let c = 0;
    for (const td of cellsOf(tr)) {
      while (grid[r][c] !== undefined) c++;
      const cols = spanOf(td, 'colspan', 1000);
      const down = spanOf(td, 'rowspan', rows.length - r);
      for (let i = 0; i < down; i++)
        for (let k = 0; k < cols; k++) grid[r + i][c + k] = i ? MERGE_UP : k ? MERGE_LEFT : td;
      c += cols;
    }
  });
  const width = Math.max(0, ...grid.map((row) => row.length));
  return { grid: grid.map((row) => Array.from({ length: width }, (_, c) => row[c] ?? null)), width };
}

/** Does a row span run down out of this grid row into the next? */
const spansDown = (grid, r) => Boolean(grid[r + 1] && grid[r + 1].some((x) => x === MERGE_UP));

/**
 * A line that is already a row of a table inside this cell keeps its own line
 * breaks one level down, `<br\>`, so they stay inside the inner cell.
 */
const nestBreaks = (text) => text.replace(/<br(\\*)>/gi, '<br$1\\>');

/**
 * A cell's content as one line of a pipe table. A cell holding paragraphs or
 * a list keeps them, joined with `<br>` — a blank line between paragraphs
 * becomes two.
 */
function joinCell(lines) {
  let out = '';
  let gap = false;
  for (const item of lines) {
    if (item.kind === 'gap') {
      gap = out.length > 0;
      continue;
    }
    const text = nestBreaks(item.text.replace(/\s+$/, '')).replace(/\n/g, '<br>');
    if (!text.trim()) continue;
    if (out) out += gap ? '<br><br>' : '<br>';
    out += text;
    gap = false;
  }
  // A cell that is one list item still needs its `<br>`: without one,
  // "- Plan" in a cell is read as the words "- Plan".
  if (!out.includes('<br>') && LIST_START.test(out)) out += '<br>';
  return out.replace(/\|/g, '\\|');
}

const hasBlockContent = (td) =>
  [...td.querySelectorAll('p, ul, ol, div, h1, h2, h3, h4, h5, h6, blockquote, table')].some(
    (el) => (el.tagName.toUpperCase() === 'TABLE' ? el.parentElement.closest('td, th') : el.closest('td, th')) === td
  );

/** The attribute line a table's widths and rule colour are written back as. */
function tableAttrLine(widths, border) {
  const parts = [];
  if (widths && widths.length) parts.push(`widths="${widths.join(' ')}"`);
  if (border) parts.push(`border="${border}"`);
  return parts.length ? `{: ${parts.join(' ')}}` : null;
}

function walkTable(table, lines) {
  const rows = rowsOf(table);
  if (!rows.length) return false;
  const { grid, width } = tableGrid(rows);
  if (width < 2) return false;

  const cell = (td) => {
    if (td == null) return '';
    if (typeof td === 'string') return td;
    if (!hasBlockContent(td)) return cellText(td);
    const inner = [];
    walk(td, inner, 0, null);
    return joinCell(inner);
  };
  const row = (cells) => `| ${cells.map(cell).join(' | ')} |`;

  // A table OnePage drew without a header row is written back without one,
  // and so is one whose first row runs down into the next: a header row
  // stands apart from the body, so no cell of it can.
  const headless = table.getAttribute('data-head') === '0' || spansDown(grid, 0);
  lines.push({ kind: 'gap' });
  lines.push({ kind: 'block', text: headless ? `|${'  |'.repeat(width)}` : row(grid[0]) });
  lines.push({ kind: 'block', text: `|${' --- |'.repeat(width)}` });
  for (const cells of headless ? grid : grid.slice(1)) lines.push({ kind: 'block', text: row(cells) });
  const widths = (table.getAttribute('data-widths') || '').trim().split(/\s+/).filter(Boolean).map(Number);
  const attrs = tableAttrLine(widths.length === width ? widths : null, toHex(table.getAttribute('data-border')));
  if (attrs) lines.push({ kind: 'block', text: attrs });
  lines.push({ kind: 'gap' });
  return true;
}

/* --- The walk ------------------------------------------------------------- */

/**
 * @param {Node} node
 * @param {object[]} lines — {kind:'block'|'text'|'gap', text, record?, join?}
 * @param {number} depth
 * @param {object|null} ctx — a Word context (see `htmlToRich`), or null for
 *   structure only
 */
function walk(node, lines, depth, ctx) {
  // Word's lists are runs of sibling paragraphs; their nesting is worked out
  // across the run, so it is tracked per container.
  const run = { stack: [] };
  for (const child of node.childNodes) {
    if (child.nodeType === 3) {
      const t = collapse(child.nodeValue);
      if (t.trim()) lines.push({ kind: 'text', text: t });
      continue;
    }
    if (child.nodeType !== 1) continue;

    const tag = child.tagName.toUpperCase();
    if (SKIP.has(tag)) continue;
    if (ctx && !isWordList(child, ctx)) run.stack = [];

    if (/^H[1-6]$/.test(tag)) {
      if (ctx) {
        const p = ctx.paragraph(child);
        // A heading is one line: a line break inside it becomes a space.
        const md = p.md.replace(/\\\n/g, ' ');
        const record = { ...p.record, plain: p.record.plain.replace(/\n/g, ' ') };
        if (p.plain) lines.push({ kind: 'block', text: `${'#'.repeat(Math.min(Number(tag[1]), 3))} ${md}`, record });
        continue;
      }
      const level = Math.min(Number(tag[1]), 3);
      const t = textOf(child).trim();
      if (t) lines.push({ kind: 'block', text: `${'#'.repeat(level)} ${t}` });
      continue;
    }

    if (tag === 'HR') {
      lines.push({ kind: 'block', text: '---' });
      continue;
    }

    // OnePage's own pages come back through here when the preview is edited
    // in place. An image is a reference on a line of its own, wherever the
    // editing left the figure — inside a paragraph, it still becomes a block.
    const imageId = tag === 'FIGURE' ? child.getAttribute('data-img') : null;
    if (imageId && /^[A-Za-z0-9_-]{1,40}$/.test(imageId)) {
      lines.push({ kind: 'gap' });
      lines.push({ kind: 'block', text: `[image:${imageId}]` });
      lines.push({ kind: 'gap' });
      continue;
    }

    // A picture Word placed: its bytes came from the RTF flavour.
    if (ctx && tag === 'FIGURE' && child.hasAttribute('data-op-pic')) {
      const id = ctx.picture(child);
      if (id) {
        lines.push({ kind: 'gap' });
        lines.push({ kind: 'block', text: `[image:${id}]` });
        lines.push({ kind: 'gap' });
      }
      continue;
    }

    // A Word text box or shape: its content, carrying the box's fill.
    if (ctx && tag === 'DIV' && child.hasAttribute('data-op-box')) {
      const fill = toHex(child.getAttribute('data-fill'));
      lines.push({ kind: 'gap', hard: true });
      ctx.inBox(fill, () => walk(child, lines, depth, ctx));
      lines.push({ kind: 'gap', hard: true });
      continue;
    }

    if (tag === 'UL' || tag === 'OL') {
      walkList(child, lines, depth);
      lines.push({ kind: 'gap' });
      continue;
    }

    if (tag === 'TABLE') {
      if (ctx ? ctx.table(child, lines) : walkTable(child, lines)) continue;
    }

    if (tag === 'BLOCKQUOTE') {
      const t = textOf(child).trim();
      if (t) {
        lines.push({ kind: 'gap' });
        for (const l of t.split('\n')) if (l.trim()) lines.push({ kind: 'block', text: `> ${l.trim()}` });
        lines.push({ kind: 'gap' });
      }
      continue;
    }

    if (tag === 'PRE') {
      const t = (child.textContent || '').replace(/\s+$/, '');
      if (t.trim()) {
        lines.push({ kind: 'gap' });
        lines.push({ kind: 'block', text: '```' });
        for (const l of t.split('\n')) lines.push({ kind: 'block', text: l });
        lines.push({ kind: 'block', text: '```' });
        lines.push({ kind: 'gap' });
      }
      continue;
    }

    if (tag === 'BR') {
      lines.push({ kind: 'gap' });
      continue;
    }

    if (BLOCK.has(tag)) {
      // A block that holds only inline content is a paragraph; one that holds
      // other blocks is just a wrapper and should not add its own text.
      const hasBlockChild = [...child.children].some((c) => BLOCK.has(c.tagName.toUpperCase()));
      // A picture that sat in a line of Word text: the line, then the picture.
      const pictures = ctx ? [...child.children].filter((c) => c.hasAttribute('data-op-pic')) : [];
      const onlyPictures =
        pictures.length && [...child.children].every((c) => !BLOCK.has(c.tagName.toUpperCase()) || pictures.includes(c));
      if (ctx && (!hasBlockChild || onlyPictures)) {
        ctx.block(child, lines, run);
        for (const fig of pictures) walk({ childNodes: [fig] }, lines, depth, ctx);
      } else if (!hasBlockChild) {
        const raw = textOf(child);
        const wordList = isWordListItem(child) || BULLET_GLYPH.test(raw);
        const t = wordList ? stripLeadingBullet(raw) : raw.trim();

        if (t.trim()) {
          if (wordList) {
            // Word's own numbered lists already carry "1." in the text, so
            // only a bulleted item needs a marker putting back.
            const numbered = /^\d{1,3}[.)]\s/.test(t);
            lines.push({ kind: 'block', text: numbered ? t : `- ${t}` });
          } else {
            lines.push({ kind: 'gap' });
            for (const l of t.split('\n')) if (l.trim()) lines.push({ kind: 'block', text: l.trim() });
            lines.push({ kind: 'gap' });
          }
        }
      } else {
        lines.push({ kind: 'gap' });
        walk(child, lines, depth, ctx);
        lines.push({ kind: 'gap' });
      }
      continue;
    }

    walk(child, lines, depth, ctx);
  }
}

/** Lines -> text, never more than one blank line between blocks. */
function assemble(lines) {
  const out = [];
  const records = [];
  let pendingGap = false;
  for (const item of lines) {
    if (item.kind === 'gap') {
      pendingGap = out.length > 0;
      continue;
    }
    const text = item.text.replace(/\s+$/, '');
    if (!text.trim()) continue;
    if (pendingGap) out.push('');
    pendingGap = false;
    out.push(text);
    if (item.record) records.push(item.record);
    if (item.records) records.push(...item.records);
  }
  return { text: out.join('\n').replace(/\n{3,}/g, '\n\n').trim(), records };
}

/**
 * @param {string} html
 * @returns {string} structured plain text
 */
export function htmlToText(html) {
  if (!html || !html.trim()) return '';
  const doc = new DOMParser().parseFromString(html, 'text/html');
  doc.querySelectorAll('style, script, xml, head meta, link, title').forEach((n) => n.remove());
  // Word wraps the real payload in fragment comments; the body is enough.
  const lines = [];
  walk(doc.body, lines, 0, null);
  return assemble(lines).text;
}

/* --- Word, with its formatting ------------------------------------------- */

/** Is this element one of Word's list paragraphs? */
const isWordList = (el, ctx) => ctx && /mso-list\s*:\s*l\d+/i.test(el.getAttribute('style') || '');

/** The family a font-family value leads with, unquoted and lower-cased. */
const leadFamily = (f) => String(f || '').split(',')[0].replace(/["']/g, '').trim().toLowerCase();

/** Fonts the typeface menu offers, by the name Word writes. */
const TYPEFACE_OF = {
  'times new roman': 'times',
  calibri: 'sans',
  arial: 'arial',
  georgia: 'serif',
  'arial narrow': 'narrow',
  'courier new': 'mono',
};

/** Bullets Word types as letters in a symbol font, and what they draw. */
const WORD_BULLETS = { o: '◦', '§': '▪' };

/** Near enough to black to be the ordinary ink, which needs no mark. */
const isInk = (hex) => !hex || luminance(hex) < 0.02;

/**
 * Everything the Word walk needs to read formatting: the stylesheet, a cache
 * of parsed style attributes, the pictures, and the statistics that decide the
 * base size and the typeface.
 */
function wordContext(styles, pictures, newImageId) {
  const parsed = new WeakMap();
  const styleOf = (el) => {
    if (!parsed.has(el)) parsed.set(el, declarations(el.getAttribute('style')));
    return parsed.get(el);
  };
  const classRule = (el) => {
    const tag = el.tagName.toLowerCase();
    for (const cls of (el.getAttribute('class') || '').split(/\s+/).filter(Boolean)) {
      const rule = styles.classes.get(`${tag}.${cls.toLowerCase()}`) || styles.classes.get(`.${cls.toLowerCase()}`);
      if (rule) return rule;
    }
    return styles.classes.get(tag) || null;
  };
  const normal = styles.classes.get('p.msonormal') || new Map();

  /** A property from the element up to its paragraph, then the stylesheet. */
  const prop = (el, para, name) => {
    for (let n = el; n; n = n.parentElement) {
      const v = n.nodeType === 1 ? styleOf(n).get(name) : null;
      if (v) return v;
      if (n === para) break;
    }
    const rule = classRule(para);
    return (rule && rule.get(name)) || normal.get(name) || null;
  };

  /** Bold, italic or underline: the nearest element that decides wins. */
  const flag = (el, para, tags, name, on, off) => {
    for (let n = el; n; n = n.parentElement) {
      if (tags.includes(n.tagName)) return true;
      const v = (styleOf(n).get(name) || '').toLowerCase();
      if (on.test(v)) return true;
      if (off.test(v)) return false;
      if (n === para) break;
    }
    const rule = classRule(para);
    const v = ((rule && rule.get(name)) || '').toLowerCase();
    return on.test(v);
  };

  const sizes = new Map();
  const families = new Map();
  const count = (map, key, n) => key && map.set(key, (map.get(key) || 0) + n);

  /** The runs of one paragraph: text and how it looks, whitespace collapsed as a browser would. */
  function runsOf(para) {
    const raw = [];
    const visit = (node) => {
      for (const child of node.childNodes) {
        if (child.nodeType === 3) {
          const el = child.parentElement;
          const family = prop(el, para, 'font-family');
          let text = child.nodeValue;
          if (symbolFont(family) || /[-]/.test(text)) text = mapSymbols(text, family);
          const link = el.closest('a[href]');
          const href = link && para.contains(link) && LINKABLE.test(link.getAttribute('href')) ? link.getAttribute('href') : null;
          const color = toHex(prop(el, para, 'color'));
          let bg = null;
          for (let n = el; n && n !== para; n = n.parentElement) {
            const st = styleOf(n);
            bg = toHex(st.get('mso-highlight')) || toHex(st.get('background')) || toHex(st.get('background-color'));
            if (bg) break;
          }
          raw.push({
            text,
            bold: flag(el, para, ['B', 'STRONG'], 'font-weight', /^(bold|[6-9]00)$/, /^(normal|[1-5]00)$/),
            italic: flag(el, para, ['I', 'EM'], 'font-style', /^italic|oblique/, /^normal/),
            underline: !href && flag(el, para, ['U'], 'text-decoration', /underline/, /^none/),
            pt: toPt(prop(el, para, 'font-size')) || 11,
            fg: href && isInk(color) ? styles.linkColor : isInk(color) ? null : color,
            bg: bg && bg !== '#ffffff' ? bg : null,
            href,
            family: symbolFont(family) ? null : leadFamily(family),
          });
          continue;
        }
        if (child.nodeType !== 1) continue;
        const tag = child.tagName.toUpperCase();
        if (SKIP.has(tag) || tag === 'FIGURE') continue;
        if (tag === 'BR') {
          raw.push({ text: '\n', br: true });
          continue;
        }
        visit(child);
      }
    };
    visit(para);

    // Collapse whitespace across run boundaries; a line break eats the
    // spaces either side of it.
    const runs = [];
    let prevSpace = true;
    for (const r of raw) {
      if (r.br) {
        if (runs.length) runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/ +$/, '');
        runs.push({ ...r, text: '\n' });
        prevSpace = true;
        continue;
      }
      let t = r.text.replace(/[\s ]+/g, ' ');
      if (prevSpace) t = t.replace(/^ /, '');
      if (!t) continue;
      prevSpace = t.endsWith(' ');
      runs.push({ ...r, text: t });
    }
    while (runs.length && /^[ \n]*$/.test(runs[runs.length - 1].text)) runs.pop();
    if (runs.length) runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/ +$/, '');
    while (runs.length && runs[0].text === '\n') runs.shift();

    for (const r of runs) {
      const n = r.text.replace(/\s/g, '').length;
      count(sizes, Math.round(r.pt * 2) / 2, n);
      count(families, r.family, n);
    }
    return runs;
  }

  let box = null;
  const images = {};

  const ctx = {
    images,
    sizes,
    families,

    inBox(fill, fn) {
      const outer = box;
      box = { fill };
      try {
        fn();
      } finally {
        box = outer;
      }
    },

    /** A paragraph's runs as source text and as a record of how it looked. */
    paragraph(para, extra = {}) {
      const runs = runsOf(para);
      const built = toMarkdown(runs);
      const own = toHex(styleOf(para).get('background')) || toHex(styleOf(para).get('background-color'));
      const band = extra.band || (own && own !== '#ffffff' ? own : null) || (box && box.fill) || null;
      const align = alignOf(para, prop);
      const record = {
        plain: built.plain,
        runs: built.runs,
        unit: { band, align, valign: extra.valign || null, dark: Boolean(band && luminance(band) < 0.35) },
      };
      return { md: built.md, plain: built.plain, record };
    },

    /** A paragraph or one of Word's list paragraphs, as lines. */
    block(para, lines, run) {
      const p = ctx.paragraph(para, ctx.cell || {});
      if (!p.plain.trim()) {
        // An empty paragraph is Word's spacer: a gap nothing joins across.
        run.stack = [];
        lines.push({ kind: 'gap', hard: true });
        return;
      }
      if (isWordList(para, ctx)) {
        const item = listItem(para, run, styles, styleOf, classRule);
        lines.push({
          kind: 'block',
          list: true,
          text: `${'  '.repeat(item.level)}${item.marker} ${p.md.replace(/\\\n/g, ' ')}`,
          record: { ...p.record, plain: p.record.plain.replace(/\n/g, ' ') },
        });
        return;
      }
      const spacing = (name) => {
        const v = styleOf(para).get(name) || ((classRule(para) || new Map()).get(name)) || normal.get(name);
        return toPt(v);
      };
      lines.push({ kind: 'gap' });
      lines.push({
        kind: 'block',
        text: p.md,
        record: p.record,
        join: {
          next: spacing('margin-bottom') === 0,
          prev: !spacing('margin-top'),
          key: `${p.record.unit.align}|${p.record.unit.band}`,
        },
      });
      lines.push({ kind: 'gap' });
    },

    /** A picture placed by `expandVml`, registered as an image of the document. */
    picture(fig) {
      const pic = pictures[Number(fig.getAttribute('data-op-pic'))];
      if (!pic || !pic.src) return null;
      const id = newImageId();
      images[id] = {
        src: pic.src,
        alt: pic.alt || '',
        w: pic.pxWidth || 0,
        h: pic.pxHeight || 0,
        width: Math.max(10, Math.min(100, Number(fig.getAttribute('data-width')) || 40)),
        align: ['left', 'center', 'right'].includes(fig.getAttribute('data-align')) ? fig.getAttribute('data-align') : 'center',
        wrap: fig.getAttribute('data-wrap') === '1',
      };
      return id;
    },

    table(table, lines) {
      return wordTable(table, lines, ctx, styleOf);
    },
  };
  return ctx;
}

/** A paragraph's alignment, in the four words CSS and the model use. */
function alignOf(para, prop) {
  const v = String(para.getAttribute('align') || prop(para, para, 'text-align') || 'left').toLowerCase();
  return v === 'justify' || v === 'both' ? 'justify' : v === 'center' || v === 'right' ? v : 'left';
}

/**
 * One of Word's list paragraphs: its marker, and how deep it sits.
 *
 * Word's own `level1`/`level2` is the level within *that* list, but a page
 * often nests one list inside another — a checklist under a bullet — so the
 * depth a reader sees is the indent. Depth is therefore read from the left
 * margin, across the run of list paragraphs it belongs to.
 */
function listItem(para, run, styles, styleOf, classRule) {
  const spec = (styleOf(para).get('mso-list') || '').match(/(l\d+)\s+(level\d+)/i);
  const def = spec ? styles.lists.get(`${spec[1].toLowerCase()}:${spec[2].toLowerCase()}`) || new Map() : new Map();
  const margin =
    toPt(styleOf(para).get('margin-left')) ??
    toPt(def.get('margin-left')) ??
    toPt((classRule(para) || new Map()).get('margin-left')) ??
    0;

  const stack = run.stack;
  if (!stack.length || margin > stack[stack.length - 1] + 1) stack.push(margin);
  else while (stack.length > 1 && margin < stack[stack.length - 1] - 1) stack.pop();
  const level = stack.length - 1;

  const format = (def.get('mso-level-number-format') || '').toLowerCase();
  const text = para.getAttribute('data-op-marker') || '';
  const font = para.getAttribute('data-op-marker-font') || def.get('font-family') || '';

  const number = text.match(/^\(?(\d{1,3})[.)]?$/);
  if (number) return { level, marker: `${number[1]}.` };
  // Roman numerals, and legal numbering such as "1.2.", are written as plain
  // numbers: the model has no roman lists, and "i." would read as the letter.
  const roman = text.match(/^\(?([ivxlcdm]+)[.)]$/i);
  if (roman && (/roman/.test(format) || roman[1].length > 1)) return { level, marker: `${romanValue(roman[1])}.` };
  const legal = text.match(/^(\d{1,3}\.)+\d{1,3}\.?$/);
  if (legal) return { level, marker: `${text.replace(/\.$/, '').split('.').pop()}.` };
  const letter = text.match(/^\(?([a-zA-Z])[.)]$/);
  if (letter && format !== 'bullet') return { level, marker: `${letter[1]}.` };
  if (format && format !== 'bullet' && !text) return { level, marker: '1.' };

  const shown = symbolFont(font) ? mapSymbols(text, font) : WORD_BULLETS[text] || text;
  const glyph = shown.length === 1 && LIST_GLYPHS.includes(shown) ? shown : null;
  return { level, marker: glyph || '-' };
}

/** "iv" -> 4. */
function romanValue(text) {
  const values = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
  const digits = text.toLowerCase().split('').map((c) => values[c] || 0);
  return Math.max(1, digits.reduce((sum, v, k) => sum + (v < (digits[k + 1] || 0) ? -v : v), 0));
}

/**
 * Runs -> source text and the record of what it looked like.
 *
 * Bold and italic go into the text as `**` and `*` where the tokeniser will
 * read them back exactly; anything it would not — italic inside a word, bold
 * and italic together, emphasis around a link — is left to a mark instead. The
 * result is checked by tokenising it, so the text and the record cannot
 * disagree about where a word starts.
 */
function toMarkdown(runs) {
  let plain = '';
  const spans = runs.map((r) => {
    const start = plain.length;
    plain += r.text;
    return { ...r, start, end: plain.length };
  });

  const stretches = (test) => {
    const out = [];
    for (let i = 0; i < spans.length; i++) {
      if (!test(spans[i])) continue;
      let j = i;
      while (j + 1 < spans.length && test(spans[j + 1]) && !spans[j + 1].br) j++;
      out.push([i, j]);
      i = j;
    }
    return out;
  };

  const opens = new Map();
  const closes = new Map();
  const put = (map, at, s) => map.set(at, (map.get(at) || '') + s);

  // Links first: a link cannot sit inside emphasis markers.
  const linked = new Set();
  for (const [i, j] of stretches((s) => Boolean(s.href))) {
    const href = spans[i].href;
    if (!spans.slice(i, j + 1).every((s) => s.href === href)) continue;
    // Brackets and spaces at the edges of a link's words stay outside it, so
    // "[Internet]" linked whole still reads as a link on "Internet".
    const whole = plain.slice(spans[i].start, spans[j].end);
    const lead = whole.length - whole.replace(/^[\s[\]]+/, '').length;
    const words = whole.slice(lead).replace(/[\s[\]]+$/, '');
    if (!words || /[[\]\n]/.test(words)) continue;
    put(opens, spans[i].start + lead, '[');
    put(closes, spans[i].start + lead + words.length, `](${href.replace(/^[a-z]+:/i, (x) => x.toLowerCase()).replace(/[()\s]/g, encodeURIComponent)})`);
    for (let k = i; k <= j; k++) linked.add(k);
  }

  const mdBold = new Set();
  for (const [i, j] of stretches((s) => s.bold && !s.br)) {
    if ([...Array(j - i + 1).keys()].some((k) => linked.has(i + k))) continue;
    const text = plain.slice(spans[i].start, spans[j].end);
    const core = text.trim();
    if (!core || /[*\n]/.test(core)) continue;
    const from = spans[i].start + (text.length - text.trimStart().length);
    const to = from + core.length;
    put(opens, from, '**');
    put(closes, to, '**');
    for (let k = i; k <= j; k++) mdBold.add(k);
  }

  const mdItalic = new Set();
  for (const [i, j] of stretches((s) => s.italic && !s.br)) {
    if ([...Array(j - i + 1).keys()].some((k) => linked.has(i + k) || spans[i + k].bold)) continue;
    const text = plain.slice(spans[i].start, spans[j].end);
    const core = text.trim();
    if (!core || /[*\n]/.test(core)) continue;
    const from = spans[i].start + (text.length - text.trimStart().length);
    const to = from + core.length;
    const before = plain[from - 1];
    const after = plain[to];
    if (before !== undefined && !/[\s(]/.test(before)) continue;
    if (after !== undefined && !/[\s).,;:!?]/.test(after)) continue;
    put(opens, from, '*');
    put(closes, to, '*');
    for (let k = i; k <= j; k++) mdItalic.add(k);
  }

  let md = '';
  for (let at = 0; at <= plain.length; at++) {
    if (closes.has(at)) md += closes.get(at);
    if (opens.has(at)) md += opens.get(at);
    if (at < plain.length) md += plain[at] === '\n' ? '\\\n' : plain[at];
  }

  // Check: what the tokeniser reads back must be exactly these words, with
  // exactly the emphasis the markers were meant to carry.
  const want = (k, set) => (k >= 0 ? set.has(k) : false);
  const spanAt = (pos) => spans.findIndex((s) => pos >= s.start && pos < s.end);
  let read = '';
  let ok = true;
  for (const t of tokenizeInline(md.replace(/\\\n/g, '\n'))) {
    for (let c = 0; c < t.text.length && ok; c++) {
      const k = spanAt(read.length + c);
      if (Boolean(t.bold) !== want(k, mdBold) && plain[read.length + c] !== ' ') ok = false;
      if (Boolean(t.italic) !== want(k, mdItalic) && plain[read.length + c] !== ' ') ok = false;
    }
    read += t.text;
  }
  if (!ok || read !== plain) {
    // Fall back to words only; every emphasis becomes a mark.
    mdBold.clear();
    mdItalic.clear();
    md = plain.replace(/\n/g, '\\\n');
  }

  return {
    md,
    plain,
    runs: spans
      .filter((s) => !s.br)
      .map((s, k) => ({
        start: s.start,
        end: s.end,
        bold: s.bold,
        italic: s.italic,
        underline: s.underline,
        pt: s.pt,
        fg: s.fg,
        bg: s.bg,
        href: s.href,
        mdBold: mdBold.has(spans.indexOf(s)),
        mdItalic: mdItalic.has(spans.indexOf(s)),
      })),
  };
}

/**
 * A Word table: a pipe table whose cells keep their paragraphs and lists, a
 * header row only when the table has one, and the cell shading, alignment and
 * column widths as records and attributes.
 */
function wordTable(table, lines, ctx, styleOf) {
  const rows = rowsOf(table);
  if (!rows.length) return false;

  const { grid, width } = tableGrid(rows);
  if (width < 2) return false;

  const cellOut = grid.map((r) =>
    r.map((td) => {
      if (!td) return { text: '', records: [], allBold: false, empty: true };
      if (typeof td === 'string') return { text: td, records: [], allBold: false, empty: true };
      const st = styleOf(td);
      const bg = toHex(st.get('background')) || toHex(st.get('background-color')) || toHex(td.getAttribute('bgcolor'));
      const band = bg && bg !== '#ffffff' ? bg : null;
      const v = (td.getAttribute('valign') || st.get('vertical-align') || 'middle').toLowerCase();
      const valign = v === 'top' || v === 'bottom' ? v : 'middle';

      const inner = [];
      const outer = ctx.cell;
      ctx.cell = { band, valign };
      try {
        walk(td, inner, 0, ctx);
      } finally {
        ctx.cell = outer;
      }
      const merged = joinParagraphs(inner);
      const blocks = merged.filter((l) => l.kind === 'block');
      // A table inside the cell carries its rows' records as `records`.
      const records = merged.flatMap((l) => (l.record ? [l.record] : l.records || []));
      const allBold =
        records.length > 0 &&
        records.every((r) => r.runs.every((x) => x.bold || !r.plain.slice(x.start, x.end).trim()));
      if (!blocks.length) return { text: '', records: [], allBold: false, empty: true };

      // One plain paragraph is an ordinary cell; anything more is a cell that
      // holds a small document, which a `<br>` marks even when it is one item.
      if (blocks.length === 1 && !blocks[0].list && !blocks[0].text.includes('\n')) {
        return { text: blocks[0].text.replace(/\|/g, '\\|'), records, allBold, empty: false, simple: true };
      }
      return { text: joinCell(merged), records, allBold, empty: false };
    })
  );

  // A header row is one whose every filled cell is bold — Word does not say.
  const firstRow = rows[0];
  // A first row that runs down into the next cannot be a header row.
  const header =
    !spansDown(grid, 0) &&
    (cellsOf(firstRow).some((c) => c.tagName.toUpperCase() === 'TH') ||
      (rows.length > 1 && cellOut[0].some((c) => !c.empty) && cellOut[0].every((c) => c.empty || (c.allBold && c.simple))));

  // Each column's width from a cell that starts in it and spans no other, as
  // percentages — a merged cell's width is several columns' together.
  const pt = (td) => Number(td.getAttribute('width')) * 0.75 || toPt(styleOf(td).get('width')) || 0;
  const single = (td) => td && typeof td !== 'string' && spanOf(td, 'colspan', 1000) === 1;
  const w = Array.from({ length: width }, (_, j) => {
    for (const r of grid) if (single(r[j]) && pt(r[j]) > 0) return pt(r[j]);
    return 0;
  });
  let widths = null;
  if (w.every((x) => x > 0)) {
    const total = w.reduce((a, b) => a + b, 0);
    widths = w.map((x) => Math.round((x / total) * 1000) / 10);
  }

  // The rule colour most cells use.
  const colours = new Map();
  for (const r of grid)
    for (const td of r) {
      if (!td || typeof td === 'string') continue;
      const border = styleOf(td).get('border') || styleOf(td).get('border-top') || styleOf(td).get('border-bottom') || '';
      for (const word of border.split(/\s+/)) {
        const hex = toHex(word);
        if (hex) colours.set(hex, (colours.get(hex) || 0) + 1);
      }
    }
  const border = [...colours.entries()].sort((a, b) => b[1] - a[1])[0];

  const row = (cells) => `| ${cells.map((c) => c.text).join(' | ')} |`;
  const recordsOf = (cells) => cells.flatMap((c) => c.records);

  lines.push({ kind: 'gap' });
  if (header) lines.push({ kind: 'block', text: row(cellOut[0]), records: recordsOf(cellOut[0]) });
  else lines.push({ kind: 'block', text: `|${'  |'.repeat(width)}` });
  lines.push({ kind: 'block', text: `|${' --- |'.repeat(width)}` });
  for (const cells of header ? cellOut.slice(1) : cellOut) lines.push({ kind: 'block', text: row(cells), records: recordsOf(cells) });
  const attrs = tableAttrLine(widths, border ? border[0] : null);
  if (attrs) lines.push({ kind: 'block', text: attrs });
  lines.push({ kind: 'gap' });
  return true;
}

/**
 * Paragraphs Word sets with no space between them read as lines of one
 * block — a letterhead, an address — and are joined with hard line breaks, so
 * they keep that look rather than gaining a paragraph's gap each.
 */
function joinParagraphs(lines) {
  const out = [];
  let last = null;
  for (const line of lines) {
    if (line.kind === 'gap') {
      if (line.hard) last = null;
      out.push(line);
      continue;
    }
    if (line.kind === 'block' && line.join && last && last.join && last.join.next && line.join.prev && last.join.key === line.join.key) {
      while (out.length && out[out.length - 1].kind === 'gap') out.pop();
      const shift = last.record.plain.length + 1;
      last.text = `${last.text}\\\n${line.text}`;
      last.record = {
        ...last.record,
        plain: `${last.record.plain}\n${line.record.plain}`,
        runs: [...last.record.runs, ...line.record.runs.map((r) => ({ ...r, start: r.start + shift, end: r.end + shift }))],
      };
      last.join = { ...last.join, next: line.join.next };
      continue;
    }
    out.push(line);
    last = line.kind === 'block' && line.join ? line : null;
  }
  return out;
}

/** The most frequent key of a count map. */
const mode = (map) => [...map.entries()].sort((a, b) => b[1] - a[1])[0];

/**
 * Word's HTML, with the formatting kept.
 *
 * @param {string} html — the text/html flavour
 * @param {{rtf?: string, newImageId: () => string}} opts — `rtf` is the
 *   text/rtf flavour, which is where Word's pictures are
 * @returns {null | {text:string, records:object[], images:object, basePt:number, typeface:string|null}}
 *   null when the HTML is not Word's
 */
export function htmlToRich(html, opts) {
  if (!isWordHtml(html)) return null;
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const styles = readWordStyles(doc);
  const { pictures } = expandVml(doc, styles);
  stripFallbacks(doc);
  doc.querySelectorAll('style, script, xml, head meta, link, title').forEach((n) => n.remove());

  // Each picture takes the next RTF picture of about the same shape; RTF and
  // HTML list them in the same order, and the shape check catches the case
  // where one flavour has a picture the other does not.
  const blips = extractRtfPictures(opts.rtf || '');
  let next = 0;
  for (const pic of pictures) {
    const ratio = pic.heightPt > 0 ? pic.widthPt / pic.heightPt : 0;
    for (let k = next; k < blips.length; k++) {
      const b = blips[k];
      const r = b.heightPt > 0 ? b.widthPt / b.heightPt : 0;
      if (!ratio || !r || Math.abs(r - ratio) / ratio < 0.2) {
        pic.src = `data:${b.mime};base64,${b.base64}`;
        pic.pxWidth = b.pxWidth;
        pic.pxHeight = b.pxHeight;
        next = k + 1;
        break;
      }
    }
  }

  const ctx = wordContext(styles, pictures, opts.newImageId);
  const lines = [];
  walk(doc.body, lines, 0, ctx);
  const { text, records } = assemble(joinParagraphs(lines));
  if (!text) return null;

  const size = mode(ctx.sizes);
  const face = mode(ctx.families);
  return {
    text,
    records,
    images: ctx.images,
    basePt: size ? size[0] : 11,
    typeface: face ? TYPEFACE_OF[face[0]] || null : null,
  };
}

/* --- Records -> marks ------------------------------------------------------ */

/** How large an element is drawn, as a multiple of the document size. */
function elementScale(unit) {
  const table = SCALE.table.size;
  if (unit.kind === 'cell' || unit.kind === 'th') return table;
  const own = SCALE[unit.kind] ? SCALE[unit.kind].size : 1;
  return own * (unit.inTable ? table : 1);
}

/** What an element already looks like with no marks on it. */
function defaultsOf(unit) {
  return {
    bold: /^h[123]$/.test(unit.kind) || unit.kind === 'th',
    italic: unit.kind === 'quote',
    align: unit.kind === 'p' ? 'justify' : 'left',
  };
}

/** One record as marks on its unit: what differs from how the unit draws anyway. */
function recordMarks(unit, rec, basePt) {
  const len = unit.plain.length;
  if (!len) return [];
  const exact = unit.plain === rec.plain;
  const def = defaultsOf(unit);
  let marks = [];

  const whole = {};
  if (rec.unit.band) whole.band = rec.unit.band;
  if (rec.unit.align && rec.unit.align !== def.align) whole.align = rec.unit.align;
  if (unit.inTable && rec.unit.valign && rec.unit.valign !== 'top') whole.valign = rec.unit.valign;
  if (Object.keys(whole).length) marks = formatRange(marks, 0, len, whole, len);
  if (!exact) return marks;

  const scale = elementScale(unit);
  for (const r of rec.runs) {
    const f = {};
    const em = r.pt && basePt ? r.pt / basePt / scale : 1;
    if (Math.abs(em - 1) > 0.07) f.size = Math.round(em * 100) / 100;
    if (Boolean(r.bold) !== Boolean(def.bold || r.mdBold)) f.bold = Boolean(r.bold);
    if (Boolean(r.italic) !== Boolean(def.italic || r.mdItalic)) f.italic = Boolean(r.italic);
    if (r.underline) f.underline = true;
    const fg = r.fg || (rec.unit.dark ? '#ffffff' : unit.kind === 'h3' ? '#000000' : null);
    if (fg) f.fg = fg;
    if (r.bg) f.bg = r.bg;
    if (Object.keys(f).length) marks = formatRange(marks, r.start, r.end, f, len);
  }
  return marks;
}

/**
 * Put a rich paste's formatting onto the document it was pasted into.
 *
 * Each record is matched to the unit with the same words, in order, starting
 * where the paste began — so a paragraph that happens to repeat one above the
 * paste is not the one that gets coloured.
 *
 * @param {import('./marks.js').Unit[]} units — the document after the paste
 * @param {object[]} records — from `htmlToRich`
 * @param {number} from — the first unit the paste can have produced
 * @param {number} basePt — the paste's body size, which becomes "1"
 * @returns {import('./marks.js').StoredMark[]}
 */
export function marksFromRecords(units, records, from, basePt) {
  const out = [];
  const same = (a, b) => a.replace(/\s+/g, ' ').trim() === b.replace(/\s+/g, ' ').trim();
  let at = Math.max(0, from);
  for (const rec of records) {
    let hit = -1;
    for (let k = at; k < units.length; k++) {
      if (same(units[k].plain, rec.plain)) {
        hit = k;
        break;
      }
    }
    if (hit < 0) continue;
    at = hit + 1;
    for (const m of recordMarks(units[hit], rec, basePt)) out.push(storedMark(units[hit], m));
  }
  return out;
}

/* --- Paste events ----------------------------------------------------------- */

/**
 * Pick the best available flavour from a paste event.
 * @param {ClipboardEvent} event
 * @returns {string}
 */
export function textFromPaste(event) {
  const dt = event.clipboardData;
  if (!dt) return '';
  const plain = dt.getData('text/plain') || '';
  const html = dt.getData('text/html') || '';
  if (!html.trim()) return plain;

  const derived = htmlToText(html);
  // Trust the HTML-derived version only when it actually captured the content;
  // some apps put a decorative HTML wrapper on the clipboard with little text.
  const words = (s) => (s.match(/\S+/g) || []).length;
  if (words(derived) >= words(plain) * 0.8) return derived;
  return plain;
}

/**
 * A Word paste with its formatting, or null when the clipboard holds anything
 * else (which then goes through `textFromPaste`).
 *
 * @param {ClipboardEvent} event
 * @param {{newImageId: () => string}} opts
 */
export function richFromPaste(event, opts) {
  const dt = event.clipboardData;
  if (!dt) return null;
  const html = dt.getData('text/html') || '';
  if (!isWordHtml(html)) return null;
  return htmlToRich(html, { rtf: dt.getData('text/rtf') || '', newImageId: opts.newImageId });
}
