/**
 * render.js — document model -> HTML for the A4 page.
 *
 * Every tag emitted here is one this file chose; all user text is escaped on
 * the way in. Nothing from the clipboard reaches the DOM as markup, so there
 * is no sanitiser to get wrong.
 */

import { tokenizeMarked, plainText } from './inline.js';
import { LIST_GLYPHS } from './parse.js';

const ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

export const escapeHtml = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ESCAPES[c]);

const HEX = /^#[0-9a-f]{6}$/i;
const ALIGNMENTS = new Set(['left', 'center', 'right']);
const TEXT_ALIGNS = new Set(['left', 'center', 'right', 'justify']);
const VERTICAL_ALIGNS = new Set(['top', 'middle', 'bottom']);
const LIST_STYLES = { 'lower-alpha': 'a', 'upper-alpha': 'A' };

/** Only http(s) and mailto survive; anything else becomes plain text. */
function safeHref(href) {
  return /^(https?:|mailto:)/i.test(href) ? escapeHtml(href) : null;
}

/** An image is only ever one the app itself read in, or a plain web URL. */
function safeSrc(src) {
  const s = String(src || '');
  return /^(data:image\/[a-z0-9.+-]+;base64,|https?:\/\/)/i.test(s) ? escapeHtml(s) : '';
}

const safeColor = (c) => (HEX.test(String(c || '')) ? String(c).toLowerCase() : null);

/**
 * A run's formatting as inline CSS.
 *
 * Only these five properties exist, every colour must be a plain hex and the
 * size is a multiplier expressed in em — so a formatted run still scales with
 * whatever size the fitter lands on, and nothing a user typed can become a
 * style declaration.
 */
function styleFor(fmt) {
  if (!fmt) return null;
  const decls = [];
  const bg = safeColor(fmt.bg);
  const fg = safeColor(fmt.fg);
  const size = Number(fmt.size);
  if (bg) decls.push(`background:${bg}`);
  if (fg) decls.push(`color:${fg}`);
  if (size > 0.2 && size < 5 && size !== 1) decls.push(`font-size:${Math.round(size * 1000) / 1000}em`);
  if (fmt.bold === true) decls.push('font-weight:700');
  else if (fmt.bold === false) decls.push('font-weight:400');
  if (fmt.italic === true) decls.push('font-style:italic');
  else if (fmt.italic === false) decls.push('font-style:normal');
  if (fmt.underline) decls.push('text-decoration:underline');
  if (!decls.length) return null;
  return { css: decls.join(';'), marked: Boolean(bg) };
}

/**
 * The full-width band behind a whole line, if it has one.
 *
 * It lives on the block element rather than on a span, because that is the
 * only way a background reaches the edge of the column instead of stopping
 * where the words do.
 */
function bandOf(marks) {
  if (!marks) return null;
  for (const m of marks) {
    const color = m.fmt && safeColor(m.fmt.band);
    if (color) return color;
  }
  return null;
}

/**
 * A size every character of a list item shares. It is drawn on the item
 * rather than on its words, so the number or bullet the item draws for itself
 * is that size too — a list of references set small has small numbers.
 */
function uniformSize(marks, length) {
  if (!marks || !marks.length || !length) return null;
  const size = marks[0].fmt && Number(marks[0].fmt.size);
  if (!(size > 0.2 && size < 5) || size === 1) return null;
  let at = 0;
  for (const m of marks) {
    if (m.start !== at || !m.fmt || Number(m.fmt.size) !== size) return null;
    at = m.end;
  }
  return at >= length ? size : null;
}

/** The same marks without their size, once the element carries it. */
const withoutSize = (marks) =>
  marks.map((m) => {
    const fmt = { ...m.fmt };
    delete fmt.size;
    return { ...m, fmt };
  });

/** A whole-unit property — alignment, vertical alignment — if one is set. */
function unitProp(marks, key, allowed) {
  if (!marks) return null;
  for (const m of marks) {
    const v = m.fmt && m.fmt[key];
    if (allowed.has(v)) return v;
  }
  return null;
}

/** Width as a percentage of the column, clamped to something drawable. */
export const clampWidth = (w) => Math.max(10, Math.min(100, Math.round(Number(w) || 60)));

/** One token as HTML, with its inline emphasis, code and link wrappers. */
/**
 * A hard line break is a newline in the text, and it stays one in the page: a
 * preserved newline in a span of its own rather than a <br>. A <br> has no
 * text, so every character offset after it — which is what a selection, a
 * mark and the caret are measured in — would be one short.
 */
const LINE_BREAK = '<span class="doc-br">\n</span>';

function tokenHtml(t) {
  let html = escapeHtml(t.text).replace(/\n/g, LINE_BREAK);
  if (t.code) html = `<code>${html}</code>`;
  if (t.bold) html = `<strong>${html}</strong>`;
  if (t.italic) html = `<em>${html}</em>`;
  if (t.href) {
    const href = safeHref(t.href);
    if (href) html = `<a href="${href}">${html}</a>`;
  }
  return html;
}

/**
 * @param {string} text
 * @param {{start:number, end:number, fmt:object}[]} [marks]
 */
function inline(text, marks) {
  const tokens = tokenizeMarked(text, marks);
  const out = [];

  for (let i = 0; i < tokens.length; i++) {
    const style = styleFor(tokens[i].fmt);
    if (!style) {
      out.push(tokenHtml(tokens[i]));
      continue;
    }

    // Consecutive tokens with identical formatting become one element, so a
    // highlight that spans a bold word is a single unbroken band rather than
    // three abutting ones.
    const run = [tokenHtml(tokens[i])];
    while (i + 1 < tokens.length) {
      const next = styleFor(tokens[i + 1].fmt);
      if (!next || next.css !== style.css) break;
      run.push(tokenHtml(tokens[++i]));
    }

    const tag = style.marked ? 'mark' : 'span';
    out.push(`<${tag} style="${style.css}">${run.join('')}</${tag}>`);
  }

  return out.join('');
}

/**
 * @param {import('./parse.js').Block[]} blocks — resolved by marks.js, so
 *   `_marks` and `_u` may be present
 * @param {{editor?: boolean}} [opts] — `editor` stamps each markable run with
 *   its unit index, which is how the preview turns a selection back into
 *   formatting, and makes images atomic so typing around one cannot get inside
 *   it. Output needs neither.
 * @returns {string} HTML
 */
export function render(blocks, opts = {}) {
  const out = [];
  const marksOf = (b, sub) => (b._marks ? b._marks[sub] : null);
  const idOf = (b, sub) =>
    opts.editor && b._u && b._u[sub] != null ? ` data-u="${b._u[sub]}"` : '';

  /**
   * The declarations a unit's element carries for its whole-unit marks.
   *
   * Inside a shaded cell the shading is the cell's, so the units in it keep
   * their alignment but not a band of their own — a band's inset would draw a
   * second, slightly larger box inside the first.
   */
  const unitCss = (marks, size) => {
    const decls = [];
    const band = opts.shaded ? null : bandOf(marks);
    const align = unitProp(marks, 'align', TEXT_ALIGNS);
    const valign = unitProp(marks, 'valign', VERTICAL_ALIGNS);
    if (band) decls.push(`background:${band}`);
    if (align) decls.push(`text-align:${align}`);
    if (valign) decls.push(`vertical-align:${valign}`);
    if (size) decls.push(`font-size:${Math.round(size * 1000) / 1000}em`);
    return { band, css: decls.join(';') };
  };

  /** A list item's size, when the whole of it is one size; see uniformSize. */
  const sizeOf = (b, sub, value) =>
    b.type === 'ul' || b.type === 'ol' ? uniformSize(marksOf(b, sub), plainText(value).length) : null;

  /** The attributes a unit's element carries: its id, its band, its alignment and size. */
  const attrs = (b, sub, value) => {
    const size = value == null ? null : sizeOf(b, sub, value);
    const { band, css } = unitCss(marksOf(b, sub), size);
    return `${idOf(b, sub)}${band ? ' class="doc-band"' : ''}${css ? ` style="${css}"` : ''}`;
  };

  const text = (b, sub, value) => {
    const marks = marksOf(b, sub);
    return inline(value, sizeOf(b, sub, value) ? withoutSize(marks) : marks);
  };

  /**
   * A cell. A cell that holds paragraphs and lists is rendered as the small
   * document it is; when every unit inside it carries the same band, that
   * band is the cell's shading and colours the whole cell, as Word draws it.
   */
  const cell = (b, sub, value, header) => {
    const tag = header ? 'th' : 'td';
    const inner = b.cells && b.cells[sub];
    if (!inner) return `<${tag}${attrs(b, sub, value)}>${text(b, sub, value)}</${tag}>`;

    const shared = sharedCellProps(inner);
    const decls = [];
    if (shared.band) decls.push(`background:${shared.band}`);
    if (shared.valign) decls.push(`vertical-align:${shared.valign}`);
    const style = decls.length ? ` style="${decls.join(';')}"` : '';
    return `<${tag} class="doc-cell"${style}>${render(inner, { ...opts, shaded: Boolean(shared.band) })}</${tag}>`;
  };

  for (const b of blocks) {
    switch (b.type) {
      case 'h1':
      case 'h2':
      case 'h3':
        out.push(`<${b.type}${attrs(b, 'text', b.text)}>${text(b, 'text', b.text)}</${b.type}>`);
        break;
      case 'p':
        out.push(`<p${attrs(b, 'text', b.text)}>${text(b, 'text', b.text)}</p>`);
        break;
      case 'quote':
        out.push(`<blockquote${attrs(b, 'text', b.text)}>${text(b, 'text', b.text)}</blockquote>`);
        break;
      case 'code':
        out.push(`<pre><code>${escapeHtml(b.text)}</code></pre>`);
        break;
      case 'hr':
        out.push('<hr>');
        break;
      case 'ul':
      case 'ol':
        out.push(renderList(b, (j) => `<li${attrs(b, `items.${j}`, b.items[j])}${glyphAttr(b, j)}>${text(b, `items.${j}`, b.items[j])}`));
        break;
      case 'table': {
        const head =
          !b.headless && b.head && b.head.length
            ? `<thead><tr>${b.head.map((c, j) => cell(b, `head.${j}`, c, true)).join('')}</tr></thead>`
            : '';
        const body = (b.rows || [])
          .map((r, ri) => `<tr>${r.map((c, j) => cell(b, `rows.${ri}.${j}`, c, false)).join('')}</tr>`)
          .join('');
        const cols =
          b.widths && b.widths.length
            ? `<colgroup>${b.widths.map((w) => `<col style="width:${clampPercent(w)}%">`).join('')}</colgroup>`
            : '';
        const border = safeColor(b.border);
        const tableAttrs = [
          b.headless ? ' data-head="0"' : '',
          b.widths && b.widths.length ? ` data-widths="${b.widths.map(clampPercent).join(' ')}"` : '',
          border ? ` data-border="${border}" style="--table-rule:${border}"` : '',
        ].join('');
        out.push(`<table${tableAttrs}>${cols}${head}<tbody>${body}</tbody></table>`);
        break;
      }
      case 'image': {
        const src = safeSrc(b.src);
        if (!src) break;
        const align = ALIGNMENTS.has(b.align) ? b.align : 'center';
        const alt = escapeHtml(b.alt || '');
        const caption = b.caption ? `<figcaption>${inline(b.caption)}</figcaption>` : '';
        out.push(
          `<figure class="doc-image" data-img="${escapeHtml(b.id || '')}" data-align="${align}"` +
            ` data-wrap="${b.wrap ? '1' : '0'}" style="--img-w:${clampWidth(b.width)}"` +
            `${opts.editor ? ' contenteditable="false"' : ''}>` +
            `<img src="${src}" alt="${alt}">${caption}</figure>`
        );
        break;
      }
      default:
        break;
    }
  }
  return out.join('\n');
}

const clampPercent = (w) => Math.max(1, Math.min(100, Math.round(Number(w) * 100) / 100));

/** The band and vertical alignment every unit in a cell agrees on, if they do. */
function sharedCellProps(blocks) {
  const all = [];
  for (const b of blocks) {
    const subs = b.type === 'ul' || b.type === 'ol' ? b.items.map((_, j) => `items.${j}`) : ['text'];
    for (const sub of subs) all.push(b._marks ? b._marks[sub] : null);
  }
  const agree = (read) => {
    const first = read(all[0]);
    return first && all.every((m) => read(m) === first) ? first : null;
  };
  return {
    band: all.length ? agree(bandOf) : null,
    valign: all.length ? agree((m) => unitProp(m, 'valign', VERTICAL_ALIGNS)) : null,
  };
}

/** A drawn bullet: only a glyph the parser itself recognises reaches the page. */
const glyphOf = (b, j) => {
  const g = b.glyphs ? b.glyphs[j] : null;
  return g && g.length === 1 && LIST_GLYPHS.includes(g) ? g : null;
};
const glyphAttr = (b, j) => (glyphOf(b, j) ? ` data-glyph="${escapeHtml(glyphOf(b, j))}"` : '');

/**
 * A list, nested by each item's level. Items stay one flat array in the model
 * — `items.N` addresses an item however deep it sits — and the nesting is
 * rebuilt here: a deeper item opens a list inside the item above it, and a
 * change of kind at the same depth closes one list and opens the other.
 *
 * @param {import('./parse.js').ListBlock} b
 * @param {(j:number) => string} itemOpen — the item's opening tag and text
 */
function renderList(b, itemOpen) {
  const n = b.items.length;
  const levelOf = (j) => (b.levels ? b.levels[j] : 0);
  const kindOf = (j) => (b.kinds ? b.kinds[j] : b.type);

  const openList = (j) => {
    const kind = kindOf(j);
    if (kind !== 'ol') return '<ul>';
    const start = b.nums ? b.nums[j] : b.start;
    const type = LIST_STYLES[b.styles ? b.styles[j] : null];
    return `<ol${start && start !== 1 ? ` start="${start}"` : ''}${type ? ` type="${type}"` : ''}>`;
  };

  let html = '';
  const stack = [];
  const close = () => {
    const top = stack.pop();
    html += `${top.open ? '</li>' : ''}</${top.kind}>`;
  };

  for (let j = 0; j < n; j++) {
    const level = levelOf(j);
    while (stack.length > level + 1) close();
    if (stack.length === level + 1 && stack[level].kind !== kindOf(j)) close();
    while (stack.length < level + 1) {
      html += openList(j);
      stack.push({ kind: kindOf(j), open: false });
    }
    const top = stack[stack.length - 1];
    if (top.open) html += '</li>';
    html += itemOpen(j);
    top.open = true;
  }
  while (stack.length) close();
  return html;
}
