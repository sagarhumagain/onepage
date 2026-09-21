/**
 * render.js — document model -> HTML for the A4 page.
 *
 * Every tag emitted here is one this file chose; all user text is escaped on
 * the way in. Nothing from the clipboard reaches the DOM as markup, so there
 * is no sanitiser to get wrong.
 */

import { tokenizeMarked } from './inline.js';

const ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

export const escapeHtml = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ESCAPES[c]);

const HEX = /^#[0-9a-f]{6}$/i;
const ALIGNMENTS = new Set(['left', 'center', 'right']);

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
  if (fmt.bold) decls.push('font-weight:700');
  if (fmt.italic) decls.push('font-style:italic');
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

/** Width as a percentage of the column, clamped to something drawable. */
export const clampWidth = (w) => Math.max(10, Math.min(100, Math.round(Number(w) || 60)));

/** One token as HTML, with its inline emphasis, code and link wrappers. */
function tokenHtml(t) {
  let html = escapeHtml(t.text);
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

  /** The attributes a unit's element carries: its id and its band, if any. */
  const attrs = (b, sub) => {
    const band = bandOf(marksOf(b, sub));
    return `${idOf(b, sub)}${band ? ` class="doc-band" style="background:${band}"` : ''}`;
  };

  const cell = (b, sub, text, header) =>
    `<${header ? 'th' : 'td'}${attrs(b, sub)}>${inline(text, marksOf(b, sub))}</${header ? 'th' : 'td'}>`;

  for (const b of blocks) {
    switch (b.type) {
      case 'h1':
      case 'h2':
      case 'h3':
        out.push(`<${b.type}${attrs(b, 'text')}>${inline(b.text, marksOf(b, 'text'))}</${b.type}>`);
        break;
      case 'p':
        out.push(`<p${attrs(b, 'text')}>${inline(b.text, marksOf(b, 'text'))}</p>`);
        break;
      case 'quote':
        out.push(`<blockquote${attrs(b, 'text')}>${inline(b.text, marksOf(b, 'text'))}</blockquote>`);
        break;
      case 'code':
        out.push(`<pre><code>${escapeHtml(b.text)}</code></pre>`);
        break;
      case 'hr':
        out.push('<hr>');
        break;
      case 'ul':
      case 'ol': {
        const start = b.type === 'ol' && b.start && b.start !== 1 ? ` start="${b.start}"` : '';
        const items = b.items
          .map((it, j) => `<li${attrs(b, `items.${j}`)}>${inline(it, marksOf(b, `items.${j}`))}</li>`)
          .join('');
        out.push(`<${b.type}${start}>${items}</${b.type}>`);
        break;
      }
      case 'table': {
        const head =
          b.head && b.head.length
            ? `<thead><tr>${b.head.map((c, j) => cell(b, `head.${j}`, c, true)).join('')}</tr></thead>`
            : '';
        const body = (b.rows || [])
          .map((r, ri) => `<tr>${r.map((c, j) => cell(b, `rows.${ri}.${j}`, c, false)).join('')}</tr>`)
          .join('');
        out.push(`<table>${head}<tbody>${body}</tbody></table>`);
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
