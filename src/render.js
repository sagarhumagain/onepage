/**
 * render.js — document model -> HTML for the A4 page.
 *
 * Every tag emitted here is one this file chose; all user text is escaped on
 * the way in. Nothing from the clipboard reaches the DOM as markup, so there
 * is no sanitiser to get wrong.
 */

import { tokenizeInline } from './inline.js';
import { HIGHLIGHT } from './markup.js';

const ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

export const escapeHtml = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ESCAPES[c]);

/** Only http(s) and mailto survive; anything else becomes plain text. */
function safeHref(href) {
  return /^(https?:|mailto:)/i.test(href) ? escapeHtml(href) : null;
}

/** A colour the user picked, reduced to something safe to put in a style. */
const safeColor = (c) => (/^#[0-9a-fA-F]{3,8}$/.test(String(c)) ? String(c) : null);

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
 * Consecutive tokens sharing a highlight become a single <mark>.
 *
 * Emitting one mark per token would be visually wrong rather than merely
 * verbose: padding would be applied at every internal boundary, so a
 * highlighted sentence containing one bold word would render as three boxes
 * with gaps between them instead of one continuous band.
 */
function inline(text) {
  const tokens = tokenizeInline(text);
  const out = [];

  for (let i = 0; i < tokens.length; i++) {
    const bg = safeColor(tokens[i].bg);
    if (!bg) {
      out.push(tokenHtml(tokens[i]));
      continue;
    }

    const pad = Math.max(0, Math.min(99, Number(tokens[i].pad) || 0));
    const run = [];
    while (i < tokens.length && safeColor(tokens[i].bg) === bg && (Number(tokens[i].pad) || 0) === pad) {
      run.push(tokenHtml(tokens[i]));
      i++;
    }
    i--;

    // Vertical padding is deliberately a fraction of the horizontal: on an
    // inline box it spills over the neighbouring lines instead of pushing
    // them apart, so a large value would overlap the text above and below.
    const style = `background:${bg}` + (pad ? `;padding:${Math.round(pad * 3.5) / 100}em ${pad / 10}em` : '');
    out.push(`<mark class="doc-hl" style="${style}">${run.join('')}</mark>`);
  }

  return out.join('');
}

/**
 * A highlight that covers a whole block is rendered as a padded box on the
 * block itself rather than as an inline mark.
 *
 * The distinction is not cosmetic. Vertical padding on an inline box does not
 * grow the line it sits on, so it overlaps the lines above and below; on a
 * block it is real space that the fitter measures and the page layout
 * accounts for. Highlighting a whole paragraph is the case where the user
 * asked for a padded background, so that case gets the padding that works.
 */
const WHOLE_HIGHLIGHT = new RegExp(`^${HIGHLIGHT.source}$`);

function blockHighlight(text) {
  const m = String(text == null ? '' : text).trim().match(WHOLE_HIGHLIGHT);
  if (!m) return null;
  const color = safeColor(m[1]);
  // A closer inside the run means this is two highlights, not one box.
  if (!color || m[3].includes('{=}')) return null;
  const pad = Math.max(0, Math.min(99, Number(m[2]) || 0));
  return { color, pad, inner: m[3] };
}

/** `<tag>` carrying a block highlight when the whole block is highlighted. */
function block(tag, text, attrs = '') {
  const hl = blockHighlight(text);
  if (!hl) return `<${tag}${attrs}>${inline(text)}</${tag}>`;
  const style = `background:${hl.color}` + (hl.pad ? `;padding:${hl.pad / 10}em ${(hl.pad / 10) * 1.3}em` : '');
  return `<${tag}${attrs} class="doc-hl-block" style="${style}">${inline(hl.inner)}</${tag}>`;
}

const cell = (c) => `<td>${inline(c)}</td>`;
const headCell = (c) => `<th>${inline(c)}</th>`;

/**
 * @param {import('./parse.js').Block[]} blocks
 * @returns {string} HTML
 */
export function render(blocks) {
  const out = [];
  for (const b of blocks) {
    switch (b.type) {
      case 'h1':
      case 'h2':
      case 'h3':
        out.push(block(b.type, b.text));
        break;
      case 'p':
        out.push(block('p', b.text));
        break;
      case 'quote':
        out.push(block('blockquote', b.text));
        break;
      case 'code':
        out.push(`<pre><code>${escapeHtml(b.text)}</code></pre>`);
        break;
      case 'hr':
        out.push('<hr>');
        break;
      case 'ul':
        out.push(`<ul>${b.items.map((it) => block('li', it)).join('')}</ul>`);
        break;
      case 'ol': {
        const start = b.start && b.start !== 1 ? ` start="${b.start}"` : '';
        out.push(`<ol${start}>${b.items.map((it) => block('li', it)).join('')}</ol>`);
        break;
      }
      case 'table': {
        const head = b.head && b.head.length ? `<thead><tr>${b.head.map(headCell).join('')}</tr></thead>` : '';
        const cols = b.head ? b.head.length : 0;
        const body = (b.rows || [])
          .map((r) => {
            const cells = r.slice(0, cols || r.length);
            while (cols && cells.length < cols) cells.push('');
            return `<tr>${cells.map(cell).join('')}</tr>`;
          })
          .join('');
        out.push(`<table>${head}<tbody>${body}</tbody></table>`);
        break;
      }
      case 'image': {
        // A block whose bytes were never resolved would render as a broken
        // image icon in the middle of the page; leaving it out is quieter and
        // the source text still shows the reference.
        if (!b.src) break;
        const align = b.align === 'left' || b.align === 'right' ? b.align : 'center';
        const width = Math.max(5, Math.min(100, Number(b.widthPct) || 62));
        const alt = escapeHtml(b.alt || '');
        const caption = b.alt ? `<figcaption>${alt}</figcaption>` : '';
        out.push(
          `<figure class="doc-image" data-align="${align}" data-ref="${escapeHtml(b.ref || '')}" style="--img-width:${width}%">` +
            `<img src="${escapeHtml(b.src)}" alt="${alt}" draggable="false">${caption}</figure>`
        );
        break;
      }
      default:
        break;
    }
  }
  return out.join('\n');
}
