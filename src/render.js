/**
 * render.js — document model -> HTML for the A4 page.
 *
 * Every tag emitted here is one this file chose; all user text is escaped on
 * the way in. Nothing from the clipboard reaches the DOM as markup, so there
 * is no sanitiser to get wrong.
 */

import { tokenizeInline } from './inline.js';

const ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

export const escapeHtml = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ESCAPES[c]);

/** Only http(s) and mailto survive; anything else becomes plain text. */
function safeHref(href) {
  return /^(https?:|mailto:)/i.test(href) ? escapeHtml(href) : null;
}

function inline(text) {
  return tokenizeInline(text)
    .map((t) => {
      let html = escapeHtml(t.text);
      if (t.code) html = `<code>${html}</code>`;
      if (t.bold) html = `<strong>${html}</strong>`;
      if (t.italic) html = `<em>${html}</em>`;
      if (t.href) {
        const href = safeHref(t.href);
        if (href) html = `<a href="${href}">${html}</a>`;
      }
      return html;
    })
    .join('');
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
        out.push(`<${b.type}>${inline(b.text)}</${b.type}>`);
        break;
      case 'p':
        out.push(`<p>${inline(b.text)}</p>`);
        break;
      case 'quote':
        out.push(`<blockquote>${inline(b.text)}</blockquote>`);
        break;
      case 'code':
        out.push(`<pre><code>${escapeHtml(b.text)}</code></pre>`);
        break;
      case 'hr':
        out.push('<hr>');
        break;
      case 'ul':
        out.push(`<ul>${b.items.map((it) => `<li>${inline(it)}</li>`).join('')}</ul>`);
        break;
      case 'ol': {
        const start = b.start && b.start !== 1 ? ` start="${b.start}"` : '';
        out.push(`<ol${start}>${b.items.map((it) => `<li>${inline(it)}</li>`).join('')}</ol>`);
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
        const align = b.align || 'center';
        const alt = escapeHtml(b.alt || '');
        const caption = b.alt ? `<div class="doc-image-caption">${alt}</div>` : '';
        out.push(`<div class="doc-image" data-align="${align}"><img src="${escapeHtml(b.src)}" alt="${alt}" />${caption}</div>`);
        break;
      }
      default:
        break;
    }
  }
  return out.join('\n');
}
