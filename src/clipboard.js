/**
 * clipboard.js — turn whatever the clipboard offers into clean structured text.
 *
 * The app's canonical input is plain text, because that is what the user
 * edits and what every exporter reads. But Word's and a browser's plain-text
 * clipboard flavour throws away list nesting and headings, while their HTML
 * flavour keeps that structure buried in a mound of `mso-` noise.
 *
 * So when HTML is on the clipboard we walk it for *structure only* and emit
 * the lightweight text conventions that parse.js already understands. No
 * source styling survives, which is exactly the intent.
 */

/** Elements whose text content is markup metadata, never document content. */
const SKIP = new Set(['STYLE', 'SCRIPT', 'HEAD', 'META', 'LINK', 'TITLE', 'NOSCRIPT', 'XML', 'O:P']);
const BLOCK = new Set([
  'P', 'DIV', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'MAIN', 'ASIDE',
  'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI', 'BLOCKQUOTE',
  'PRE', 'TABLE', 'TR', 'HR', 'FIGURE', 'FIGCAPTION', 'DL', 'DT', 'DD',
]);

const collapse = (s) => s.replace(/[\s ]+/g, ' ');

/**
 * Word does not emit <ul>/<li> for bulleted lists. It emits a flat run of
 * <p class=MsoListParagraph style='mso-list:l0 level1 lfo1'> elements, each
 * beginning with a literal bullet glyph in a Symbol-font span (and, in Word's
 * older output, the letter "o" or "v" for the second and third levels).
 *
 * Detecting that and re-emitting a real list marker is the difference between
 * a pasted Word document keeping its bullets and silently losing them.
 */
const BULLET_GLYPH = /^[\s ]*([·•●▪◦‣§-]|o|v)[\s ]+/;

const isWordListItem = (el) => {
  const cls = el.getAttribute('class') || '';
  const style = el.getAttribute('style') || '';
  return /MsoList/i.test(cls) || /mso-list\s*:/i.test(style);
};

const stripLeadingBullet = (s) => s.replace(BULLET_GLYPH, '').trim();

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
    if (tag === 'STRONG' || tag === 'B') {
      const inner = textOf(child);
      out.push(inner.trim() ? `**${inner.trim()}**` : '');
      continue;
    }
    if (tag === 'EM' || tag === 'I') {
      const inner = textOf(child);
      out.push(inner.trim() ? `*${inner.trim()}*` : '');
      continue;
    }
    if (tag === 'CODE' || tag === 'KBD' || tag === 'SAMP') {
      const inner = textOf(child);
      out.push(inner.trim() ? `\`${inner.trim()}\`` : '');
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

function walk(node, lines, depth) {
  for (const child of node.childNodes) {
    if (child.nodeType === 3) {
      const t = collapse(child.nodeValue);
      if (t.trim()) lines.push({ kind: 'text', text: t });
      continue;
    }
    if (child.nodeType !== 1) continue;

    const tag = child.tagName.toUpperCase();
    if (SKIP.has(tag)) continue;

    if (/^H[1-6]$/.test(tag)) {
      const level = Math.min(Number(tag[1]), 3);
      const t = textOf(child).trim();
      if (t) lines.push({ kind: 'block', text: `${'#'.repeat(level)} ${t}` });
      continue;
    }

    if (tag === 'HR') {
      lines.push({ kind: 'block', text: '---' });
      continue;
    }

    if (tag === 'UL' || tag === 'OL') {
      const ordered = tag === 'OL';
      let n = Number(child.getAttribute('start')) || 1;
      const pad = '  '.repeat(depth);
      for (const li of child.children) {
        if (li.tagName.toUpperCase() !== 'LI') continue;
        // Nested lists are rendered by the recursive call below, so take
        // only this item's own text here.
        const own = document.createElement('div');
        for (const kid of li.childNodes) {
          const t = kid.nodeType === 1 ? kid.tagName.toUpperCase() : '';
          if (t === 'UL' || t === 'OL') continue;
          own.appendChild(kid.cloneNode(true));
        }
        const t = stripLeadingBullet(textOf(own));
        if (t) lines.push({ kind: 'block', text: `${pad}${ordered ? `${n++}.` : '-'} ${t}` });
        for (const kid of li.children) {
          const kt = kid.tagName.toUpperCase();
          if (kt === 'UL' || kt === 'OL') walk(li, lines, depth + 1);
        }
      }
      lines.push({ kind: 'gap' });
      continue;
    }

    if (tag === 'TABLE') {
      const rows = [...child.querySelectorAll('tr')];
      if (rows.length) {
        const cellsOf = (tr) => [...tr.children].filter((c) => /^T[HD]$/.test(c.tagName.toUpperCase()));
        const headRow = cellsOf(rows[0]);
        const width = Math.max(...rows.map((r) => cellsOf(r).length));
        if (width > 1) {
          lines.push({ kind: 'gap' });
          lines.push({ kind: 'block', text: `| ${headRow.map(cellText).join(' | ')} |` });
          lines.push({ kind: 'block', text: `|${' --- |'.repeat(width)}` });
          for (const tr of rows.slice(1)) {
            const cells = cellsOf(tr).map(cellText);
            while (cells.length < width) cells.push('');
            lines.push({ kind: 'block', text: `| ${cells.join(' | ')} |` });
          }
          lines.push({ kind: 'gap' });
          continue;
        }
      }
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
      if (!hasBlockChild) {
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
        walk(child, lines, depth);
        lines.push({ kind: 'gap' });
      }
      continue;
    }

    walk(child, lines, depth);
  }
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
  walk(doc.body, lines, 0);

  const out = [];
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
  }

  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

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
