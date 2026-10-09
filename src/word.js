/**
 * word.js — Word's clipboard HTML, made into ordinary HTML first.
 *
 * Word does not write the HTML a browser would. Three things in it matter here:
 *
 *   - Text boxes, shapes and floating pictures are VML inside conditional
 *     comments (`<!--[if gte vml 1]>...<![endif]-->`). To a browser that is a
 *     comment, so a title bar, a "Key facts" box or a logo simply is not there;
 *     the fallback Word offers instead is a picture of the box, at a file path
 *     no page can read. The VML is parsed here and its content put back into
 *     the document as real elements, in the order a reader meets them.
 *
 *   - Lists are not <ul>/<li>. They are paragraphs carrying `mso-list:l3
 *     level1`, whose marker is literal text in a hidden span and whose
 *     indentation lives in `@list` rules in the <style> block.
 *
 *   - Most formatting is in class rules (`p.MsoNormal {font-size:11pt}`), so
 *     an element's size or font is only known by reading the stylesheet.
 *
 * Nothing here decides what the document *is*; clipboard.js still does that,
 * by walking the DOM this leaves behind.
 */

export const isWordHtml = (html) =>
  /urn:schemas-microsoft-com:office:word|<meta[^>]+content=["']?Word\.Document|class=["']?Mso/i.test(String(html || ''));

/* --- CSS, as Word writes it ---------------------------------------------- */

/** `a:b; c:d` -> Map. Keys lower-cased; values trimmed. */
export function declarations(text) {
  const out = new Map();
  for (const part of String(text || '').split(';')) {
    const at = part.indexOf(':');
    if (at < 0) continue;
    const key = part.slice(0, at).trim().toLowerCase();
    if (key) out.set(key, part.slice(at + 1).trim());
  }
  return out;
}

/** A CSS length in points. Bare numbers are returned as-is (VML group units). */
export function toPt(value) {
  const m = String(value || '').trim().match(/^(-?[\d.]+)\s*(pt|px|in|cm|mm|pc)?$/i);
  if (!m) return null;
  const n = Number(m[1]);
  switch ((m[2] || '').toLowerCase()) {
    case 'px':
      return n * 0.75;
    case 'in':
      return n * 72;
    case 'cm':
      return (n / 2.54) * 72;
    case 'mm':
      return (n / 25.4) * 72;
    case 'pc':
      return n * 12;
    default:
      return n;
  }
}

const NAMED = {
  black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff',
  yellow: '#ffff00', gray: '#808080', grey: '#808080', silver: '#c0c0c0', maroon: '#800000',
  navy: '#000080', olive: '#808000', purple: '#800080', teal: '#008080', aqua: '#00ffff',
  cyan: '#00ffff', fuchsia: '#ff00ff', magenta: '#ff00ff', lime: '#00ff00', orange: '#ffa500',
  windowtext: '#000000', darkblue: '#00008b', darkred: '#8b0000', darkgreen: '#006400',
  darkcyan: '#008b8b', darkmagenta: '#8b008b', darkyellow: '#808000', lightgray: '#d3d3d3',
  darkgray: '#a9a9a9',
};

/** Any colour Word writes, as `#rrggbb`, or null for none / automatic. */
export function toHex(value) {
  const v = String(value || '').trim().toLowerCase().split(/\s+/)[0];
  if (!v || v === 'auto' || v === 'none' || v === 'transparent') return null;
  if (/^#[0-9a-f]{6}$/.test(v)) return v;
  if (/^#[0-9a-f]{3}$/.test(v)) return `#${v[1]}${v[1]}${v[2]}${v[2]}${v[3]}${v[3]}`;
  const rgb = v.match(/^rgb\((\d+),(\d+),(\d+)\)$/);
  if (rgb) return `#${rgb.slice(1).map((n) => Number(n).toString(16).padStart(2, '0')).join('')}`;
  return NAMED[v] || null;
}

/** Relative luminance, 0 (black) to 1 (white). */
export function luminance(hex) {
  const n = parseInt(hex.slice(1), 16);
  const c = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((x) => {
    const s = x / 255;
    return s <= 0.03928 ? s / 12.92 : Math.pow((s + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
}

/**
 * The class rules, list levels and page box out of Word's <style> block.
 *
 * @param {Document} doc
 */
export function readWordStyles(doc) {
  const classes = new Map();
  const lists = new Map();
  let page = null;

  const css = [...doc.querySelectorAll('style')].map((s) => s.textContent).join('\n').replace(/\/\*[\s\S]*?\*\//g, '');
  for (const m of css.matchAll(/([^{}]+)\{([^}]*)\}/g)) {
    const selectors = m[1].replace(/<!--|-->/g, '').trim();
    const decls = declarations(m[2]);
    const list = selectors.match(/^@list\s+(l\d+):(level\d+)$/i);
    if (list) {
      lists.set(`${list[1].toLowerCase()}:${list[2].toLowerCase()}`, decls);
      continue;
    }
    if (/^@page\s+WordSection1$/i.test(selectors)) {
      page = decls;
      continue;
    }
    for (const sel of selectors.split(',')) {
      const s = sel.trim().toLowerCase();
      if (/^[a-z0-9]*\.[a-z0-9_-]+$/.test(s)) classes.set(s, decls);
    }
  }

  // The text column: page width less its side margins. A4 with an inch each
  // side when Word did not say.
  let contentWidthPt = 451;
  let marginLeftPt = 72;
  if (page) {
    const size = (page.get('size') || '').split(/\s+/).map(toPt);
    const margin = (page.get('margin') || '').split(/\s+/).map(toPt);
    const [top, right = top, , left = right] = margin;
    if (size[0] > 0) {
      marginLeftPt = left == null ? 72 : left;
      contentWidthPt = size[0] - marginLeftPt - (right == null ? 72 : right);
    }
  }

  const link = toHex(doc.body && doc.body.getAttribute('link'));
  return { classes, lists, contentWidthPt, marginLeftPt, linkColor: link };
}

/* --- Symbol fonts --------------------------------------------------------- */

/*
 * Word writes a symbol-font character as its code in that font: a tick is a
 * "ü" in Wingdings, a bullet a "·" in Symbol — or the same code moved into
 * the private-use area at U+F000. Shown in any other font they are the wrong
 * letters, so they are mapped to the Unicode character they draw.
 */
const WINGDINGS = {
  0x6c: '●', 0x6e: '■', 0x71: '❑', 0x75: '◆', 0x76: '❖', 0x9f: '•',
  0xa7: '▪', 0xa8: '□', 0xd8: '➢', 0xe0: '→', 0xe8: '➔', 0xf0: '⇨',
  0xfb: '✗', 0xfc: '✓', 0xfd: '☒', 0xfe: '☑',
};
const SYMBOL = {
  0xb7: '•', 0xae: '→', 0xac: '←', 0xb3: '≥', 0xa3: '≤', 0xb1: '±',
  0xb4: '×', 0xb8: '÷', 0xb0: '°', 0xa5: '∞', 0xb9: '≠', 0xbb: '≈',
  0xd6: '√', 0x61: 'α', 0x62: 'β', 0x67: 'γ', 0x64: 'δ', 0x65: 'ε',
  0x71: 'θ', 0x6c: 'λ', 0x6d: 'μ', 0x70: 'π', 0x73: 'σ', 0x74: 'τ',
  0x66: 'φ', 0x77: 'ω', 0x44: 'Δ', 0x53: 'Σ', 0x57: 'Ω',
};

/** Which symbol font, if any, a font-family names. */
export function symbolFont(family) {
  const f = String(family || '').toLowerCase();
  if (f.includes('wingdings')) return WINGDINGS;
  if (/(^|["'\s,])symbol(["'\s,]|$)/.test(f)) return SYMBOL;
  return null;
}

/** Text in a symbol font, as the characters it actually shows. */
export function mapSymbols(text, family) {
  const table = symbolFont(family);
  let out = '';
  for (const c of String(text)) {
    let code = c.codePointAt(0);
    const pua = code >= 0xf020 && code <= 0xf0ff;
    if (pua) code -= 0xf000;
    if (table && table[code]) out += table[code];
    else if (pua) out += (WINGDINGS[code] || SYMBOL[code] || '');
    else out += c;
  }
  return out;
}

/* --- VML ------------------------------------------------------------------ */

const SHAPES = new Set(['v:shape', 'v:rect', 'v:roundrect', 'v:oval', 'v:group', 'v:image']);
const BLOCKS = new Set(['P', 'DIV', 'LI', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'TD', 'TH', 'BODY']);

const childrenNamed = (el, name) => [...el.children].filter((c) => c.localName === name);

/** A shape's fill, if it has a visible one. White is the page and counts as none. */
function fillOf(el) {
  if (el.getAttribute('filled') === 'f') return null;
  const child = childrenNamed(el, 'v:fill')[0];
  const raw = (child && child.getAttribute('color')) || el.getAttribute('fillcolor');
  const hex = toHex(raw);
  return hex && hex !== '#ffffff' ? hex : null;
}

/**
 * The floating things in one VML comment, flattened: a group becomes its
 * pictures and boxes, each placed in the coordinates of the page's text column.
 */
function shapeItems(el, frame, out) {
  const st = declarations(el.getAttribute('style'));
  const name = el.localName;

  let x;
  let y;
  let w;
  let h;
  if (frame.group) {
    const g = frame.group;
    x = g.x + ((toPt(st.get('left')) || 0) - g.ox) * g.sx;
    y = g.y + ((toPt(st.get('top')) || 0) - g.oy) * g.sy;
    w = (toPt(st.get('width')) || 0) * g.sx;
    h = (toPt(st.get('height')) || 0) * g.sy;
  } else {
    x = toPt(st.get('margin-left')) || toPt(st.get('left')) || 0;
    y = toPt(st.get('margin-top')) || toPt(st.get('top')) || 0;
    w = toPt(st.get('width')) || 0;
    h = toPt(st.get('height')) || 0;
  }

  const base = frame.group
    ? frame.base
    : {
        floating: st.get('position') === 'absolute',
        hRel: (st.get('mso-position-horizontal-relative') || 'text').toLowerCase(),
        vRel: (st.get('mso-position-vertical-relative') || 'text').toLowerCase(),
      };

  if (name === 'v:group') {
    const [cw, ch] = (el.getAttribute('coordsize') || '1000,1000').split(',').map(Number);
    const [ox, oy] = (el.getAttribute('coordorigin') || '0,0').split(',').map(Number);
    const group = { x, y, ox: ox || 0, oy: oy || 0, sx: cw ? w / cw : 1, sy: ch ? h / ch : 1 };
    for (const child of el.children) {
      if (SHAPES.has(child.localName)) shapeItems(child, { group, base }, out);
    }
    return;
  }

  const textbox = childrenNamed(el, 'v:textbox')[0];
  const image = childrenNamed(el, 'v:imagedata')[0] || (name === 'v:image' ? el : null);
  if (textbox) {
    // Word wraps the box's content in a layout table for other browsers, and
    // the real content in a <div>.
    const holder = textbox.querySelector('div') || textbox;
    out.push({ type: 'text', x, y, w, h, fill: fillOf(el), nodes: [...holder.childNodes], ...base });
  } else if (image) {
    const alt = el.getAttribute('alt') || image.getAttribute('o:title') || '';
    out.push({ type: 'pic', x, y, w, h, alt: alt.split(/\n/)[0].trim(), ...base });
  }
}

/** Vertical extents overlap enough to call two things one row. */
const overlaps = (a, b) => Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y) > Math.min(a.h, b.h) * 0.3;

/**
 * Floating items in reading order.
 *
 * Items positioned against the same thing (the paragraph, or the page) are
 * sorted top to bottom; items positioned against different things keep the
 * order Word wrote them in, because there is no common ruler to compare them
 * with. Items side by side are one row: a picture beside a box of text is a
 * picture the text wraps around.
 */
function readingOrder(items, contentWidthPt) {
  const frameOf = (it) => (it.vRel === 'margin' || it.vRel === 'page' ? 'page' : 'text');
  const order = items.map((it, i) => ({ it, i }));
  for (const frame of ['text', 'page']) {
    const slots = order.map((o, k) => (frameOf(o.it) === frame ? k : -1)).filter((k) => k >= 0);
    const sorted = slots.map((k) => order[k]).sort((a, b) => a.it.y - b.it.y || a.i - b.i);
    slots.forEach((k, n) => (order[k] = sorted[n]));
  }

  const rows = [];
  for (const { it } of order) {
    const row = rows[rows.length - 1];
    if (row && frameOf(row[0]) === frameOf(it) && row.some((o) => overlaps(o, it))) row.push(it);
    else rows.push([it]);
  }

  const out = [];
  for (const row of rows) {
    const side = (it) => {
      const centre = (it.hRel === 'page' ? it.x - 0 : it.x) + it.w / 2;
      const r = centre / contentWidthPt;
      return r < 0.42 ? 'left' : r > 0.58 ? 'right' : 'center';
    };
    const shared = row.length > 1;
    for (const it of row) {
      it.align = side(it);
      // A picture with something beside it floats, so that thing can sit
      // next to it; one on its own line is a block.
      it.wrap = shared && it.type === 'pic';
      if (it.wrap && it.align === 'center') it.align = it.x + it.w / 2 < contentWidthPt / 2 ? 'left' : 'right';
    }
    // Floats first, right before left, so both are in place before the text
    // that runs between them.
    const rank = (it) => (it.type !== 'pic' || !it.wrap ? 2 : it.align === 'right' ? 0 : 1);
    out.push(...row.slice().sort((a, b) => rank(a) - rank(b) || a.x - b.x));
  }
  return out;
}

/**
 * Turn every VML comment into real elements, and drop Word's fallbacks.
 *
 * @param {Document} doc
 * @param {{contentWidthPt:number, marginLeftPt:number}} styles
 * @returns {{pictures: {alt:string, widthPt:number, heightPt:number}[]}} the
 *   pictures found, in document order; each is placed as a
 *   `<figure data-op-pic="N">` in the document
 */
export function expandVml(doc, styles) {
  const pictures = [];
  const comments = [];
  const walker = doc.createTreeWalker(doc.body, 128 /* comments */);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) comments.push(n);

  // Anchor paragraph -> the floating items it carries, in source order.
  const anchors = new Map();

  for (const comment of comments) {
    const data = comment.data || '';
    if (!/^\[if gte vml 1\]>/i.test(data)) continue;

    const holder = doc.createElement('div');
    holder.innerHTML = data.replace(/^\[if gte vml 1\]>/i, '').replace(/<!\[endif\]\s*$/i, '');
    const items = [];
    for (const el of holder.children) if (SHAPES.has(el.localName)) shapeItems(el, {}, items);

    for (const it of items) {
      if (it.type === 'pic') {
        it.index = pictures.length;
        pictures.push({ alt: it.alt, widthPt: it.w, heightPt: it.h });
      }
    }

    // A picture that is part of the line stays where it is in the line.
    const inline = items.filter((it) => !it.floating);
    for (const it of inline) comment.parentNode.insertBefore(element(doc, it, styles), comment);

    const floating = items.filter((it) => it.floating);
    if (floating.length) {
      let anchor = comment.parentNode;
      while (anchor && !BLOCKS.has(anchor.tagName)) anchor = anchor.parentNode;
      if (!anchors.has(anchor)) anchors.set(anchor, []);
      anchors.get(anchor).push(...floating);
    }
    comment.remove();
  }

  for (const [anchor, items] of anchors) {
    const ordered = readingOrder(items, styles.contentWidthPt);
    // Above the paragraph's own line only what sits above it; the rest follows.
    const before = ordered.filter((it) => it.vRel !== 'margin' && it.vRel !== 'page' && it.y + it.h <= 0);
    const after = ordered.filter((it) => !before.includes(it));
    if (!anchor || anchor.tagName === 'BODY' || anchor.tagName === 'TD' || anchor.tagName === 'TH') {
      const host = anchor || doc.body;
      for (const it of [...before, ...after]) host.appendChild(element(doc, it, styles));
      continue;
    }
    for (const it of before) anchor.parentNode.insertBefore(element(doc, it, styles), anchor);
    let at = anchor.nextSibling;
    for (const it of after) anchor.parentNode.insertBefore(element(doc, it, styles), at);
  }

  return { pictures };
}

/** The element an item becomes: a box of content, or a placed picture. */
function element(doc, it, styles) {
  if (it.type === 'text') {
    const box = doc.createElement('div');
    box.setAttribute('data-op-box', '1');
    if (it.fill) box.setAttribute('data-fill', it.fill);
    for (const n of it.nodes) box.appendChild(n);
    return box;
  }
  const fig = doc.createElement('figure');
  fig.setAttribute('data-op-pic', String(it.index));
  fig.setAttribute('data-align', it.align || 'center');
  fig.setAttribute('data-wrap', it.wrap ? '1' : '0');
  fig.setAttribute('data-width', String(Math.round(Math.min(100, (it.w / styles.contentWidthPt) * 100) * 10) / 10));
  return fig;
}

/**
 * Remove the content of Word's other conditional sections: `[if !vml]` (a
 * picture of each shape, at an unreadable path) and `[if !supportLists]` (a
 * list marker typed out as text). The marker is not thrown away — it is the
 * best record of what the list looked like — but kept on its paragraph.
 *
 * @param {Document} doc
 */
export function stripFallbacks(doc) {
  const comments = [];
  const walker = doc.createTreeWalker(doc.body, 128);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) comments.push(n);

  for (const c of comments) {
    const kind = (c.data || '').trim().toLowerCase();
    if (kind !== '[if !vml]' && kind !== '[if !supportlists]') continue;
    if (!c.parentNode) continue;
    const removed = [];
    for (let n = c.nextSibling; n; ) {
      const next = n.nextSibling;
      if (n.nodeType === 8 && /^\[endif\]$/i.test((n.data || '').trim())) {
        n.remove();
        break;
      }
      removed.push(n);
      n.remove();
      n = next;
    }
    if (kind === '[if !supportlists]') {
      let para = c.parentNode;
      while (para && !BLOCKS.has(para.tagName)) para = para.parentNode;
      const marker = readMarker(removed);
      if (para && marker) {
        para.setAttribute('data-op-marker', marker.text);
        if (marker.font) para.setAttribute('data-op-marker-font', marker.font);
      }
    }
    c.remove();
  }
}

/** The marker text in a removed `[if !supportLists]` section, and its font. */
function readMarker(nodes) {
  for (const n of nodes) {
    if (n.nodeType !== 1) continue;
    const ignore = n.matches('[style*="mso-list"]') ? n : n.querySelector('[style*="mso-list"]');
    if (!ignore) continue;
    let text = '';
    for (const k of ignore.childNodes) if (k.nodeType === 3) text += k.nodeValue;
    text = text.replace(/[\s ]+/g, '');
    if (!text) continue;
    let font = null;
    for (let el = ignore; el && el !== n.parentNode; el = el.parentElement) {
      const f = declarations(el.getAttribute('style')).get('font-family');
      if (f) {
        font = f;
        break;
      }
    }
    return { text, font };
  }
  return null;
}
