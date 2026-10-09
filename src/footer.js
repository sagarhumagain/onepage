/**
 * footer.js — the two pictures pinned to the foot of the page.
 *
 * A fact sheet carries the same logos at its foot whatever the text above
 * says, so they are page furniture, like the background colour: not lines in
 * the source, not blocks the text flows around, and not something a paste or
 * an edit can move. Two slots, left and right; each holds a picture or nothing.
 *
 * The footer takes a fixed strip off the bottom of the text area, and only
 * when it has a picture in it — with both slots empty the page is exactly the
 * page it was before there was a footer.
 */

/** Height of the footer strip, and the gap above it, in mm. Mirrored in document.css. */
export const FOOTER_MM = 20;
export const FOOTER_GAP_MM = 3;

export const SLOTS = ['left', 'right'];

/** @typedef {{src:string, alt?:string, w?:number, h?:number, preset?:boolean}} FooterImage */
/** @typedef {{left: FooterImage|null, right: FooterImage|null}} Footer */

export const hasFooter = (footer) => Boolean(footer && (footer.left || footer.right));

/** Only a picture the app read in itself — a data URL — is drawn. */
const isImage = (img) => Boolean(img && /^data:image\/[a-z0-9.+-]+;base64,/i.test(String(img.src || '')));

const escapeAttr = (s) =>
  String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);

/**
 * The footer's content: both slots, always, so the right-hand picture stays on
 * the right when the left one is gone. An empty slot is an empty box here; the
 * editor's own stylesheet is what turns it into a "+ Image" target, so no
 * placeholder can reach a printed page.
 *
 * @param {Footer} footer
 */
export function footerHtml(footer) {
  return SLOTS.map((slot) => {
    const img = footer && footer[slot];
    if (!isImage(img)) return `<figure class="footer-slot" data-slot="${slot}" data-empty="1"></figure>`;
    return `<figure class="footer-slot" data-slot="${slot}"><img src="${escapeAttr(img.src)}" alt="${escapeAttr(img.alt)}"></figure>`;
  }).join('');
}

/**
 * The footer as it is saved: a bundled default is stored as the word
 * "default" rather than as a data URL, which keeps the saved state small and
 * lets a newer default replace an older one.
 */
export function serializeFooter(footer) {
  const out = {};
  for (const slot of SLOTS) {
    const img = footer && footer[slot];
    out[slot] = !isImage(img) ? null : img.preset ? 'default' : { src: img.src, alt: img.alt || '', w: img.w || 0, h: img.h || 0 };
  }
  return out;
}

/**
 * A saved footer back into one the page can draw. No saved footer at all —
 * a first run, or state saved before footers existed — means the defaults; a
 * slot saved as empty stays empty.
 *
 * @param {unknown} saved
 * @param {Footer} defaults
 * @returns {Footer}
 */
export function restoreFooter(saved, defaults) {
  if (!saved || typeof saved !== 'object') return { ...defaults };
  const out = {};
  for (const slot of SLOTS) {
    const v = saved[slot];
    out[slot] = v === 'default' ? defaults[slot] || null : isImage(v) ? { src: v.src, alt: v.alt || '', w: v.w || 0, h: v.h || 0 } : null;
  }
  return out;
}
