/**
 * rtf.js — the pictures inside an RTF clipboard flavour.
 *
 * Word's HTML flavour does not carry its pictures. It points at temporary
 * files (`file:///.../msohtmlclip/clip_image004.png`) that a web page cannot
 * read and that, on macOS, sit inside Office's own sandboxed container. The
 * RTF flavour Word writes alongside it carries the same pictures inline, as
 * hex-encoded PNG or JPEG, in document order — so that is where they are read
 * from.
 *
 * Only PNG and JPEG are taken. Word also writes a WMF copy of each picture for
 * older readers (`\nonshppict`), which no browser can draw.
 */

/**
 * `widthPt`/`heightPt` is the size Word showed it at; `pxWidth`/`pxHeight`
 * the picture's own pixels.
 * @typedef {{mime:string, base64:string, widthPt:number, heightPt:number,
 *   pxWidth:number, pxHeight:number, uid:string}} RtfPicture
 */

/** Pixel size from a PNG's header or a JPEG's frame marker; zero if unreadable. */
function pixelSize(bytes, mime) {
  const u16 = (i) => (bytes[i] << 8) | bytes[i + 1];
  const u32 = (i) => ((bytes[i] << 24) | (bytes[i + 1] << 16) | (bytes[i + 2] << 8) | bytes[i + 3]) >>> 0;
  if (mime === 'image/png' && bytes.length > 24) return { w: u32(16), h: u32(20) };
  if (mime === 'image/jpeg') {
    for (let i = 2; i + 9 < bytes.length; ) {
      if (bytes[i] !== 0xff) break;
      const marker = bytes[i + 1];
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { w: u16(i + 7), h: u16(i + 5) };
      }
      i += 2 + u16(i + 2);
    }
  }
  return { w: 0, h: 0 };
}

/** Base64 of a run of bytes, in chunks so a large picture cannot overflow the call stack. */
function toBase64(bytes) {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

/** The index just past the group that opens at `open`, honouring escaped braces. */
function groupEnd(rtf, open) {
  let depth = 0;
  for (let i = open; i < rtf.length; i++) {
    const c = rtf[i];
    if (c === '\\') {
      i++; // an escaped brace or backslash is text, not structure
      continue;
    }
    if (c === '{') depth++;
    else if (c === '}' && --depth === 0) return i + 1;
  }
  return rtf.length;
}

/** Is `at` inside a `{\nonshppict ...}` group — the legacy copy of a picture? */
function inLegacyCopy(rtf, at, legacy) {
  return legacy.some(([from, to]) => at > from && at < to);
}

/**
 * @param {string} rtf
 * @returns {RtfPicture[]} pictures in document order, each one once
 */
export function extractRtfPictures(rtf) {
  if (!rtf || !rtf.includes('\\pict')) return [];

  const legacy = [];
  for (const m of rtf.matchAll(/\{\\nonshppict/g)) legacy.push([m.index, groupEnd(rtf, m.index)]);

  const out = [];
  const seen = new Set();
  for (const m of rtf.matchAll(/\{\\pict/g)) {
    const start = m.index;
    if (inLegacyCopy(rtf, start, legacy)) continue;
    const end = groupEnd(rtf, start);
    const group = rtf.slice(start, end);

    const mime = /\\pngblip/.test(group) ? 'image/png' : /\\jpegblip/.test(group) ? 'image/jpeg' : null;
    if (!mime) continue;

    const uid = (group.match(/\{\\\*\\blipuid\s+([0-9a-fA-F]+)\}/) || [])[1] || '';
    if (uid && seen.has(uid)) continue;

    // The picture data is the hex after the last control word, once nested
    // groups such as the blip uid are taken out. A removed group leaves a
    // space behind: `\bliptag255{...}89504e` would otherwise read as one
    // control word with the number 25589504.
    let body = group.slice(1, -1);
    body = body.replace(/\{[^{}]*\}/g, ' ');
    const hex = body.replace(/\\[a-zA-Z]+-?\d* ?/g, '').replace(/[^0-9a-fA-F]/g, '');
    if (hex.length < 16 || hex.length % 2) continue;

    const bytes = new Uint8Array(hex.length / 2);
    for (let i = 0; i < bytes.length; i++) bytes[i] = parseInt(hex.substr(i * 2, 2), 16);

    const num = (word) => {
      const n = group.match(new RegExp(`\\\\${word}(-?\\d+)`));
      return n ? Number(n[1]) : 0;
    };
    const scaleX = num('picscalex') || 100;
    const scaleY = num('picscaley') || 100;
    // \picwgoal is twips; 20 twips to the point.
    const widthPt = (num('picwgoal') * scaleX) / 100 / 20;
    const heightPt = (num('pichgoal') * scaleY) / 100 / 20;

    if (uid) seen.add(uid);
    const px = pixelSize(bytes, mime);
    out.push({ mime, base64: toBase64(bytes), widthPt, heightPt, pxWidth: px.w, pxHeight: px.h, uid });
  }
  return out;
}
