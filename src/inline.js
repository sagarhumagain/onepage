/**
 * inline.js — inline formatting, tokenised once and consumed twice.
 *
 * The preview renders these tokens as HTML and the Word exporter renders the
 * same tokens as OOXML runs. Sharing the tokeniser is what keeps a bold word
 * bold in both, and stops the two renderers from disagreeing about where a
 * span starts.
 */

/** @typedef {{text:string, bold?:boolean, italic?:boolean, code?:boolean, href?:string, fmt?:object}} Token */

const PATTERNS = [
  { re: /`([^`\n]+)`/, make: (m) => ({ text: m[1], code: true }) },
  { re: /\*\*([^*\n]+)\*\*/, make: (m) => ({ text: m[1], bold: true }) },
  { re: /__([^_\n]+)__/, make: (m) => ({ text: m[1], bold: true }) },
  {
    re: /(?<=^|[\s(])\*([^*\n]+)\*(?=[\s).,;:!?]|$)/,
    make: (m) => ({ text: m[1], italic: true }),
  },
  {
    re: /(?<=^|[\s(])_([^_\n]+)_(?=[\s).,;:!?]|$)/,
    make: (m) => ({ text: m[1], italic: true }),
  },
  {
    // [text](url) — a link whose words are not its address, as Word's are.
    re: /\[([^[\]\n]+)\]\(((?:https?:\/\/|mailto:)[^\s()<>]+)\)/,
    make: (m) => ({ text: m[1], href: m[2] }),
  },
  {
    re: /\bhttps?:\/\/[^\s<>()]+/,
    make: (m) => ({ text: m[0], href: m[0] }),
  },
  {
    re: /\bwww\.[^\s<>()]+/,
    make: (m) => ({ text: m[0], href: `https://${m[0]}` }),
  },
  {
    re: /\b[\w.+-]+@[\w-]+\.[\w.-]+\b/,
    make: (m) => ({ text: m[0], href: `mailto:${m[0]}` }),
  },
];

/**
 * @param {string} text
 * @returns {Token[]}
 */
export function tokenizeInline(text) {
  /** @type {Token[]} */
  const out = [];
  let rest = String(text == null ? '' : text);

  while (rest) {
    let best = null;
    for (const pattern of PATTERNS) {
      const m = pattern.re.exec(rest);
      if (m && (best === null || m.index < best.m.index)) best = { pattern, m };
      if (best && best.m.index === 0) break; // cannot do better
    }

    if (!best) {
      out.push({ text: rest });
      break;
    }

    if (best.m.index > 0) out.push({ text: rest.slice(0, best.m.index) });
    out.push(best.pattern.make(best.m));
    rest = rest.slice(best.m.index + best.m[0].length);
  }

  return out.filter((t) => t.text !== '');
}

/** Plain text of a tokenised string, markers removed. */
export const plainText = (text) =>
  tokenizeInline(text)
    .map((t) => t.text)
    .join('');

/**
 * Tokenise, then cut the tokens at the boundaries of formatted runs.
 *
 * Mark ranges are offsets into the *plain* text — what the reader sees, with
 * the inline markers already removed — because that is what a selection in the
 * preview reports. Cutting here, once, is what keeps a half-highlighted bold
 * phrase identical in the preview, the PDF and the Word file.
 *
 * @param {string} text
 * @param {{start:number, end:number, fmt:object}[]} [marks]
 * @returns {Token[]} tokens, each carrying `fmt` when it falls inside a range
 */
export function tokenizeMarked(text, marks) {
  const tokens = tokenizeInline(text);
  if (!marks || !marks.length) return tokens;

  const out = [];
  let pos = 0;
  for (const token of tokens) {
    const start = pos;
    const end = pos + token.text.length;
    pos = end;

    const cuts = new Set([start, end]);
    for (const m of marks) {
      if (m.start > start && m.start < end) cuts.add(m.start);
      if (m.end > start && m.end < end) cuts.add(m.end);
    }

    const points = [...cuts].sort((a, b) => a - b);
    for (let i = 0; i < points.length - 1; i++) {
      const a = points[i];
      const b = points[i + 1];
      const hit = marks.find((m) => m.start <= a && m.end >= b);
      const piece = { ...token, text: token.text.slice(a - start, b - start) };
      if (hit) piece.fmt = hit.fmt;
      out.push(piece);
    }
  }
  return out.filter((t) => t.text !== '');
}
