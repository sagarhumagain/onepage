/**
 * inline.js — inline formatting, tokenised once and consumed twice.
 *
 * The preview renders these tokens as HTML and the Word exporter renders the
 * same tokens as OOXML runs. Sharing the tokeniser is what keeps a bold word
 * bold in both, and stops the two renderers from disagreeing about where a
 * span starts.
 */

/**
 * @typedef {{text:string, bold?:boolean, italic?:boolean, code?:boolean,
 *   href?:string, bg?:string, pad?:number}} Token
 */

import { HIGHLIGHT } from './markup.js';

const PATTERNS = [
  /*
   * A highlight wraps other formatting rather than competing with it, so it
   * re-enters the tokeniser for its contents and stamps the colour onto every
   * token that comes back. That is what keeps a bold word inside a highlight
   * both bold and highlighted.
   */
  {
    re: HIGHLIGHT,
    make: (m) => ({
      nested: tokenizeInline(m[3]).map((t) => ({ ...t, bg: m[1].toLowerCase(), pad: Number(m[2]) || 0 })),
    }),
  },
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
    const made = best.pattern.make(best.m);
    if (made.nested) out.push(...made.nested);
    else out.push(made);
    rest = rest.slice(best.m.index + best.m[0].length);
  }

  return out.filter((t) => t.text !== '');
}

/** Plain text of a tokenised string, markers removed. */
export const plainText = (text) =>
  tokenizeInline(text)
    .map((t) => t.text)
    .join('');
