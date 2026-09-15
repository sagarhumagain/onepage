/**
 * scale.js — the typographic scale, in one place.
 *
 * document.css renders the preview and the printed page; the DOCX exporter
 * has to reproduce the same proportions in OOXML twips. If the two drift, a
 * document that fits on screen stops fitting in Word. So the ratios live
 * here, and test/scale.test.js asserts that document.css still matches.
 *
 * `size` is a multiple of the document base size.
 * `lineHeight` is a multiple of that element's own size (null = inherit).
 * `before`/`after` are margins in em of the element's own size, matching CSS.
 */
export const SCALE = {
  h1: { size: 1.9, lineHeight: 1.12, before: 0, after: 0.5, bold: true },
  h2: { size: 1.24, lineHeight: 1.2, before: 1.15, after: 0.34, bold: true },
  h3: { size: 1.06, lineHeight: 1.25, before: 0.95, after: 0.28, bold: true },
  p: { size: 1, lineHeight: null, before: 0, after: 0.62 },
  li: { size: 1, lineHeight: null, before: 0, after: 0.24 },
  quote: { size: 1, lineHeight: null, before: 0.7, after: 0.7, italic: true },
  code: { size: 0.9, lineHeight: 1.32, before: 0.7, after: 0.7, mono: true },
  table: { size: 0.95, lineHeight: null, before: 0.7, after: 0.7 },
};

/** Gap between columns, mm. Mirrors `column-gap` in document.css. */
export const COLUMN_GAP_MM = 7;

/**
 * Typefaces. The preview stack must LEAD with the same family the DOCX names,
 * because Word re-runs line breaking with the real font metrics — if the
 * preview measured Helvetica and Word lays out Calibri, every line break
 * moves and a one-page fit can become two pages.
 *
 * `probe` is checked at startup with document.fonts.check(); when the family
 * is missing the app falls back to `fallback` for BOTH preview and export, so
 * the two always agree.
 */
export const TYPEFACES = {
  sans: {
    label: 'Calibri',
    docx: 'Calibri',
    css: `Calibri, Carlito, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif`,
    probe: 'Calibri',
    fallback: 'arial',
  },
  arial: {
    label: 'Arial',
    docx: 'Arial',
    css: `Arial, 'Helvetica Neue', Helvetica, sans-serif`,
    probe: 'Arial',
    fallback: null,
  },
  serif: {
    label: 'Georgia',
    docx: 'Georgia',
    css: `Georgia, 'Times New Roman', Times, serif`,
    probe: 'Georgia',
    fallback: 'times',
  },
  times: {
    label: 'Times New Roman',
    docx: 'Times New Roman',
    css: `'Times New Roman', Times, serif`,
    probe: 'Times New Roman',
    fallback: null,
  },
  narrow: {
    label: 'Arial Narrow',
    docx: 'Arial Narrow',
    css: `'Arial Narrow', 'Liberation Sans Narrow', Arial, sans-serif`,
    probe: 'Arial Narrow',
    fallback: 'arial',
  },
  mono: {
    label: 'Courier New',
    docx: 'Courier New',
    css: `'Courier New', Courier, monospace`,
    probe: 'Courier New',
    fallback: null,
  },
};

/** Resolve a typeface key to one whose font is actually installed. */
export function resolveTypeface(key, isAvailable) {
  let face = TYPEFACES[key] || TYPEFACES.sans;
  const seen = new Set();
  while (face.fallback && !seen.has(face.label) && !isAvailable(face.probe)) {
    seen.add(face.label);
    face = TYPEFACES[face.fallback];
  }
  return face;
}

/** 1pt = 20 twips. OOXML measures paragraph spacing and line height in twips. */
export const ptToTwip = (pt) => Math.round(pt * 20);

/**
 * Font sizes in OOXML are half-points, and this rounds DOWN on purpose.
 * Rounding to nearest would turn a solved 9.76pt into 20 half-points = 10.0pt
 * — 2.5% larger than the size that was measured to fit, which spends the
 * reflow slack the one-page guarantee depends on. Smaller is always safe.
 */
export const ptToHalfPoint = (pt) => Math.max(1, Math.floor(pt * 2));

/** A4 in twips. The literals Word itself writes — see note in docx.js. */
export const A4_TWIP = { width: 11906, height: 16838 };

export const mmToTwip = (mm) => Math.round((mm / 25.4) * 1440);
