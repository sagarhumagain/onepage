/**
 * fit.js — the engine that forces arbitrary content onto exactly one A4 page.
 *
 * The result of fitting is a *style budget* — real numbers (base font size in
 * pt, line-height ratio, column count, margins) — not a visual transform.
 * That matters because `transform: scale()` and `zoom` cannot be represented
 * in OOXML at all, so a scaled preview could never be exported to Word. Every
 * size in document.css derives from --doc-font-pt, so one number rescales the
 * whole document and maps 1:1 onto half-points in DOCX.
 *
 * Content is never truncated. When the comfortable size overflows, the engine
 * climbs a ladder — more columns, tighter leading, narrower margins, and only
 * then below the readable floor — and reports which rung it had to use.
 */

const MM_PER_IN = 25.4;
const CSS_PX_PER_IN = 96;

export const A4 = {
  widthMm: 210,
  heightMm: 297,
  widthPx: (210 / MM_PER_IN) * CSS_PX_PER_IN, // 793.70
  heightPx: (297 / MM_PER_IN) * CSS_PX_PER_IN, // 1122.52
};

export const DEFAULTS = {
  marginMm: 15,
  /** Margins the ladder may reclaim before shrinking type below the floor. */
  tightMarginMm: 10,
  minMarginMm: 7,
  /** Smallest size used before reaching for columns or tighter leading. */
  comfortPt: 9,
  /** The readable floor. Crossed only to avoid losing content. */
  minPt: 7,
  /**
   * Text this small is barely legible in print, but the product promise is
   * that content is never silently cut, so the engine keeps shrinking and the
   * UI says loudly what it had to do.
   */
  emergencyPt: 2.5,
  /**
   * And below even that, rather than give up: one page is the whole product,
   * so "this does not fit" is not an outcome the engine is allowed to choose
   * while there is any size left to try. The status bar says how small it had
   * to go; the sheet is still one sheet.
   */
  absoluteMinPt: 0.4,
  /**
   * Growth caps. Above roughly 13-14pt an A4 page stops reading as a document
   * and starts reading as a large-print flyer, so slack is spent on leading
   * before it is spent on type size.
   */
  fillMaxPt: 13.5,
  naturalMaxPt: 11.5,
  lineHeight: 1.38,
  maxLineHeight: 1.55,
  tightLineHeight: 1.16,
  maxColumns: 3,
  columns: 'auto',
  fillPage: true,
  /**
   * Images do not shrink when type does, so on a page that is mostly picture
   * the type ladder alone ends at 3pt text beside a full-size photograph.
   * These are the multipliers the fitter is allowed to apply to every image,
   * tried in order and only once the readable type sizes are exhausted.
   */
  imageSteps: [0.82, 0.66, 0.5, 0.35],
  hasImages: false,
  /**
   * The height budget is shaved slightly. Print rasterisation rounds
   * differently from screen layout, and a single sub-pixel of overflow
   * becomes a second sheet of paper.
   */
  safety: 0.988,
  tolerancePx: 0.5,
  maxIterations: 16,
};

/**
 * A multi-column box must not be created for a single column.
 *
 * `column-count: 1` still makes the element a multi-column container, and a
 * multicol container with a definite height does not grow downwards when it
 * runs out of room — it fragments sideways into further columns. That makes
 * the vertical overflow probe silently useless. `auto` keeps a single-column
 * page an ordinary block box that grows down, which is what we measure.
 */
export const cssColumns = (n) => (Number(n) > 1 ? String(n) : 'auto');

/**
 * Measures how far content extends past its box, in CSS pixels.
 *
 * A zero-size sentinel appended after the content is the only reliable probe.
 * `scrollHeight` is an integer and silently ignores a collapsed bottom margin
 * on the last child; `scrollWidth` does not report spilled multicol columns
 * at all, reading equal to clientWidth whether the content fits in three
 * columns or needs five. Either flaw hides overflow that the print engine
 * turns into a second page.
 *
 * Both axes are always checked: one column overflows downwards, two or three
 * overflow to the right as extra column boxes.
 *
 * @returns {number} positive when overflowing
 */
function overflowBy(body, sentinel, safety, tol) {
  const box = body.getBoundingClientRect();
  const end = sentinel.getBoundingClientRect();

  // The measured box can carry padding (the inset background), and padding is
  // not room the text may use. Measure the content box, not the border box.
  const cs = body.ownerDocument.defaultView.getComputedStyle(body);
  const padTop = parseFloat(cs.paddingTop) || 0;
  const padLeft = parseFloat(cs.paddingLeft) || 0;
  const innerHeight = box.height - padTop - (parseFloat(cs.paddingBottom) || 0);
  const innerWidth = box.width - padLeft - (parseFloat(cs.paddingRight) || 0);

  const heightOver = end.bottom - (box.top + padTop) - innerHeight * safety;
  // The sentinel has no width, so its inline position is the position of the
  // column the content ended in.
  const widthOver = end.right - (box.left + padLeft) - innerWidth - tol;

  return Math.max(heightOver, widthOver);
}

/**
 * Largest value in [lo, hi] that does not overflow.
 *
 * `fits(pt)` is only *almost* monotone: `text-wrap: pretty`, `hyphens: auto`
 * and `break-inside: avoid` blocks mean a slightly smaller font can
 * occasionally produce a taller box by changing where lines wrap. The caller
 * verifies the winner and steps down if the bisection was fooled.
 */
function bisect(probe, lo, hi, maxIterations) {
  let iterations = 0;

  if (probe(hi) <= 0) return { value: hi, fits: true, iterations: ++iterations };
  iterations++;

  if (probe(lo) > 0) return { value: lo, fits: false, iterations: ++iterations };
  iterations++;

  let best = lo;
  while (hi - lo > 0.02 && iterations < maxIterations) {
    const mid = (lo + hi) / 2;
    iterations++;
    if (probe(mid) > 0) hi = mid;
    else {
      lo = mid;
      best = mid;
    }
  }
  return { value: best, fits: true, iterations };
}

/**
 * Ordered attempts, most readable first. Each rung is reached only because
 * every rung above it failed to hold the content.
 */
function buildLadder(cfg) {
  const top = cfg.fillPage ? cfg.fillMaxPt : cfg.naturalMaxPt;
  const allowed =
    cfg.columns === 'auto'
      ? [1, 2, 3].filter((c) => c <= cfg.maxColumns)
      : [Math.min(Math.max(Number(cfg.columns) || 1, 1), cfg.maxColumns)];

  const widest = allowed[allowed.length - 1];
  const margin = cfg.marginMm;
  const tighter = Math.min(margin, cfg.tightMarginMm);
  const tightest = Math.min(margin, cfg.minMarginMm);
  const rung = (cols, lineHeight, marginMm, lo, hi, emergency, imageScale) => ({
    cols,
    lineHeight,
    marginMm,
    lo,
    hi,
    emergency: Boolean(emergency),
    imageScale: imageScale == null ? 1 : imageScale,
  });

  const steps = cfg.hasImages ? cfg.imageSteps : [];
  const img = (i) => (steps.length ? steps[Math.min(i, steps.length - 1)] : 1);

  const ladder = allowed.map((cols) => rung(cols, cfg.lineHeight, margin, cfg.comfortPt, top));

  // With pictures on the page the first thing to give is a little of the
  // picture: a photograph at 82% still reads, 7pt body text is already a
  // squint. With no pictures every `img()` is 1 and this is the original
  // four-rung ladder, unchanged.
  if (steps.length) ladder.push(rung(widest, cfg.lineHeight, margin, cfg.comfortPt, top, false, img(0)));

  // Below the comfortable size: tighten the leading, then reclaim the
  // margins, and only then go under the readable floor.
  ladder.push(rung(widest, cfg.lineHeight, margin, cfg.minPt, cfg.comfortPt, false, img(0)));
  ladder.push(rung(widest, cfg.tightLineHeight, margin, cfg.minPt, cfg.comfortPt, false, img(1)));
  ladder.push(rung(widest, cfg.tightLineHeight, tighter, cfg.minPt, cfg.comfortPt, false, img(1)));
  ladder.push(rung(widest, cfg.tightLineHeight, tightest, cfg.emergencyPt, cfg.minPt, true, img(2)));
  ladder.push(rung(widest, cfg.tightLineHeight, tightest, cfg.absoluteMinPt, cfg.emergencyPt, true, img(3)));
  return ladder;
}

/**
 * @param {{page: HTMLElement, body: HTMLElement}} els
 *   `page` carries the CSS custom properties; `body` is the measured box.
 */
export function createFitter({ page, body }) {
  const doc = page.ownerDocument;

  // A permanent end-of-content marker, sized to nothing so it cannot itself
  // affect layout or the position it is being used to report.
  const sentinel = doc.createElement('div');
  sentinel.setAttribute('aria-hidden', 'true');
  // The page is editable, and nothing the user types may end up inside the
  // marker the fit is measured against.
  sentinel.setAttribute('contenteditable', 'false');
  // `clear: both` is load-bearing: a floated (wrapped) image can hang below
  // the last line of text, and an uncleared sentinel sits above it and reports
  // that everything fits while the picture is being cut off.
  sentinel.style.cssText =
    'display:block;clear:both;width:0;height:0;margin:0;padding:0;border:0;font-size:0;line-height:0';

  const applyState = (fontPt, columns, lineHeight, marginMm, imageScale) => {
    page.style.setProperty('--doc-font-pt', String(round(fontPt, 3)));
    page.style.setProperty('--doc-columns', cssColumns(columns));
    page.style.setProperty('--doc-line-height', String(lineHeight));
    if (marginMm != null) page.style.setProperty('--doc-margin-mm', String(marginMm));
    page.style.setProperty('--doc-image-scale', String(round(imageScale == null ? 1 : imageScale, 3)));
  };

  /**
   * @param {Partial<typeof DEFAULTS>} [options]
   * @returns {Promise<{fontPt:number, columns:number, lineHeight:number, marginMm:number,
   *   imageScale:number, overflow:boolean, belowFloor:boolean, marginsReduced:boolean,
   *   imagesShrunk:boolean, iterations:number}>}
   */
  async function fit(options = {}) {
    const fontScale = options.fontScale || 1.0;
    const cfg = { ...DEFAULTS, ...options };

    // Measuring before fonts settle yields a size that shifts once they load.
    // The page lives in the sheet iframe, so use *its* font registry.
    if (doc.fonts && doc.fonts.ready) await doc.fonts.ready;

    // Always re-append: replacing the content removed it.
    body.appendChild(sentinel);

    // One write then one read per probe, in that order, in the same task: the
    // read flushes pending layout, so the number can never be stale.
    const probeFor = (r) => (pt) => {
      applyState(pt * fontScale, r.cols, r.lineHeight, r.marginMm, r.imageScale);
      return overflowBy(body, sentinel, cfg.safety, cfg.tolerancePx);
    };

    const ladder = buildLadder(cfg);
    let iterations = 0;
    let last = null;

    for (const r of ladder) {
      const probe = probeFor(r);
      const search = bisect(probe, r.lo, r.hi, cfg.maxIterations);
      iterations += search.iterations;
      last = r;
      if (!search.fits) continue;

      // Verify the winner. If a non-monotone wrap fooled the bisection, walk
      // down in small steps rather than shipping a layout that overflows.
      let pt = search.value;
      let ok = probe(pt) <= 0;
      iterations++;
      for (let i = 0; i < 4 && !ok && pt > r.lo; i++) {
        pt = Math.max(r.lo, pt - 0.25);
        ok = probe(pt) <= 0;
        iterations++;
      }
      if (!ok) continue; // this rung cannot hold the content after all

      // Short content: spend the leftover room on leading, not on type size,
      // so a half-page memo still reads as a document rather than a poster.
      let lineHeight = r.lineHeight;
      if (cfg.fillPage && !r.emergency && r.hi === cfg.fillMaxPt && pt >= r.hi - 0.05) {
        const grown = bisect(
          (lh) => {
            applyState(pt * fontScale, r.cols, lh, r.marginMm, r.imageScale);
            return overflowBy(body, sentinel, cfg.safety, cfg.tolerancePx);
          },
          r.lineHeight,
          cfg.maxLineHeight,
          10
        );
        iterations += grown.iterations;
        lineHeight = grown.fits ? grown.value : r.lineHeight;
        applyState(pt * fontScale, r.cols, lineHeight, r.marginMm, r.imageScale);
      }

      return {
        fontPt: round(pt * fontScale, 2),
        columns: r.cols,
        lineHeight: round(lineHeight, 3),
        marginMm: r.marginMm,
        imageScale: r.imageScale,
        overflow: false,
        belowFloor: r.emergency,
        marginsReduced: r.marginMm < cfg.marginMm,
        imagesShrunk: r.imageScale < 1,
        iterations,
      };
    }

    // Nothing fits, even at the emergency size. Keep the smallest layout and
    // report it, so the UI can say so instead of silently clipping.
    applyState(last.lo * fontScale, last.cols, last.lineHeight, last.marginMm, last.imageScale);
    return {
      fontPt: round(last.lo * fontScale, 2),
      columns: last.cols,
      lineHeight: last.lineHeight,
      marginMm: last.marginMm,
      imageScale: last.imageScale,
      overflow: true,
      belowFloor: true,
      marginsReduced: last.marginMm < cfg.marginMm,
      imagesShrunk: last.imageScale < 1,
      iterations,
    };
  }

  return { fit, applyState };
}

export const round = (n, places) => {
  const f = Math.pow(10, places);
  return Math.round(n * f) / f;
};

/** Usable text area in mm — used by the DOCX exporter. */
export function textAreaMm(marginMm) {
  return { width: A4.widthMm - marginMm * 2, height: A4.heightMm - marginMm * 2 };
}
