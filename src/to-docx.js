/**
 * to-docx.js — document model -> .docx, at the fitted size.
 *
 * Word does not have a "fit to one page" feature and OOXML has no element
 * that means it: Word re-runs its own line breaking and pagination. So the
 * job here is to remove every source of disagreement between the preview and
 * Word's layout engine:
 *
 *   - exact page geometry (A4 as the literal twips Word itself writes)
 *   - `lineRule: EXACTLY`, so a line box is an absolute height and cannot
 *     grow for a tall glyph or a substituted font
 *   - widow/orphan control off, so Word will not push two lines to page 2
 *   - fixed table layout, so tables cannot re-autofit and change row heights
 *   - a font that is present on the machine, named identically to the one the
 *     preview measured
 *
 * Even then the match is not exact, so the export is taken a few percent
 * below the fitted size. See WORD_REFLOW_SLACK.
 */

import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  HorizontalPositionAlign,
  HorizontalPositionRelativeFrom,
  ImageRun,
  LevelFormat,
  LineRuleType,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  TextWrappingSide,
  TextWrappingType,
  VerticalAlignTable,
  VerticalPositionRelativeFrom,
  WidthType,
} from 'docx';

import { plainText, tokenizeMarked } from './inline.js';
import { LIST_GLYPHS } from './parse.js';
import { clampWidth } from './render.js';
import { A4_TWIP, COLUMN_GAP_MM, SCALE, mmToTwip, ptToHalfPoint, ptToTwip } from './scale.js';

/**
 * Word's line breaking never matches Chromium's exactly — kerning,
 * hyphenation and justification all differ slightly. Exporting a few percent
 * below the fitted size absorbs that residual error, which is the difference
 * between "usually one page" and "reliably one page".
 */
const WORD_REFLOW_SLACK = 0.97;

/** 1 CSS px = 1/96 in = 15 twips. 1 pt = 12700 EMU. */
const TWIPS_PER_PX = 15;
const EMU_PER_PT = 12700;

const RULE_COLOR = 'C8CCD4';
const MUTED = '565C66';

/** OOXML has no leading '#', and Word is fussy about the case. */
const fillOf = (color) => String(color || '').replace('#', '').toUpperCase();

/**
 * A run's formatting, in OOXML terms.
 *
 * A background becomes run shading rather than Word's `highlight` attribute:
 * `highlight` is a fixed palette of sixteen named colours, so any colour the
 * user actually picked would be snapped to the nearest marker pen. A size is a
 * multiplier of whatever the fitter chose for that element, exactly as the em
 * value in the preview is, so the two stay in step.
 */
function runStyle(fmt, base) {
  const f = fmt || {};
  return {
    // `false` is a choice too: a heading set in the regular weight.
    bold: f.bold === false ? false : base.bold || f.bold === true,
    italics: f.italic === false ? false : base.italic || f.italic === true,
    underline: f.underline ? {} : undefined,
    size: f.size > 0 ? Math.max(2, Math.floor(base.size * f.size)) : base.size,
    color: f.fg ? fillOf(f.fg) : base.color,
    shading: f.bg ? { type: ShadingType.CLEAR, color: 'auto', fill: fillOf(f.bg) } : undefined,
  };
}

const ALIGN = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
};
const VALIGN = { top: VerticalAlignTable.TOP, middle: VerticalAlignTable.CENTER, bottom: VerticalAlignTable.BOTTOM };

/** A whole-unit property from a unit's marks: alignment, vertical alignment, band. */
const unitProp = (marks, key) => {
  const hit = (marks || []).find((m) => m.fmt && m.fmt[key]);
  return hit ? hit.fmt[key] : undefined;
};

/** A paragraph's alignment: its own, if it has one, otherwise the element's. */
const alignmentOf = (marks, fallback) => ALIGN[unitProp(marks, 'align')] || fallback;

/**
 * A size every character of the unit shares, as render.js draws it on the
 * element. Word has no element to put it on, so it reaches a list's number
 * through the list's level instead.
 */
function uniformSize(marks, text) {
  const length = plainText(text).length;
  if (!marks || !marks.length || !length) return null;
  const size = Number(marks[0].fmt && marks[0].fmt.size);
  if (!(size > 0.2 && size < 5) || size === 1) return null;
  let at = 0;
  for (const m of marks) {
    if (m.start !== at || Number(m.fmt && m.fmt.size) !== size) return null;
    at = m.end;
  }
  return at >= length ? size : null;
}

/** The same marks without their size, once the paragraph carries it. */
const withoutSize = (marks) =>
  marks.map((m) => {
    const fmt = { ...m.fmt };
    delete fmt.size;
    return { ...m, fmt };
  });

/**
 * `lineRule: EXACTLY` is what stops Word growing a line box and paginating
 * differently from the preview — but an exact line box also *clips* anything
 * taller than itself, and an enlarged run is exactly that. Those paragraphs,
 * and only those, ask for at least the same height instead, which is what the
 * preview's line box does too.
 */
const lineRuleFor = (marks) =>
  (marks || []).some((m) => m.fmt && m.fmt.size > 1) ? LineRuleType.AT_LEAST : LineRuleType.EXACTLY;

/**
 * A band is paragraph shading, not run shading: in Word as in the preview,
 * that is the difference between a colour that fills the line and one that
 * stops at the last letter.
 */
const bandShading = (marks) => {
  const hit = (marks || []).find((m) => m.fmt && m.fmt.band);
  return hit ? { type: ShadingType.CLEAR, color: 'auto', fill: fillOf(hit.fmt.band) } : undefined;
};

/**
 * Runs for a token. A hard line break is a newline in the text, and in Word a
 * break before the next run — never a new paragraph, which would add the
 * paragraph's spacing the preview does not have.
 */
const linesOf = (text, props) =>
  String(text)
    .split('\n')
    .map((line, k) => new TextRun({ ...props, text: line, break: k ? 1 : undefined }));

const runsFor = (text, base, marks) =>
  tokenizeMarked(text, marks).flatMap((t) => {
    const style = runStyle(t.fmt, base);
    const props = {
      ...style,
      bold: Boolean(t.bold) || style.bold,
      italics: Boolean(t.italic) || style.italics,
      font: t.code ? 'Courier New' : base.font,
    };
    if (!t.href) return linesOf(t.text, props);
    return [
      new ExternalHyperlink({
        link: t.href,
        children: linesOf(t.text, { ...props, font: base.font, underline: {} }),
      }),
    ];
  });

/**
 * Turn a scale entry into the concrete OOXML numbers for a given base size.
 * The ratios come from scale.js, which document.css also follows.
 */
function metrics(key, basePt, docLineHeight) {
  const s = SCALE[key];
  const sizePt = basePt * s.size;
  const lh = s.lineHeight == null ? docLineHeight : s.lineHeight;
  return {
    sizePt,
    halfPoints: ptToHalfPoint(sizePt),
    line: ptToTwip(sizePt * lh),
    before: ptToTwip(sizePt * (s.before || 0)),
    after: ptToTwip(sizePt * (s.after || 0)),
    bold: Boolean(s.bold),
    italic: Boolean(s.italic),
    mono: Boolean(s.mono),
  };
}

/**
 * @param {object} opts
 * @param {import('./parse.js').Block[]} opts.blocks
 * @param {{fontPt:number, columns:number, lineHeight:number, marginMm:number,
 *   imageScale?:number}} opts.layout
 * @param {{docx:string, label:string}} opts.face
 * @param {string} [opts.title]
 * @returns {Promise<Blob>}
 */
export async function buildDocx({ blocks, layout, face, title }) {
  const basePt = layout.fontPt * WORD_REFLOW_SLACK;
  const docLineHeight = layout.lineHeight;
  const imageScale = layout.imageScale == null ? 1 : layout.imageScale;
  const marginTwip = mmToTwip(layout.marginMm);
  const columns = Math.max(1, Math.min(3, layout.columns || 1));
  const gapTwip = mmToTwip(COLUMN_GAP_MM);

  const usableWidth = A4_TWIP.width - marginTwip * 2;
  const columnWidth = Math.floor((usableWidth - gapTwip * (columns - 1)) / columns);

    const fontOf = (mm) => (mm.mono ? 'Courier New' : face.docx);

  const mOf = (key, scale = 1) => metrics(key, basePt * scale, docLineHeight);

  /**
   * One paragraph. `scale` is the size it sits at relative to the document —
   * a table cell's 0.95. A list item whose every word shares one size has that
   * folded in too, as render.js puts it on the item: its number, line height
   * and spacing follow the size, exactly as `em` margins and a unitless
   * line-height do.
   */
  const para = (key, text, extra = {}) => {
    const uniform = key === 'li' ? uniformSize(extra.marks, text) : null;
    const marks = uniform ? withoutSize(extra.marks) : extra.marks;
    const mm = mOf(key, (extra.scale || 1) * (uniform || 1));
    return new Paragraph({
      children: runsFor(
        text,
        {
          bold: extra.bold == null ? mm.bold : extra.bold,
          italic: mm.italic,
          font: fontOf(mm),
          size: mm.halfPoints,
          color: extra.color,
        },
        marks
      ),
      spacing: {
        before: extra.before == null ? mm.before : ptToTwip(mm.sizePt * extra.before),
        after: extra.after == null ? mm.after : ptToTwip(mm.sizePt * extra.after),
        line: mm.line,
        lineRule: lineRuleFor(marks),
      },
      shading: extra.band === false ? undefined : bandShading(extra.marks),
      // Word's default is to push a short tail to the next page rather than
      // split it — a guaranteed second sheet even when the text geometrically
      // fits. Off, always, for one-page output.
      widowControl: false,
      keepNext: extra.keepNext,
      alignment: alignmentOf(extra.marks, extra.alignment),
      indent: extra.indent,
      border: extra.border,
      numbering: extra.numbering,
    });
  };

  const listIndent = (level = 0, scale = 1) => {
    const mm = mOf('li', scale);
    return { left: ptToTwip(mm.sizePt * 1.35 * (level + 1)), hanging: ptToTwip(mm.sizePt * 0.95) };
  };

  /** Highlights resolved onto this block by marks.js, addressed as it does. */
  const marksOf = (b, sub) => (b._marks ? b._marks[sub] : undefined);

  /**
   * An image, at the size the preview fitted it to.
   *
   * `transformation` is in CSS pixels, not twips — the two were confused here
   * once, and a column-wide picture went into Word a hundred inches across.
   * The aspect ratio comes from the file itself; the height cap mirrors the
   * one in document.css so the two layouts agree about a tall photograph.
   */
  const imageRun = (b) => {
    const src = String(b.src || '');
    if (!src.startsWith('data:image/')) return null;

    let data;
    try {
      data = Uint8Array.from(atob(src.slice(src.indexOf(',') + 1)), (c) => c.charCodeAt(0));
    } catch {
      return null; // an image we cannot decode is skipped, never guessed at
    }

    const px = (twips) => twips / TWIPS_PER_PX;
    const ratio = b.w > 0 && b.h > 0 ? b.h / b.w : 0.62;
    let width = px(columnWidth) * (clampWidth(b.width) / 100) * imageScale;
    let height = width * ratio;
    const maxHeight = px(A4_TWIP.height - marginTwip * 2) * 0.62 * imageScale;
    if (height > maxHeight) {
      width *= maxHeight / height;
      height = maxHeight;
    }

    const transformation = { width: Math.max(1, Math.round(width)), height: Math.max(1, Math.round(height)) };
    const altText = { title: b.alt || 'Image', description: b.alt || 'Image', name: b.alt || 'image' };

    // Wrapped: Word anchors the picture to this (empty, zero-height) paragraph
    // and runs the following text around it, which is as close as OOXML gets
    // to the float the preview shows.
    const gap = Math.round(basePt * 0.6 * EMU_PER_PT);
    const floating =
      b.wrap && b.align !== 'center'
        ? {
            horizontalPosition: {
              relative: HorizontalPositionRelativeFrom.COLUMN,
              align: b.align === 'right' ? HorizontalPositionAlign.RIGHT : HorizontalPositionAlign.LEFT,
            },
            verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: 0 },
            wrap: { type: TextWrappingType.SQUARE, side: TextWrappingSide.BOTH_SIDES },
            margins: { top: 0, bottom: gap, left: b.align === 'right' ? gap : 0, right: b.align === 'right' ? 0 : gap },
            allowOverlap: false,
            layoutInCell: true,
          }
        : undefined;

    return {
      run: new ImageRun({ data, transformation, altText, floating }),
      floating: Boolean(floating),
      heightPt: (transformation.height * 72) / 96,
      alignment:
        b.align === 'left'
          ? AlignmentType.LEFT
          : b.align === 'right'
            ? AlignmentType.RIGHT
            : AlignmentType.CENTER,
    };
  };

  const numberingConfigs = [];

  /**
   * A list, with one numbering definition of its own so it restarts where the
   * preview restarts it. Each level takes its marker — a number, a letter, a
   * bullet or a drawn glyph — and its size from the first item at that depth.
   */
  const list = (b, scope) => {
    const reference = `list-${numberingConfigs.length}`;
    const levelOf = (j) => (b.levels ? b.levels[j] : 0);
    const kindOf = (j) => (b.kinds ? b.kinds[j] : b.type);
    const depth = Math.max(...b.items.map((_, j) => levelOf(j)));
    const levels = [];
    for (let level = 0; level <= depth; level++) {
      const j = b.items.findIndex((_, k) => levelOf(k) === level);
      const uniform = j >= 0 ? uniformSize(marksOf(b, `items.${j}`), b.items[j]) : null;
      const mm = mOf('li', scope.scale * (uniform || 1));
      const run = { size: mm.halfPoints, font: face.docx };
      const indent = listIndent(level, scope.scale);
      if (j >= 0 && kindOf(j) === 'ol') {
        const style = b.styles ? b.styles[j] : null;
        const start = b.nums && b.nums[j] ? b.nums[j] : level === 0 && b.start > 0 ? b.start : 1;
        levels.push({
          level,
          format: style === 'lower-alpha' ? LevelFormat.LOWER_LETTER : style === 'upper-alpha' ? LevelFormat.UPPER_LETTER : LevelFormat.DECIMAL,
          text: `%${level + 1}.`,
          alignment: AlignmentType.LEFT,
          start,
          style: { paragraph: { indent }, run },
        });
      } else {
        const glyph = j >= 0 && b.glyphs && LIST_GLYPHS.includes(b.glyphs[j] || '') ? b.glyphs[j] : '•';
        levels.push({
          level,
          format: LevelFormat.BULLET,
          text: glyph,
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent }, run },
        });
      }
    }
    numberingConfigs.push({ reference, levels });
    return b.items.map((item, j) =>
      para('li', item, {
        scale: scope.scale,
        numbering: { reference, level: levelOf(j) },
        marks: marksOf(b, `items.${j}`),
        band: scope.shaded ? false : undefined,
      })
    );
  };

  /** The band and vertical alignment every unit in a cell's content agrees on. */
  const sharedCell = (blocks) => {
    const all = [];
    for (const b of blocks) {
      const subs = b.type === 'ul' || b.type === 'ol' ? b.items.map((_, j) => `items.${j}`) : ['text'];
      for (const sub of subs) all.push(marksOf(b, sub));
    }
    const agree = (key) => {
      const first = unitProp(all[0], key);
      return first && all.every((mk) => unitProp(mk, key) === first) ? first : undefined;
    };
    return { band: all.length ? agree('band') : undefined, valign: all.length ? agree('valign') : undefined };
  };

  const table = (b, scope) => {
    const mm = mOf('table');
    const tableScale = SCALE.table.size;
    const colCount = Math.max(1, (b.head && b.head.length) || (b.rows[0] || []).length);
    const total = b.widths && b.widths.length === colCount ? b.widths.reduce((a, w) => a + w, 0) : 0;
    const widths = Array.from({ length: colCount }, (_, j) =>
      Math.floor(total ? (scope.width * b.widths[j]) / total : scope.width / colCount)
    );
    const ruleColor = b.border ? fillOf(b.border) : RULE_COLOR;
    const ruleSize = b.border ? 6 : 4;
    const borders = {
      top: { style: BorderStyle.SINGLE, size: ruleSize, color: ruleColor },
      bottom: { style: BorderStyle.SINGLE, size: ruleSize, color: ruleColor },
      left: { style: BorderStyle.SINGLE, size: ruleSize, color: ruleColor },
      right: { style: BorderStyle.SINGLE, size: ruleSize, color: ruleColor },
    };

    const cell = (sub, text, header, j) => {
      const inner = b.cells && b.cells[sub];
      let children;
      let band;
      let valign;
      if (inner) {
        const shared = sharedCell(inner);
        band = shared.band;
        valign = shared.valign;
        children = blockChildren(inner, { scale: tableScale, width: widths[j], inCell: true, shaded: Boolean(band) });
      } else {
        const marks = marksOf(b, sub);
        band = unitProp(marks, 'band');
        valign = unitProp(marks, 'valign');
        children = [para('p', text, { scale: tableScale, before: 0, after: 0, bold: header, marks, band: false, alignment: AlignmentType.LEFT })];
      }
      return new TableCell({
        width: { size: widths[j], type: WidthType.DXA },
        borders,
        verticalAlign: VALIGN[valign],
        shading: band ? { type: ShadingType.CLEAR, color: 'auto', fill: fillOf(band) } : header ? { fill: 'F4F5F7' } : undefined,
        children: children.length ? children : [new Paragraph({ children: [] })],
      });
    };

    const rows = [];
    if (!b.headless && b.head && b.head.length) {
      rows.push(new TableRow({ tableHeader: true, children: b.head.map((c, j) => cell(`head.${j}`, c, true, j)) }));
    }
    (b.rows || []).forEach((r, ri) => {
      const cells = r.slice(0, colCount);
      while (cells.length < colCount) cells.push('');
      rows.push(new TableRow({ children: cells.map((c, j) => cell(`rows.${ri}.${j}`, c, false, j)) }));
    });
    if (!rows.length) return [];
    return [
      new Table({
        rows,
        width: { size: scope.width, type: WidthType.DXA },
        // Fixed layout stops Word re-autofitting the columns, which
        // would change row heights and the total page height with them.
        layout: TableLayoutType.FIXED,
        columnWidths: widths,
      }),
      new Paragraph({
        children: [],
        spacing: { before: 0, after: ptToTwip(mm.sizePt * 0.4), line: 1, lineRule: LineRuleType.EXACTLY },
      }),
    ];
  };

  /**
   * Blocks -> Word's paragraphs and tables. A cell that holds paragraphs and
   * lists is the same thing at the table's size, with a cell's spacing: its
   * paragraphs a little closer, nothing after the last.
   *
   * @param {import('./parse.js').Block[]} list
   * @param {{scale:number, width:number, inCell?:boolean, shaded?:boolean}} scope
   */
  function blockChildren(list_, scope) {
    const out = [];
    list_.forEach((b, index) => {
      const last = scope.inCell && index === list_.length - 1;
      const band = scope.shaded ? false : undefined;
      switch (b.type) {
        case 'h1':
        case 'h2':
          out.push(para(b.type, b.text, { scale: scope.scale, keepNext: true, marks: marksOf(b, 'text'), band }));
          break;
        case 'h3':
          out.push(para('h3', b.text, { scale: scope.scale, keepNext: true, color: MUTED, marks: marksOf(b, 'text'), band }));
          break;
        case 'p':
          out.push(
            para('p', b.text, {
              scale: scope.scale,
              alignment: AlignmentType.JUSTIFIED,
              marks: marksOf(b, 'text'),
              after: scope.inCell ? (last ? 0 : 0.4) : undefined,
              band,
            })
          );
          break;
        case 'quote':
          out.push(
            para('quote', b.text, {
              scale: scope.scale,
              color: MUTED,
              marks: marksOf(b, 'text'),
              indent: { left: ptToTwip(mOf('quote', scope.scale).sizePt * 0.85) },
              border: { left: { style: BorderStyle.SINGLE, size: 12, space: 6, color: RULE_COLOR } },
              band,
            })
          );
          break;
        case 'hr':
          out.push(
            new Paragraph({
              children: [],
              spacing: { before: ptToTwip(basePt * 0.5), after: ptToTwip(basePt * 0.5), line: ptToTwip(basePt * 0.4), lineRule: LineRuleType.EXACTLY },
              border: { bottom: { style: BorderStyle.SINGLE, size: 6, space: 1, color: RULE_COLOR } },
            })
          );
          break;
        case 'code': {
          const mm = mOf('code', scope.scale);
          for (const line of b.text.split('\n')) {
            out.push(
              new Paragraph({
                children: [new TextRun({ text: line || ' ', font: 'Courier New', size: mm.halfPoints })],
                spacing: { before: 0, after: 0, line: mm.line, lineRule: LineRuleType.EXACTLY },
                widowControl: false,
              })
            );
          }
          break;
        }
        case 'ul':
        case 'ol':
          out.push(...list(b, scope));
          break;
        case 'table':
          out.push(...table(b, scope));
          break;
        case 'image': {
          const image = imageRun(b);
          if (!image) break;
          out.push(
            new Paragraph({
              children: [image.run],
              alignment: image.alignment,
              // An inline image is clipped to the line box, and the document
              // default is `lineRule: EXACTLY` at body-text height — which would
              // slice every picture down to one line. Images, and only images,
              // get a line box at least as tall as they are.
              spacing: image.floating
                ? { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACTLY }
                : {
                    before: ptToTwip(basePt * 0.5),
                    after: ptToTwip(basePt * 0.5),
                    line: ptToTwip(image.heightPt),
                    lineRule: LineRuleType.AT_LEAST,
                  },
              widowControl: false,
            })
          );
          break;
        }
        default:
          break;
      }
    });
    return out;
  }

  const children = blockChildren(blocks, { scale: 1, width: columnWidth });

  if (!children.length) children.push(new Paragraph({ children: [] }));

  const doc = new Document({
    title: title || 'Document',
    creator: 'OnePage',
    description: 'Fitted to a single A4 page',
    // Compatibility flags that remove Word's discretionary extra spacing —
    // each one is a place Word could otherwise add height we did not predict.
    compatibility: {
      doNotUseHTMLParagraphAutoSpacing: true,
      noExtraLineSpacing: true,
      doNotAutofitConstrainedTables: true,
      doNotBreakWrappedTables: true,
    },
    styles: {
      default: {
        document: {
          run: { font: face.docx, size: ptToHalfPoint(basePt) },
          paragraph: {
            spacing: { line: ptToTwip(basePt * docLineHeight), lineRule: LineRuleType.EXACTLY, after: 0 },
            widowControl: false,
          },
        },
      },
    },
    numbering: numberingConfigs.length ? { config: numberingConfigs } : undefined,
    sections: [
      {
        properties: {
          page: {
            // Literal twips, not a mm helper: docx 9.7.1's
            // convertMillimetersToTwip truncates and returns 11905 x 16837,
            // one twip short of the A4 that Word itself writes. On a
            // borderline layout that single twip is a second page.
            size: { width: A4_TWIP.width, height: A4_TWIP.height, orientation: 'portrait' },
            margin: {
              top: marginTwip,
              right: marginTwip,
              bottom: marginTwip,
              left: marginTwip,
              header: 0,
              footer: 0,
              gutter: 0,
            },
          },
          column: columns > 1 ? { count: columns, space: gapTwip, separate: false } : undefined,
        },
        children,
      },
    ],
  });

  return Packer.toBlob(doc);
}
