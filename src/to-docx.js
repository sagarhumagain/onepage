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
  VerticalPositionRelativeFrom,
  WidthType,
} from 'docx';

import { tokenizeMarked } from './inline.js';
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
    bold: base.bold || f.bold === true,
    italics: base.italic || f.italic === true,
    size: f.size > 0 ? Math.max(2, Math.floor(base.size * f.size)) : base.size,
    color: f.fg ? fillOf(f.fg) : base.color,
    shading: f.bg ? { type: ShadingType.CLEAR, color: 'auto', fill: fillOf(f.bg) } : undefined,
  };
}

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

const runsFor = (text, base, marks) =>
  tokenizeMarked(text, marks).flatMap((t) => {
    const style = runStyle(t.fmt, base);
    const run = new TextRun({
      text: t.text,
      ...style,
      bold: Boolean(t.bold) || style.bold,
      italics: Boolean(t.italic) || style.italics,
      font: t.code ? 'Courier New' : base.font,
    });
    if (!t.href) return [run];
    return [
      new ExternalHyperlink({
        link: t.href,
        children: [
          new TextRun({
            text: t.text,
            ...style,
            bold: Boolean(t.bold) || style.bold,
            italics: Boolean(t.italic) || style.italics,
            font: base.font,
            underline: {},
          }),
        ],
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

  const m = (key) => metrics(key, basePt, docLineHeight);
  const fontOf = (mm) => (mm.mono ? 'Courier New' : face.docx);

  const para = (key, text, extra = {}) => {
    const mm = m(key);
    return new Paragraph({
      children: runsFor(
        text,
        {
          bold: mm.bold,
          italic: mm.italic,
          font: fontOf(mm),
          size: mm.halfPoints,
          color: extra.color,
        },
        extra.marks
      ),
      spacing: { before: mm.before, after: mm.after, line: mm.line, lineRule: lineRuleFor(extra.marks) },
      shading: bandShading(extra.marks),
      // Word's default is to push a short tail to the next page rather than
      // split it — a guaranteed second sheet even when the text geometrically
      // fits. Off, always, for one-page output.
      widowControl: false,
      keepNext: extra.keepNext,
      alignment: extra.alignment,
      indent: extra.indent,
      border: extra.border,
      numbering: extra.numbering,
      bullet: extra.bullet,
    });
  };

  const listIndent = () => {
    const mm = m('li');
    return { left: ptToTwip(mm.sizePt * 1.35), hanging: ptToTwip(mm.sizePt * 0.95) };
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

  const children = [];
  const numberingConfigs = [];
  let olInstance = 0;

  for (const b of blocks) {
    switch (b.type) {
      case 'h1':
        children.push(para('h1', b.text, { keepNext: true, marks: marksOf(b, 'text') }));
        break;
      case 'h2':
        children.push(para('h2', b.text, { keepNext: true, marks: marksOf(b, 'text') }));
        break;
      case 'h3':
        children.push(para('h3', b.text, { keepNext: true, color: MUTED, marks: marksOf(b, 'text') }));
        break;
      case 'p':
        children.push(para('p', b.text, { alignment: AlignmentType.JUSTIFIED, marks: marksOf(b, 'text') }));
        break;
      case 'quote':
        children.push(
          para('quote', b.text, {
            color: MUTED,
            marks: marksOf(b, 'text'),
            indent: { left: ptToTwip(m('quote').sizePt * 0.85) },
            border: { left: { style: BorderStyle.SINGLE, size: 12, space: 6, color: RULE_COLOR } },
          })
        );
        break;
      case 'hr':
        children.push(
          new Paragraph({
            children: [],
            spacing: { before: ptToTwip(basePt * 0.5), after: ptToTwip(basePt * 0.5), line: ptToTwip(basePt * 0.4), lineRule: LineRuleType.EXACTLY },
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, space: 1, color: RULE_COLOR } },
          })
        );
        break;
      case 'code': {
        const mm = m('code');
        for (const line of b.text.split('\n')) {
          children.push(
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
        b.items.forEach((item, j) => {
          children.push(
            para('li', item, { bullet: { level: 0 }, indent: listIndent(), marks: marksOf(b, `items.${j}`) })
          );
        });
        break;
      case 'ol': {
        // A fresh instance per list, otherwise Word continues the previous
        // list's count instead of restarting at the block's own start value.
        const reference = `ol-${olInstance}`;
        const mm = m('li');
        numberingConfigs.push({
          reference,
          levels: [
            {
              level: 0,
              format: 'decimal',
              text: '%1.',
              alignment: AlignmentType.LEFT,
              start: b.start && b.start > 0 ? b.start : 1,
              style: { paragraph: { indent: listIndent() }, run: { size: mm.halfPoints, font: face.docx } },
            },
          ],
        });
        b.items.forEach((item, j) => {
          children.push(
            para('li', item, { numbering: { reference, level: 0 }, marks: marksOf(b, `items.${j}`) })
          );
        });
        olInstance++;
        break;
      }
      case 'table': {
        const mm = m('table');
        const colCount = Math.max(1, (b.head && b.head.length) || (b.rows[0] || []).length);
        const cellWidth = Math.floor(columnWidth / colCount);
        const borders = {
          top: { style: BorderStyle.SINGLE, size: 4, color: RULE_COLOR },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE_COLOR },
          left: { style: BorderStyle.SINGLE, size: 4, color: RULE_COLOR },
          right: { style: BorderStyle.SINGLE, size: 4, color: RULE_COLOR },
        };

        const makeRow = (cells, header, sub) =>
          new TableRow({
            tableHeader: header,
            children: cells.map(
              (text, j) =>
                new TableCell({
                  width: { size: cellWidth, type: WidthType.DXA },
                  borders,
                  shading: bandShading(marksOf(b, `${sub}.${j}`)) || (header ? { fill: 'F4F5F7' } : undefined),
                  children: [
                    new Paragraph({
                      children: runsFor(
                        text,
                        { bold: header, font: face.docx, size: mm.halfPoints },
                        marksOf(b, `${sub}.${j}`)
                      ),
                      spacing: { before: 0, after: 0, line: mm.line, lineRule: lineRuleFor(marksOf(b, `${sub}.${j}`)) },
                      widowControl: false,
                    }),
                  ],
                })
            ),
          });

        const rows = [];
        if (b.head && b.head.length) rows.push(makeRow(b.head, true, 'head'));
        (b.rows || []).forEach((r, ri) => {
          const cells = r.slice(0, colCount);
          while (cells.length < colCount) cells.push('');
          rows.push(makeRow(cells, false, `rows.${ri}`));
        });

        if (rows.length) {
          children.push(
            new Table({
              rows,
              width: { size: columnWidth, type: WidthType.DXA },
              // Fixed layout stops Word re-autofitting the columns, which
              // would change row heights and the total page height with them.
              layout: TableLayoutType.FIXED,
              columnWidths: Array(colCount).fill(cellWidth),
            })
          );
          children.push(
            new Paragraph({
              children: [],
              spacing: { before: 0, after: ptToTwip(mm.sizePt * 0.4), line: 1, lineRule: LineRuleType.EXACTLY },
            })
          );
        }
        break;
      }
      case 'image': {
        const image = imageRun(b);
        if (!image) break;
        children.push(
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
  }

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
