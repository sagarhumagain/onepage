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
  ImageRun,
  LineRuleType,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';

import { tokenizeInline } from './inline.js';
import { A4_TWIP, COLUMN_GAP_MM, SCALE, TWIP_PER_PX, mmToTwip, ptToHalfPoint, ptToTwip } from './scale.js';

/**
 * Word's line breaking never matches Chromium's exactly — kerning,
 * hyphenation and justification all differ slightly. Exporting a few percent
 * below the fitted size absorbs that residual error, which is the difference
 * between "usually one page" and "reliably one page".
 */
const WORD_REFLOW_SLACK = 0.97;

const RULE_COLOR = 'C8CCD4';
const MUTED = '565C66';

/**
 * Word expresses a run's background as a shading fill: six hex digits, no
 * leading hash. A highlight's padding has no OOXML equivalent at all, so the
 * colour carries over and the padding does not.
 */
const shadingFor = (t) => {
  if (!t.bg || !/^#[0-9a-fA-F]{6}$/.test(t.bg)) return undefined;
  return { fill: t.bg.slice(1).toUpperCase() };
};

const runsFor = (text, base) =>
  tokenizeInline(text).flatMap((t) => {
    const shading = shadingFor(t);
    const run = new TextRun({
      text: t.text,
      bold: Boolean(t.bold) || base.bold,
      italics: Boolean(t.italic) || base.italic,
      font: t.code ? 'Courier New' : base.font,
      size: base.size,
      color: base.color,
      shading,
    });
    if (!t.href) return [run];
    return [
      new ExternalHyperlink({
        link: t.href,
        children: [
          new TextRun({
            text: t.text,
            bold: Boolean(t.bold) || base.bold,
            italics: Boolean(t.italic) || base.italic,
            font: base.font,
            size: base.size,
            color: base.color,
            shading,
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
 * @param {{fontPt:number, columns:number, lineHeight:number, marginMm:number}} opts.layout
 * @param {{docx:string, label:string}} opts.face
 * @param {string} [opts.title]
 * @returns {Promise<Blob>}
 */
export async function buildDocx({ blocks, layout, face, title }) {
  const basePt = layout.fontPt * WORD_REFLOW_SLACK;
  const docLineHeight = layout.lineHeight;
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
      children: runsFor(text, {
        bold: mm.bold,
        italic: mm.italic,
        font: fontOf(mm),
        size: mm.halfPoints,
        color: extra.color,
      }),
      spacing: { before: mm.before, after: mm.after, line: mm.line, lineRule: LineRuleType.EXACTLY },
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

  const children = [];
  const numberingConfigs = [];
  let olInstance = 0;

  for (const b of blocks) {
    switch (b.type) {
      case 'h1':
        children.push(para('h1', b.text, { keepNext: true }));
        break;
      case 'h2':
        children.push(para('h2', b.text, { keepNext: true }));
        break;
      case 'h3':
        children.push(para('h3', b.text, { keepNext: true, color: MUTED }));
        break;
      case 'p':
        children.push(para('p', b.text, { alignment: AlignmentType.JUSTIFIED }));
        break;
      case 'quote':
        children.push(
          para('quote', b.text, {
            color: MUTED,
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
        for (const item of b.items) {
          children.push(para('li', item, { bullet: { level: 0 }, indent: listIndent() }));
        }
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
        for (const item of b.items) {
          children.push(para('li', item, { numbering: { reference, level: 0 } }));
        }
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

        const makeRow = (cells, header) =>
          new TableRow({
            tableHeader: header,
            children: cells.map(
              (text) =>
                new TableCell({
                  width: { size: cellWidth, type: WidthType.DXA },
                  borders,
                  shading: header ? { fill: 'F4F5F7' } : undefined,
                  children: [
                    new Paragraph({
                      children: runsFor(text, { bold: header, font: face.docx, size: mm.halfPoints }),
                      spacing: { before: 0, after: 0, line: mm.line, lineRule: LineRuleType.EXACTLY },
                      widowControl: false,
                    }),
                  ],
                })
            ),
          });

        const rows = [];
        if (b.head && b.head.length) rows.push(makeRow(b.head, true));
        for (const r of b.rows || []) {
          const cells = r.slice(0, colCount);
          while (cells.length < colCount) cells.push('');
          rows.push(makeRow(cells, false));
        }

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
        /*
         * docx sizes an image in pixels, but every other measurement in this
         * file is in twips. At 96dpi one CSS pixel is 15 twips, and the image
         * keeps its own aspect ratio rather than an assumed one — a portrait
         * photo squashed to 3:5 was the most visible way this export used to
         * disagree with the preview.
         */
        if (!b.src || !b.src.startsWith('data:')) break;
        try {
          const image = m('image');
          const base64 = b.src.split(',')[1];
          const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));

          const columnPx = columnWidth / TWIP_PER_PX;
          const pct = Math.max(5, Math.min(100, Number(b.widthPct) || 62)) / 100;
          const width = Math.max(1, Math.round(columnPx * pct));
          const ratio = b.naturalWidth && b.naturalHeight ? b.naturalHeight / b.naturalWidth : 0.62;
          const height = Math.max(1, Math.round(width * ratio));

          children.push(
            new Paragraph({
              children: [
                new ImageRun({
                  data: bytes,
                  transformation: { width, height },
                  altText: { title: b.alt || 'Image', description: b.alt || 'Image', name: b.alt || 'image' },
                }),
              ],
              alignment:
                b.align === 'left'
                  ? AlignmentType.LEFT
                  : b.align === 'right'
                    ? AlignmentType.RIGHT
                    : AlignmentType.CENTER,
              spacing: { before: image.before, after: b.alt ? 0 : image.after },
              widowControl: false,
            })
          );

          if (b.alt) {
            const cap = m('caption');
            children.push(
              new Paragraph({
                children: runsFor(b.alt, { font: face.docx, size: cap.halfPoints, color: MUTED }),
                alignment:
                  b.align === 'left'
                    ? AlignmentType.LEFT
                    : b.align === 'right'
                      ? AlignmentType.RIGHT
                      : AlignmentType.CENTER,
                spacing: { before: cap.before, after: image.after },
              })
            );
          }
        } catch {
          // An image that cannot be decoded is left out rather than failing
          // the whole export.
        }
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
