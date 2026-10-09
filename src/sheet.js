/**
 * sheet.js — owns the A4 iframe.
 *
 * The iframe is the app's single rendering surface: what you see in the
 * preview is the exact DOM the fitter measures, and the standalone HTML it
 * exports is the same markup with the same stylesheet inlined. Nothing is
 * re-laid-out for output, so "it fits on screen" and "it fits on paper"
 * cannot drift.
 *
 * document.css is imported with `?raw` rather than fetched. Fetching it would
 * work in a production build but not in dev, where the dev server answers a
 * request for a stylesheet with a JavaScript module that installs it — so the
 * exported file would silently carry JS in its <style> tag and render with no
 * styles at all. Importing the text means the preview and the export are
 * guaranteed to be the same bytes.
 */

import documentCssText from './document.css?raw';
import { cssColumns } from './fit.js';

/*
 * Affordances that exist only in the preview: the hover and selection rings on
 * an image, and the cursor that says a picture can be picked up. They live in
 * their own <style>, never in document.css, so there is no route by which an
 * editing outline could reach a printed page or an exported file.
 */
const EDITOR_CHROME = `
.doc-image { cursor: pointer; }
.doc-image:hover { outline: 1.5px dashed rgba(15,118,110,0.7); outline-offset: 2px; }
.doc-image[data-selected='1'] { outline: 2px solid #0f766e; outline-offset: 2px; }
.doc-image[data-selected='1'] img { pointer-events: none; }
::selection { background: rgba(15,118,110,0.22); }
.footer-slot { cursor: pointer; }
.footer-slot:hover { outline: 1.5px dashed rgba(15,118,110,0.7); outline-offset: 2px; }
.footer-slot[data-selected='1'] { outline: 2px solid #0f766e; outline-offset: 2px; }
.footer-slot[data-empty='1'] { min-width: 24mm; justify-content: center; border: 1.2px dashed rgba(15,118,110,0.55); border-radius: 1.5mm; box-sizing: border-box; }
.footer-slot[data-empty='1']::after { content: '+ Image'; font: 600 8pt system-ui, sans-serif; color: #0f766e; }
/* With no picture at all the footer takes no room; its empty slots wait in the bottom margin. */
.page:not([data-footer='1']) .page-footer { display: flex; bottom: max(0.5mm, calc(var(--doc-margin-mm) * 0.5mm - 3mm)); height: 6mm; }
`;

const SHELL = `<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<style>${documentCssText}</style>
<style>${EDITOR_CHROME}</style>
</head><body class="print-root">
<div class="page" id="page"><div class="page-body" id="page-body"></div><div class="page-footer" id="page-footer" contenteditable="false"></div></div>
</body></html>`;

/**
 * @param {HTMLIFrameElement} iframe
 * @returns {Promise<{doc:Document, win:Window, page:HTMLElement, body:HTMLElement, footer:HTMLElement}>}
 */
export function mountSheet(iframe) {
  return new Promise((resolve) => {
    iframe.addEventListener(
      'load',
      () => {
        const doc = iframe.contentDocument;
        resolve({
          doc,
          win: iframe.contentWindow,
          page: doc.getElementById('page'),
          body: doc.getElementById('page-body'),
          footer: doc.getElementById('page-footer'),
        });
      },
      { once: true }
    );
    iframe.srcdoc = SHELL;
  });
}

/**
 * A standalone HTML file carrying the sheet and its styles with the fitted
 * layout baked in. This is what gets printed and turned into a PDF.
 *
 * @param {string} bodyHtml
 * @param {{fontPt:number, columns:number, lineHeight:number, marginMm:number,
 *   fontFamily:string, title?:string, bgColor?:string, imageScale?:number,
 *   inset?:boolean, footerHtml?:string}} layout — `footerHtml` is the
 *   footer's content, given only when it has a picture in it
 */
export function standaloneHtml(bodyHtml, layout) {
  const title = String(layout.title || 'Document').replace(/[<&]/g, ' ');
  const vars = [
    `--doc-font-pt:${layout.fontPt}`,
    `--doc-columns:${cssColumns(layout.columns)}`,
    `--doc-line-height:${layout.lineHeight}`,
    `--doc-margin-mm:${layout.marginMm}`,
    `--doc-font-family:${layout.fontFamily}`,
    `--doc-bg:${layout.bgColor || '#ffffff'}`,
    `--doc-image-scale:${layout.imageScale == null ? 1 : layout.imageScale}`,
  ].join(';');

  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>${title}</title>
<style>${documentCssText}</style>
</head><body class="print-root">
<div class="page${layout.inset ? ' inset' : ''}"${layout.footerHtml ? ' data-footer="1"' : ''} style="${vars}">
<div class="page-body">
${bodyHtml}
</div>${layout.footerHtml ? `<div class="page-footer">${layout.footerHtml}</div>` : ''}</div></body></html>`;
}
