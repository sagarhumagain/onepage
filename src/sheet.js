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

const SHELL = `<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<style>${documentCssText}</style>
</head><body class="print-root">
<div class="page" id="page"><div class="page-body" id="page-body"></div></div>
</body></html>`;

/**
 * @param {HTMLIFrameElement} iframe
 * @returns {Promise<{doc:Document, win:Window, page:HTMLElement, body:HTMLElement}>}
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
 *   fontFamily:string, title?:string}} layout
 */
export function standaloneHtml(bodyHtml, layout) {
  const title = String(layout.title || 'Document').replace(/[<&]/g, ' ');
  const vars = [
    `--doc-font-pt:${layout.fontPt}`,
    `--doc-columns:${cssColumns(layout.columns)}`,
    `--doc-line-height:${layout.lineHeight}`,
    `--doc-margin-mm:${layout.marginMm}`,
    `--doc-font-family:${layout.fontFamily}`,
  ].join(';');

  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><title>${title}</title>
<style>${documentCssText}</style>
</head><body class="print-root">
<div class="page" style="${vars}">
<div class="page-body">
${bodyHtml}
</div></div></body></html>`;
}
