/**
 * exporters.js — print, PDF and Word, with a browser fallback.
 *
 * The desktop shell exposes `window.onepage`; when it is absent (running the
 * renderer in a plain browser during development) every path degrades to
 * something that still works. No file here imports Electron.
 */

import { buildDocx } from './to-docx.js';

const bridge = () => (typeof window !== 'undefined' && window.onepage) || null;

/** File-system-safe name derived from the document's first heading. */
export function suggestName(title) {
  const base = String(title || 'document')
    .replace(/[\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 60);
  return base || 'document';
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/**
 * Print the fitted sheet.
 *
 * In the desktop shell the sheet HTML is rendered in an offscreen window and
 * printed through Chromium at exactly A4 and 100% scale, so no "fit to page"
 * setting in a dialog can undo the layout. In a browser we fall back to
 * printing the preview iframe directly.
 *
 * @param {{html: string, iframe?: HTMLIFrameElement}} opts
 */
export async function printSheet({ html, iframe }) {
  const api = bridge();
  if (api) return api.print(html);

  if (iframe && iframe.contentWindow) {
    iframe.contentWindow.focus();
    iframe.contentWindow.print();
    return { printed: true };
  }
  window.print();
  return { printed: true };
}

/**
 * @param {{html: string, title: string}} opts
 */
export async function exportPdf({ html, title }) {
  const api = bridge();
  const name = suggestName(title);

  if (api) {
    const result = await api.exportPdf(html, name);
    return result;
  }

  // A browser cannot write a PDF directly; the print dialog's "Save as PDF"
  // is the honest fallback, and the page is already sized for it.
  const w = window.open('', '_blank');
  if (!w) throw new Error('Allow pop-ups to export a PDF from the browser');
  w.document.write(html);
  w.document.close();
  await new Promise((r) => setTimeout(r, 300));
  w.focus();
  w.print();
  return { saved: false, viaDialog: true };
}

/**
 * @param {{blocks: any[], layout: object, face: object, title: string}} opts
 */
export async function exportDocx({ blocks, layout, face, title }) {
  const blob = await buildDocx({ blocks, layout, face, title });
  const name = suggestName(title);
  const api = bridge();

  if (api) {
    const bytes = new Uint8Array(await blob.arrayBuffer());
    return api.exportDocx(bytes, name);
  }

  downloadBlob(blob, `${name}.docx`);
  return { saved: true };
}
