/**
 * main.js — Electron main process.
 *
 * The shell is deliberately thin. All document logic lives in the renderer as
 * plain ES modules with no Electron imports, so the same code runs in a browser
 * during development and could move to another shell without changes.
 *
 * Main owns exactly three things the renderer cannot do: native save dialogs,
 * writing files, and driving Chromium's print pipeline.
 */

import { app, BrowserWindow, dialog, ipcMain, shell } from 'electron';
import { writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DEV_URL = process.env.ONEPAGE_DEV_URL;

/** @type {BrowserWindow | null} */
let mainWindow = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 720,
    minHeight: 520,
    show: false,
    backgroundColor: '#e8e9ec',
    title: 'OnePage',
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: false,
    },
  });

  mainWindow.once('ready-to-show', () => mainWindow.show());

  if (DEV_URL) mainWindow.loadURL(DEV_URL);
  else mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));

  // Links to the outside world open in the real browser, never in the app.
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

/**
 * Render a standalone HTML sheet in an offscreen window.
 *
 * The exported HTML contains only the A4 page and document.css — no app
 * chrome — so what Chromium paginates here is exactly what the preview
 * showed. The caller decides whether to turn it into a PDF or a print job.
 *
 * @param {string} html
 * @param {(contents: Electron.WebContents) => Promise<T>} use
 * @returns {Promise<T>}
 * @template T
 */
async function withRenderedSheet(html, use) {
  const win = new BrowserWindow({
    show: false,
    webPreferences: { offscreen: false, contextIsolation: true, nodeIntegration: false, sandbox: true },
  });

  try {
    const loaded = new Promise((resolve, reject) => {
      win.webContents.once('did-finish-load', resolve);
      win.webContents.once('did-fail-load', (_e, code, desc) => reject(new Error(`${desc} (${code})`)));
    });
    await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    await loaded;
    // Fonts must be resolved before pagination, or the first line breaks move.
    await win.webContents.executeJavaScript('document.fonts.ready.then(() => true)');
    return await use(win.webContents);
  } finally {
    if (!win.isDestroyed()) win.destroy();
  }
}

/**
 * Chromium's PDF options. Both `pageSize` and `preferCSSPageSize` are set on
 * purpose: preferCSSPageSize makes the page's own `@page { size: A4 }` rule
 * authoritative, and pageSize is the belt-and-braces fallback if that rule
 * ever goes missing. Without them Chromium silently defaults to US Letter and
 * rescales the content, which would undo the fit.
 *
 * Margins are zero because the A4 page element owns its own margins as
 * padding — the fitter measured them that way.
 */
const PDF_OPTIONS = {
  pageSize: 'A4',
  preferCSSPageSize: true,
  printBackground: true,
  scale: 1,
  margins: { top: 0, bottom: 0, left: 0, right: 0 },
  displayHeaderFooter: false,
  generateDocumentOutline: false,
  // The last belt. Everything upstream — the fitter, the overflow clamp, the
  // zeroed margins — is meant to make a second page impossible; this makes it
  // impossible regardless. Sub-pixel rounding, a stray break, or a future
  // Chromium change cannot emit page 2 or a trailing blank sheet.
  // Deliberately NOT used in test/e2e.js, so the tests still measure the true
  // page count rather than a truncated one.
  pageRanges: '1',
};

ipcMain.handle('onepage:export-pdf', async (_event, { html, suggestedName }) => {
  const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
    title: 'Save PDF',
    defaultPath: `${suggestedName || 'document'}.pdf`,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  });
  if (canceled || !filePath) return { saved: false };

  const buffer = await withRenderedSheet(html, (contents) => contents.printToPDF(PDF_OPTIONS));
  await writeFile(filePath, buffer);
  return { saved: true, filePath };
});

ipcMain.handle('onepage:export-docx', async (_event, { data, suggestedName }) => {
  const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
    title: 'Save Word document',
    defaultPath: `${suggestedName || 'document'}.docx`,
    filters: [{ name: 'Word document', extensions: ['docx'] }],
  });
  if (canceled || !filePath) return { saved: false };

  await writeFile(filePath, Buffer.from(data));
  return { saved: true, filePath };
});

ipcMain.handle('onepage:print', async (_event, { html }) => {
  return withRenderedSheet(
    html,
    (contents) =>
      new Promise((resolve, reject) => {
        contents.print(
          {
            // Electron's print path never sets Chromium's scaling type, so it
            // stays at "actual size" — no silent shrink-to-fit can undo the
            // layout we computed.
            pageSize: 'A4',
            printBackground: true,
            scaleFactor: 100,
            margins: { marginType: 'none' },
            // Same belt as the PDF path: only ever print the first sheet.
            pageRanges: [{ from: 0, to: 0 }],
          },
          (success, failureReason) => {
            if (success) resolve({ printed: true });
            else if (/cancel/i.test(failureReason || '')) resolve({ printed: false, canceled: true });
            else reject(new Error(failureReason || 'Print failed'));
          }
        );
      })
  );
});

ipcMain.handle('onepage:reveal', async (_event, filePath) => {
  if (filePath) shell.showItemInFolder(filePath);
});

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
