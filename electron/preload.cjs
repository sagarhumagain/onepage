/**
 * preload.cjs — the only bridge between the document code and the OS.
 *
 * CommonJS on purpose: preload scripts run in a sandboxed context, which does
 * not support ES modules. The surface is four calls, all of which the renderer
 * treats as optional so the same renderer runs in a plain browser too.
 */

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('onepage', {
  available: true,
  platform: process.platform,

  /** @param {string} html standalone sheet HTML @param {string} suggestedName */
  exportPdf: (html, suggestedName) => ipcRenderer.invoke('onepage:export-pdf', { html, suggestedName }),

  /** @param {Uint8Array} data @param {string} suggestedName */
  exportDocx: (data, suggestedName) => ipcRenderer.invoke('onepage:export-docx', { data, suggestedName }),

  /** @param {string} html standalone sheet HTML */
  print: (html) => ipcRenderer.invoke('onepage:print', { html }),

  /** @param {string} filePath */
  reveal: (filePath) => ipcRenderer.invoke('onepage:reveal', filePath),
});
