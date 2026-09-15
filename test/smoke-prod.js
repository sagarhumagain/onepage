/**
 * smoke-prod.js — verify the PACKAGED renderer, not the dev server.
 *
 * The production bundle differs from dev in two ways that matter: document.css
 * is inlined via `?raw` instead of served, and the test hook is stripped by
 * `import.meta.env.DEV`. So this drives the UI the way a person does — typing
 * into the textarea — and reads back only what the DOM exposes.
 *
 * Usage: npm run build && npx electron test/smoke-prod.js
 */

import { app, BrowserWindow } from 'electron';
import { execFileSync } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import os from 'node:os';
import path from 'node:path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(os.tmpdir(), 'onepage-smoke');

const SAMPLE = [
  'Quarterly Field Report',
  '',
  'EXECUTIVE SUMMARY',
  'The programme reached 12,400 households across three districts this quarter, exceeding the',
  'revised target by 6%. Coverage improved in every district, and the cost per household fell',
  'for the third consecutive quarter. Supervision visits remain the strongest single predictor',
  'of district performance.',
  '',
  '1. Methodology',
  'Data were collected using the standard household survey instrument between June and August.',
  'Enumerators were retrained in May following the coverage dip recorded in the first quarter.',
  '',
  'Key findings:',
  '- Coverage rose to 87%, up from 81% in the previous quarter',
  '- Dropout fell by four percentage points, with the improvement sustained across all months',
  '- Cost per household fell 12% to NPR 340',
  '- Two districts now exceed the national benchmark',
  '',
  'Priorities for the next quarter:',
  '1. Restore the full schedule of supervision visits',
  '2. Re-train enumerators in the two lagging districts',
  '3. Close the remaining data-entry backlog',
  '',
  '> Sustained gains require continued supervision. Where visits lapsed, coverage fell within',
  '> two months.',
  '',
  '| District | Reached | Target | Variance |',
  '|---|---|---|---|',
  '| Bara | 4,200 | 4,000 | +5% |',
  '| Parsa | 3,900 | 4,200 | -7% |',
  '| Rautahat | 4,300 | 4,000 | +8% |',
  '',
  'Full detail is published at https://example.org/reports and questions may be sent to',
  'the programme office. The **next review** is scheduled for the second week of October.',
].join('\n');

async function run() {
  mkdirSync(OUT, { recursive: true });

  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    show: false,
    backgroundColor: '#e8e9ec',
    webPreferences: {
      preload: path.join(__dirname, '..', 'electron', 'preload.cjs'),
      contextIsolation: true,
      sandbox: true,
    },
  });

  await win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  await win.webContents.executeJavaScript(
    `new Promise(r => { const t = setInterval(() => {
       if (document.documentElement.dataset.ready === 'true') { clearInterval(t); r(true); }
     }, 40); })`
  );

  const probe = await win.webContents.executeJavaScript(`(async () => {
    const src = document.getElementById('source');
    src.value = ${JSON.stringify(SAMPLE)};
    src.dispatchEvent(new Event('input', { bubbles: true }));
    await new Promise(r => setTimeout(r, 1200));

    const doc = document.getElementById('sheet').contentDocument;
    const page = doc.getElementById('page');
    const body = doc.getElementById('page-body');
    const cs = getComputedStyle(body);
    const ps = getComputedStyle(page);
    const box = body.getBoundingClientRect();
    const kids = [...body.children].filter(el => el.style.height !== '0px');
    const last = kids[kids.length - 1];

    return {
      testHookStripped: typeof window.__onepage === 'undefined',
      bridgePresent: Boolean(window.onepage && window.onepage.available),
      // If document.css failed to inline, these come back as browser defaults.
      stylesApplied: Math.abs(parseFloat(ps.width) - 793.7) < 0.5 && cs.overflow === 'hidden',
      pageWidth: ps.width,
      pageHeight: ps.height,
      fontSize: cs.fontSize,
      columns: cs.columnCount,
      overflowPx: Math.round((last.getBoundingClientRect().bottom - box.bottom) * 10) / 10,
      blocks: kids.length,
      tables: body.querySelectorAll('table').length,
      lists: body.querySelectorAll('ul,ol').length,
      links: body.querySelectorAll('a').length,
      status: {
        words: document.getElementById('stat-words').textContent,
        size: document.getElementById('stat-size').textContent,
        layout: document.getElementById('stat-layout').textContent,
        note: document.getElementById('stat-note').textContent,
      },
    };
  })()`);

  win.show();
  await new Promise((r) => setTimeout(r, 600));
  const shot = await win.webContents.capturePage();
  const shotFile = path.join(OUT, 'app.png');
  writeFileSync(shotFile, shot.toPNG());

  console.log('\n  PRODUCTION BUNDLE SMOKE TEST');
  console.log('  ' + '-'.repeat(60));
  console.log('  dev test hook stripped :', probe.testHookStripped);
  console.log('  preload bridge present :', probe.bridgePresent);
  console.log('  document.css inlined   :', probe.stylesApplied, `(page ${probe.pageWidth} x ${probe.pageHeight})`);
  console.log('  fitted font size       :', probe.fontSize);
  console.log('  columns                :', probe.columns);
  console.log('  overflow past page box :', probe.overflowPx + 'px');
  console.log('  blocks/tables/lists/links:', probe.blocks, probe.tables, probe.lists, probe.links);
  console.log('  status bar             :', Object.values(probe.status).join('  |  '));
  console.log('  screenshot             :', shotFile);

  const ok =
    probe.testHookStripped &&
    probe.bridgePresent &&
    probe.stylesApplied &&
    probe.overflowPx <= 1.5 &&
    probe.tables === 1 &&
    probe.lists === 2 &&
    probe.links >= 1;
  console.log('  ' + '-'.repeat(60));
  console.log(`  ${ok ? 'PASS' : 'FAIL'}\n`);
  app.exit(ok ? 0 : 1);
}

app.whenReady().then(() => run().catch((e) => { console.error(e); app.exit(1); }));
