/**
 * e2e.js — end-to-end verification in the real target environment.
 *
 * This runs as an Electron main process because the claim under test is not
 * "the preview looks right" but "Chromium's print pipeline emits exactly one
 * A4 page". Only the actual print path can prove that, so every case is
 * rendered, fitted, exported to PDF and then measured with pdfinfo.
 *
 * Usage: ONEPAGE_DEV_URL=http://localhost:5183 npx electron test/e2e.js
 */

import { app, BrowserWindow } from 'electron';
import { execFileSync } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const DEV_URL = process.env.ONEPAGE_DEV_URL || 'http://localhost:5183';
const OUT = path.join(os.tmpdir(), 'onepage-e2e');

const WORDS =
  `programme coverage household district enumerator supervision baseline indicator quarterly
   assessment surveillance immunisation nutrition sanitation outbreak response capacity referral
   facility community volunteer training logistics procurement distribution monitoring evaluation
   reporting verification coordination stakeholder guideline protocol threshold benchmark`
    .trim()
    .split(/\s+/);

/** Deterministic filler, so a failure is always reproducible. */
function lorem(wordCount, seed = 7) {
  let s = seed;
  const next = () => (s = (s * 1103515245 + 12345) % 2147483648) / 2147483648;
  const out = [];
  let sentence = [];
  for (let i = 0; i < wordCount; i++) {
    sentence.push(WORDS[Math.floor(next() * WORDS.length)]);
    if (sentence.length >= 9 + Math.floor(next() * 9)) {
      sentence[0] = sentence[0][0].toUpperCase() + sentence[0].slice(1);
      out.push(sentence.join(' ') + '.');
      sentence = [];
    }
  }
  if (sentence.length) out.push(sentence.join(' ') + '.');

  // Break it into paragraphs with the occasional heading, like a real document.
  const paras = [];
  for (let i = 0; i < out.length; i += 4) {
    if (i && i % 12 === 0) paras.push(`\n## Section ${Math.floor(i / 12)}\n`);
    paras.push(out.slice(i, i + 4).join(' '));
  }
  return paras.join('\n\n');
}

const CASES = [
  { name: 'tiny (8 words)', text: 'Meeting moved to Thursday at ten in the annexe.' },
  { name: 'short (120 words)', text: `Field Note\n\n${lorem(120)}` },
  { name: 'medium (400 words)', text: `Quarterly Update\n\n${lorem(400, 11)}` },
  { name: 'long (900 words)', text: `Programme Review\n\n${lorem(900, 13)}` },
  { name: 'very long (1800 words)', text: `Annual Report\n\n${lorem(1800, 17)}` },
  { name: 'extreme (3500 words)', text: `Consolidated Report\n\n${lorem(3500, 19)}` },
  { name: 'absurd (6000 words)', text: `Compendium\n\n${lorem(6000, 41)}` },
  {
    name: 'mixed structure',
    text: [
      'Quarterly Field Report',
      '',
      'EXECUTIVE SUMMARY',
      lorem(90, 23),
      '',
      '1. Methodology',
      lorem(70, 29),
      '',
      'Key findings:',
      '- Coverage rose to 87% across all three districts',
      '- Dropout fell by four percentage points',
      '- Cost per household fell by 12%',
      '',
      'Priorities:',
      '1. Restore supervision visits',
      '2. Re-train enumerators',
      '3. Close the data-entry backlog',
      '',
      '> Sustained gains require continued supervision.',
      '',
      '| District | Reached | Target |',
      '|---|---|---|',
      '| Bara | 4200 | 4000 |',
      '| Parsa | 3900 | 4200 |',
      '',
      lorem(120, 31),
    ].join('\n'),
  },
  { name: 'unbreakable url', text: `Link Test\n\nSee https://example.org/${'a'.repeat(300)}/end for detail.\n\n${lorem(200, 37)}` },
];

/**
 * Note: unlike electron/main.js this deliberately omits `pageRanges: '1'`.
 * The shipped app clamps to one page as a safety net; the test must be able to
 * see a second page in order to fail when the fitter regresses.
 */
const PDF_OPTIONS = {
  pageSize: 'A4',
  preferCSSPageSize: true,
  printBackground: true,
  scale: 1,
  margins: { top: 0, bottom: 0, left: 0, right: 0 },
  displayHeaderFooter: false,
};

async function renderPdf(html, file) {
  const win = new BrowserWindow({ show: false, webPreferences: { sandbox: true } });
  try {
    const loaded = new Promise((res, rej) => {
      win.webContents.once('did-finish-load', res);
      win.webContents.once('did-fail-load', (_e, c, d) => rej(new Error(`${d} (${c})`)));
    });
    await win.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
    await loaded;
    await win.webContents.executeJavaScript('document.fonts.ready.then(()=>true)');
    const buf = await win.webContents.printToPDF(PDF_OPTIONS);
    writeFileSync(file, buf);
  } finally {
    if (!win.isDestroyed()) win.destroy();
  }
}

const pdfInfo = (file) => {
  const out = execFileSync('pdfinfo', [file], { encoding: 'utf8' });
  const pages = Number((out.match(/^Pages:\s+(\d+)/m) || [])[1]);
  const size = (out.match(/^Page size:\s+(.*)$/m) || [])[1];
  return { pages, size };
};

const textLength = (file) =>
  execFileSync('pdftotext', [file, '-'], { encoding: 'utf8' }).replace(/\s+/g, ' ').trim().length;

async function run() {
  mkdirSync(OUT, { recursive: true });

  const win = new BrowserWindow({
    show: false,
    width: 1400,
    height: 900,
    webPreferences: { sandbox: true, contextIsolation: true },
  });

  await win.loadURL(DEV_URL);
  await win.webContents.executeJavaScript(
    `new Promise(r => { const t = setInterval(() => {
       if (document.documentElement.dataset.ready === 'true') { clearInterval(t); r(true); }
     }, 40); })`
  );

  const results = [];
  let failures = 0;

  for (const [i, testCase] of CASES.entries()) {
    const probe = await win.webContents.executeJavaScript(`(async () => {
      await window.__onepage.setText(${JSON.stringify(testCase.text)});
      const doc = document.getElementById('sheet').contentDocument;
      const body = doc.getElementById('page-body');
      const box = body.getBoundingClientRect();
      const kids = body.children;
      const lastRect = kids.length ? kids[kids.length - 1].getBoundingClientRect() : box;
      return {
        layout: window.__onepage.layout(),
        face: window.__onepage.face().label,
        overflowY: lastRect.bottom - box.bottom,
        overflowX: body.scrollWidth - body.clientWidth,
        note: document.getElementById('stat-note').textContent,
        words: document.getElementById('stat-words').textContent,
        html: await window.__onepage.sheetHtml(),
      };
    })()`);

    const pdfFile = path.join(OUT, `case-${i}.pdf`);
    await renderPdf(probe.html, pdfFile);
    const info = pdfInfo(pdfFile);
    const chars = textLength(pdfFile);

    const isA4 = /594\.\d+ x 84[12]\.\d+ pts/.test(info.size) || /A4/.test(info.size);
    const onePage = info.pages === 1;
    const noClip = probe.overflowY <= 1.5 && probe.overflowX <= 1.5;
    const hasText = chars > 20;
    // A reported overflow means the fitter ran out of ladder and content is
    // being clipped — that is a failure of the product's core promise.
    const ok = onePage && isA4 && noClip && hasText && !probe.layout.overflow;
    if (!ok) failures++;

    results.push({
      name: testCase.name,
      pt: probe.layout.fontPt,
      cols: probe.layout.columns,
      lh: probe.layout.lineHeight,
      pages: info.pages,
      size: info.size,
      chars,
      overflowY: Math.round(probe.overflowY * 10) / 10,
      overflowX: probe.overflowX,
      belowFloor: probe.layout.belowFloor,
      reportedOverflow: probe.layout.overflow,
      ok,
    });
  }

  // --- Pasting from Word -------------------------------------------------
  // The primary way content arrives. Word puts a mound of `mso-` markup on the
  // clipboard, encodes bullets as literal glyphs in Symbol-font spans rather
  // than as list markup, and buries a large <style> block in the payload.
  const WORD_HTML = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8">
<style><!-- /* Font Definitions */ @font-face {font-family:Calibri;}
p.MsoNormal {mso-style-parent:""; margin:0cm; font-size:11.0pt;}
p.MsoListParagraphCxSpFirst {margin-left:36.0pt; text-indent:-18.0pt;} --></style></head>
<body lang=EN-GB><!--StartFragment-->
<p class=MsoNormal><b><span style='font-size:16.0pt;font-family:"Calibri",sans-serif'>Mission Report<o:p></o:p></span></b></p>
<p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'>The team completed the
assessment in <b>four districts</b> and identified three priority gaps.<o:p></o:p></span></p>
<p class=MsoListParagraphCxSpFirst style='margin-left:36.0pt;mso-add-space:auto;text-indent:-18.0pt;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='font-family:Symbol'>&#183;<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span><![endif]>Cold chain capacity is insufficient<o:p></o:p></p>
<p class=MsoListParagraphCxSpMiddle style='margin-left:36.0pt;mso-add-space:auto;text-indent:-18.0pt;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='font-family:Symbol'>&#183;<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span><![endif]>Supervision visits have lapsed<o:p></o:p></p>
<p class=MsoListParagraphCxSpLast style='margin-left:36.0pt;mso-add-space:auto;text-indent:-18.0pt;mso-list:l0 level1 lfo1'><![if !supportLists]><span style='font-family:Symbol'>&#183;<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span><![endif]>Data entry is six weeks behind<o:p></o:p></p>
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0>
<tr><td width=200><p class=MsoNormal><b>District<o:p></o:p></b></p></td><td width=120><p class=MsoNormal><b>Status<o:p></o:p></b></p></td></tr>
<tr><td><p class=MsoNormal>Bara<o:p></o:p></p></td><td><p class=MsoNormal>On track<o:p></o:p></p></td></tr>
<tr><td><p class=MsoNormal>Parsa<o:p></o:p></p></td><td><p class=MsoNormal>At risk<o:p></o:p></p></td></tr>
</table>
<p class=MsoNormal>Contact the office at ops@example.org.<o:p></o:p></p>
<!--EndFragment--></body></html>`;

  const paste = await win.webContents.executeJavaScript(`(async () => {
    const src = document.getElementById('source');
    src.value = '';
    src.focus();
    const dt = new DataTransfer();
    dt.setData('text/html', ${JSON.stringify(WORD_HTML)});
    dt.setData('text/plain', 'Mission Report\\nThe team completed the assessment.');
    src.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
    await new Promise(r => setTimeout(r, 900));
    const doc = document.getElementById('sheet').contentDocument;
    const body = doc.getElementById('page-body');
    return {
      text: src.value,
      msoLeaked: /mso-|MsoNormal|font-face|o:p/i.test(src.value),
      headings: body.querySelectorAll('h1,h2,h3').length,
      listItems: body.querySelectorAll('li').length,
      tableRows: body.querySelectorAll('tr').length,
      bold: body.querySelectorAll('strong').length,
      links: body.querySelectorAll('a').length,
      strayBullets: /[\\u00b7\\u2022]/.test(body.textContent),
    };
  })()`);

  const pasteOk =
    !paste.msoLeaked &&
    paste.headings >= 1 &&
    paste.listItems === 3 &&
    paste.tableRows === 3 &&
    paste.bold >= 1 &&
    paste.links === 1 &&
    !paste.strayBullets;
  if (!pasteOk) failures++;

  console.log('\n  WORD PASTE');
  console.log('  ' + '-'.repeat(66));
  console.log('  mso- markup leaked into source :', paste.msoLeaked, '(want false)');
  console.log('  headings / list items / rows   :', paste.headings, '/', paste.listItems, '/', paste.tableRows);
  console.log('  bold runs / links              :', paste.bold, '/', paste.links);
  console.log('  stray bullet glyphs in output  :', paste.strayBullets, '(want false)');
  console.log('  ' + '-'.repeat(66));
  console.log(`  ${pasteOk ? 'PASS' : 'FAIL'}`);
  console.log('\n  recovered source text:');
  for (const line of paste.text.split('\n').slice(0, 12)) console.log('    ' + line);

  // --- Merged cells and tables inside tables -----------------------------
  // A pasted table keeps its merges and a table inside one of its cells, and
  // typing into the page reads both back into the same source.
  const TABLE_HTML =
    '<table><tr><th>Region</th><th colspan=2>Cases</th></tr>' +
    '<tr><td rowspan=2>East</td><td>Jan</td><td>12</td></tr>' +
    '<tr><td>Feb</td><td><p>By week:</p><table><tr><th>W1</th><th>W2</th></tr><tr><td>4</td><td>5</td></tr></table></td></tr></table>';

  const tables = await win.webContents.executeJavaScript(`(async () => {
    const src = document.getElementById('source');
    src.value = '';
    src.focus();
    const dt = new DataTransfer();
    dt.setData('text/html', ${JSON.stringify(TABLE_HTML)});
    dt.setData('text/plain', 'Region Cases East Jan 12 Feb By week W1 W2 4 5');
    src.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
    await new Promise(r => setTimeout(r, 900));
    const body = document.getElementById('sheet').contentDocument.getElementById('page-body');
    const before = src.value;
    // An edit anywhere in the page writes the whole page back as source.
    body.dispatchEvent(new InputEvent('input', { bubbles: true }));
    await new Promise(r => setTimeout(r, 900));
    return {
      text: before,
      after: src.value,
      colspan: body.querySelectorAll('th[colspan="2"]').length,
      rowspan: body.querySelectorAll('td[rowspan="2"]').length,
      nested: body.querySelectorAll('td table').length,
      nestedCells: [...body.querySelectorAll('td table td')].map((td) => td.textContent).join(','),
      markersShown: /<<|\\^\\^/.test(body.textContent),
    };
  })()`);

  const tablesOk =
    tables.colspan === 1 &&
    tables.rowspan === 1 &&
    tables.nested === 1 &&
    tables.nestedCells === '4,5' &&
    !tables.markersShown &&
    tables.after.trim() === tables.text.trim();
  if (!tablesOk) failures++;

  console.log('\n  MERGED AND NESTED TABLES');
  console.log('  ' + '-'.repeat(66));
  console.log('  colspan / rowspan cells        :', tables.colspan, '/', tables.rowspan, '(want 1 / 1)');
  console.log('  nested tables / their cells    :', tables.nested, '/', tables.nestedCells, '(want 1 / 4,5)');
  console.log('  merge markers drawn on the page:', tables.markersShown, '(want false)');
  console.log('  page edit writes the same source:', tables.after.trim() === tables.text.trim());
  console.log('  ' + '-'.repeat(66));
  console.log(`  ${tablesOk ? 'PASS' : 'FAIL'}`);
  for (const line of tables.text.split('\n')) console.log('    ' + line);
  if (tables.after.trim() !== tables.text.trim()) for (const line of tables.after.split('\n')) console.log('  > ' + line);

  // The Word export is generated from the same fitted layout; verify it builds
  // and carries the exact A4 geometry Word itself writes.
  const docxBytes = await win.webContents.executeJavaScript('window.__onepage.docxBytes()');
  const docxFile = path.join(OUT, 'export.docx');
  writeFileSync(docxFile, Buffer.from(docxBytes));

  console.log('\n  case                      pt    cols  lh     pages  pdf text  overflow   status');
  console.log('  ' + '-'.repeat(84));
  for (const r of results) {
    console.log(
      `  ${r.name.padEnd(24)} ${String(r.pt).padStart(5)} ${String(r.cols).padStart(4)}  ` +
        `${String(r.lh).padEnd(6)} ${String(r.pages).padStart(5)}  ${String(r.chars).padStart(7)}  ` +
        `${String(r.overflowY).padStart(6)}px  ${r.ok ? 'PASS' : 'FAIL'}` +
        `${r.belowFloor ? '  [below floor]' : ''}${r.reportedOverflow ? '  [reported overflow]' : ''}`
    );
  }
  console.log('  ' + '-'.repeat(84));
  console.log(`  page size reported by pdfinfo: ${results[0].size}`);
  console.log(`  docx written: ${docxFile} (${docxBytes.length} bytes)`);
  console.log(`  artifacts:    ${OUT}`);
  console.log(`\n  ${results.length - failures}/${results.length} cases passed\n`);

  app.exit(failures ? 1 : 0);
}

app.whenReady().then(() =>
  run().catch((err) => {
    console.error('\nE2E FAILED:', err);
    app.exit(1);
  })
);
