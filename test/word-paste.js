/**
 * word-paste.js — a Word document pasted into the real app keeps its look.
 *
 * Word's clipboard is not ordinary HTML: text boxes and pictures are VML in
 * conditional comments, lists are paragraphs with their markers in hidden
 * spans and their indents in the stylesheet, and the pictures themselves are
 * only in the RTF flavour. None of that can be checked without a browser and a
 * real paste, so this puts Word-shaped HTML and RTF on the system clipboard,
 * pastes it into the app, and checks the document, the page, the PDF and the
 * .docx that come out.
 *
 * The fixture mirrors the structures of a real one-page Word fact sheet: a
 * shaded title bar in a shape, a text box with a bulleted list and ticks
 * nested under it, a table with a shaded label column and lists in its cells,
 * and a letterhead of a logo and four lines beside a second logo.
 *
 * It uses the system clipboard and puts back the text that was on it.
 *
 * Usage: ONEPAGE_DEV_URL=http://localhost:5183 npx electron test/word-paste.js
 */

import { app, BrowserWindow, clipboard, ClipboardItem } from 'electron';
import { mkdirSync, writeFileSync } from 'node:fs';
import { inflateRawSync } from 'node:zlib';
import os from 'node:os';
import path from 'node:path';

const DEV_URL = process.env.ONEPAGE_DEV_URL || 'http://localhost:5183';
const OUT = path.join(os.tmpdir(), 'onepage-word-paste');

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok: Boolean(ok), detail });
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? `\n          ${detail}` : ''}`);
};

/* --- The fixture: what Word puts on the clipboard -------------------------- */

const STYLE = `<style><!--
@page WordSection1 {size:612.0pt 792.0pt; margin:0cm 36.0pt 18.0pt 72.0pt;}
p.MsoNormal, li.MsoNormal, div.MsoNormal {margin-top:0cm; margin-bottom:8.0pt; font-size:11.0pt; font-family:"Calibri",sans-serif;}
p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph {margin:0cm 0cm 0cm 36.0pt; font-size:12.0pt; font-family:"Times New Roman",serif;}
p.MsoListParagraphCxSpFirst, li.MsoListParagraphCxSpFirst {margin:0cm 0cm 0cm 36.0pt; font-size:12.0pt; font-family:"Times New Roman",serif;}
p.MsoListParagraphCxSpMiddle, li.MsoListParagraphCxSpMiddle {margin:0cm 0cm 0cm 36.0pt; font-size:12.0pt; font-family:"Times New Roman",serif;}
@list l1:level1 {mso-level-number-format:bullet; margin-left:18.0pt; font-family:Symbol;}
@list l2:level1 {mso-level-number-format:bullet; margin-left:54.0pt; font-family:Wingdings;}
@list l3:level1 {margin-left:18.0pt;}
@list l4:level1 {mso-level-number-format:bullet; font-family:Symbol;}
--></style>`;

const T = `font-family:"Times New Roman",serif`;
const marker = (glyph, font) =>
  `<![if !supportLists]><span style='font-size:9.0pt;${font ? `font-family:${font}` : ''}'><span style='mso-list:Ignore'>${glyph}<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp; </span></span></span><![endif]>`;
const item = (list, margin, glyph, font, text) =>
  `<p class=MsoListParagraphCxSpMiddle style='${margin}text-align:justify;text-indent:-18.0pt;mso-list:${list} level1 lfo1'>${marker(glyph, font)}<span style='font-size:9.0pt;${T}'>${text}</span></p>`;

const TEXTBOX = (style, attrs, inner) =>
  `<!--[if gte vml 1]><v:shape type="#_x0000_t202" style='${style}' ${attrs}><v:textbox><![if !mso]><table cellpadding=0 cellspacing=0 width="100%"><tr><td><![endif]><div>${inner}</div><![if !mso]></td></tr></table><![endif]></v:textbox></v:shape><![endif]--><![if !vml]><img width=500 height=50 src="file:////tmp/msohtmlclip/clip_image001.png" alt="Text Box"><![endif]>`;

const HTML = `<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta name=ProgId content=Word.Document>${STYLE}</head>
<body lang=EN-US link="#0563C1"><!--StartFragment-->
<p class=MsoNormal>${TEXTBOX(
  'position:absolute;margin-left:-36pt;margin-top:52pt;width:540pt;height:90pt;mso-position-horizontal-relative:margin;mso-position-vertical-relative:text',
  'stroked="f"',
  `<p class=MsoNormal style='margin-bottom:0cm'><b><u><span style='font-size:8.5pt;${T}'>Key facts:</span></u></b></p>` +
    item('l1', 'margin-left:18.0pt;', '·', 'Symbol', 'Surveillance is weekly, and <b>every case</b> is reported.') +
    item('l2', 'margin-left:54.0pt;', 'ü', 'Wingdings', 'Suspected') +
    item('l2', 'margin-left:54.0pt;', 'ü', 'Wingdings', 'Confirmed') +
    item('l1', 'margin-left:18.0pt;', '·', 'Symbol', 'There is no vaccine in routine use.')
)}${TEXTBOX(
  'position:absolute;margin-left:-36pt;margin-top:7pt;width:555pt;height:43pt;mso-position-horizontal-relative:text;mso-position-vertical-relative:text',
  'fillcolor="#538135 [2409]" stroked="f"',
  `<p class=MsoNormal align=center style='text-align:center'><span style='font-size:28.0pt;${T}'>Outbreak | Field Guide</span></p>`
)}<span><o:p></o:p></span></p>
<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=726 style='border-collapse:collapse'>
 <tr>
  <td width=102 style='width:76.5pt;border:solid black 1.0pt;background:#9BBB59'>
   <p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:9.0pt;${T};color:black'>Spread</span></b></p></td>
  <td width=624 valign=top style='width:468.0pt;border:solid black 1.0pt'>
   ${item('l4', 'margin-left:18.0pt;', '·', 'Symbol', 'Spread is mainly by the bite of an infected mosquito.')}</td>
 </tr>
 <tr>
  <td width=102 style='width:76.5pt;border:solid black 1.0pt;background:#9BBB59'>
   <p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:9.0pt;${T};color:black'>Signs</span></b></p></td>
  <td width=624 valign=top style='width:468.0pt;border:solid black 1.0pt'>
   <p class=MsoNormal style='margin-bottom:0cm;text-align:justify'><b><span style='font-size:9.0pt;${T}'>Incubation:</span></b><span style='font-size:9.0pt;${T}'> 4-10 days.</span></p>
   <p class=MsoNormal style='text-align:justify'><span style='font-size:9.0pt;${T}'>There are three phases:</span></p>
   ${item('l3', 'margin-left:18.0pt;', '1.', '', '<b>Febrile:</b> fever for 2-7 days.')}
   ${item('l3', 'margin-left:18.0pt;', '2.', '', '<b>Critical:</b> <i>watch for warning signs</i> closely.')}
   ${item('l4', '', '·', 'Symbol', 'Platelets fall.')}</td>
 </tr>
</table>
<p class=MsoNormal>${TEXTBOX(
  'position:absolute;margin-left:-35pt;margin-top:635pt;width:567pt;height:67pt;mso-position-horizontal-relative:margin;mso-position-vertical-relative:margin',
  'stroked="f"',
  `<p class=MsoNormal><span style='font-size:7.0pt;${T}'>Reference:</span></p>` +
    `<p class=MsoListParagraphCxSpFirst style='margin-left:9.0pt;text-indent:-9.0pt;mso-list:l3 level1 lfo2'>${marker('1.', '')}<span style='font-size:7.0pt;${T}'>Field manual [</span><a href="https://example.org/manual"><span style='font-size:7.0pt'>Internet</span></a><span style='font-size:7.0pt;${T}'>]. 2025.</span></p>`
)}<!--[if gte vml 1]><v:group style='position:absolute;margin-left:-40pt;margin-top:80pt;width:265pt;height:70pt' coordsize="33733,8839"><v:shape type="#_x0000_t202" style='position:absolute;left:8734;top:1023;width:24999;height:7233' fillcolor="white [3201]" stroked="f"><v:textbox><![if !mso]><table cellpadding=0 cellspacing=0 width="100%"><tr><td><![endif]><div>${[
  'Ministry of Health',
  'Department of Health Services',
  'Curative Service Division',
]
  .map((l) => `<p class=MsoNormal style='margin-bottom:0cm'><span style='font-size:10.0pt;${T};color:red'>${l}</span></p>`)
  .join('')}</div><![if !mso]></td></tr></table><![endif]></v:textbox></v:shape><v:shape alt="Emblem" style='position:absolute;width:8756;height:8839'><v:imagedata src="file:////tmp/msohtmlclip/clip_image004.png" o:title="Emblem"/></v:shape></v:group><![endif]--><![if !vml]><span><img width=268 height=69 v:shapes="Group_x0020_1"></span><![endif]><!--[if gte vml 1]><v:shape style='position:absolute;margin-left:379pt;margin-top:82pt;width:136pt;height:70pt;mso-position-horizontal-relative:text;mso-position-vertical-relative:text'><v:imagedata src="file:////tmp/msohtmlclip/clip_image005.png" o:title=""/></v:shape><![endif]--><![if !vml]><img width=137 height=70 src="file:////tmp/msohtmlclip/clip_image005.png"><![endif]><span style='font-size:8.0pt;${T};color:black'>*Report every case.</span></p>
<!--EndFragment--></body></html>`;

/** A small PNG in a solid colour, so each picture is recognisably itself. */
function png(width, height, rgb) {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc = (buf) => {
    let c = 0xffffffff;
    for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const out = Buffer.alloc(12 + data.length);
    out.writeUInt32BE(data.length, 0);
    out.write(type, 4, 'latin1');
    data.copy(out, 8);
    out.writeUInt32BE(crc(Buffer.concat([Buffer.from(type, 'latin1'), data])), 8 + data.length);
    return out;
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr.set([8, 2, 0, 0, 0], 8);
  const raw = Buffer.alloc((width * 3 + 1) * height);
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) raw.set(rgb, y * (width * 3 + 1) + 1 + x * 3);
  return Buffer.concat([
    Buffer.from('89504e470d0a1a0a', 'hex'),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflate(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

/** zlib's stored (uncompressed) form: no dependency, and exactly what a PNG reader expects. */
function deflate(raw) {
  const blocks = [];
  for (let i = 0; i < raw.length || i === 0; i += 65535) {
    const part = raw.subarray(i, i + 65535);
    const head = Buffer.alloc(5);
    head[0] = i + 65535 >= raw.length ? 1 : 0;
    head.writeUInt16LE(part.length, 1);
    head.writeUInt16LE(~part.length & 0xffff, 3);
    blocks.push(head, part);
  }
  let a = 1;
  let b = 0;
  for (const x of raw) {
    a = (a + x) % 65521;
    b = (b + a) % 65521;
  }
  const adler = Buffer.alloc(4);
  adler.writeUInt32BE(((b << 16) | a) >>> 0);
  return Buffer.concat([Buffer.from([0x78, 0x01]), ...blocks, adler]);
}

const EMBLEM = png(40, 40, [200, 30, 30]);
const LOGO = png(80, 40, [30, 120, 200]);

const blip = (bytes, w, h, uid) =>
  `{\\pict\\picscalex100\\picscaley100\\picw${w}\\pich${h}\\picwgoal${w * 20}\\pichgoal${h * 20}\\pngblip\\bliptag255{\\*\\blipuid ${uid}}${bytes.toString('hex')}}`;

// The emblem sits in the group; the second logo is a shape of its own. Word's
// RTF carries each as a PNG, and a WMF copy for old readers.
const RTF = `{\\rtf1\\ansi {\\shp{\\sp{\\sn pib}{\\sv ${blip(EMBLEM, 69, 70, 'aa01')}}}}{\\nonshppict{\\pict\\wmetafile8 0102}} {\\shp{\\sp{\\sn pib}{\\sv ${blip(LOGO, 136, 70, 'aa02')}}}}}`;

/* --- Driving the app --------------------------------------------------------- */

const ready = (win) =>
  win.webContents.executeJavaScript(
    `new Promise(r => { const t = setInterval(() => {
       if (document.documentElement.dataset.ready === 'true') { clearInterval(t); r(true); }
     }, 40); })`
  );

const run = (win, code) => win.webContents.executeJavaScript(`(async () => { ${code} })()`);

async function freshApp(win) {
  await win.webContents.executeJavaScript('localStorage.clear(); true');
  await win.reload();
  await ready(win);
}

async function paste(win, target) {
  await clipboard.write([new ClipboardItem({ 'text/html': HTML, 'text/rtf': RTF, 'text/plain': 'Key facts:' })]);
  win.focus();
  await run(
    win,
    target === 'page'
      ? `document.getElementById('sheet').contentDocument.getElementById('page-body').focus(); return true;`
      : `document.getElementById('source').focus(); return true;`
  );
  win.webContents.paste();
  await new Promise((r) => setTimeout(r, 1500));
}

const READ = `
  const o = window.__onepage;
  const doc = document.getElementById('sheet').contentDocument;
  const body = doc.getElementById('page-body');
  const style = (el) => el ? el.getAttribute('style') || '' : '';
  const title = body.querySelector('h1');
  const td = (r, c) => body.querySelectorAll('table tr')[r].children[c];
  const figures = [...body.querySelectorAll('.doc-image')];
  const { htmlToText } = await import('/clipboard.js');
  const { parse } = await import('/parse.js');
  const { render } = await import('/render.js');
  // Two sources are the same document when they draw the same page: "# Title"
  // and an inferred title, or one mixed list and two adjacent ones, differ
  // only in how they are written. Image ids are random per paste.
  const shape = (text) => render(parse(text.replace(/\\[image:[A-Za-z0-9_-]+\\]/g, '[image:x]'))).replace(/>\\s+</g, '><');
  const roundTrip = htmlToText(body.innerHTML);
  return {
    source: o.source(),
    roundTrip,
    shape: shape(o.source()),
    roundTripShape: shape(roundTrip),
    layout: o.layout(),
    typeface: document.getElementById('typeface').value,
    columns: document.getElementById('columns').value,
    order: [...body.children].filter((el) => el.getAttribute('aria-hidden') !== 'true').map((el) => el.tagName.toLowerCase()),
    title: title ? { text: title.textContent, cls: title.className, style: style(title), inner: title.innerHTML } : null,
    ticks: [...body.querySelectorAll('li[data-glyph]')].map((li) => li.dataset.glyph + ' ' + li.textContent),
    keyFacts: body.querySelector('p') ? body.querySelector('p').innerHTML : '',
    label: { text: td(0, 0).textContent, style: style(td(0, 0)) },
    table: { head: Boolean(body.querySelector('thead')), cols: [...body.querySelectorAll('col')].map(style), rule: style(body.querySelector('table')) },
    signs: td(1, 1).innerHTML,
    figures: figures.map((f) => ({ align: f.dataset.align, wrap: f.dataset.wrap, width: f.style.getPropertyValue('--img-w'), w: f.querySelector('img').naturalWidth })),
    letterhead: [...body.querySelectorAll('p')].map((p) => p.innerHTML).find((h) => h.includes('Ministry of Health')) || '',
    reference: [...body.querySelectorAll('li')].map((li) => li.outerHTML).find((h) => h.includes('Field manual')) || '',
    html: await o.sheetHtml(),
    docx: await o.docxBytes(),
  };
`;

async function main() {
  mkdirSync(OUT, { recursive: true });
  // Electron's clipboard calls are asynchronous.
  const saved = await clipboard.readText();

  const win = new BrowserWindow({ show: true, width: 1440, height: 900, webPreferences: { sandbox: true, contextIsolation: true } });
  await win.loadURL(DEV_URL);
  await ready(win);

  try {
    await freshApp(win);
    await paste(win, 'source');
    const got = await run(win, READ);
    writeFileSync(path.join(OUT, 'source.txt'), got.source);
    writeFileSync(path.join(OUT, 'round-trip.txt'), got.roundTrip);
    writeFileSync(path.join(OUT, 'shape.html'), got.shape.replace(/></g, '>\n<'));
    writeFileSync(path.join(OUT, 'round-trip-shape.html'), got.roundTripShape.replace(/></g, '>\n<'));

    console.log('\n  pasting a Word document\n  ' + '-'.repeat(70));
    check('the shapes come back as content, in the order a reader meets them', got.order.slice(0, 4).join(',') === 'h1,p,ul,table', got.order.join(' '));
    check(
      'the title keeps its bar: the shape fill as a band, white type, centred, regular weight',
      got.title && got.title.cls === 'doc-band' && /background:#538135/.test(got.title.style) && /text-align:center/.test(got.title.style) &&
        /color:#ffffff/.test(got.title.inner) && /font-weight:400/.test(got.title.inner),
      got.title && `${got.title.style} | ${got.title.inner.slice(0, 120)}`
    );
    check(
      'and its size relative to the body: 28pt over 9pt',
      got.title && /font-size:1\.64em/.test(got.title.inner),
      got.title && got.title.inner.slice(0, 80)
    );
    check('Key facts: bold and underlined', /<strong>/.test(got.keyFacts) && /text-decoration:underline/.test(got.keyFacts), got.keyFacts);
    check('the ticks are ticks, nested under the item above them', got.ticks.join('|') === '✓ Suspected|✓ Confirmed', got.ticks.join(' | '));
    check('the table has no header row, as in Word', !got.table.head);
    check('its column widths and black rules survive', got.table.cols.join(',') === 'width:14%,width:86%' && /--table-rule:#000000/.test(got.table.rule), `${got.table.cols} ${got.table.rule}`);
    check(
      'the label cell is shaded, centred and vertically centred',
      /background:#9bbb59/.test(got.label.style) && /text-align:center/.test(got.label.style) && /vertical-align:middle/.test(got.label.style),
      got.label.style
    );
    check(
      'a cell keeps its paragraphs and its numbered and bulleted lists',
      /<p[^>]*><strong>Incubation:<\/strong> 4-10 days\.<span class="doc-br">\n<\/span>There are three phases:<\/p>/.test(got.signs) &&
        /<ol><li[^>]*><strong>Febrile:<\/strong>/.test(got.signs) && /<ul><li[^>]*>Platelets fall\.<\/li><\/ul>/.test(got.signs),
      got.signs.slice(0, 300)
    );
    check('italic inside a numbered item survives', /font-style:italic|<em>watch for warning signs<\/em>/.test(got.signs));
    check(
      'both pictures arrive from the RTF, the second logo right and the emblem left, both floated',
      got.figures.length === 2 && got.figures[0].align === 'right' && got.figures[1].align === 'left' && got.figures.every((f) => f.wrap === '1'),
      JSON.stringify(got.figures)
    );
    check('each at its own pixels, so the right bytes went to the right place', got.figures[0].w === 80 && got.figures[1].w === 40, JSON.stringify(got.figures));
    check(
      'the letterhead is one block of red lines, not three headings',
      (got.letterhead.match(/doc-br/g) || []).length === 2 && /color:#ff0000/.test(got.letterhead),
      got.letterhead.slice(0, 160)
    );
    check('a link keeps its address and Word\'s link colour', /<a href="https:\/\/example\.org\/manual">Internet<\/a>/.test(got.reference) && /#0563c1/.test(got.reference), got.reference.slice(0, 200));
    check('a pasted document brings its typeface and its single column', got.typeface === 'times' && got.columns === '1', `${got.typeface} / ${got.columns}`);
    check('it fits one page in a single column', got.layout.columns === 1 && !got.layout.overflow, JSON.stringify(got.layout));

    check('typing on the page writes the same document back to the source', got.roundTripShape === got.shape, firstDifference(got.shape, got.roundTripShape));

    const pdf = await renderPdf(got.html);
    writeFileSync(path.join(OUT, 'word-paste.pdf'), pdf);
    const pages = pdfPageCount(pdf);
    check('the printed sheet is exactly one A4 page', pages === 1, `${pages} page(s)`);

    const docx = Buffer.from(got.docx);
    writeFileSync(path.join(OUT, 'word-paste.docx'), docx);
    const xml = readZipEntry(docx, 'word/document.xml') || '';
    const numbering = readZipEntry(docx, 'word/numbering.xml') || '';
    check('the .docx keeps the column widths', /<w:gridCol w:w="\d+"\/><w:gridCol w:w="\d+"\/>/.test(xml) && (() => {
      const [a, b] = [...xml.matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => Number(m[1]));
      return Math.abs(a / (a + b) - 0.14) < 0.01;
    })());
    check('and the shading, the vertical centring and the black rules', /w:fill="9BBB59"/.test(xml) && /<w:vAlign w:val="center"\/>/.test(xml) && /w:color="000000"/.test(xml));
    check('and the title bar as paragraph shading', /<w:shd [^>]*w:fill="538135"/.test(xml));
    check('and the ticks as a bullet level of their own', numbering.includes('w:val="✓"'));
    check('and both pictures, anchored right and left', (xml.match(/<wp:anchor/g) || []).length === 2 && /<wp:align>right<\/wp:align>/.test(xml) && /<wp:align>left<\/wp:align>/.test(xml));
    check('and the letterhead\'s line breaks as breaks, not paragraphs', (xml.match(/<w:br\/>/g) || []).length >= 2);

    console.log('\n  pasting into the page instead\n  ' + '-'.repeat(70));
    await freshApp(win);
    await paste(win, 'page');
    const page = await run(win, READ);
    writeFileSync(path.join(OUT, 'page-source.txt'), page.source);
    check('pasting into the page gives the same document as pasting into the source', page.shape === got.shape, firstDifference(got.shape, page.shape));
    check('with the same formatting on it', page.title && page.title.style === got.title.style && page.label.style === got.label.style);
  } finally {
    await clipboard.writeText(saved);
  }

  const failed = results.filter((r) => !r.ok);
  console.log(`\n  ${results.length - failed.length}/${results.length} passed\n  output: ${OUT}\n`);
  app.exit(failed.length ? 1 : 0);
}

/* --- Helpers, as in interaction.js ------------------------------------------- */

/** Where two strings part company, for a failure worth reading. */
function firstDifference(a, b) {
  if (a === b) return '';
  let i = 0;
  while (i < a.length && a[i] === b[i]) i++;
  return `at ${i}: ${JSON.stringify(a.slice(i - 20, i + 40))} vs ${JSON.stringify(b.slice(i - 20, i + 40))}`;
}

async function renderPdf(html) {
  const file = path.join(OUT, 'sheet.html');
  writeFileSync(file, html);
  const win = new BrowserWindow({ show: false, webPreferences: { sandbox: true, contextIsolation: true } });
  try {
    await win.loadFile(file);
    return await win.webContents.printToPDF({ pageSize: 'A4', preferCSSPageSize: true, printBackground: true, margins: { marginType: 'none' } });
  } finally {
    win.destroy();
  }
}

/** Page count without pdfinfo: the page tree's own /Count. */
function pdfPageCount(buffer) {
  const text = buffer.toString('latin1');
  const counts = [...text.matchAll(/\/Type\s*\/Pages\b[^>]*?\/Count\s+(\d+)/g)].map((m) => Number(m[1]));
  if (counts.length) return Math.max(...counts);
  return (text.match(/\/Type\s*\/Page\b(?!s)/g) || []).length;
}

/** One file out of a zip, without a zip library. */
function readZipEntry(buf, wanted) {
  let at = 0;
  while (at + 30 <= buf.length && buf.readUInt32LE(at) === 0x04034b50) {
    const method = buf.readUInt16LE(at + 8);
    const size = buf.readUInt32LE(at + 18);
    const nameLength = buf.readUInt16LE(at + 26);
    const extraLength = buf.readUInt16LE(at + 28);
    const name = buf.toString('utf8', at + 30, at + 30 + nameLength);
    const start = at + 30 + nameLength + extraLength;
    if (name === wanted) {
      const data = buf.subarray(start, start + size);
      return (method === 8 ? inflateRawSync(data) : data).toString('utf8');
    }
    at = start + size;
  }
  return null;
}

app.whenReady().then(() =>
  main().catch((err) => {
    console.error(err);
    app.exit(1);
  })
);
