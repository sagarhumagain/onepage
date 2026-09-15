# OnePage

Paste anything. Get exactly one A4 page. Print it, or save it as PDF or Word.

Content length varies every time — a two-line note, a three-thousand-word
report — and OnePage rebuilds it as a clean document and scales the typography
until it lands on a single sheet. Nothing is ever silently cut.

Runs on macOS and Windows.

---

## How it works

```
paste  ->  parse  ->  render  ->  fit  ->  print / PDF / Word
```

**Paste** keeps structure, not styling. When the clipboard carries HTML (Word,
Google Docs, a browser) OnePage walks it for headings, lists, tables and
emphasis, and throws away every font, colour and `mso-` artefact. The result is
plain text you can edit in the left pane, using conventions the parser already
understands (`#`, `-`, `1.`, `>`, `|` tables, `**bold**`).

**Fit** is a binary search, not a zoom. Every size in `document.css` derives
from one custom property, `--doc-font-pt`, so a single number rescales the whole
document. The engine searches for the largest value that does not overflow a
fixed 210×297mm box. The result is a *style budget* — real point sizes — because
`transform: scale()` and `zoom` cannot be represented in a Word file at all, so
a scaled preview could never be exported.

**Output** comes from the same DOM you are looking at. The preview is an iframe
containing only the page and `document.css`; printing and PDF export render that
same markup, with the same stylesheet, through Chromium. There is no second
layout pass that could disagree with the first.

### When it does not fit

The fitter climbs a ladder, most readable rung first, and only moves down when
the rung above genuinely failed:

| # | Columns | Leading | Margins | Size range |
|---|---------|---------|---------|------------|
| 1 | 1 | 1.38 | as set | 9 – 13.5pt |
| 2 | 2 | 1.38 | as set | 9 – 13.5pt |
| 3 | 3 | 1.38 | as set | 9 – 13.5pt |
| 4 | 3 | 1.38 | as set | 7 – 9pt |
| 5 | 3 | 1.16 | as set | 7 – 9pt |
| 6 | 3 | 1.16 | 10mm | 7 – 9pt |
| 7 | 3 | 1.16 | 7mm | 2.5 – 7pt |

Short content is handled the other way round: leftover room is spent on leading
before type size, and growth stops at 13.5pt. Past that an A4 page stops reading
as a document and starts reading as a large-print flyer.

The status bar always reports what it had to do — the fitted size, the column
count, the margins actually used, and a warning when the type went below a
comfortable reading size.

---

## Running it

```sh
npm install
npm start          # dev server + app, with live reload
```

Other scripts:

```sh
npm test           # unit tests: parsing, tokenising, CSS/DOCX agreement
npm run build      # bundle the renderer into dist/
npm run dist:mac   # .dmg  (arm64 + x64)
npm run dist:win   # .exe installer (NSIS)
```

### Verification

The interesting claim is not "the preview looks right", it is "Chromium's print
pipeline emits exactly one A4 page". So the end-to-end test renders each case,
exports a real PDF and measures it with `pdfinfo`:

```sh
npm run build && npx vite &            # dev server on :5183
ONEPAGE_DEV_URL=http://localhost:5183 npx electron test/e2e.js
npx electron test/smoke-prod.js        # same checks against the built bundle
```

Current results — every case one page, at true A4, with selectable text and no
content lost:

```
case                      pt    cols  lh     pages  pdf text
tiny (8 words)            13.5    1  1.55       1       47
short (120 words)         13.5    1  1.55       1     1289
medium (400 words)       10.04    1  1.38       1     4310
long (900 words)          7.03    3  1.38       1     9742
very long (1800 words)    5.68    3  1.16       1    19348
extreme (3500 words)      3.89    3  1.16       1    37524
absurd (6000 words)       2.89    3  1.16       1    50615
mixed structure           9.81    1  1.38       1     3360
unbreakable url          13.24    1  1.38       1     2493
```

---

## Layout of the code

| File | Responsibility |
|------|----------------|
| `src/parse.js` | text → document blocks (headings, lists, tables, quotes) |
| `src/inline.js` | inline tokens, shared by the HTML and Word renderers |
| `src/render.js` | blocks → HTML, escaping everything on the way in |
| `src/fit.js` | the one-page engine |
| `src/document.css` | the page itself; the only stylesheet that reaches print |
| `src/scale.js` | the typographic scale, shared by CSS and Word |
| `src/to-docx.js` | blocks → `.docx` at the fitted size |
| `src/sheet.js` | the A4 iframe, and the standalone file that gets printed |
| `src/exporters.js` | print / PDF / Word, with browser fallbacks |
| `electron/main.js` | window, save dialogs, Chromium's print pipeline |

The renderer imports nothing from Electron. It runs in a plain browser, which is
also how it is developed, and the desktop shell is a four-call preload bridge.

---

## Things that are true and non-obvious

Each of these was a bug first, and each is load-bearing.

**`column-count: 1` is not "one column".** It still makes the element a
multi-column container, and a multicol container with a definite height does not
grow downwards when it runs out of room — it fragments *sideways* into further
columns. A vertical overflow probe on such a box reads "fits" forever. Single
column therefore uses `column-count: auto`, an ordinary block box.

**`scrollWidth` does not see spilled columns.** It reads equal to `clientWidth`
whether the content fits in three columns or needs five. Overflow is measured
with a zero-size sentinel appended after the content: its inline position *is*
the column the content ended in.

**`scrollHeight` is an integer** and ignores a collapsed bottom margin on the
last child. Either flaw hides the one pixel that becomes a second sheet, which
is why the vertical probe is `getBoundingClientRect()` on the same sentinel.

**The fit is not strictly monotone.** `text-wrap: pretty`, `hyphens: auto` and
`break-inside: avoid` mean a slightly *smaller* font can occasionally produce a
*taller* box, because the wrap points move. The winner of each bisection is
re-measured and stepped down if the search was fooled.

**Chromium defaults to US Letter** in `printToPDF` and rescales content to fit
it. Both `pageSize: 'A4'` and `preferCSSPageSize: true` are set, and the page
declares `@page { size: A4; margin: 0 }`, so all three agree.

**A stylesheet cannot be `fetch()`ed in dev.** The dev server answers a request
for one with a JavaScript module that installs it, so the exported file would
carry JS in its `<style>` tag and render unstyled. `document.css` is imported
with `?raw` instead, which also guarantees the preview and the export are the
same bytes.

**`docx`'s millimetre helper truncates.** `convertMillimetersToTwip(210)`
returns 11905, not the 11906 Word itself writes; on a borderline layout that one
twip is a second page. The A4 literals are hard-coded.

**Word half-points round down, deliberately.** Rounding to nearest turns a
solved 9.76pt into 10.0pt — 2.5% larger than the size measured to fit, spending
the very slack the guarantee depends on.

---

## The Word export

Word has no "fit to one page" feature and OOXML has no element that means it:
Word re-runs its own line breaking and pagination. The export therefore removes
every source of disagreement it can — exact A4 geometry, `lineRule="exactly"` so
a line box cannot grow for a tall glyph, widow/orphan control off, fixed table
layout, and a font named identically to the one the preview measured — and then
exports 3% below the fitted size to absorb what is left.

That is the honest position: the PDF and the print output are exact, and the
`.docx` is a faithful, deliberately conservative reproduction. If a document
must be pixel-identical, send the PDF.

---

## Licence

MIT.
