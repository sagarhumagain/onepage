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

**Paste** keeps structure, and from Word it keeps the look too. When the
clipboard carries HTML from a browser, OnePage walks it for headings, lists,
tables, links and emphasis, and throws away every font and colour: a web
page's CSS is the site's design, not the author's. The result is plain text
you can edit in the left pane, using conventions the parser already
understands (`#`, `-`, `1.`, `>`, `|` tables, `**bold**`, `[text](url)`).

From Word the formatting *is* the author's, so it is kept: a shaded title bar
drawn as a shape, a "Key facts" text box, ticks nested under a bullet, a table
with a shaded label column and lists inside its cells, column widths, black
rules, sizes, colours, alignment and the logos. The text still arrives as
editable source; everything a reader would call styling arrives as marks — the
same content-anchored formatting the format bar writes — so a pasted fact sheet
looks like the one that was copied and can be edited like anything typed. A
document pasted into an empty page also brings its typeface and its single
column with it.

**Fit** is a binary search, not a zoom. Every size in `document.css` derives
from one custom property, `--doc-font-pt`, so a single number rescales the whole
document. The engine searches for the largest value that does not overflow a
fixed 210×297mm box. The result is a *style budget* — real point sizes — because
`transform: scale()` and `zoom` cannot be represented in a Word file at all, so
a scaled preview could never be exported.

**Images** are placed, not appended. Pasting a picture, or clicking *Image*,
drops a reference on a line of its own at the cursor — `[image:a1b2c3]` — so it
is a block of the document like any paragraph, with text above it and text
below it. Moving an image is moving that line. Click the picture in the page to
size it by its corner handles, align it, or let the text wrap around it; Esc
puts it back if a drag went wrong. Width is a percentage of the column, so it
means the same thing in one column as in three.

**The page is the editor.** Type straight into it and the edited page is walked
back into source text — by the same walker that reads pasted Word HTML, so
there is only one place where markup becomes text — and the left pane updates
as you go. *Insert* in the header drops in a table, a bullet or numbered list,
a heading, a quote, a divider or a picture at the cursor; each one is built
from the same conventions the source pane uses, so an inserted table is
identical to a pasted one.

**Formatting** works on a selection in the page, not on the source. Select any
run of words and a bar appears over it: bold, italic, bigger, smaller, text
colour, highlight, **band** and clear. A highlight hugs the words; a band fills
the whole width of the line, which is what a section header bar is. Sizes are
multipliers, never point sizes — the fitter owns absolute size, so an
emphasised line grows and shrinks with the rest of the page and the one-page
guarantee still holds. Formatting is stored against the words it was applied to
rather than against a position, so editing another paragraph does not slide
every colour onto the wrong sentence.

**One history covers all of it.** Ctrl/Cmd+Z steps back through typing, an
inserted table, a resized picture, a colour — and Ctrl/Cmd+Shift+Z forward
again. The browser's own undo for the editable page is deliberately taken over:
it would put DOM back without putting the model back.

**Output** comes from the same DOM you are looking at. The preview is an iframe
containing only the page and `document.css`; printing and PDF export render that
same markup, with the same stylesheet, through Chromium. There is no second
layout pass that could disagree with the first.

### When it does not fit

The fitter climbs a ladder, most readable rung first, and only moves down when
the rung above genuinely failed:

| # | Columns | Leading | Margins | Images | Size range |
|---|---------|---------|---------|--------|------------|
| 1 | 1 | 1.38 | as set | 100% | 9 – 13.5pt |
| 2 | 2 | 1.38 | as set | 100% | 9 – 13.5pt |
| 3 | 3 | 1.38 | as set | 100% | 9 – 13.5pt |
| 4\* | 3 | 1.38 | as set | 82% | 9 – 13.5pt |
| 5 | 3 | 1.38 | as set | 82% | 7 – 9pt |
| 6 | 3 | 1.16 | as set | 66% | 7 – 9pt |
| 7 | 3 | 1.16 | 10mm | 66% | 7 – 9pt |
| 8 | 3 | 1.16 | 7mm | 50% | 2.5 – 7pt |
| 9 | 3 | 1.16 | 7mm | 35% | 0.4 – 2.5pt |

\* Rungs that scale images exist only when the document contains one; without
pictures this is the same four-rung ladder it always was. A photograph at 82%
still reads, and 7pt body text beside a full-size photograph does not — so the
picture gives first.

The last rung is there because "this does not fit" is not an outcome the engine
is allowed to choose while there is any size left to try. One page is the whole
product; the status bar says how small it had to go, and the sheet is still one
sheet.

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
ONEPAGE_DEV_URL=http://localhost:5183 npx electron test/interaction.js
ONEPAGE_DEV_URL=http://localhost:5183 npx electron test/word-paste.js
npx electron test/smoke-prod.js        # same checks against the built bundle
```

`e2e.js` proves the fitter; `interaction.js` proves the editing gestures, which
cannot be checked anywhere else because each one is a DOM selection, a pointer
drag and a re-render. It places an image between two paragraphs, drags a corner
handle, cancels a drag with Esc, formats a selection, bands a heading, types
into the page and checks the source pane received it, formats a line that was
just typed, undoes and redoes, inserts a table from the header and reads it
back out of the source, and then checks that a full page of text plus a large
picture is still a single A4 sheet — with the .docx unzipped and its XML
inspected for the image size, the text wrap and the shading.

`word-paste.js` puts Word-shaped HTML and RTF on the system clipboard, pastes
it into the app — into the source pane and into the page — and checks the
document, the page, the PDF and the .docx that come out. It borrows the
clipboard and puts back the text that was on it.

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
| `src/parse.js` | text → document blocks (headings, nested lists, tables, quotes, images) |
| `src/clipboard.js` | clipboard HTML → source text; Word's formatting → marks |
| `src/word.js` | Word's clipboard HTML made ordinary: VML shapes, list paragraphs, class styles |
| `src/rtf.js` | the pictures in an RTF clipboard flavour |
| `src/inline.js` | inline tokens, shared by the HTML and Word renderers |
| `src/marks.js` | formatting on a run of words, anchored to the words |
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

**The edited page is never diffed back into the model.** Typing into the
contenteditable page serialises the whole page to source text through
`htmlToText` — the clipboard walker, run in the other direction — and that text
is the document from then on. A second serialiser, or a DOM-to-model patcher,
would be a second place for the two representations to disagree. The page is
only re-rendered from the text when something actually needs the model (a
format, a click away, an export), which is what stops the structure
rearranging itself under the caret mid-sentence.

**A band is paragraph shading, not a highlight.** A background on a run of text
stops where the letters do, whatever its width; only a background on the block
itself reaches the edge of the column. The two are different properties, and in
Word they are different elements — run shading and paragraph shading.

**A float is invisible to a sentinel that does not clear it.** A wrapped image
can hang below the last line of text; an uncleared end-of-content marker sits
above it and reports that everything fits while the picture is being cut off.
The sentinel is `clear: both`.

**An image must shrink when the type does, or the type shrinks alone.** Image
widths are a percentage of the column multiplied by `--doc-image-scale`, which
is the fitter's last lever before it goes under the readable floor. Without it a
page that is mostly picture ends at 3pt text beside a full-size photograph.

**A mark cannot be stored as a position.** "Block 4, characters 10-25" is wrong
the moment a word is inserted above it. Formatting remembers the plain text of
the run it was on, which occurrence of that text it was, and the marked
substring, and re-anchors in that order; a mark that cannot find its words is
dropped rather than applied to text nobody chose.

**Word measures images in pixels, not twips.** `ImageRun`'s `transformation`
is CSS pixels (1px = 15 twips). Passing a twip width straight through made a
column-wide picture a hundred inches across — the file opened, so nothing
failed loudly.

**`lineRule: EXACTLY` clips an inline image to the line box.** The document
default is exact leading at body-text size, which slices every picture down to
one line of height. Image paragraphs, and only those, use `AT_LEAST` at the
picture's own height.

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

**Half of a Word document is in comments.** Text boxes, shapes and floating
pictures arrive as VML inside `<!--[if gte vml 1]>` conditional comments, so
to a browser a title bar or a "Key facts" box simply is not there; the
fallback Word offers instead is a picture of the box at a file path no page
can read. word.js parses the comments and puts their content back into the
document, ordered top to bottom the way a reader meets it.

**Word's pictures are only in the RTF.** The HTML flavour points at temporary
files in Office's own sandboxed container. The RTF flavour Word writes beside
it carries the same pictures as hex PNG, in the same order — and a WMF copy of
each for old readers, which is skipped.

**A line break is a character, not an element.** A `<br>` has no text, so
every offset after one — which is what a selection, a mark and the caret are
measured in — would be one short. A hard break is a preserved newline in a
span of its own, and counts as the one character it is in the source.

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

Formatting survives as it should: a highlight becomes run shading (not Word's
`highlight` attribute, which snaps any colour to one of sixteen marker pens), a
size becomes the same multiple of the fitted size that the preview uses, and a
wrapped image becomes a floating drawing with square text wrapping.

That is the honest position: the PDF and the print output are exact, and the
`.docx` is a faithful, deliberately conservative reproduction. If a document
must be pixel-identical, send the PDF.

---

## Licence

MIT.
