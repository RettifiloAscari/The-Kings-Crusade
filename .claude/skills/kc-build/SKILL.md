---
name: kc-build
description: The King's Crusade production pipeline — editing the docx-js generators in scripts/, the escape and table conventions, the authoring kit, and the build and verification loop. Use whenever a task will touch scripts/, run tools/build.sh, add a document to the corpus, or debug a rendering problem in documents/. Not needed for lore questions, canon discussion, or design drafts that write no code.
---

# The King's Crusade — Production Pipeline

`CLAUDE.md` holds the canon and the campaign specs; this holds the mechanics. It is the
single copy of them — do not restate these rules in `CLAUDE.md`, because two copies drift.

## The one-paragraph model

Eighteen generators in `scripts/` are the source of truth. `tools/build.sh` runs each one
twice — once through `tools/docx-md-shim/` to emit Markdown into `corpus/`, once through
real docx-js to emit `.docx`, which then goes through `transplant.py` (the visual
template), LibreOffice (PDF), Ghostscript, and `normalize_pdf.py` (byte reproducibility).
The `.docx` is a build intermediate in `.stage/` and is gitignored. Nothing in `corpus/`
or `documents/` is ever hand-edited; the next build silently discards it.

## Order of work

1. **Read the corpus first.** Never generate against remembered canon.
2. Edit `scripts/` with `tools/anchor.py` — read, assert the count, replace.
3. `python3 tools/normalize_escapes.py scripts/*.js`
4. `tools/verify.sh` (add `--full` before committing)
5. Render the pages you changed and **look at them**.
6. Commit `scripts/`, `corpus/`, and `documents/` **together**.

## Prerequisites

All checked by the build, and every one has bitten this project.

| Requirement | Install | Why |
|---|---|---|
| Node + `docx` | `npm install docx` | The generators |
| Python 3 | stdlib only | `transplant.py`, `normalize_pdf.py` |
| LibreOffice **Writer** | `apt-get install libreoffice-writer` | `libreoffice-core` alone loads *nothing* and every document fails with "source file could not be loaded" — a content-shaped error with an environment cause. This container has shipped with core only. |
| Ghostscript | `apt-get install ghostscript` | Reproducible PDF |
| poppler-utils | `apt-get install poppler-utils` | `pdftotext`, `pdffonts`, `pdftoppm` |
| Alegreya SC, Alegreya Sans SC, Lato | TTFs into `~/.local/share/fonts`, then `fc-cache -f` | Missing fonts **substitute silently and change pagination**. `fonts.google.com/download` is blocked here — pull static TTFs from `raw.githubusercontent.com/google/fonts/main/ofl/{alegreyasc,alegreyasanssc,lato}/`. |

## Conventions that fail silently

### Compose in real characters, then normalize

Write prose with real typographic characters and run `tools/normalize_escapes.py`
afterward. Hand-escaping while composing is slow and is where the doubled-backslash bug
comes from: a run written `\\u2019` compiles clean, passes the non-ASCII scanner, and
leaks the literal text `’` into the PDF. The tool collapses a doubled escape,
converts non-ASCII to `\uXXXX`, and curls a straight apostrophe **only between two word
characters** — `(\w)'(\w)`, narrow on purpose, because any wider rule matches
`require('docx')` and `path.join(__dirname, '..', 'images')` and corrupts the generator.
It is idempotent; `--check` reports without writing. The collapse step caught seventeen
doubled escapes across four files during the expansion pass, and the narrow apostrophe
rule converted 176 without touching a single `require`.

Use `grep -Pl` (list offending files) rather than `-Pc` for the source scan: with
eighteen generators, `-Pc` prints a `file:count` line per script and there is no single
number to read.

### `node --check` is not sufficient

It validates syntax, not identifiers. A call to a helper the file does not define, or a
push onto the wrong document's array, passes `--check` and throws at build time. **Run
every generator.** `tools/verify.sh` does.

### Tables

- **`columnWidths` is not decoration.** docx-js emits per-cell `tcW` percentages but **no
  `<w:tblGrid>` unless `columnWidths` is given**, and without a grid LibreOffice discards
  the percentages and distributes every column evenly. `table()` passes `columnWidths`
  (twips, from the percentage `widths` against a 9026 nominal) and
  `layout: TableLayoutType.FIXED`. Supplying the grid tightened the whole set by two pages
  and is the single largest legibility win this pipeline has had. `TableLayoutType` must
  also be exported by `tools/docx-md-shim/` — **the shim must export everything the
  generators destructure from `docx`**, or the Markdown half throws while the PDF half
  succeeds.
- **Every table sits in the column flow. There are no full-width tables, deliberately.**
  `transplant.py` still carries the machinery — a `KCFullWidth` marker style — and nothing
  uses it. Do not reach for it. It was added because tables "wrapped to three or four words
  a line", but that was the missing `tblGrid`, not the column width. **If a table looks
  cramped, the fix is the widths, the header wording, or fewer columns — never breaking
  the column.**
- **Three columns, four at the outside.** Six-column encounter-scaling tables were the
  worst offenders and broke headers mid-word as `Multiplie`/`r`. The one six-column table
  left is in the DM Reference Guide, which is single-column and has the whole page.
- **A column must hold its longest unbreakable word.** Run `tools/check_columns.py`; it
  reports starved columns banded as **BREAKS** (certain), **likely** (render and look),
  and **marginal** (inside the measurement spread). It only fails the build on a certain
  break, because a check that fails on findings the reader will dismiss is a check the
  reader stops running. It splits on hyphens, because `Censor-Captain` wraps quite legally.
  Fix in this order: shorten the cell text; widen taking the difference from the **widest**
  column only (never redistribute proportionally to slack — prose columns have short words,
  measure as slack, and get gutted); or restructure.
- **Tables must not tear.** `row()` sets `cantSplit`, header rows carry `tableHeader` so
  they repeat across a break, and the ability-score row in `SB()` uses `keepNext` so values
  stay with their labels.
- **Prose after a table gets its gap from `transplant.py`, not the generator.** A table
  carries no space-after in OOXML, so prose immediately following one sat flush against its
  bottom border; headings always looked right because their style supplies
  `spacing.before`. `gap_after_tables()` gives the first paragraph after each table a
  180-twip before-gap, skipping headings, any paragraph whose author set a `before`
  deliberately, and blank spacer paragraphs — an empty paragraph is already the gap, and
  that test must be on the text content, since docx-js emits an empty run as `<w:t></w:t>`
  rather than omitting it. Fixed in `transplant.py` because `table()` returns a single
  Table, so a per-call fix would mean touching all seventy-one call sites forever after.
- **Cell padding is 60 twips** (left/right), and cells kill the inherited first-line indent
  with `indent: { firstLine: 0 }` — the template's default `firstLine=180` otherwise leaks
  in and every wrapped cell gets a ragged left edge. Padding lives in `margins`, *not* in
  an indent; the leaked indent looks like padding until a cell wraps.

### Bullets hang at 260 twips, not Word's 720

The numbering config's `indent: { left, hanging }` governs every bullet in the set, and the
docx-js default of `left: 720, hanging: 360` is sized for a 6.5in page. In this repository's
**3.28in column** that spends half an inch — about 15% of the measure — on indent, on every
line of every bullet, wrapped lines included. Bullet blocks visibly floated in a well beside
body paragraphs that used the full column.

`left: 260, hanging: 260` puts the glyph at the text margin and the text about one em in
(body text is 10pt, so an em is 200 twips). The five-entry block under *What They Cannot
Agree On* went from four rendered lines per entry to three.

One value is correct everywhere, because **every bullet in the set is in a two-column
document**: the DM Reference Guide, the only single-column one, is entirely tables and
contains no bullets at all. Check that before assuming a per-document value is needed.

**This is the third full-page default found applied to a narrow column**, after
`columnWidths` and cell padding. When something looks loose or cramped in a two-column
document, suspect a measurement inherited from a 6.5in page before suspecting the content.
The sister repository carries the identical 720/360 config at the same column measure, so
the same fix applies there and has not been made.

### DM markers are bold book-red, never italic

`const DM = (t) => ({ t, b: true, c: "5B1F1F" })`, used as
`PS([DM("DM Only: "), { t: "the note." }])`. Colour is preattentive and leaves the body
roman, which matters because these notes run 100–200 words. **Colour the marker, not the
prose.** Italic is reserved for read-aloud, quotations, and epigraphs. The Markdown shim
appends a space after every bold run, so the marker carries its own trailing space and the
following segment never begins with one. Sections already titled `(DM Only)` need no
inline marker.

### The authoring kit

The helper preamble at the top of each generator: `P`, `PS`, `DM`, `H1/H2/H3`, `BULLET`,
`B`, `BUL`, `BOX`, `cell/row/table`, `mod/abCell/SB`, `IMG`. Copy it into each new
generator.

- **Output path.** Never hardcode a stage directory:
  ```js
  const { stagePath } = require('./stage');
  fs.writeFileSync(stagePath("KC_Sourcebook.docx"), buf);
  ```
  `stagePath` resolves `$KC_STAGE` (set by `build.sh`) and falls back to `<repo>/.stage`
  so a generator runs standalone.
- **Artwork.** `IMG(file, widthPt, heightPt, alt)` places an image from `images/`. Points
  at 72/inch, so 288 is four inches. PNG or JPEG only, and real alt text on every one.
- **Every helper forwards its `opts`.** `BULLET` silently ignored its second argument for
  a whole build-out, which made a `keepNext` do nothing and look like a LibreOffice bug.
- **Headings carry `keepNext`**, so a heading can never sit alone at the foot of a column,
  and each module binds its final Loot bullet to the Refrain that follows — without it the
  closing verse strands itself on a blank last page, which three of eleven modules did.

### Edit with an anchor and an assertion

```bash
printf 'children.push(H1("The Muster"));\n' > /tmp/a.txt
tools/anchor.py before scripts/KC_Module01_TheMuster.js /tmp/a.txt /tmp/new.js
```

A silent zero-match quietly does nothing and the build still passes. `anchor.py` refuses
on any count mismatch and writes nothing. It strips a heredoc's trailing newline from the
replacement when the anchor is single-line — without that the newline lands inside a JS
string literal and breaks the file.

## Verification

`tools/verify.sh` runs every generator, the three escape greps, `check_columns.py`, the
build, a drift check, and the player-facing leak scan. `--full` adds three builds and a
byte comparison — the Ghostscript trailer-`/ID` failure mode is intermittent, so one
comparison is not enough.

**Reproducibility is within an environment, not across them.** `normalize_pdf.py` strips
per-run randomness (timestamps, trailer `/ID`, XMP UUIDs, font-subset tags); it does not
strip LibreOffice or Ghostscript *version* differences. A rebuild on a different toolchain
produces a different byte stream from identical sources with identical text, layout, page
count and fonts. If every `corpus/*.md` is unchanged and only PDFs differ, that is render
churn, not drift: **`git checkout -- documents/` rather than committing it.** `verify.sh`
says so when it sees that shape.

Then **render and look**. Rendering bugs are invisible in source and to every grep:

```bash
tools/find_page.py "the Tenth Work"
pdftoppm -r 110 -png -f 9 -l 9 documents/<doc>.pdf /tmp/page
```

For a precise answer on whether a word broke, read glyph extents rather than squinting:
`pdftotext -bbox -f N -l N doc.pdf -` lists every word with its box, so a word absent from
that list is a word that got split.

The escape check must use a **doubled backslash** — `grep -c '\\u'`. The single-quoted
`'\u'` form matches the plain letter *u* and can never return zero on real prose.

## Visual template

The /u/YaAlex-derived 5e style — Alegreya SC Medium headings in deep book-red `5B1F1F`,
Alegreya Sans SC and Lato body, A4, page-number footers. The book-red is load-bearing in
**both** `scripts/style_template_encoded.md` and every generator's `DM` constant; changing
it means changing both together. `transplant.py` carries **four** filename references to
the encoded template; renaming it means updating all four.

## Adding a document

1. New generator in `scripts/`, writing via `stagePath()`.
2. Add its basename to `GENERATORS=(...)` in `tools/build.sh`.
3. Add a row to the README index and to the document table in `CLAUDE.md`.
4. If it is player-facing, add it to `PLAYER_FACING` in `tools/pipeline.conf`.

## Reference

Everything repository-specific lives in `tools/pipeline.conf`. The five tools below are
byte-identical to their copies in The Qilvayas Symphony, so a fix to one is a copy rather
than a re-derivation.

- `tools/anchor.py` — assert-then-edit
- `tools/normalize_escapes.py` — real characters in, `\uXXXX` out; `--check` to report
- `tools/check_columns.py` — starved table columns, banded by confidence; `-v` for marginals
- `tools/find_page.py` — which PDF page a string is on
- `tools/check_tearing.py` — stat-block ability rows torn from their header
- `tools/verify.sh` — everything in one call; `--full` adds reproducibility
- `tools/build.sh` — the build; `--no-verify` only when the render toolchain is unavailable
