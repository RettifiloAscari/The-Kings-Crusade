# The King's Crusade — Project Instructions

> **STATUS: BOOTSTRAPPED, NOT STARTED.** The document pipeline is installed and proven.
> **The campaign itself does not exist yet** — no theme, no setting, no canon, no NPCs.
>
> **Before generating any content, read `drafts/NEW-CAMPAIGN-HANDOFF.md` in full**, then
> walk the user through the decisions in its §7 as short multiple-choice questions, one at
> a time. **Do not invent the setting.** The premise, tone, historical spine, and document
> set are the user's to define, and guessing at them will produce work that has to be
> thrown away.
>
> Once those decisions are made, replace this banner and the placeholder sections below
> with the real thing, using the Qilvayas Symphony `CLAUDE.md` as the structural model
> (handoff §8 lists the sections that worked).

---

## The Ruleset — Confirm This First

The Qilvayas Symphony, which this pipeline comes from, is built entirely on the **2014**
D&D 5th Edition ruleset, and the SRD validation here reads `src/2014/en/...` on purpose.
"D&D 5e" is ambiguous now that the 2024 revision exists.

**Pin this before writing a single stat block** — it is decision #1 in the handoff's §7.
If this campaign stays on 2014, replace this section with the standing constraint: *race*
not *species*, no weapon masteries, Bastions, Epic Boons, Heroic Inspiration, revised
exhaustion, or reworked grapple rules, and monster math from the 2014 DMG. If it moves to
2024, that is a different SRD path, a different monster table, and a different vocabulary —
decide once, up front, because retrofitting is a migration across every stat block and
every player-facing document.

## Repository Layout

The generator scripts are the source of truth. Everything else in this repository is
either input to them or output from them, with the single exception of `images/`.

- `scripts/` — **the canon.** The docx-js generators, `transplant.py`, and the
  base64-encoded visual template. Editing canon means editing a script here.
- `corpus/` — **generated Markdown**, one file per document. Diffable and greppable.
  Every file carries a DO-NOT-EDIT banner.
- `documents/` — **generated PDF**, styled and ready to read. Committed deliberately,
  against the usual rule about build output, so the corpus opens on any device without a
  build step. PDF embeds its fonts, so it renders identically everywhere.
- `tools/` — `build.sh` (regenerates everything and verifies it), `docx-md-shim/`
  (a stub of the `docx` package that emits Markdown instead of OOXML), and
  `normalize_pdf.py` (makes the rendered PDF byte-reproducible).
- `drafts/` — design drafts awaiting sign-off. Not canon; never read as canon.
- `reference/` — mirrored instructions, if a Claude Chat project reads this repo.
- `images/` — artwork, named for what it depicts. **The one directory the build does not
  touch.** Artwork cannot currently reach the published documents: the generators emit
  `.docx` via docx-js and the Markdown shim has no image path. If this campaign wants art
  in its documents, build that support *before* writing content — retrofitting it across
  seven generators is far worse than adding it to one.
- `README.md` — repository index.

`corpus/` and `documents/` come from the same untouched scripts in the same build, so they
cannot drift apart. The `.docx` is a build intermediate and is never committed.

## Working in This Repository

These rules are theme-independent and apply from the first commit.

- **Read before you write.** Ground yourself in `corpus/` at the start of any task. Do not
  generate against remembered canon — check the file.
- **Never hand-edit generated output.** Nothing in `corpus/` or `documents/` is ever
  edited directly; the next build silently discards it. Edit the script that produces it.
- **Regenerate in the same commit as the script change.** Run `tools/build.sh` and commit
  `scripts/`, `corpus/`, and `documents/` together. A corpus that disagrees with its
  generator is worse than no corpus.
- **Canon changes propagate in one pass.** A change to the sourcebook that affects a
  reference guide, a player guide, or a session module updates all of them in the same
  commit.
- **Large passes go through a draft first.** Any systemic layer, region set, NPC roster, or
  mechanical conversion gets a design draft in `drafts/` with per-item sign-off flags, an
  explicit split between natural extension and new invention, a checklist, and a
  propagation plan. Only after approval does it fold into canon in one consolidated pass.
  Rename the draft `*.RESOLVED.md` afterward with a header recording what was decided.
- **Commit at meaningful boundaries** — one consolidated pass per commit. Write commit
  messages describing what changed in the fiction, not the mechanics of the edit.

## Production Practice

**One command rebuilds everything: `tools/build.sh`.** It regenerates `corpus/` and
`documents/`, applies the template, renders each document to PDF, and **fails the build on
escape leaks or font substitution** — the two errors that are invisible in source.

Prerequisites, all checked by the build:

| Requirement | Install | Why |
|---|---|---|
| Node + `docx` | `npm install docx` | The generators |
| Python 3 | stdlib only | `transplant.py`, `normalize_pdf.py` |
| LibreOffice **Writer** | `apt-get install libreoffice-writer` | `libreoffice-core` alone loads *nothing* and every document fails with "source file could not be loaded" |
| Ghostscript | `apt-get install ghostscript` | Reproducible PDF |
| poppler-utils | `apt-get install poppler-utils` | `pdftotext`, `pdffonts`, `pdftoppm` for verification |
| Alegreya SC, Alegreya Sans SC, Lato | Google Fonts TTFs into `~/.local/share/fonts`, then `fc-cache -f` | Missing fonts **substitute silently and change pagination**; verifying layout without them is meaningless |

**Verification, every time:**

```bash
grep -Pc '[^\x00-\x7F]' scripts/*.js                  # MUST be 0 — no literal non-ASCII
pdftotext documents/<doc>.pdf - | grep -c '\\u'       # MUST be 0 — no escape leaks
pdffonts documents/<doc>.pdf | grep -c DejaVu         # MUST be 0 — no font substitution
```

**Note the doubled backslash** in the escape check. Single-quoted `'\u'` matches the plain
letter *u* and reports every ordinary word containing one, so it can never return 0 on real
prose. Then build twice and `cmp` — an unchanged document must rebuild byte-identical.

**All prose lives in the generators as `\uXXXX` escapes**, never as literal typographic
characters. Normalize immediately after inserting new text, then re-run the greps.

**`node --check` is not sufficient.** It validates syntax, not identifiers — a call to a
helper that file does not define passes `--check` and throws at build time. Helper sets
genuinely differ between generators. Grep for the definition, and actually run the script.

**DM-only markers are bold book-red, never italic:** `const DM = (t) => ({ t, b: true,
c: "5B1F1F" })`, used as `PS([DM("DM Only: "), { t: "the note." }])`. Colour the marker,
not the prose. Italic is reserved for read-aloud, quotations, and epigraphs.

## Current State

- `scripts/smoke.js` — a throwaway one-page document exercising every rendering path.
  **It is not canon.** Delete it, and its `corpus/` and `documents/` output, once real
  generators build clean. Until then it is what keeps `tools/build.sh` runnable.
- `tools/build.sh` — `GENERATORS=(smoke)` and `SINGLE_COL_MATCH="__NONE__"` are the two
  lines to update as real documents arrive.
- Everything else in `scripts/` and `tools/` is the campaign-agnostic pipeline, carried
  over intact from The Qilvayas Symphony.
