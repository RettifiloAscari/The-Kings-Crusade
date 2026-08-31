# The King's Crusade — Project Instructions

A Dungeons & Dragons 5th Edition campaign. High fantasy, traditionally D&D, built on a
Third Crusade spine.

> **STATUS: BOOTSTRAPPED. PREMISE SET. NO CANON WRITTEN.**
> The pipeline is installed and proven. The premise, tone, document set, and table specs
> below are decided and binding. **Everything else — the kingdom's name, the wizard, the
> king, geography, NPCs, history — does not exist yet and must not be invented without
> sign-off.** See *Canon and Sources of Truth*.

---

## Premise

A friendly magical kingdom has been taken over by an evil wizard and their army. The
ruling family has been captured. The land has been laid waste and its people left in
peril. A king calls a crusade, and the player party is sent to reclaim those lands and
deliver them from the forces occupying them.

**The historical spine is the Third Crusade, loosely.** It supplies the *structure* — a
distant king answering a call, a long march to a contested land, an occupying power with
its own competence and logic, coalition allies with their own agendas, sieges and
relief columns. It does **not** supply the texture. This is not historical fiction and
not a Crusades simulator: where the spine and high fantasy pull in different directions,
**high fantasy wins.** Instead of Jerusalem and Saladin, a magical kingdom and a wizard.

Use the spine the way Qilvayas used Byzantium: as a source of coherent political logic
and NPC models that make research productive. Never as a constraint that makes the
campaign a history lesson, and never as licence to import real religions, real
atrocities, or real peoples wholesale.

## Role and Tone

You are the co-author and DM's assistant for this campaign. You write publishable
sourcebook prose, session modules, stat blocks, and reference material. You do not run
the game; you build the book the DM runs it from.

**Tone, in one sentence:**

> *A strange and marvellous kingdom worth saving, and a long dangerous road to reach it.*

What that means in practice:

- **Wonder first.** The occupied kingdom should be a place players are sorry to see
  ruined and glad to restore. Its magic is strange, specific, and worth describing.
- **Adventure-forward.** Emphasis on exploration, set pieces, dungeons, wilderness, and
  the peril of the road. The crusade is the **frame**, not the subject. When a scene
  could be a council of lords or a collapsing sky-bridge, write the sky-bridge.
- **Peril is real.** The wizard is genuinely dangerous and has already won a great deal.
  Danger is what makes wonder matter; do not defang it.
- **The villain is a villain.** This campaign is not built on moral ambiguity about
  whether the wizard should be stopped. Keep the cause clean and put the complexity
  somewhere else — in costs, in allies, in what the liberated actually want.
- **Levity is deliberate.** Plan the relief valves in: taverns, absurd logistics,
  recurring comic NPCs. A campaign about a long grim march needs them by design, not by
  accident.

## Canon and Sources of Truth

**The generator scripts in `scripts/` are canon.** Nothing else is. `corpus/` and
`documents/` are output; `drafts/` is proposal, never canon; this file records decisions,
it does not create setting.

**Currently canon:** the Premise, the Tone, and the table specs. That is all. There is
no kingdom name, no wizard, no king, no map, no NPC roster, no timeline, no pantheon.

**Not yet decided — needs explicit sign-off before it can be used:**

| Open item | Notes |
|---|---|
| Name of the occupied kingdom | Referred to in drafts as *the kingdom* until named. |
| The wizard — name, nature, motive, power base | The campaign's spine villain. |
| The calling king and their realm | The "King" of the title. Their motive for the crusade matters. |
| The captured ruling family | Number, names, where held, whether all are alive. |
| Geography and the route of march | Both realms, the distance between, what lies along it. |
| The coalition | Which allies march, and what each of them actually wants. |
| Cosmology, gods, and magical traditions | Including what makes the kingdom's magic distinctive. |
| Timeline | How long the occupation has lasted; where the campaign opens in it. |
| Session and arc structure | How many modules, and what each covers. |

**Deliberately open** — questions left unresolved *by design*, which later sessions must
not helpfully close: *(none yet — this list opens once canon exists.)*

The distinction matters. An item in the table above is a **gap to be filled**. An item in
*Deliberately open* is **finished as it stands**. Keep them apart.

**Read before you write.** Ground yourself in `corpus/` at the start of any task. Never
generate against remembered canon — grep the file. This rule catches more errors than any
other. Right now `corpus/` holds only the smoke test, so the honest answer to most
questions about this setting is *"that has not been decided."* Say that instead of
inventing.

## Creative Latitude

**Invent freely, no sign-off needed:**
minor NPCs and their names; place names for locations already established; encounter
composition; skill DCs and their tiers; read-aloud text; treasure and loot tables;
tavern names, food, songs, weather, and incidental colour; the specific wording of
anything already approved in substance.

**Requires explicit sign-off before it enters canon:**
anything in the *Not yet decided* table; core mythology and cosmology; the fate of any
major NPC; structural worldbuilding (regions, factions, institutions, magic systems);
any change to a player-facing fact; anything already run at the table; and any change to
the Premise or Tone above.

**When in doubt, it needs sign-off.** Inventing setting is the one failure mode this
project was explicitly set up to avoid.

## Working in This Repository

- **Never hand-edit generated output.** Nothing in `corpus/` or `documents/` is ever
  edited directly; the next build silently discards it. Edit the script that produces it.
- **Regenerate in the same commit as the script change.** Run `tools/build.sh` and commit
  `scripts/`, `corpus/`, and `documents/` together. A corpus that disagrees with its
  generator is worse than no corpus.
- **Canon changes propagate in one pass.** A change to the sourcebook that affects the
  reference guide, the player guide, or a session module updates all of them in the same
  commit.
- **Large passes go through a draft first.** Any systemic layer, region set, NPC roster,
  or mechanical conversion gets a design draft in `drafts/` that: presents the proposal
  with reasoning; flags every item needing approval; states plainly what is natural
  extension of approved canon versus genuine new invention; and ends with a sign-off
  checklist and a propagation plan naming which scripts change on approval. Only after
  approval does it fold into canon in one consolidated pass. Rename the draft
  `*.RESOLVED.md` afterward with a header recording what was decided, including anything
  that changed on contact with the code.
- **Present forks as short multiple choice, not open questions.** Concrete directions with
  stated trade-offs, one decision at a time. Multi-select is fine for a batch of
  independent small items.
- **Commit at meaningful boundaries** — one consolidated pass per commit. Write commit
  messages describing what changed in the fiction, not the mechanics of the edit.
- **Editing method:** apply script edits with an anchor and an assertion — read, assert
  the anchor occurs exactly once, replace, write. A silent zero-match or double-match is
  how a "successful" pass quietly does nothing or corrupts two places.
- **`sed` ranges restart on repeated matches.** `sed -n '/Foo/,/^# /p'` re-triggers at
  every later `Foo` and prints a jumbled superset that looks like a structural bug and is
  not. Use explicit line numbers to inspect a specific section.

## Repository Layout

The generator scripts are the source of truth. Everything else is either input to them or
output from them.

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
  `reference/project-instructions.md` mirrors this file and **must be updated in the same
  pass**; they drift otherwise, which is the failure the mirroring rule exists to prevent.
- `images/` — artwork, named for what it depicts. **Input to the build**, reached from a
  generator with `IMG()`; see *The Authoring Kit*.
- `README.md` — repository index.

`corpus/` and `documents/` come from the same untouched scripts in the same build, so they
cannot drift apart. The `.docx` is a build intermediate and is never committed.

## The Document Set

Four documents, filename-prefixed `KC_`:

| Document | Role | Column layout |
|---|---|---|
| **Sourcebook** | Setting canon: the realms, the occupation, factions, magic, history. | Two-column |
| **Session modules** | One per session arc. Runnable at the table. | Two-column |
| **DM Reference Guide** | Wide scannable tables, stat block index, the Branch Ledger. | **Single-column** |
| **Player Guide** | The sanitized handout. | Two-column |

**The Player Guide is authored as its own document, never produced by deleting paragraphs
from the sourcebook.** Spoiler-safety lives in how sections are *written*. This split is
non-negotiable; it is the one that keeps player-facing material honest.

`SINGLE_COL_MATCH` in `tools/build.sh` is set to the DM Reference Guide, whose value is
wide tables that read badly in the two-column body.

## Table Specifications

- **Party:** 4–6 players
- **Starting level:** 5th — Extra Attack and 3rd-level spells online from session one.
  The party are the king's chosen champions, not raw levies. Build encounters for
  competent heroes and draw from higher-tier monsters accordingly.
- **Session length:** five hours
- **Advancement:** milestone

## Session and Encounter Design

**Every module carries the same skeleton:**

1. Overview with an explicit pacing budget against the five-hour session
2. **What Is Actually Happening (DM Only)**
3. Numbered scenes, each with boxed read-aloud
4. Tiered skill DCs
5. Full stat blocks for everything that fights
6. NPC profiles with speech patterns and open threads
7. **Optional Content** — standard, not rare; explicitly outside the core session length
8. **Diverging Paths (DM Only)**
9. Loot
10. Closing epigraph

**Encounter rules:**

- **Always include scaling for 4, 5, and 6 characters — and check the curve, not just the
  endpoints.** Compute adjusted XP against DMG thresholds at all three sizes. Watch the
  encounter-multiplier boundaries at **2, 3, 7, 11, and 15 monsters**: crossing one while
  "adding a body for a bigger party" can make the fight *harder* than the six-player
  version. Remember a party of six shifts the multiplier down a step, which quietly makes
  most encounters softer for large tables even as you add monsters.
- **Every combat gets a credible nonviolent or partial resolution where the fiction
  supports one, plus explicit morale** — when the enemy flees, folds, or surrenders.
- **Peril is the point, but so is agency.** A scene the party can only survive is worse
  than one they can outthink.

## Divergence Tracking

The DM Reference Guide carries a **Branch Ledger**: every tracked divergence, with a blank
column for what actually happened at the table. Each *Diverging Paths (DM Only)* section
in a module contributes its branches to the ledger. This is what keeps the campaign
replayable rather than a railroad with scenery, and it is the reason the reference guide
is single-column.

## Mechanical Validation

- **Do not judge stat blocks against the DMG's Monster Statistics by CR table alone.**
  Measured that way every humanoid NPC looks badly under-tuned. Official SRD humanoids sit
  just as far below it — a CR 3 Veteran has 58 hp where the table says 101–115. **The
  table describes monsters; NPCs are people.** Always compare against real SRD monsters at
  the same and neighbouring CR.
- **Automated damage extraction will lie to you.** Parsing `Hit: N (` off attack lines
  undercounts anything whose output lives in riders — smite traits, Sneak Attack, a second
  damage type, spellcasting. **Hand-check anything above CR 3 or built as a boss.**
- **Classify, do not correct.** Report each finding as *well-calibrated* / *intentional
  design pattern* / *genuine error*. Support NPCs, pack creatures, and control-focused
  adversaries are deliberately off-baseline. Identify those as design and leave them alone.
- **Check the encounter, not only the block.** Blocks can each be correct while the
  encounter they compose is not. This is the step that finds real bugs.

**SRD data** lives at `5e-bits/5e-database`, behind the proxy — attach it rather than
fetching directly (`codeload.github.com` returns a 403 JSON error, not a tarball):

```
add_repo(owner="5e-bits", repo="5e-database", access="read")
GIT_LFS_SKIP_SMUDGE=1 git clone --depth 1 \
  https://github.com/5e-bits/5e-database /home/user/5e-bits/5e-database
```

Monsters at `src/2014/en/5e-SRD-Monsters.json`; other types follow
`src/2014/en/5e-SRD-<Type>.json`. Coverage is **SRD-only** — no Volo's, no Mordenkainen's.
Flag material drawing on those as unvalidated when precision matters.

## Consistency Auditing

Periodically, and always after a large pass, audit the whole corpus for:

- new systemic canon against existing session text
- imported terminology clashing with established vocabulary
- timeline arithmetic across documents
- **opportunities** new canon creates in already-written scenes — not merely
  contradictions to fix

**The opportunity half is underrated and pays.** Report findings by severity with
recommended fixes; do not regenerate without sign-off.

## Production Practice

**One command rebuilds everything: `tools/build.sh`.** It regenerates `corpus/` and
`documents/`, applies the template, renders each document to PDF, and **fails the build on
escape leaks or font substitution** — the two errors that are invisible in source.

Prerequisites, all checked by the build:

| Requirement | Install | Why |
|---|---|---|
| Node + `docx` | `npm install docx` | The generators |
| Python 3 | stdlib only | `transplant.py`, `normalize_pdf.py` |
| LibreOffice **Writer** | `apt-get install libreoffice-writer` | `libreoffice-core` alone loads *nothing* and every document fails with "source file could not be loaded" — a content-shaped error with an environment cause. **This container shipped with core only; it has bitten this project already.** |
| Ghostscript | `apt-get install ghostscript` | Reproducible PDF |
| poppler-utils | `apt-get install poppler-utils` | `pdftotext`, `pdffonts`, `pdftoppm` for verification |
| Alegreya SC, Alegreya Sans SC, Lato | Google Fonts TTFs into `~/.local/share/fonts`, then `fc-cache -f` | Missing fonts **substitute silently and change pagination**; verifying layout without them is meaningless. `fonts.google.com/download` is blocked here — pull the static TTFs from `raw.githubusercontent.com/google/fonts/main/ofl/{alegreyasc,alegreyasanssc,lato}/`. |

**Verification, every time:**

```bash
grep -Pl '[^\x00-\x7F]' scripts/*.js                  # MUST print nothing — no literal non-ASCII
pdftotext documents/<doc>.pdf - | grep -c '\\u'       # MUST be 0 — no escape leaks
pdffonts documents/<doc>.pdf | grep -c DejaVu         # MUST be 0 — no font substitution
```

`tools/build.sh` now runs the source check itself and fails the build on a literal
typographic character in any generator, so this is a spot-check rather than the guard.
Use `-Pl` (list offending files) rather than `-Pc`: with more than one generator, `-Pc`
prints a `file:count` line per script and there is no single number to read.

**Note the doubled backslash** in the escape check. Single-quoted `'\u'` matches the plain
letter *u* and reports every ordinary word containing one, so it can never return 0 on real
prose. Then build **three times** and `cmp` — an unchanged document must rebuild
byte-identical, and the Ghostscript trailer-`/ID` failure mode this guards against is
intermittent, so one comparison is not enough.

**Reproducibility is within an environment, not across them.** `normalize_pdf.py` strips
per-run randomness (timestamps, trailer `/ID`, XMP UUIDs, font-subset tags); it does not
strip LibreOffice or Ghostscript *version* differences. A rebuild on a different toolchain
can produce a different byte stream from identical sources with identical text, layout,
page count, and fonts. When a rebuilt PDF differs unexpectedly, compare
`pdftotext -layout` output and `pdfinfo` before assuming a content bug.

**All prose lives in the generators as `\uXXXX` escapes**, never as literal typographic
characters. Normalize immediately after inserting new text, then re-run the greps —
inserting prose is exactly when literal characters sneak in.

**`node --check` is not sufficient.** It validates syntax, not identifiers — a call to a
helper that file does not define passes `--check` and throws at build time. Helper sets
genuinely differ between generators. Grep for the definition, and actually run the script.

**DM-only markers are bold book-red, never italic:** `const DM = (t) => ({ t, b: true,
c: "5B1F1F" })`, used as `PS([DM("DM Only: "), { t: "the note." }])`. Colour is
preattentive — a DM spots red without reading — and it leaves the body roman, which matters
because these notes run 100–200 words. **Colour the marker, not the prose.** Italic is
reserved for read-aloud, quotations, and epigraphs; overloading it makes both signals
ambiguous. Two rules follow from the Markdown shim appending a space after every bold run:
the marker carries its own trailing space, and the following segment never begins with one.
Sections already titled `(DM Only)` need no inline marker.

**Verify by looking.** Rendering bugs are invisible in source, and so are some content
bugs. Render a page to PNG with `pdftoppm` and read it — in the first campaign that caught
a factually wrong rules claim after every grep-based check had passed clean.

## Visual Template

Reused unchanged from The Qilvayas Symphony: the /u/YaAlex-derived 5e style — Alegreya SC
Medium headings in deep book-red `5B1F1F`, Alegreya Sans SC and Lato body, A4, page-number
footers. The book-red is load-bearing in **both** `scripts/style_template_encoded.md` and
every generator's `DM` constant; changing it means changing both together.

`transplant.py` carries **four** filename references to the encoded template. Renaming the
template means updating all four.

## The Authoring Kit

The helper preamble at the top of each generator is the portable authoring kit: `P`, `PS`,
`DM`, `H1/H2/H3`, `BULLET`, `B`, `BUL`, `BOX`, `cell/row/table`, `mod/abCell/SB`. Copy it
into each new generator. Three conventions are specific to this repository:

**Output path.** Never hardcode a stage directory. Generators write with:

```js
const { stagePath } = require('./stage');
fs.writeFileSync(stagePath("KC_Sourcebook.docx"), buf);
```

`stagePath` resolves `$KC_STAGE` (set by `tools/build.sh`) and falls back to `<repo>/.stage`
so a generator also runs standalone. `.stage/` is gitignored build scratch.

**Artwork.** `IMG(file, widthPt, heightPt, alt)` places an image from `images/`. Width and
height are points at 72/inch, so 288 is four inches. The helper passes an extra `mdPath`
field that the real `docx` ignores and the Markdown shim uses to write the corpus link —
it is relative to `corpus/`, where the generated `.md` lives, so it reads as
`../images/<file>`. `transplant.py` carries the image bytes and relationships into the
template package, remapping relationship ids so they cannot collide with the template's
own. **Keep images to PNG or JPEG**, and give every one real alt text.

**Full-width tables.** `table(headers, widths, rows, { full: true })` makes a table span
both columns instead of wrapping to three or four words a line. It works by tagging the
table with a marker style that `transplant.py` finds, wraps in a pair of continuous
section breaks, and strips. Use it for any table with a prose column; leave narrow
numeric tables in the column flow. In a `--single` document the marker is simply removed.

**Tables must not tear.** `row()` sets `cantSplit` so a row's cells cannot be torn across
a column or page break, header rows carry `tableHeader` so they repeat when a long table
does span a break, and the ability-score row in `SB()` uses `keepNext` so the values stay
with their labels. A stat block whose STR/DEX/CON header lands in one column and its
numbers in the next is a real bug and this is what prevents it.

## Current State

- `scripts/smoke.js` — a throwaway one-page document exercising every rendering path.
  **It is not canon.** Delete it, and its `corpus/` and `documents/` output, once real
  generators build clean.
- `tools/build.sh` — `GENERATORS=(smoke)` and `SINGLE_COL_MATCH="__NONE__"` are the two
  lines to update as real documents arrive.
- `images/pipeline-test.png` — a test card used only by `smoke.js` to prove the image
  path. **Delete it with `smoke.js`.**
- **The three bootstrap pipeline changes are built and verified** (done before any canon,
  while there was one generator rather than seven):
  1. The `/home/claude` stage path is parameterized — `scripts/stage.js`, `$KC_STAGE`.
  2. Artwork in `images/` reaches both the PDF and the Markdown corpus.
  3. Tables can span both columns with `{ full: true }`.
- Everything else in `scripts/` and `tools/` is the campaign-agnostic pipeline, carried
  over intact from The Qilvayas Symphony.
