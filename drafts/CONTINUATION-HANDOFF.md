# Handoff — Continuing The King's Crusade

*For a fresh Claude Code session picking this campaign up. Not a summary of `CLAUDE.md`,
which you get automatically and which is the authority on canon. This is the part that
is not written down anywhere else: where the work actually stands, what is worth doing
next and why, and the specific ways a well-meaning session can damage this repository
while trying to help.*

Distinct from `drafts/NEW-CAMPAIGN-HANDOFF.md`, which is about bootstrapping a *different*
campaign on the same machinery. This one is about continuing *this* one.

---

## Read before you write

1. **`CLAUDE.md`** — canon, the sign-off gates, the tone. It is long and it is load-bearing.
2. **The `kc-build` skill** (`.claude/skills/kc-build/`) — every production mechanic:
   the escape convention, table rules, the authoring kit, the anchored-edit discipline,
   the verification loop. Invoke it before touching `scripts/`. It is the single copy of
   those rules; do not restate them elsewhere and do not rediscover them.
3. **`corpus/`** — grep it before generating anything. Never write against remembered
   canon. This rule catches more errors than any other in the repository.

Then run `tools/verify.sh` once, cold, before changing anything. It takes about a minute
and it tells you the set is intact and the toolchain works.

## What you inherit

**Eighteen documents, 101 pages, ~84,000 words**, all verified clean and byte-reproducible.
Six core volumes (Sourcebook, Gazetteer, Bestiary, Character Options, DM Reference Guide,
Player Guide) plus twelve module files covering eleven session slots.

The setting's foundation is signed off and binding: the kingdom, the wizard and his motive,
the calling king, the occupation as a hired legion, the Nine Works and the schism, House
Ysolde by name, Oksitan and Norvatch's agendas, and the shape of the war across an approved
eleven-module arc. `CLAUDE.md` states all of it precisely.

## The scale question, honestly

The DM asked for something at published scale — 224–320 pages for a setting book, 50–100
for an adventure module. **We are at 101 pages total.** That gap is real and it should be
stated plainly rather than quietly ignored.

The arithmetic, at this template's density of roughly 830 words per page:

| Target | Words needed | Still to write |
|---|---|---|
| 224 pages | ~186,000 | ~102,000 |
| 320 pages | ~266,000 | ~182,000 |

The first expansion pass closed the *structural* gaps — there was no gazetteer, no pantheon,
no timeline, no factions, no character options, no bestiary, and no puzzles at all. Those
now exist. **What remains is depth on existing bones, not new bones**, and that is a
different and slower kind of work: more settlements with more keyed detail, more stat
blocks, per-region adventure seeds, and substantially longer scene text in the modules,
which are 4–6 pages each against a published module's 50.

Do not promise a page count you will not reach in one pass. Say what you actually did.

## Where the next pass should go, ranked

1. **Settle what Auberitz wants.** It is the last coalition member with no agenda, and it
   blocks more than it looks: Auberitz owns the siege train, the engineers and the
   quartermasters, so it appears in nearly every module and is currently run by role only.
   Every other power in this war has a motive — Harrowmark the Promise, Oksitan the Willing
   Road it can never have, Norvatch the market, the Tenth Work the tenth work. Auberitz
   builds the engines and wants nothing. **Put it to the DM as short multiple choice with a
   recommendation, then propagate in one pass.**
2. **The modules are the thinnest thing in the set relative to published work.** Each has a
   puzzle and a set piece now; what they lack is keyed depth — more locations, more scene
   text, more NPC lines, and per-module encounter tables. This is where the largest number
   of pages lives and it needs no new canon.
3. **The Gazetteer's regions can carry adventure seeds.** Thirty keyed places exist with a
   hook apiece. Three or four seeds per region, tied to the encounter tables already there,
   is high value for the effort and again needs no sign-off.
4. **Maps.** Tracked as an open item because it is a canon question, not only a production
   one: a drawn map commits to coastline, adjacency and scale in ways the prose has
   deliberately avoided, and the gazetteer's travel times are the only binding spatial facts.
   `IMG()` and the transplant path are built and tested. Needs geography sign-off first.

## The gates — do not walk through these

**Needs explicit sign-off** (the `Not yet decided` table in `CLAUDE.md`): what Auberitz
wants; whether Harrowmark was ever otherwise; what the deepest vaults contain; whether Vale
is still human; cosmology and the planes; maps and artwork.

**Deliberately open — finished as they stand, and a helpful session will try to close them:**

- **What the Willing Road measures.** Nobody has ever established it, the campaign never
  will, and the DM guidance says not to invent a criterion *even privately in your notes*.
  The moment it measures something nameable it becomes an alignment detector with a
  travel-time output and the strangest thing in the setting is gone.
- **Whether the party holds Elduvaine or turns back.** Both endings are genuinely supported.
  All three touchstones converge on the settlement ending and the campaign still declines to
  favour it. Do not reweight this.
- **What Elduvaine should be after the war.** Five royals hold five positions and the
  campaign adjudicates none of them.

The distinction matters: the first list is **gaps to fill**, the second is **finished**.
Keep them apart.

## Loose threads left open

- **A 3.1 MB PNG sits in git history.** The title banner was committed as PNG, then
  re-encoded to a 540 KB JPEG. The blob costs ~3 MB on a fresh clone and nothing on page
  views. Purging it needs a history rewrite and a force-push; that was attempted, verified
  locally (10 MB → 6.9 MB, tree hash identical), and then abandoned because the force-push
  was blocked. The DM said they would replace the image manually in due course. **Do not
  rewrite published history without being asked.**
- **`drafts/MODULE-BREAKDOWN.md` is stale.** It is still headed "PROPOSAL. NOT CANON —
  redline it and I proceed", and the modules it proposes have all been written. It should
  either be renamed `*.RESOLVED.md` with a decision record, per the repository's own
  convention, or deleted.
- **The sister repository has drifted on tooling.** `tools/pipeline.conf` states that
  `anchor.py`, `normalize_escapes.py`, `check_columns.py`, `find_page.py` and `verify.sh`
  are byte-identical between this repository and The Qilvayas Symphony. They are not:
  Qilvayas has only `build.sh`, `docx-md-shim`, `normalize_escapes.py` and
  `normalize_pdf.py`. The claim in `pipeline.conf` is currently aspirational.
- **Qilvayas has the same bullet-indent defect** fixed here, at the same column measure.
  Its equivalent value differs because its body type is 11pt rather than 10pt.

## Three things learned the hard way

All three are in the `kc-build` skill in full; they are flagged here because they are the
mistakes most likely to be repeated by someone who has not read it yet.

- **Full-page defaults keep turning up in a narrow column.** Three so far: table
  `columnWidths` (docx-js emits no `<w:tblGrid>` without them, so LibreOffice threw the
  percentages away), table cell padding, and bullet indent. When a two-column page looks
  loose or cramped, **suspect a measurement inherited from a 6.5in sheet before suspecting
  the writing.** Check the fourth one before assuming there isn't one.
- **The drift check has a false negative.** A layout-only change alters no text, so every
  `corpus/*.md` stays byte-identical while all the PDFs differ — the exact shape of
  cross-container render churn, which you are told to discard. Prove it before discarding:
  `pdftotext -bbox` both versions and diff the word boxes. Same word count with many boxes
  moved is a real layout change and must be committed.
- **Compose in real characters and normalize afterward.** Hand-escaping while writing is
  slow and produces the doubled-backslash bug, which compiles clean, passes the non-ASCII
  scanner, and leaks literal `’` into the PDF. `tools/normalize_escapes.py` handles it.

## Working rules that are easy to miss

- **Never hand-edit `corpus/` or `documents/`.** Edit the generator; the next build discards
  everything else.
- **Commit `scripts/`, `corpus/` and `documents/` together.** A corpus that disagrees with
  its generator is worse than no corpus.
- **`reference/project-instructions.md` mirrors `CLAUDE.md`** and must be regenerated in the
  same pass. It has silently drifted twice.
- **`tools/pipeline.conf` lists the player-facing documents** and the leak patterns that
  guard them. Character Options is player-facing; the Player Guide is authored separately
  and never by trimming the sourcebook.
- **Large passes go through a draft in `drafts/` first**, with a sign-off checklist and a
  propagation plan, renamed `*.RESOLVED.md` afterward with a record of what changed on
  contact with the code.
- **Verify with `tools/verify.sh --full` before committing.** `--full` adds three builds and
  a byte comparison, because the Ghostscript trailer failure mode is intermittent.
- **Then render the pages you changed and look at them.** Every layout bug in this
  repository's history was invisible in source and to every grep.
