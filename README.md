# The King's Crusade

A Dungeons & Dragons 5th Edition campaign setting.

**Status: bootstrapped, premise set, no canon written.** The document pipeline is
installed and proven end to end. The premise, tone, document set, and table specifications
are decided. The setting itself — the kingdom, the wizard, the king, the map, the NPCs —
does not exist yet.

## The campaign

A friendly magical kingdom has fallen to an evil wizard and their army. The ruling family
is captive, the land laid waste. A king calls a crusade, and the party is sent to reclaim
those lands and deliver them.

High fantasy, traditionally D&D, on a **Third Crusade** spine — the structure of a distant
king answering a call, a long march, an occupying power, coalition allies with their own
agendas. Where the spine and high fantasy disagree, high fantasy wins.

> *A strange and marvellous kingdom worth saving, and a long dangerous road to reach it.*

Four documents: a sourcebook, session modules, a single-column DM Reference Guide, and a
separately-authored Player Guide. Built for 4–6 players starting at 5th level, five-hour
sessions, milestone advancement.

## Start here

Read [`CLAUDE.md`](CLAUDE.md) — it carries the premise, the tone, what is canon, what is
explicitly undecided, and the production practice. Then
[`drafts/NEW-CAMPAIGN-HANDOFF.md`](drafts/NEW-CAMPAIGN-HANDOFF.md), which carries the
method and the environment gotchas from the first campaign built on this pipeline
([The Qilvayas Symphony](https://github.com/RettifiloAscari/The-Qilvayas-Symphony)).

## How this repository works

The **generator scripts are the canon.** Everything else is input to them or output from
them — with one exception, `images/`. To change the campaign, edit a script in `scripts/`
and run `tools/build.sh`; never edit the generated files directly, because the next build
discards those edits.

| Directory | What it is |
|---|---|
| `scripts/` | **The canon.** docx-js generators, `transplant.py`, and the encoded visual template. |
| `corpus/` | **Generated Markdown** — readable and greppable on any device. |
| `documents/` | **Generated PDF** — styled, embeds its fonts, reads on any device. |
| `tools/` | `build.sh` regenerates and verifies everything; `docx-md-shim/` emits the Markdown; `normalize_pdf.py` makes the PDF reproducible. |
| `drafts/` | Design drafts awaiting sign-off. **Not canon.** |
| `reference/` | Mirrored instructions — `project-instructions.md` mirrors `CLAUDE.md`; update both together. |
| `images/` | Artwork. **Input to the build** — generators place it with `IMG()`. |

`corpus/` and `documents/` are produced from the same untouched scripts in the same build,
so the Markdown and the published documents cannot drift apart.

## Building

```bash
npm install docx
apt-get install -y libreoffice-writer ghostscript poppler-utils
#            ^ libreoffice-writer, NOT libreoffice-core: core alone loads nothing and
#              every document fails with "source file could not be loaded".
# Alegreya SC, Alegreya Sans SC, Lato -> ~/.local/share/fonts && fc-cache -f
#   fonts.google.com/download may be blocked; the static TTFs are at
#   raw.githubusercontent.com/google/fonts/main/ofl/{alegreyasc,alegreyasanssc,lato}/
tools/build.sh
```

The build fails on escape leaks, font substitution, or a missing Writer module — the errors
that are invisible in source or that masquerade as content bugs. An unchanged document
rebuilds byte-identical **within one toolchain**; LibreOffice and Ghostscript version
differences change the byte stream without changing the document.

## What is here now

`scripts/smoke.js` is a throwaway document that exercises every rendering path — headings,
DM markers, bullets, tables, read-aloud, a stat block, an embedded image, and a full-width
table spanning both columns. It exists so the pipeline could be proved before any canon was
written. Delete it, `images/pipeline-test.png`, and their generated output once real
generators build clean.
