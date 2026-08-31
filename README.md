# The King's Crusade

A Dungeons & Dragons 5th Edition campaign setting.

**Status: bootstrapped, not started.** The document pipeline is installed and proven end
to end. The campaign itself — theme, setting, canon — has not been defined yet.

## Start here

Read [`drafts/NEW-CAMPAIGN-HANDOFF.md`](drafts/NEW-CAMPAIGN-HANDOFF.md). It carries the
method, the conventions, and the environment gotchas from the first campaign built on this
pipeline ([The Qilvayas Symphony](https://github.com/RettifiloAscari/The-Qilvayas-Symphony)),
and §7 lists the decisions that have to be made before any content is written.

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
| `reference/` | Mirrored instructions, if a Claude Chat project reads this repo. |
| `images/` | Artwork. **Outside the build** — see `CLAUDE.md`. |

`corpus/` and `documents/` are produced from the same untouched scripts in the same build,
so the Markdown and the published documents cannot drift apart.

## Building

```bash
npm install docx
apt-get install -y libreoffice-writer ghostscript poppler-utils
# Alegreya SC, Alegreya Sans SC, Lato -> ~/.local/share/fonts && fc-cache -f
tools/build.sh
```

The build fails on escape leaks or font substitution — the two errors invisible in source.
An unchanged document rebuilds byte-identical.

## What is here now

`scripts/smoke.js` is a throwaway one-page document that exercises every rendering path
(headings, DM markers, bullets, tables, read-aloud, a stat block). It exists so the
pipeline could be proved before any canon was written. Delete it once real generators
build clean.
