# The King's Crusade

A Dungeons & Dragons 5th Edition campaign setting.

**Status: foundation signed off, sourcebook v1 built.** The kingdom, the wizard, the
king, the occupation and the shape of the war are canon. The coalition's members, the map,
the royal names and the module breakdown are not.

## The campaign

**Elduvaine** is the Living Realm, where magic is resident rather than worked: roads
shorten for travellers who mean well, rivers keep what was said on their banks, woods hold
the season they were planted in. Three years ago **Maedoc Vale**, Keeper of the Ysolde
Archive, opened every ward he had spent his life maintaining and let an army through them.
The kingdom fell to a key, not a siege.

**Xavier III of Harrowmark** has called a crusade and is going himself. The party are the
champions he sends ahead of the army. He is not yet called the Wyvernheart; he earns that
name during the war, and the players are meant to be there when he does.

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

`scripts/KC_Sourcebook.js` — the sourcebook, and the only canon written so far. The
session modules, the DM Reference Guide and the Player Guide are still to come. The
throwaway smoke test that proved the pipeline has been retired now that a real document
builds clean.
