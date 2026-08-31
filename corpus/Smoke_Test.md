<!-- GENERATED FILE - DO NOT EDIT.
     Source:     scripts/smoke.js
     Regenerate: tools/build.sh
     Hand edits here are overwritten and never reach the published documents.
-->

**Pipeline Smoke Test**

*Delete this document, and scripts/smoke.js, once the real generators build clean.*

# What This Page Proves

If this document rendered, the whole chain works: docx-js produced a .docx, transplant.py applied the visual template, LibreOffice converted it to PDF, Ghostscript rewrote it, and normalize_pdf.py stripped the per-run randomness. If tools/build.sh also reported “all documents verified clean,” then no escape sequences leaked into the output and no template font was silently substituted.

Build it twice and compare the bytes. They must be identical — that is what makes the git history stay clean when a document has not actually changed.

## Every Rendering Path, Exercised

- **Bullets:** this is one, produced by the BUL helper.
- **Bold leads:** the lead is bold, the rest is roman, and the helper puts the space in the right place.
- A bullet with no lead, for continuation points.

### The DM Marker Convention

**DM Only:** the marker is bold and book-red; the prose after it stays roman. Colour is preattentive — a DM spots red without reading — and leaving the body roman matters because these notes run long. Colour the marker, not the prose. Italic is reserved for read-aloud, quotations, and epigraphs, and overloading it makes both signals ambiguous. Note also that the marker segment carries its own trailing space and the segment after it never begins with one, because the Markdown shim appends a space after every bold run.

## Read-Aloud

> *Boxed text renders in italic on a tinted ground, indented from both margins. This is where the players hear the scene, and it is the one place italic is doing its intended job.*

## A Table

| **Check** | **What it catches** | **Must be** |
| --- | --- | --- |
| grep -Pc non-ASCII | Literal typographic characters in the generator source | 0 |
| pdftotext \| grep -c escape | Escape sequences leaking into the rendered output | 0 |
| pdffonts \| grep -c DejaVu | A template font missing and silently substituted | 0 |
| Second build, cmp | Per-run randomness the normalizer failed to strip | identical |

## A Stat Block

**Test Construct**

*Medium construct, unaligned — exists only to prove the SB helper renders*

**Armor Class:** 15 (plated)

**Hit Points:** 27 (5d8 + 5)

**Speed:** 30 ft.

| **STR** | **DEX** | **CON** | **INT** | **WIS** | **CHA** |
| --- | --- | --- | --- | --- | --- |
| 14 (+2) | 12 (+1) | 13 (+1) | 3 (−4) | 10 (+0) | 1 (−5) |

**Senses:** blindsight 30 ft., passive Perception 10

**Languages:** understands its instructions; cannot speak

**Challenge:** 1 (200 XP)

***Immutable Form.*** The construct is immune to any spell or effect that would alter its form.

**ACTIONS**

***Slam.*** Melee Weapon Attack: +4 to hit, reach 5 ft., one target. Hit: 6 (1d8 + 2) bludgeoning damage.

~

*Pipeline proved. Now go and decide what the campaign is about.*
