# Design Draft — Module Breakdown

**Status: PROPOSAL. NOT CANON. This is the plan, not a menu — redline it and I proceed.**

`CLAUDE.md` lists the module breakdown as the one remaining item that everything else
waits on: session content, the DM Reference Guide's Branch Ledger, and the Player Guide
all need a shape to hang from. This draft sets that shape, sized against a **40–60 hour**
target at five-hour sessions — eleven modules, one of them a two-variant fork, landing
at roughly 55 core hours before a single Optional Content section is counted, which gives
real headroom under the 60-hour ceiling and comfortable room over the 40-hour floor even
if a table skips every optional scene.

Unlike the setting and draconic drafts, this one is not built as branching options — the
arc's shape is already locked (`CLAUDE.md`: *the march; the siege of Vindana; the field
battle that proves the enemy beatable; the approach; the decision at the gates*), so the
work here is filling in what sits between those five fixed points. I've made the calls
myself, as invited. Flag anything that's wrong; silence reads as approval to proceed.

---

## 0. What this does not touch

Two items already signed off constrain this draft and are not reopened: the two roads
split at the end of Module 1 and the unchosen road is met only as aftermath, never as a
full parallel module; and the turn-back-or-hold ending stays genuinely undecided until
the table decides it in Module 11.

One item is **not yet signed off** and this draft routes around it rather than assuming
it: the Wyvernheart battle and the draconic layer (`drafts/DRACONIC-LAYER.md`). Module 7
is written to be the natural home for that scene, per the draft's own recommendation, but
nothing there is treated as locked. If the draconic draft is marked differently, Module 7
changes; nothing else in this breakdown does.

## 1. The eleven modules

| # | Title | Core hrs | Carries |
|---|---|---|---|
| 1 | The Muster | 5 | The call reaches the party in Harrowmark; Xavier chooses them in person; a last Harrowmark trial establishes the land, its wyverns, and its unimpressed people. Closes on the road choice — **Branch Ledger entry 1**. |
| 2A / 2B | The Sea Road / The Mountain Road | 5 | Whichever the table picked, written in full; the other exists only as the aftermath folded into Module 5. A side-conquest (2A) or an attrition march through nominally friendly territory (2B). |
| 3 | Landfall | 5 | First contact with the Living Realm — a working habit, then a failing one, in the same session. First occupation soldiers and collaborators. Uses the Four Voices for the first time. |
| 4 | The Coalition | 5 | Oksitan and Auberitz in camp; the Promise's tension surfaces between allies; Norvatch offered as a leverage option, not an ally; the first rescue thread opens (one held royal). |
| 5 | The Road to Vindana | 5 | The approach march. The drowned king's aftermath is delivered here via the Four Voices, regardless of which road the party took — **Branch Ledger entry 1, resolved**. |
| 6 | The Siege of Vindana — Investment | 5 | The city invested, siege engines, first assault, Harrowmark riders introduced. Stakes rising, nothing decisive yet. |
| 7 | The Siege of Vindana — The Breaking | 5 | The siege's turning point. Pending the draconic draft, this is where Xavier's title-earning scene lands. The campaign's structural centre. |
| 8 | Held Ground | 5 | Aftermath and consolidation. Second rescue thread. A deliberate relief-valve module — taverns, absurd logistics, the coalition's comic figures get real room. |
| 9 | The Field Battle | 5 | Arsuf-keyed: the battle that proves Vale's army beatable. Third rescue thread or a real loss: the balance shifts visibly toward the party. |
| 10 | The Approach | 5 | The final march on Caer Ysolde. Last quiet chances to turn back, weighed narratively rather than mechanically. The coalition's cost is fully visible by now. |
| 11 | The Decision at the Gates | 5 | Final module. Turn back or hold — decided at the table, both endings fully written. **The Refrain's last line changes here, and only here.** |

Eleven slots, twelve files (Module 2's fork). 55 core hours. Optional Content in each
module — explicitly standard, not rare, per `CLAUDE.md` — is where the range extends
toward 60 for a table that takes it, and where it compresses toward 40 for one that
doesn't.

## 2. Why this shape, briefly

**Three modules before the party reaches Elduvaine, not one.** The Muster, the road, and
the road itself all before landfall. This is the direct instruction from the Kingdom of
Heaven touchstone — *the road as transformation*, and the note that the march "must not be
hurried through" — taken literally rather than gestured at.

**The rescue threads are staggered, not batched.** One royal at Module 4, one at Module 8,
the third folded into Module 9's stakes rather than run as its own errand. This keeps the
captured family a running thread across the whole campaign instead of a single fetch
quest, and gives the *mild version* of their internal friction room to surface gradually
rather than all at once.

**Module 8 exists on purpose and does no plot work.** `CLAUDE.md` is explicit that levity
is planned, not accidental — "a campaign about a long grim march needs them by design, not
by accident." A module with no required beat, placed right after the campaign's biggest
set piece, is where that design shows up structurally rather than as scattered asides.

**Vindana gets two modules, nothing else does.** It is named as the campaign's central
siege, and two modules is what "central" should cost against an eleven-module arc — investment,
then the breaking. Everything else gets one.

## 3. What this draft does not cover

Deliberately out of scope: the DM Reference Guide's full Branch Ledger (built as each
module writes its own Diverging Paths section, per `CLAUDE.md`); the Player Guide (its own
document, written after the modules exist to draw the spoiler-safe line from); and any
scene-level content, which is Creative Latitude territory once a module's slot is
approved and does not need its own sign-off.

## 4. Sign-off

| Item | Recommendation | Mark |
|---|---|---|
| Eleven-module shape, ~55 core hours | Approve as written | |
| Staggered rescue threads (4 / 8 / 9) | Approve as written | |
| Module 8 as a dedicated levity module | Approve as written | |
| Module 7 provisionally hosting the Wyvernheart scene | Approve, contingent on the draconic draft | |

## 5. Propagation plan on approval

1. `CLAUDE.md` — the *Module breakdown* row moves out of *Not yet decided* into canon,
   naming the eleven-module shape. Mirror to `reference/project-instructions.md`.
2. `tools/build.sh` — `GENERATORS` grows by one entry per module as each is written,
   in play order.
3. `scripts/KC_Module01_TheMuster.js`, then each module in sequence — built and committed
   one at a time, not as a single pass, so each is verified against the pipeline before
   the next is written. This draft does not assume all eleven land in one session of work.
4. `scripts/KC_DM_Reference_Guide.js` — started once two or three modules exist, so the
   Branch Ledger has real entries to open with rather than an empty table.
5. This draft renamed `MODULE-BREAKDOWN.RESOLVED.md` once the shape is confirmed, with a
   header recording anything that changed on contact with the writing.
