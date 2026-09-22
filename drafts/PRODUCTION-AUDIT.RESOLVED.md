# Production Audit — RESOLVED

> **Status: resolved and applied.** A full audit of the nineteen-document set before it goes
> to the table: formatting, grammar, spelling, continuity, arithmetic and structure. Every
> change listed here is in the generators, rebuilt, and verified; nothing in it is new canon.
> The two lists at the end record what was deliberately *not* changed, so the next session
> does not "fix" it.

Result: **nineteen documents, 112 pages, ~96,500 words.** `tools/verify.sh --full` passes
clean, and every document is byte-identical across three builds.

---

## What was checked

| Check | How | Result |
|---|---|---|
| Doubled words, stray spaces, doubled punctuation, doubled dashes | Regex scan of every `corpus/*.md` | No defects. The remaining hits are `ft.,` in stat blocks and empty ledger cells, both correct |
| Straight quotes and apostrophes | Regex scan | Three fixed (below) |
| Spelling | `aspell` (en_GB) over the whole corpus | No typos. Real issue was **mixed US/UK spelling**, fixed (below) |
| Module skeleton | Every module against the ten-part skeleton in `CLAUDE.md` | Five gaps found and filled (below) |
| The Refrain | Verse compared byte for byte across all twelve module files | Identical in eleven; only Module Eleven's last line changes, as canon requires |
| Wyvernheart gating | Every use of the name, by document | None in narration before Module Seven; none in the Player Guide |
| Name and rank consistency | Every named NPC and title | Consistent (Serjeant Hoth, Marshal Ossian Drell, General Ilyana Voss, Baron Osgar Vell, Warden Ivor Thane) |
| Encounter arithmetic | Every scaling table recomputed against 2014 DMG thresholds and multipliers | Three errors found, including one in the previous pass's own work |
| Numeric continuity | Travel times, populations, ages, years, tonnages across documents | Five contradictions found |
| Cross-references | Every "Module N" reference and every handoff | Style unified. The last pass already fixed the broken handoffs |
| Branch Ledger coverage | Each module's *Diverging Paths* against the ledger | Ledger is a strict superset. Two campaign-spanning entries added |
| Stat Block Index | Index against the blocks actually printed and called for | Stale. Rewritten |
| Bestiary integration | Each Bestiary block against the modules that use it | **Five of six resident-magic creatures appeared in no module.** Placed |
| Spoiler safety | Leak scan of the two player-facing books | Clean |
| Layout | Column starvation, torn stat blocks, stranded headings, spot renders | Clean. Four *tight* findings, all ordinary typography (below) |
| Project instructions | `CLAUDE.md` against the canon it records | Six stale statements corrected. Mirror in sync |

## Changes made

### Spelling and punctuation

- **House style is British, and the corpus now follows it throughout.** The sourcebook,
  gazetteer and bestiary were already British; the modules drifted. There were
  more than eighty corrections in total — colour, flavour, recognise, centre, travelled, rumour, humour,
  organise, favour, defence, summarise, armoury, harbour, judgement, realised, stabilise,
  maths, centrepiece, practised, outmanoeuvred, and the one "waggon" among thirteen "wagon".
- **5e rules terms were deliberately kept in American form:** *Armor Class*, SRD armour
  names inside AC parentheses, and *mage armor*. Those are rules text and must match the
  rulebook.
- **Three straight apostrophes curled.** The escape normaliser cannot reach these, because
  it only converts an apostrophe between two word characters: *commons’ version*,
  *Will-o’-wisp*, and *three days’ ration*.
- **"Module 5" and "Module Five" were both in use.** Every prose reference is now spelled
  out, matching the majority. Module 2A and 2B keep their letters.

### Arithmetic

- **Module Five, the Held Winter:** two winter wolves and a troll were counted as "four
  monsters". They are three, and the text now says so. The six-player row said "extremely
  dangerous"; the actual figure is Hard, and it now says that instead.
- **Module Nine, the scaling table added last pass:** the Hard column used 1,050 XP per
  character at 7th level. The 2014 DMG figure is 1,100, so the column now reads 4,400,
  5,500 and 6,600. The difficulty readings were checked again and did not change.
- **Module One's muster roll** read "four hundred people out of sixteen hundred and sixty of
  them were promised…", which parses as 1,660. A comma was missing.

### Continuity

- **Greywatch was three days from Duncarrow in Module One and four in the gazetteer.** It is
  four, and the muster yard the party returns to is now "ten days ago" rather than three.
- **Stannock:** the gazetteer's population of 2,400 contradicted both its own "a quarter of
  everyone it had" and Module One's "four hundred out of sixteen hundred". The population
  is now 1,600, which makes both statements true.
- **Norvatch's ledgers:** the sourcebook says the carriage contract *triples* in Year Two,
  but Module Ten's table showed light-stone going up nearly fivefold. The table now matches
  canon: 1,400 tons, then 4,200 (tripled), then 19,200 (more than quadrupled). The
  one-third conclusion is unchanged.
- **Raimon V:** last pass had him "never reaching either road", which contradicted canon
  ("takes the road the party does not take"). Module One also shows Oksitan's camp at the
  muster. Both are now reconciled: Oksitan sent a contingent to Duncarrow, and Raimon led
  the main host out of Aurignan by the other road. The Vaskren lay across the start of that
  road, and he drowned on the ninth day, which matches the Standing Water's "on the ninth
  day".
- **Brenna Vane** was "two centuries old" in the DM Reference Guide and "in her second
  century" in Module One. The guide now matches the module.

### Structure: the module skeleton

- **Module Five** had no *Tiered Skill DCs* and no stat blocks, despite a set piece that
  calls for winter wolves and a troll. It now has a DC table, both SRD blocks reprinted
  unmodified (checked trait by trait against the 2014 JSON), and a scaling note for the
  wolves and troll met separately and together.
- **Module Four** had no scaling for Sennoch Hall. It now has a three-row table (quiet, the
  gate, the whole Hall) showing why the module steers toward quiet. Its claim that a party
  fighting all eighteen legionaries "wins" was untrue; it now says the fight is past Deadly
  and winnable only on the party's own terms.
- **Modules Six and Seven** had no scaling for First Assault or the breach. Both now do, at
  6th level.
- **Modules Seven, Eight and Ten** had no stat-block section. Each now says where its blocks
  live, or that it has none by design. **Module Ten** also gained a DC table.
- **Module Eleven** had no Optional Content. See the Veiled Sovereign, below.

### Integration: canon that no module used

- **Maelis Ysolde had no scene and no profile anywhere in the modules.** She is the
  campaign's second clock, and she is held in the very city the finale takes. Module Eleven
  now carries her NPC profile (still no stat block, as canon requires) and an optional scene
  in the Keep, built from the sourcebook's own words and the DM Reference Guide's
  *Sovereign's Veil*. She tells the party the number if asked. She does not settle the
  vaults, Vale's humanity, or the ending. Scene Four's list of royal voices now names her,
  and names Ninian, Ottoline and Aveline in place of their titles.
- **Five of the Bestiary's six resident-magic creatures appeared in no module.** Each is now
  placed as optional content where the gazetteer puts its kind:
  - the **Waystone Warden** on the Willing Road to Vindana (Module Five);
  - the **Season-Bound Stag** in Bryn Aeling (Module Five);
  - the **Light-Hollow** under Vindana's dark inner wall (Module Six);
  - a **Draining Engine** behind Voss's reserve (Module Nine), with a Branch Ledger count of
    Engines broken;
  - the **Echo of the Listening Water** at Lisswater, home of Aveline's miller (Module Ten).

  None of the five placements touches what the Willing Road measures.
- **DM Reference Guide.**
  - The Stat Block Index said the Occupation Guard was reused in five modules; it is now
    Caerwyn's alone.
  - The index gained the Legionary, the Winter Wolf and the Troll.
  - A new *Bestiary Blocks in Play* table maps every Bestiary block to every module that
    calls for it.
  - The Branch Ledger gained the Fenmarrow letter and a running count of broken Engines, for
    45 rows in total.
- **Module One's Diverging Paths** now tracks the Fenmarrow letter, which the previous pass
  planted without a ledger entry. **Module Ten's** refers to Aveline by name.
- **Module Ten's Lisswater scene** was drafted with an invented event (a royal household
  "carried out of the Keep"). It was replaced before commit with two established facts.

### Project instructions

`CLAUDE.md` and its mirror had six statements that canon has since overtaken:

- the status header still listed Oksitan's agenda and the royal names as undecided;
- the Barbarossa touchstone pointed at *Not yet decided* for the drowned king;
- the Wyvernheart bullet said "which battle… is not yet decided";
- the Skaldic Bard bullet called what the Call promises "an open question";
- the coalition bullet said Oksitan's want was undecided;
- the named-peoples list gave three of the royals by title only.

All six are corrected. Also corrected: the stale counts (nineteen recurring NPCs is now
seventeen, since two were duplicate rows; the ledger count is now 45) and "three new
generators".

## Found and deliberately left alone

- **The four *tight* heading findings:**
  - Module 2A p3 *ACTIONS*;
  - Module Seven p2 *Scene 3: The Dragon*;
  - Sourcebook p4 *Caer Ysolde and the Archive*;
  - Sourcebook p12 *The Four Voices*.

  All four were rendered. Each is a heading with one or two lines under it at a column
  break, which is ordinary typography. Binding them would buy nothing but white space.
- **General *Ilyana* Voss.** The given name is not on the sourcebook's Legion names list.
  The list is illustrative, not exhaustive, and the name is used consistently in four places.
- **Both *toward* and *towards*, *learned*, *burned*.** Both forms are correct British
  usage; normalising them would be noise.
- **The Branch Ledger has more rows than the modules' Diverging Paths.** The extra rows come
  from the puzzle sections. The ledger is meant to be the superset.
- **Everything in *Not yet decided* and *Deliberately open*.** None of it was touched:
  - Auberitz's agenda;
  - the vaults' contents;
  - Vale's humanity;
  - Harrowmark's past;
  - the planes;
  - maps;
  - what the Willing Road measures;
  - hold or turn back;
  - what Elduvaine becomes.

## Worth the DM's attention before the first session

- **Run [Session Zero](../corpus/KC_Session_Zero_Primer.md) first.** It settles the one
  question Module One cannot: whether the party already knows each other.
- **Plant the Fenmarrow letter in Module One's muster yard, and write down the son's name.**
  It pays out for the rest of the campaign, and the ledger now tracks it.
- **The campaign prescribes no milestone schedule.** Every scaling table from Module Six
  onward states the level it assumes. Recompute against your own party's level before
  running any of them.
