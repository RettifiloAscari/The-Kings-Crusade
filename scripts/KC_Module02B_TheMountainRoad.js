// KC_Module02B_TheMountainRoad.js -- Session Module Two (Mountain Road variant).
//
// Canon lives here. corpus/ and documents/ are generated from this file and are
// never edited by hand. See CLAUDE.md for the sign-off rules: anything in the
// "Not yet decided" table must not appear here until it has been approved.
//
// ESCAPE CONVENTION: all prose lives as \uXXXX escapes, never as literal
// typographic characters. tools/build.sh fails the build if that slips. If you
// hand-type an escape rather than writing the real character, use ONE
// backslash (\uXXXX) -- a doubled backslash compiles fine and passes the
// non-ASCII scanner, but leaks the literal text into the rendered PDF. See the
// commit that fixed this in KC_Module02A_TheSeaRoad.js if in doubt.
//
// This is one of two mutually exclusive Module Two variants -- run this one if
// the party chose the mountain road at the end of Module One; run 2A if they
// chose the sea road. Per CLAUDE.md, whichever road they take is written in
// full; the other is met only as aftermath, delivered in Module 5.
//
// Encounter design note: Baron Osgar Vell is the SRD Bandit Captain (CR 2, 450
// XP) renamed and reflavored, backed by SRD Scouts (CR 1/2, 100 XP each) as
// mountain skirmishers -- both pulled from 5e-bits/5e-database rather than
// from memory. As in 2A, the DMG math reads Easy-to-Medium; this road\u2019s real
// danger is attrition and terrain, not a single fight, and Scene 2\u2019s river
// crossing is written to foreshadow the drowned king without naming him.

const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType, LevelFormat } = require('docx');
const fs = require('fs');
const path = require('path');
const { stagePath } = require('./stage');

// ---------- helpers (the portable authoring kit) ----------
const P = (text, opts = {}) => new Paragraph({
  spacing: { after: 200 },
  ...opts,
  children: [new TextRun({ text, ...(opts.run || {}) })]
});

const PS = (segs, opts = {}) => new Paragraph({
  spacing: { after: 200 },
  ...opts,
  children: segs.map(s => new TextRun({ text: s.t, bold: !!s.b, italics: !!s.i, color: s.c }))
});

const DM = (t) => ({ t, b: true, c: "5B1F1F" });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });

const BULLET = (segs, opts = {}) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 120 },
  ...opts,
  children: segs.map(s => new TextRun({ text: s.t, bold: !!s.b, italics: !!s.i, color: s.c }))
});

const B = (lead, rest) => PS([{ t: lead + " ", b: true }, { t: rest }]);
const BUL = (lead, rest, opts = {}) => BULLET(lead ? [{ t: lead + " ", b: true }, { t: rest }] : [{ t: rest }], opts);

const BOX = (text) => new Paragraph({
  spacing: { before: 120, after: 160 },
  shading: { type: "clear", fill: "F3EFE4" },
  indent: { left: 220, right: 220, firstLine: 0 },   // template default firstLine=180 otherwise leaks in
  keepLines: true,   // never let a boxed passage tear across a page break
  children: [new TextRun({ text, italics: true })]
});

const VERSE = (lines) => new Paragraph({
  spacing: { before: 120, after: 160 },
  shading: { type: "clear", fill: "F3EFE4" },
  indent: { left: 220, right: 220, firstLine: 0 },   // same fix as BOX -- see its comment
  keepLines: true,   // the refrain must never split across a page break
  children: lines.map((l, i) => new TextRun({ text: l, italics: true, ...(i ? { break: 1 } : {}) }))
});

const IMG_DIR = path.join(__dirname, '..', 'images');
const IMG = (file, w, h, alt) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 140, after: 160 },
  children: [new ImageRun({
    type: path.extname(file).slice(1).toLowerCase().replace('jpeg', 'jpg'),
    data: fs.readFileSync(path.join(IMG_DIR, file)),
    transformation: { width: w, height: h },
    altText: { name: alt, description: alt, title: alt },
    mdPath: '../images/' + file
  })]
});

const { Table, TableRow, TableCell, WidthType, ShadingType, TableLayoutType } = require('docx');
const cell = (text, opts = {}) => new TableCell({ width: { size: opts.w || 20, type: WidthType.PERCENTAGE }, shading: opts.head ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, margins: { top: 55, bottom: 55, left: 60, right: 60 }, children: [new Paragraph({ spacing: { after: 0 }, indent: { firstLine: 0 }, children: [new TextRun({ text, bold: !!opts.head, size: 18 })] })] });
const row = (cells, opts = {}) => new TableRow({ children: cells, cantSplit: true, ...opts });
// Column widths, normalised to the text width in twips. Passing columnWidths is what
// makes docx-js emit a <w:tblGrid>; without one LibreOffice discards the per-cell
// percentages and distributes every column evenly. Normalising by the sum rather than
// assuming the widths total 100 is what the Qilvayas generators do, and it means a
// widths array never has to be arithmetic-checked by hand.
const CW = (w) => { const t = w.reduce((a, b) => a + b, 0); return w.map((x) => Math.round(9026 * x / t)); };
const table = (headers, widths, rows) => new Table({ layout: TableLayoutType.FIXED, columnWidths: CW(widths), width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ row(headers.map((h, i) => cell(h, { head: true, w: widths[i] })), { tableHeader: true }), ...rows.map(r => row(r.map((v, i) => cell(v, { w: widths[i] })))) ] });

const mod = (v) => { const m = Math.floor((v - 10) / 2); return (m >= 0 ? "+" : "\u2212") + Math.abs(m); };
const abCell = (text, bold) => new TableCell({ width: { size: 16.6, type: WidthType.PERCENTAGE }, shading: bold ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40, before: 40 }, indent: { firstLine: 0 }, keepNext: !!bold, children: [new TextRun({ text, bold: !!bold, size: 20 })] })] });
const SB = (d) => { const out = []; out.push(new Paragraph({ spacing: { before: 240, after: 40 }, children: [new TextRun({ text: d.name, bold: true, size: 26, color: "5B1F1F" })] })); out.push(PS([{ t: d.meta, i: true }], { spacing: { after: 120 } })); out.push(B("Armor Class:", d.ac)); out.push(B("Hit Points:", d.hp)); out.push(B("Speed:", d.speed)); out.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ new TableRow({ cantSplit: true, tableHeader: true, children: ["STR","DEX","CON","INT","WIS","CHA"].map(h => abCell(h, true)) }), new TableRow({ cantSplit: true, children: [d.str,d.dex,d.con,d.int,d.wis,d.cha].map(v => abCell(v + " (" + mod(v) + ")")) }) ] })); out.push(P("", { spacing: { after: 60 } })); if (d.saves) out.push(B("Saving Throws:", d.saves)); if (d.skills) out.push(B("Skills:", d.skills)); if (d.senses) out.push(B("Senses:", d.senses)); if (d.langs) out.push(B("Languages:", d.langs)); out.push(B("Challenge:", d.cr)); (d.traits||[]).forEach(t => out.push(PS([{ t: t.n + ". ", b: true, i: true }, { t: t.t }]))); if (d.actions && d.actions.length) { out.push(PS([{ t: "ACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.actions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } if (d.reactions && d.reactions.length) { out.push(PS([{ t: "REACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.reactions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } return out; };


// ---------- content ----------
const c = [];

c.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: "The Mountain Road", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module Two, Mountain Road Variant", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("Run this module if the party chose the mountain road at the end of Module One. If they chose the sea road, use Module 2A instead \u2014 the two are mutually exclusive, and only one is ever played at a given table. The coalition marches overland, through passes and through realms that hold nominal allegiance to the crusade and practically resent every column that crosses their land. This is the road that historically killed an emperor\u2019s army; nothing here is that scale, but the module is built from the same material \u2014 attrition, weather, and a welcome that is thinner on the ground than the coalition\u2019s letters of passage promised. Core scenes run three and a half to four hours; Optional Content fills out the rest of a five-hour session and can be cut cleanly if the table is short on time."));

c.push(H2("The Shape of an Overland War"));

c.push(P("Play the march as a slow accumulation of small costs rather than one large danger. Nobody the party meets on this road is Vale\u2019s, or has ever heard of him. A local baron who charges an illegal toll, a mountain crossing that kills through carelessness rather than malice, and a column that is measurably thinner at the far end than it was at the start \u2014 that is the whole threat, and it should feel like exactly that: nothing dramatic happening, and the column arriving worn down anyway. Save the wonder and the true danger for Elduvaine; this road\u2019s job is to make the party, and the table, feel the distance they have actually crossed to get there."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 19, 51],
  [
    ["1. The Long Column", "20\u201330 min", "Establishes the grind of the march before anything goes wrong."],
    ["2. The Ashgate Ford", "30\u201345 min", "A hazard scene, not a fight. See Tiered Skill DCs below."],
    ["3. Baron Vell\u2019s Toll", "90\u2013120 min", "Negotiation, and combat if it comes to that. DC table and stat blocks below."],
    ["4. The Cost of the Road", "30\u201345 min", "What the march has spent, and the road onward \u2014 Branch Ledger entry."],
    ["Optional Content", "30\u201345 min", "Run if the table has time; cut cleanly if not."]
  ]
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("Nothing on this road is Vale\u2019s doing. Baron Osgar Vell holds a toll-keep at the mouth of the Ashgate pass under a technicality of allegiance to Oksitan that nobody in Oksitan\u2019s own court would recognize as binding, and has been quietly taxing every column through his valley for longer than the coalition\u2019s maps have been current. The Ashgate Ford is not cursed or watched; it is simply a fast, cold, early-season river that has killed careless travellers for as long as there has been a road beside it. The column\u2019s losses this module \u2014 to weather, to the ford, to Vell\u2019s harassment if it comes to a fight \u2014 are the ordinary cost of moving an army overland, and should be played as exactly that: unglamorous, and real."));

c.push(PS([DM("DM Only: "), { t: "the Ashgate Ford in Scene 2 is deliberately written to echo the second king\u2019s death, which the party will learn about in Module 5 as news from the road they did not take. Do not explain this echo at the table. If a player nearly loses a companion to the water here and later hears that a king died the same way on the other road, let them draw that line themselves \u2014 it lands harder unexplained." }]));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: The Long Column"));

c.push(P("However the party spends their days on the march \u2014 scouting ahead, riding with the Harrowmark levies, keeping their own company \u2014 give the table a real sense of the grind before anything happens: cold breakfasts eaten walking, a supply wagon\u2019s wheel that takes half a morning to fix, boots that were dry once and no longer remember it."));

c.push(BOX("The column does not so much travel as accumulate distance. Yesterday looked like today, which will look like tomorrow \u2014 until the quartermasters start arguing in low voices about a toll-keep ahead that was not on anyone\u2019s map three years ago, and the column\u2019s pace, without anyone giving an order, gets a little more careful."));

c.push(P("Play this scene short and let it do exactly one job: establish that the march has already cost something before the party reaches either the ford or the baron. A DM who wants more texture here can pull from Optional Content below rather than slowing this scene down."));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: The Ashgate Ford"));

c.push(P("The road drops into a valley and the column\u2019s scouts report the obvious problem: the ford that should be a shin-deep crossing is running high and fast with early snowmelt, and the column\u2019s wagons and the coalition\u2019s heavier troops cannot simply walk it."));

c.push(BOX("The water is the color of old iron and moves like something with a purpose. A supply wagon\u2019s ox baulks at the bank and has to be led in blindfolded; halfway across, the current takes the wagon sideways for one long, silent moment before the ropes hold and it lurches up onto the far bank, half its load gone downstream."));

c.push(P("This is a hazard scene, not a fight. Let the party help however they choose \u2014 anchoring ropes, calming panicked animals, physically steadying wagons and people through the crossing \u2014 and use the Tiered Skill DCs below to resolve it. Nobody needs to die here for the scene to matter; a lost wagon, a soaked and furious quartermaster, or a near-miss that everyone in earshot will still be talking about that evening is enough. If a PC fails badly and the table wants real stakes, a companion NPC or a nameless soldier can be swept downstream and require a rescue (see DC table) rather than simply be lost \u2014 reserve an actual death here for a table that has clearly bought into the danger and wants it to be real."));

// -------------------------------------------------------------- Skill DCs (ford)
c.push(H3("The Ford: Tiered Skill DCs"));

c.push(table(
  ["Task", "Skill", "DC", "Tier"],
  [44, 26, 10, 20],
  [
    ["Read the ford and find the safest line across", "Survival / Nature", "13", "Moderate"],
    ["Steady a wagon or animal mid-crossing", "Athletics / Animal Handling", "13", "Moderate"],
    ["Anchor a rope line the column can cross on", "Athletics", "10", "Easy"],
    ["Rescue someone the current has taken", "Athletics (Strength save DC 13 for the one swept)", "16", "Hard"]
  ]
));

// ---------------------------------------------------------------- Scene 3
c.push(H2("Scene 3: Baron Vell\u2019s Toll"));

c.push(P("A day beyond the ford, the valley narrows to a single defensible road beneath a toll-keep flying a banner that claims allegiance to Oksitan in a way no one in the column entirely believes. Baron Osgar Vell rides out to meet the column himself, backed by mountain scouts who were watching from the ridgeline well before anyone in the column noticed them."));

c.push(BOX("\u201COksitan\u2019s allies pass free,\u201D Vell says, in the tone of a man who has said it many times and means something different by it each time. \u201CEveryone else pays the Ashgate toll. I\u2019d hate for there to be some confusion about which one you are.\u201D His scouts, unhurried, have already found the high ground on both sides of the road."));

c.push(P("Vell\u2019s scouts are a mixed company of humans, dwarves and one very bored half-orc, all of them local, all of them paid, and none of them under any illusion about what they are guarding. If the party gets one of them talking, the honest complaint is not about the coalition at all \u2014 it is about the kobolds in the old workings above the pass, who have been stealing tools for a year and whom nobody has been paid enough to deal with."));

c.push(H3("Running the Scene"));

c.push(P("Vell will negotiate, and a table that wants to avoid a fight in this narrow, badly-chosen ground should be able to. A successful Charisma (Persuasion) check against DC 16, or convincing proof of a coalition writ of passage the party can produce or talk their way into being believed to have, gets the column through for a token payment. A successful Wisdom (Insight) check against DC 13 reveals that Vell is bluffing about his numbers \u2014 he has enough scouts to make a fight costly, not enough to actually hold the pass against a coalition column, and he knows it."));

c.push(P("If the party wants to avoid Vell entirely rather than pay or fight him, a DC 16 Survival check (working from a scout\u2019s report or the party\u2019s own reconnaissance) finds a longer goat-track around the toll-keep that costs the column an extra half-day but avoids the confrontation altogether \u2014 a different flavor of nonviolent resolution than talking Vell down, and one some tables will prefer."));

c.push(P("If it comes to violence \u2014 Vell refuses reasonable terms, the party attacks, or negotiations simply fail \u2014 his scouts fight from cover on the high ground rather than committing to open melee, and break off and flee once Vell himself is dropped, captured, or clearly beaten; they are paid skirmishers, not zealots, and have no reason to die for a toll-keep. This should read as a real skirmish in bad terrain, not a formality; see Scaling the Fight for why it is not meant to be this module\u2019s hardest moment regardless."));

c.push(H3("Scaling the Fight"));

c.push(P("Baron Vell is the SRD Bandit Captain (CR 2, 450 XP) renamed and reflavored; his scouts are the SRD Scout (CR 1/2, 100 XP each), both taken from the SRD unaltered. Run with Vell plus three Scouts \u2014 four total monsters, inside the 3\u20136 monster band, so no table\u2019s party size crosses a multiplier boundary on its own."));

c.push(table(
  ["PCs", "Mult.", "Adj. XP", "Medium", "Reads as"],
  [12, 16, 16, 19, 37],
  [
    ["4", "750", "\u00D72", "1,500", "2,000", "Easy\u2013Medium"],
    ["5", "750", "\u00D72", "1,500", "2,500", "Easy\u2013Medium, softer"],
    ["6", "750", "\u00D71.5", "1,125", "3,000", "below Easy"]
  ]
));

c.push(P("As with the wyvern in Module One and Thane in Module 2A, this is not meant to be a Hard or Deadly encounter, and the six-player row is not a problem to fix by adding a fourth Scout \u2014 crossing from 4 to 5 monsters is safe, but a table that keeps adding bodies for a larger party risks crossing into the 7\u201310 monster band, which jumps the multiplier up a full step and can overcorrect badly. If a table of six wants more bite, the terrain itself is the honest lever: Vell\u2019s scouts firing from height and cover, forcing difficult ground or cover checks to close the distance, does more to raise the real difficulty than another body would."));

// ---------------------------------------------------------------- Stat Blocks
c.push(H2("Stat Blocks"));

c.push(...SB({
  name: "Baron Osgar Vell",
  meta: "Medium humanoid (human), any non-lawful alignment \u2014 SRD Bandit Captain, renamed",
  ac: "15 (studded leather armor)",
  hp: "65 (10d8 + 20)",
  speed: "30 ft.",
  str: 15, dex: 16, con: 14, int: 14, wis: 11, cha: 14,
  senses: "passive Perception 10",
  saves: "Str +4, Dex +5, Wis +2",
  skills: "Athletics +4, Deception +4",
  langs: "Common",
  cr: "2 (450 XP)",
  actions: [
    { n: "Multiattack", t: "Vell makes three melee attacks: two with his scimitar and one with his dagger. Or he makes two ranged attacks with his daggers." },
    { n: "Scimitar", t: "Melee Weapon Attack: +5 to hit, reach 5 ft., one target. Hit: 6 (1d6 + 3) slashing damage." },
    { n: "Dagger", t: "Melee or Ranged Weapon Attack: +5 to hit, reach 5 ft. or range 20/60 ft., one target. Hit: 5 (1d4 + 3) piercing damage." }
  ],
  reactions: [
    { n: "Parry", t: "Vell adds 2 to his AC against one melee attack that would hit him. To do so, he must see the attacker and be wielding a melee weapon." }
  ]
}));

c.push(...SB({
  name: "Ashgate Scout",
  meta: "Medium humanoid (any race), any alignment \u2014 SRD Scout, unmodified",
  ac: "13 (leather armor)",
  hp: "16 (3d8 + 3)",
  speed: "30 ft.",
  str: 11, dex: 14, con: 12, int: 11, wis: 13, cha: 11,
  skills: "Nature +4, Perception +5, Stealth +6, Survival +5",
  senses: "passive Perception 15",
  langs: "Common",
  cr: "1/2 (100 XP)",
  traits: [
    { n: "Keen Hearing and Sight", t: "The scout has advantage on Wisdom (Perception) checks that rely on hearing or sight." }
  ],
  actions: [
    { n: "Multiattack", t: "The scout makes two melee attacks or two ranged attacks." },
    { n: "Shortsword", t: "Melee Weapon Attack: +4 to hit, reach 5 ft., one target. Hit: 5 (1d6 + 2) piercing damage." },
    { n: "Longbow", t: "Ranged Weapon Attack: +4 to hit, range 150/600 ft., one target. Hit: 6 (1d8 + 2) piercing damage." }
  ]
}));

// -------------------------------------------------------------- Skill DCs (toll)
c.push(H2("Tiered Skill DCs"));

c.push(P("Easy 10, Moderate 13, Hard 16, matching the tiers used throughout this campaign."));

c.push(table(
  ["Task", "Skill", "DC", "Tier"],
  [44, 26, 10, 20],
  [
    ["Talk Vell down to a token toll", "Persuasion", "16", "Hard"],
    ["Read that Vell\u2019s numbers are a bluff", "Insight", "13", "Moderate"],
    ["Find a way around the toll-keep entirely", "Survival", "16", "Hard"],
    ["Spot the scouts on the ridgeline before they act", "Perception", "16", "Hard (their passive Perception is 15)"]
  ]
));

// ---------------------------------------------------------------- Scene 4
c.push(H2("Scene 4: The Cost of the Road"));

c.push(P("However the ford and the toll-keep were settled, the column presses on beyond Ashgate with a tally the quartermasters keep quietly rather than announce: a wagon lost or delayed, a handful of soldiers hurt or worse, and whatever the crossing with Vell cost in time, coin, or blood. None of it threatens the crusade. All of it is real."));

c.push(BOX("Garrick Hollow, the column\u2019s hired guide, falls into step beside the party as the toll-keep drops out of sight behind them. \u201CEvery mile past here I know less than the mile before it,\u201D he says, without much apparent concern. \u201CWhich is either a comfort or it isn\u2019t, depending on how you\u2019ve found my company so far.\u201D"));

c.push(P("This is where the module ends, and where the DM should let the weight of the road actually register \u2014 a brief moment of the column taking stock, tending its wounded, and continuing, rather than a triumphant scene. Hand off directly to Module 5 for the approach to Vindana, where the sea road\u2019s own losses \u2014 including the second king\u2019s \u2014 catch up with the party as news from the road not taken."));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("Puzzles and Set Pieces"));

c.push(P("The Back of the Charter expands the toll negotiation and adds ten minutes. The Cold Stair is how the march is run. The Signposted Corridor belongs to the Old Workings and stays in Optional Content, where it is worth sixty minutes to a table that wants a dungeon."));

c.push(H2("The Puzzle: The Back of the Charter"));

c.push(P("Baron Vell\u2019s toll is legal, and that is the whole difficulty with Baron Vell. The charter is real, it is four generations old, it hangs framed in the keep\u2019s hall, and it says what he may charge. He will show it to anybody. He is proud of it."));

c.push(P("He has never read the back."));

c.push(B("What the front says:", "the holder of Ashgate keep may levy a toll on the ford, at rates set out in a schedule, per head and per axle, in perpetuity."));

c.push(B("What the back says:", "the charter is granted in exchange for the keep maintaining the ford, the causeway and the winter marks, and for the keep passing free of toll any body of men raised under a summons of the Crown."));

c.push(P("Getting to it requires noticing that a document hanging in a frame has a reverse, which is a DC 12 Intelligence (Investigation) check or simply a player asking. Getting it out of the frame without Vell\u2019s permission is a DC 15 Dexterity (Sleight of Hand) check. Getting Vell to take it out himself is a DC 16 Charisma (Persuasion) check and is much the better scene, because he will do it, and he will read it, and the party will watch him work out what it says."));

c.push(PS([DM("DM Only: "), { t: "Vell is not a fraud and has not been cheating anybody. His great-grandfather knew about the clause; his grandfather did not; the family has collected in good faith for two generations on a document nobody had turned over. When he reads it he goes quiet, and then he honours it, because he is a man whose entire identity is that his charter is legitimate and he has just found out what legitimate costs. Let the party win this without humiliating him. If they humiliate him anyway, he still honours it, and the column still crosses, and something has been broken that did not need to be." }]));

c.push(H2("The Puzzle: The Signposted Corridor"));

c.push(P("The kobolds of the Old Workings have had two generations and nothing else to do. Every junction in the upper gallery is trapped, and every trap is signposted \u2014 carefully, legibly, and in Draconic, which they do not consider a trick and which everyone who has died here has considered one."));

c.push(P("The signs are honest. That is the puzzle. A party that can read Draconic walks through the entire complex unharmed; a party that cannot must work out that the marks are warnings rather than decoration, and then work out which is which by watching what the kobolds themselves do."));

c.push(table(
  ["The mark", "What it says", "What is actually there"],
  [22, 30, 48],
  [
    ["Three scratches, descending", "Falling", "A covered pit, 20 ft. DC 13 Dexterity save or 2d6 damage and prone."],
    ["A circle with a line through", "Do not stand", "Pressure plate. DC 14 Dexterity save or a rock fall for 3d6 bludgeoning in a 10-ft. square."],
    ["A wave", "Water", "The flooded lower gallery. Not a trap. A genuine warning, and the thing they want help with."],
    ["Two dots and a stroke", "Ours, do not touch", "A larder. It is a larder. There is nothing else to it and a paranoid party will spend twenty minutes on it."],
    ["A hand, open", "Come in, talk", "The negotiating chamber. They have prepared a speech and are hurt if nobody hears it."]
  ]
));

c.push(PS([DM("DM Only: "), { t: "the whole encounter tips on whether a party assumes signposting is malice. If they blunder through, they take real damage from traps that told them, in writing, what they were. If they stop and think, they get sixty kobolds who would very much like to be left alone, will pay in silver they have no use for, and want somebody to do something about what is in the flooded lower gallery. Reward the reading. Do not punish the fighting, but do not soften it either." }]));

c.push(H2("Set Piece: The Cold Stair in Weather"));

c.push(P("Eleven miles of switchback cut into the eastern face, wind coming across rather than along, no water for the middle four miles and no shelter for the last three. Run it as an environment that is trying to kill the column, because it is."));

c.push(BOX("\u201CThe wind does not gust here. It leans. It has been leaning against your left shoulder for six hours and you have been walking with your body at an angle for so long that when the road turns and it stops, you stagger, and the man behind you laughs, and then it turns again and he stops laughing.\u201D"));

c.push(B("The hazard.", "Each of the three days on the Stair, every character makes a DC 12 Constitution saving throw at nightfall. On a failure they gain one level of exhaustion. Characters who reach an intact refuge hut, or who are sheltered by somebody who plans, make the save with advantage; characters who pushed on past the last hut make it with disadvantage."));

c.push(B("The problem.", "Somebody has been stripping the refuge huts for roofing timber. Two of the four are open to the sky. This is a small crime at sea level and a lethal one at five thousand feet, and the party can find out who is doing it in an afternoon: it is the Kir Halloway carters, who are cold, who are also right that nobody has repaired the huts in nine years, and who are entirely prepared to be shouted at."));

c.push(B("The set piece.", "On the second night, a section of the column ahead does not make the hut. Forty people, in the open, in the dark, three quarters of a mile up the road. The party can go back for them. Every hour they spend doing it is an hour they are also in the open, and the DM should track their own exhaustion honestly and out loud."));

c.push(H1("NPC Profiles"));

c.push(H2("Garrick Hollow"));
c.push(P("A half-elf hired out of the last town before the passes, paid to know these mountains better than the coalition\u2019s own maps do, which he does. Speech: unbothered, faintly dry, prone to stating bad news as though it were simply weather. Genuinely good at his job, and genuinely indifferent to whether anyone finds that reassuring."));
c.push(P("Open thread: Garrick knows this road and, plausibly, others \u2014 a DM can bring him back as a recurring guide for any later overland stretch of the campaign, or use him as the source of a rumor or shortcut when the party needs one and has no other way to justify finding it."));

c.push(H2("Baron Osgar Vell"));
c.push(P("Comfortable in bad terrain the way a man is comfortable in a home he built himself. Speech: measured, faintly amused, treats the toll as a fact of nature rather than a crime. Not personally vicious \u2014 his scouts are paid, not fanatical, and he has no interest in a fight he does not expect to win."));
c.push(P("Open thread: if he survives \u2014 paid off, talked down, or simply outmaneuvered \u2014 Vell keeps his keep and his technicality of allegiance, and a DM can bring him back later as a minor, recurring complication on any road that passes near Ashgate, or as a source of local intelligence if the party ever needs a favor from someone who owes them one."));

// --------------------------------------------------------------- Optional
c.push(H1("Optional Content"));

c.push(H2("The Old Workings"));

c.push(P("Above the toll-keep, half a day off the road, an abandoned dwarven working has been reoccupied by kobolds \u2014 forty or so, organised, and responsible for most of what the valley has lost to theft in the past year. They are not a threat to a coalition column and know it; what they are is a warren of pit traps, tripped rockfalls and very good tunnel discipline in a space too tight for anyone to bring numbers to bear."));

c.push(P("Run six to ten kobolds (SRD, CR 1/8, 25 XP each) in terrain that does all their work for them: darkvision against a party carrying light, pack tactics wherever two can reach the same target, and a readiness to collapse a passage and leave rather than lose. The honest resolution here is not slaughter \u2014 a party that offers the kobolds the tools instead of taking them back can buy a guide through the workings and out above the toll-keep entirely, which is a third route past Vell that no one in the column has thought of."));

c.push(P("If a table wants something heavier on this road instead, an ogre (SRD, CR 2, 450 XP) has been working the same stretch of scree for a season and is a straightforward, honest, dangerous obstacle for a party of this level \u2014 the mountain road\u2019s equivalent of weather with arms."));

c.push(H2("What the Column Talks About"));
c.push(P("Before Ashgate, let the party spend some idle marching time among the coalition\u2019s ordinary soldiers \u2014 Harrowmark levies, and the first Oksitan or Auberitz troops the party has walked near long enough to actually talk to. Play this for texture: homesickness, rumors about Elduvaine that are mostly wrong, and the particular, unglamorous humor of people who have been walking together for weeks. No mechanical stakes."));

c.push(H2("The Wounded Wagon"));
c.push(P("If the table wants more to do before the ford, a supply wagon has thrown a wheel on bad road and needs help before the column can move again \u2014 a short, low-stakes problem-solving scene (Investigation or Athletics, DC 10) that adds texture and a chance for a less combat-oriented character to shine, without adding a beat the module needs."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("What the ford cost.", "A wagon, a delay, a near-miss, or an actual loss \u2014 track which, and how the party responded. This has no mechanical follow-up but is worth a callback when the drowned king\u2019s fate reaches the party in Module 5."));
c.push(BUL("How Vell\u2019s toll was settled.", "Paid, talked down, bypassed, or fought \u2014 track which. A bypassed or talked-down Vell keeps his keep and his technicality of allegiance, which a DM can raise again if the coalition\u2019s relationship with Oksitan becomes a live question later in the campaign."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BUL("Vell\u2019s writ.", "A document, genuine as far as it goes, establishing his technicality of allegiance to Oksitan. Not valuable in itself, but a useful thread if a DM wants to raise the coalition\u2019s internal reliability as a question later in the campaign."));
c.push(BUL("A mountain-forged blade.", "Taken from Vell or bought through Garrick\u2019s local contacts \u2014 a +1 weapon of its type, dwarven work out of the old valley forges, unremarkable to look at and better than anything the party marched in with."));

c.push(BUL("Kobold tunnel-charts.", "If the old workings were dealt with peaceably \u2014 scratched on slate, entirely accurate, and a standing shortcut through the Ashgate range that a DM can honour any time the campaign comes back this way."));
c.push(BUL("Salvage from the ford.", "Whatever of the lost wagon\u2019s cargo the party recovers or the column can spare from what remains \u2014 modest, practical supplies rather than treasure, playable as a small easing of the crusade\u2019s logistics rather than a coin payout at 5th level.", { keepNext: true }));

// -------------------------------------------------------------- Refrain
c.push(H1("The Refrain"));

c.push(VERSE([
  "By thought, and by word, and by deed,",
  "the king\u2019s own chosen kept their creed.",
  "Far from home, where the quiet land lay,",
  "they held the line, and would not stray."
]));

const doc = new Document({
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  styles: {
    default: { document: { run: { font: "Georgia", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, font: "Georgia", color: "3B2F2F" }, paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 0, keepNext: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: "Georgia", color: "3B2F2F" }, paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1, keepNext: true } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 22, bold: true, font: "Georgia", color: "3B2F2F" }, paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2, keepNext: true } }
    ]
  },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } }, children: c }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(stagePath("KC_Module02B_TheMountainRoad.docx"), buf);
  console.log("Written.");
});
