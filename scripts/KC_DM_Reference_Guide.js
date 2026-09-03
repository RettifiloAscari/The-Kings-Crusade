// KC_DM_Reference_Guide.js -- The DM Reference Guide.
//
// Canon lives here. corpus/ and documents/ are generated from this file and are
// never edited by hand. See CLAUDE.md for the sign-off rules: anything in the
// "Not yet decided" table must not appear here until it has been approved.
//
// ESCAPE CONVENTION: all prose lives as \uXXXX escapes, never as literal
// typographic characters. tools/build.sh fails the build if that slips. If
// hand-typing an escape, use ONE backslash -- a doubled backslash compiles
// clean and passes the non-ASCII scanner but leaks literal text into the PDF.
//
// SINGLE-COLUMN by design (SINGLE_COL_MATCH in tools/build.sh already points
// at this document\u2019s basename): wide scannable tables, a stat block index,
// and the Branch Ledger -- every Diverging Paths entry from all eleven module
// slots, compiled here with a blank column for what actually happened at a
// given table. Regenerate this file whenever a new module adds its own
// Diverging Paths section; do not let the ledger drift from the modules it
// summarizes.

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
  children: [new TextRun({ text, italics: true })]
});

const VERSE = (lines) => new Paragraph({
  spacing: { before: 120, after: 160 },
  shading: { type: "clear", fill: "F3EFE4" },
  indent: { left: 220, right: 220, firstLine: 0 },   // same fix as BOX -- see its comment
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
  children: [new TextRun({ text: "The DM Reference Guide", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("This document is the table\u2019s working reference, not a fifth telling of the setting or the story. It gathers three things a DM needs at hand and does not want to hunt for across seventeen other files: a one-page map of the whole campaign\u2019s shape, an index of every stat block already built so a name can be found without a search, a table of where every puzzle in the campaign is and what solves it, and the Branch Ledger \u2014 the compiled record of every tracked divergence, module by module, with a blank column for what actually happened at this table. Everything here is drawn from the sourcebook, the gazetteer, the bestiary and the eleven module slots; nothing here is new canon, and none of the campaign\u2019s deliberately open questions have been quietly settled to make this document tidier."));

c.push(P("Use it during play as a quick-lookup: which module a name belongs to, what CR a recurring stat block sits at, what got decided last session that this session might reference. Use the Branch Ledger after every session, not just at the end of the campaign \u2014 an entry filled in while it is fresh is worth more than one reconstructed from memory two modules later."));

// ------------------------------------------------------ Campaign at a Glance
c.push(H1("Campaign at a Glance"));

c.push(P("Eleven module slots (twelve files \u2014 Module Two forks by road and only one half is played). Starting level 5, milestone advancement, built for a party of 4\u20136. Each module\u2019s own Overview carries a scene-by-scene pacing table against its five-hour budget; this table is the one-line shape of the whole march."));

c.push(table(
  ["Module", "Title", "Core Content"],
  [12, 30, 58],
  [
    ["1", "The Muster", "Duncarrow; the wyvern at Greywatch; the road choice \u2014 sea or mountain."],
    ["2A", "The Sea Road", "Warden Thane\u2019s hold; the port of Calanthe and its fate."],
    ["2B", "The Mountain Road", "The Ashgate Ford; Baron Vell\u2019s toll-keep."],
    ["3", "Landfall", "First ground in occupied Elduvaine; the checkpoint; Wyn Alder."],
    ["4", "The Coalition", "Sennoch Hall rescue; the Ward freed; Doria Kell and Norvatch."],
    ["5", "The Road to Vindana", "The second king\u2019s drowning resolved; mourning at the Standing Water."],
    ["6", "The Siege of Vindana: Investment", "Engines raised; the supply raid; First Assault."],
    ["7", "The Siege of Vindana: The Breaking", "Marshal Drell\u2019s fate; the city breaks; Xavier becomes the Wyvernheart."],
    ["8", "Held Ground", "Securing Vindana; the Magistrate freed."],
    ["9", "The Field Battle", "General Voss\u2019s army met in the open and broken; Tam Ondry."],
    ["10", "The Approach", "The march on Caer Ysolde; the Regent; Norvatch names its price."],
    ["11", "The Decision at the Gates (Final)", "Caer Ysolde falls; Maedoc Vale confronted; hold or set down."]
  ]
));

// ---------------------------------------------------------------- Stat Blocks
c.push(H1("Peoples at a Glance"));

c.push(P("Every realm is mixed; the table is who a party actually meets there and in what proportion. The load-bearing rule is that nobody in this campaign is their species \u2014 dwarves hunt wyverns in Harrowmark and keep contracts in Norvatch, and orcs fight on both sides of this war without a single NPC remarking on it."));

c.push(table(
  ["Realm or faction", "Who you meet", "What the place makes of them"],
  [20, 32, 48],
  [
    ["Harrowmark", "Humans, dwarves, orcs; half-orcs and half-elves throughout", "Wyvern country. Dwarves on the pikes, orcs on the ropes, nine centuries unchanged. All magic is worked, professional and carried in a person, because the land does nothing."],
    ["Elduvaine", "Elves, gnomes, halflings, humans \u2014 plus dryads, sprites and satyrs", "Elves keep the Kept Season orchards; gnomes clerk the Archive and cut the light-stone; halflings hold the river parishes and supply the whole kingdom\u2019s manners."],
    ["Oksitan", "Humans, dragonborn nobility, half-elves at court", "River and horse country. Dragonborn houses have held the same fords for longer than the records go back."],
    ["Auberitz", "Humans, gnomes, halflings", "The duchy that builds things. The coalition\u2019s siege train is theirs, gnome-designed and halfling-quartermastered."],
    ["Norvatch", "Dwarves, tieflings, humans", "Guild-law country. A bargain is written, witnessed and binding, and honoured to the letter no matter what the letter turns out to have meant."],
    ["The occupation", "Hobgoblin core; orc, human, goblin auxiliaries; Elduvish clerks", "A hired legion, not a horde. Permits, a published grain levy, competent officers, no fanatics. They are employees."]
  ]
));

c.push(H1("Stat Block Index"));

c.push(P("Every named stat block built so far, in order of first appearance. \u201cSRD base\u201d names the unmodified SRD creature a block was built from \u2014 renamed blocks keep the base creature\u2019s numbers exactly and change only name, flavor text, and (where noted in the module) minor cosmetic description. The Occupation Guard is defined in Module 3, reprinted in full in Module 4 because that module\u2019s own fight needs it to hand, and pointed at by reference in Modules 6, 7, 9 and 11."));

c.push(table(
  ["Name", "SRD Base", "CR (XP)", "AC", "HP", "Module"],
  [20, 20, 14, 10, 10, 26],
  [
    ["Wyvern", "Wyvern, unmodified", "6 (2,300)", "13", "110", "1 \u2014 The Muster"],
    ["Warden Ivor Thane", "Veteran, renamed", "3 (700)", "17", "58", "2A \u2014 The Sea Road"],
    ["Thane\u2019s Guard", "Bandit, unmodified", "1/8 (25)", "12", "11", "2A \u2014 The Sea Road"],
    ["Baron Osgar Vell", "Bandit Captain, renamed", "2 (450)", "15", "65", "2B \u2014 The Mountain Road"],
    ["Ashgate Scout", "Scout, unmodified", "1/2 (100)", "13", "16", "2B \u2014 The Mountain Road"],
    ["Occupation Guard", "Guard, unmodified", "1/8 (25)", "16", "11", "3 \u2014 Landfall (reused: 4, 6, 7, 9, 11)"],
    ["Marshal Ossian Drell", "Veteran, renamed", "3 (700)", "17", "58", "6 \u2014 Vindana: Investment (returns: 7)"],
    ["General Ilyana Voss", "Gladiator, renamed", "5 (1,800)", "16", "112", "9 \u2014 The Field Battle"],
    ["Maedoc Vale", "Archmage, renamed", "12 (8,400)", "15", "99", "11 \u2014 The Decision at the Gates"]
  ]
));

c.push(P("A stat block\u2019s full traits and actions live only in its own module\u2019s script \u2014 this index is for finding the right file fast, not for running the creature from this page alone."));

c.push(H2("Creatures Called For By Reference"));

c.push(P("These are used by name in the modules without a reprinted block, because they are unmodified SRD entries a DM can look up in seconds. All are 2014 SRD."));

c.push(table(
  ["Creature", "CR (XP)", "Where it is called for"],
  [26, 16, 58],
  [
    ["Kobold", "1/8 (25)", "The old workings above Ashgate (2B, optional); Vindana\u2019s undercity (6). Trap-layers and negotiators, not a slaughter."],
    ["Winter Wolf", "3 (700)", "The Held Winter (5, optional). Hunts the wood\u2019s edges."],
    ["Troll", "5 (1,800)", "The Held Winter (5, optional). Has taken the orchard-keeper\u2019s cottage. Regeneration is the fight."],
    ["Dryad", "1 (200)", "The Held Winter (5, optional). Not an encounter \u2014 a conversation, and the campaign\u2019s grief in one NPC."],
    ["Sprite", "1/4 (50)", "Caerwyn\u2019s orchard hedge (3). Four inches of extremely rude opinion. Pure texture."],
    ["Ogre", "2 (450)", "The mountain road (2B, optional), as an alternative to the kobold workings."]
  ]
));

c.push(H2("The Elduvish Wonders"));

c.push(P("Four habits made portable, and the campaign\u2019s best treasure. Ordinary inside Elduvaine, astonishing everywhere else."));

c.push(table(
  ["Wonder", "What it does", "First available"],
  [24, 52, 24],
  [
    ["Standing-stone lamp", "Gnome-cut stone that holds daylight poured into it and gives it back for hours.", "Module 3, from Caerwyn\u2019s baker"],
    ["Flask of Listening Water", "Holds what is said at its mouth; gives it back once, in the speaker\u2019s own voice.", "Module 5, at the Standing Water"],
    ["Road-token", "Cut from a Willing Road waystone. Explains itself no more than the road does.", "DM\u2019s discretion"],
    ["Kept Season seeds", "A small stand holds whatever season the seeds were sown in.", "DM\u2019s discretion"]
  ]
));

// ---------------------------------------------------------------- Recurring NPCs
c.push(H1("Recurring NPCs"));

c.push(P("Named NPCs the campaign\u2019s own modules flag as \u201copen threads\u201d \u2014 built to be brought back, not to be used once and discarded. Three (the Ward, the Magistrate, the Regent) are members of Elduvaine\u2019s royal house and deliberately carry no proper name yet, and naming them is a decision this campaign has not yet made."));

c.push(table(
  ["Name", "Introduced", "Role"],
  [22, 14, 64],
  [
    ["Xavier III of Harrowmark", "1", "The calling king. Leads in person; not yet the Wyvernheart until Module 7."],
    ["Huntmaster Brenna Vane (dwarf)", "1", "Greywatch\u2019s huntmaster, two centuries old; runs the wyvern hunt; a recurring source of camp levity."],
    ["Sera Vosk (rock gnome)", "2A", "Auberitz quartermaster; a standing logistics and information contact."],
    ["Garrick Hollow (half-elf)", "2B", "Hired mountain guide; a standing source for overland routes and rumor."],
    ["Wyn Alder (gnome)", "3", "An occupation-era clerk at the Landfall checkpoint; the campaign\u2019s first answer to what an ordinary complicit person looks like."],
    ["Doria Kell (tiefling)", "4", "Norvatch factor, and the campaign\u2019s second-largest lever. Her house has been buying Elduvaine by weight for three years; her ledgers are the clock. Names her price in Module 10."],
    ["The Ward (half-elf)", "4", "Royal house, freed at Sennoch Hall. Precise rather than broken; a standing voice for what the liberated want."],
    ["Tam Ondry", "5", "Harrowmark courier attached to the lost column; a standing source for news arriving from elsewhere."],
    ["Marshal Ossian Drell (hobgoblin)", "6", "Commands Vindana\u2019s garrison; a legionary professional on a contract, not a fanatic. Fate resolved in Module 7."],
    ["The Magistrate (gnome)", "8", "Royal house, freed in Vindana; the Ward\u2019s great-aunt by marriage. Fought her captivity with bureaucratic warfare and won."],
    ["General Ilyana Voss (orc)", "9", "Commands Vale\u2019s field army; came up through the same legion as Drell. Fate resolved this module; a captured Voss is a real long-term asset."],
    ["Aveline Ysolde, the Regent (human)", "10", "Royal house, at large; runs what resistance survives and refuses evacuation. Has refused it in writing, twice. Stat block in the bestiary."],
    ["Maelis Ysolde, the Veiled Sovereign (elf)", "11", "Held in Caer Ysolde. Dying at the same rate as the land, because a sovereign is bound to the habits. The campaign\u2019s second clock, and the one that cannot be bought. No stat block, deliberately."],
    ["Ninian Ysolde, the Ward (half-elf)", "4", "Freed at Sennoch Hall. Heir presumptive, and the one who has done the arithmetic on the Promise."],
    ["Ottoline Vahn, the Magistrate (gnome)", "8", "Freed in Vindana. Fought three years of captivity by filing, and won. Holds the only complete record of the occupation."],
    ["Emrys Ysolde, the Envoy (elf)", "\u2014", "Held separately, and not entirely as a prisoner. Has been talking to Vale for three years. Complexity in the cost, not the cause; there is no reveal and no resolution."],
    ["Raimon V of Oksitan (human)", "5", "The second crown, seventy-one and undeterred. Takes the road the party does not and drowns fording the Vaskren. Reaches the party as rumour, then refugees, then a problem."],
    ["Serjeant Hoth (hobgoblin)", "4", "Held Sennoch Hall correctly for three years and would like that written down. Surrenders on terms, with the file."],
    ["Maedoc Vale", "11", "The wizard. Appears in person only here. Whether he is still human is deliberately unresolved."]
  ]
));

// ---------------------------------------------------------------- Branch Ledger
c.push(H1("Faith and Factions at a Glance"));

c.push(P("The Concord holds that the Works made the world and withdrew while it was unfinished. Nine of them, worshipped across Harrowmark, Oksitan, Auberitz and Norvatch. Elduvaine never built a church at all, because you do not raise a temple to ask for an answer in a country where the river answers directly."));

c.push(table(
  ["The Work", "Sphere", "Domains"],
  [26, 46, 28],
  [
    ["Ashet the Anvil", "Craft, making, the honest tool", "Knowledge, War"],
    ["Voran of the Long Road", "Travel, messengers, guest-right", "Life, Trickery"],
    ["Sennet the Witness", "Oath, contract, law, testimony", "Knowledge, Trickery"],
    ["Halevin the Hearth-Kept", "Home, harvest, healing", "Life, Nature"],
    ["Aurine the Unshuttered", "Light, truth, dawn, courage", "Light"],
    ["Duran Ninefold", "War as discipline; the held line", "War"],
    ["Threnn Greywater", "Sea, storm, river, the drowned", "Tempest"],
    ["Ossuar the Quiet Warden", "Death, the grave, remembrance", "Death"],
    ["Saveth of the Green Verge", "Wilds, beasts, the seasons", "Nature"]
  ]
));

c.push(table(
  ["Faction", "Standing, in three steps", "What the top step gets you"],
  [22, 36, 42],
  [
    ["The Crusade", "Sworn \u00b7 Lance \u00b7 Banner of the Call", "Xavier without an appointment, and the standing to ask the coalition for something it does not want to give."],
    ["The Order of the Tenth Work", "Postulant \u00b7 Hand \u00b7 Warden of the Work", "Disciplined troops, healing without price, and an expectation you will not like."],
    ["The Ysolde Remnant", "Known \u00b7 Trusted \u00b7 Named", "Safe houses across the Braid, forged permits, and Aveline answering the same day."],
    ["The Sixth Free Legion", "Noted \u00b7 Respected \u00b7 Owed", "Parley honoured, prisoners exchanged, and an officer who stays bought."],
    ["House Kell of Norvatch", "Client \u00b7 Factor\u2019s Guest \u00b7 Signatory", "Anything buyable, on time, plus three years of ledgers."],
    ["The Unbound Clerks", "Enquirer \u00b7 Reader \u00b7 Keeper\u2019s Friend", "The only people alive who can say what Vale has already read."]
  ]
));

c.push(PS([DM("DM Only: "), { t: "the Tenth Work is where this campaign keeps its complexity, and it is on the party\u2019s own side of the line. They are allies, they will die holding a wall for the party, and they intend to do something to the liberated kingdom that a great many Elduvish would call a second occupation with better manners. Do not make them hypocrites and do not give them a secret plan." }]));

c.push(H1("The Puzzle Index"));

c.push(P("Every puzzle in the campaign, where it sits, and what solves it. Solutions are listed so a DM can recognise a good answer arriving from an unexpected direction, not so that only these answers count."));

c.push(table(
  ["Module", "Puzzle", "What solves it"],
  [12, 30, 58],
  [
    ["1", "The Muster Roll", "Count the billets, walk the range at dawn, or ask a Stannock pikeman how many of his village came."],
    ["1", "The Ledge", "Rope from above, bait it off, take the shepherds first, or bring the overhang down. No preferred answer."],
    ["2A", "The Moved Light", "Four pieces of evidence. Nobody moved a light; the harbour light was two hours late and Thane\u2019s boats were already crewed."],
    ["2B", "The Back of the Charter", "Notice that a framed document has a reverse. Best outcome is Vell reading it himself."],
    ["2B", "The Signposted Corridor", "Read Draconic, or work out that the marks are honest warnings by watching what the kobolds do."],
    ["3", "Getting Past Wyn Alder", "Be boring, be somebody else, split up to defeat the head count, or go round by the creeks."],
    ["4", "What the Water Heard", "Ten minutes at the ornamental canal. Three watches, not two; the bell rope is frayed."],
    ["5", "Who Is Speaking", "Four voices, none lying, none complete. Not meant to converge."],
    ["6", "What the Kobolds Want", "Grain from the undercroft, a written promise that will actually hold, and both before the route."],
    ["7", "The Terms of the Binding", "It must be commanded aloud. Silence the relay officer, the sound, or take the horn yourself."],
    ["8", "Three Years of Filing", "Filed by issuing office, not subject. The fourth office has no docket mark, which is itself the mark."],
    ["9", "Voss\u2019s Field Orders", "Company, hour, bearing. Doubled numerals are feints; no bearing at all means reserve."],
    ["10", "Reading the Ledgers", "Ordinary cargo is flat. All the increase is light-stone, and light-stone is the medium. About a third is gone."],
    ["11", "Access by Rule", "The rule has no exception and never had one. Vale was never close."]
  ]
));

c.push(H1("The Branch Ledger"));

c.push(P("Every tracked divergence across the campaign, compiled from each module\u2019s own Diverging Paths (DM Only) section. Fill in the last column at the table, as it happens \u2014 this is the record that makes the campaign replayable rather than a railroad with scenery, and it is the reason this document is single-column: these rows need the width."));

c.push(table(
  ["Module", "Divergence", "Options", "What Actually Happened"],
  [8, 27, 33, 32],
  [
    ["1", "The wyvern\u2019s fate", "Killed \u00b7 driven off \u00b7 roped and released", ""],
    ["1", "Branch Ledger entry 1: the road choice", "Sea road (2A) \u00b7 Mountain road (2B)", ""],
    ["2A", "How Thane\u2019s hold is settled", "Paid off \u00b7 talked down \u00b7 fought", ""],
    ["2A", "Calanthe\u2019s fate", "Garrisoned \u00b7 independent \u00b7 hybrid", ""],
    ["2B", "What the ford cost", "A wagon \u00b7 a delay \u00b7 a near-miss \u00b7 an actual loss", ""],
    ["2B", "How Vell\u2019s toll was settled", "Paid \u00b7 talked down \u00b7 bypassed \u00b7 fought", ""],
    ["3", "How the party treated Wyn Alder", "Threatened \u00b7 ignored \u00b7 understood", ""],
    ["3", "Whether the checkpoint turned violent", "Yes (word reaches Caer Ysolde early) \u00b7 No", ""],
    ["4", "How Sennoch Hall was resolved", "Stealth \u00b7 violence \u00b7 a mix", ""],
    ["4", "The party\u2019s standing with Doria Kell", "Face value \u00b7 pressed for more \u00b7 suspicion", ""],
    ["5", "Branch Ledger entry 1, resolved", "The party\u2019s own crossing, set against the second king\u2019s", ""],
    ["5", "How the party engaged with the mourning at the Standing Water", "Spoke at the water \u00b7 listened only \u00b7 kept apart", ""],
    ["6", "How the engines were raised", "Smoothly \u00b7 roughly", ""],
    ["6", "The supply raid\u2019s outcome", "Clean success \u00b7 failed or messy", ""],
    ["6", "How First Assault was fought", "Clear local win \u00b7 costly draw \u00b7 real loss", ""],
    ["7", "How Marshal Drell\u2019s fate was resolved", "Surrendered \u00b7 killed \u00b7 fled", ""],
    ["7", "Scene 3: as written, or a direct dragon fight", "As written \u00b7 direct dragon fight", ""],
    ["8", "How the Magistrate was freed", "Caper \u00b7 force", ""],
    ["8", "Whether the party noticed the pressed clerk", "Yes \u00b7 No", ""],
    ["9", "Which Scene 2 was run", "Tam\u2019s death \u00b7 the third-thread alternate", ""],
    ["9", "General Voss\u2019s fate", "Killed \u00b7 captured \u00b7 fled", ""],
    ["10", "Whether Xavier signed Norvatch\u2019s contract", "Signed \u00b7 refused \u00b7 deferred", ""],
    ["10", "What the party told the Regent, and what the Regent told them", "(Private DM note \u2014 not mechanically tracked)", ""],
    ["11", "The campaign\u2019s final choice", "Held \u00b7 set down", ""],
    ["11", "Vale\u2019s fate", "Killed \u00b7 escaped \u00b7 captured", ""],
    ["1", "What the party did about the muster roll", "Reported it \u00b7 buried it \u00b7 never found it", ""],
    ["1", "How the ledge was solved", "Rope \u00b7 baited off \u00b7 shepherds first \u00b7 overhang brought down", ""],
    ["2A", "Whether the coalition was told the truth about Thane", "Told \u00b7 withheld \u00b7 never established", ""],
    ["2A", "Whether the draught team came off the ship", "Saved \u00b7 hatch closed", ""],
    ["2B", "How Vell came to honour the charter", "He read it himself \u00b7 shown it \u00b7 taken from him \u00b7 never found", ""],
    ["2B", "How the Old Workings went", "Read the marks \u00b7 fought through \u00b7 not entered", ""],
    ["3", "How the party passed the Caerwyn checkpoint", "Boring \u00b7 forged \u00b7 split up \u00b7 went round \u00b7 caught", ""],
    ["4", "What the water at Sennoch Hall gave them", "Used it \u00b7 never tried it", ""],
    ["4", "Hoth\u2019s surrender, and the file", "Terms accepted \u00b7 fought, file burned", ""],
    ["6", "Whether the party opened the under-wall route", "Negotiated \u00b7 fought for it \u00b7 never found it", ""],
    ["6", "Whether the kobolds got their written promise", "Signed and holds \u00b7 signed and worthless \u00b7 refused", ""],
    ["7", "The bound dragon", "Grounded \u00b7 killed \u00b7 neither", ""],
    ["7", "Whether the binding instrument was recovered", "Yes \u2014 names the office \u00b7 No", ""],
    ["8", "Whether the Keeper\u2019s Office drawer was found", "Found \u00b7 missed", ""],
    ["9", "Whether Voss\u2019s reserve was located before it committed", "Found \u00b7 not found", ""],
    ["10", "Whether the party worked out the number", "Yes \u2014 about a third \u00b7 No \u00b7 refused the folio", ""],
    ["11", "Whether the party established what the rule permits", "Established \u00b7 forced the doors \u00b7 never reached them", ""],
    ["11", "Whether the Sixth was bought out at term", "Withdrew \u00b7 declined to sortie \u00b7 fought \u00b7 never approached", ""]
  ]
));

c.push(P("The final choice draws on everything above it \u2014 before running Module 11\u2019s Scene 4, a DM may find it worth reading back across this whole ledger with the table."));

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
  fs.writeFileSync(stagePath("KC_DM_Reference_Guide.docx"), buf);
  console.log("Written.");
});
