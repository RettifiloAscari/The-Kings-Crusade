// KC_Module06_VindanaInvestment.js -- Session Module Six: The Siege of Vindana, Investment.
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
// NOTE: "The Cavalry Screen" (Scene 3) is Harrowmark\u2019s ordinary mounted
// scouts -- horses. The draconic layer is now signed off and canon (see
// CLAUDE.md), and the campaign\u2019s actual wyvern-riders belong to Module Seven,
// where Xavier earns "the Wyvernheart" -- they are deliberately absent here so
// that introduction lands with its own weight. Vale does not appear in this
// module either, continuing the pattern from Module Three: his officers run
// this war, and his own first appearance is saved for later, larger weight.
//
// This module is written to end on a setback -- the first assault fails or
// costs more than it should -- so that Module Seven can be the siege\u2019s actual
// turning point, per the module breakdown\u2019s own shape.

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
  children: [new TextRun({ text: "The Siege of Vindana: Investment", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module Six", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("The coalition arrives at Vindana and begins the work of a siege: surrounding the city, cutting its supply, and preparing the engines that will eventually break its walls. This module is the campaign\u2019s largest set piece so far and the first of two covering Vindana \u2014 this one ends in a real setback, not a victory, so that Module Seven can be the siege\u2019s actual turning point. Core scenes run four to four and a half hours; Optional Content fills out the rest of a five-hour session and can be cut if the table is short on time."));

c.push(H2("A City, Not a Dungeon"));

c.push(P("Play Vindana as a place with its own life going on behind its walls, not a monster to be whittled down. Its defenders are professional soldiers doing an ugly job competently, per the occupation\u2019s established character \u2014 nobody here is a fanatic, and Marshal Ossian Drell, who commands the garrison, is exactly as dangerous as a good officer with a strong position should be. Vale does not appear in this module, continuing the pattern from Landfall: his war is run by people like Drell, and that facelessness is still doing its work this many modules in."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 19, 51],
  [
    ["1. The Lines Close", "30\u201340 min", "First sight of Vindana up close; the siege lines form."],
    ["2. Raising the Engines", "45\u201360 min", "A practical, largely non-combat scene. See Tiered Skill DCs."],
    ["3. The Cavalry Screen", "45\u201360 min", "A horse-raid on Vindana\u2019s supply. DC table and stat blocks below."],
    ["4. First Assault", "60\u201390 min", "A probing attack that does not take the city. The module\u2019s setback."],
    ["Optional Content", "30\u201345 min", "Run if the table has time; cut cleanly if not."]
  ]
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("Vindana is well-garrisoned, well-supplied for the moment, and commanded by an officer who has had three years to prepare for exactly this. Marshal Drell knows the coalition cannot simply walk over his walls, and he is right \u2014 that is the entire point of Module Seven existing as its own module. Nothing that goes wrong for the party in Scene 4 is a trap or a trick; it is simply what a strong, competently held position does to an attacker who has not yet found its weakness. That weakness is Module Seven\u2019s job to find, not this one\u2019s."));

c.push(PS([DM("DM Only: "), { t: "resist ending this module on a clean coalition victory. The whole shape of the siege depends on Investment costing something real, so that the Breaking (Module Seven) has genuine stakes to turn around. A DM running these back to back should let the table feel the setback in Scene 4 before moving on \u2014 do not soften it in the moment to keep spirits up." }]));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: The Lines Close"));

c.push(P("Vindana up close is larger and better defended than it looked from the ridge in Module Five \u2014 walls thick enough to have shrugged off worse than a first coalition army, a harbor still working under the garrison\u2019s control, watch-fires along every tower by the time the coalition\u2019s own lines are staked out around it."));

c.push(BOX("By the second day, Vindana is ringed rather than merely approached \u2014 coalition tents in a wide arc from harbor-mouth to inland hill, cook-fires enough to be seen from the walls, and somewhere above the gatehouse, a banner the party does not recognize, raised deliberately high enough to be read from the coalition\u2019s own lines. Marshal Drell wants them to know exactly who they are besieging."));

c.push(P("Let the party see the siege take shape as a logistics problem as much as a military one: where the coalition\u2019s three contingents camp relative to each other, how supply lines from Caerwyn are organized, and \u2014 if Module Four was played \u2014 whether Doria Kell or her Norvatch contacts have already found a way to profit from a besieging army\u2019s needs."));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: Raising the Engines"));

c.push(P("The engines are Auberitz work and the crews raising them are Auberitz too \u2014 gnome engineers who designed the things and halfling quartermasters who move them, all of whom have strong views about how a siege ought to be conducted and none of whom have been asked. The party will get further here by taking the engineers seriously than by outranking them: the chief of the works is a gnome who has built eleven of these, has never once seen one used the way she intended, and will tell the party exactly what will go wrong an hour before it does."));

c.push(P("A siege needs engines, and building them under threat of sortie is its own kind of work. This scene is largely non-combat by design \u2014 a chance for characters who do not want to spend the whole module fighting to matter directly."));

c.push(BOX("Timber comes in faster than the engineers can shape it and slower than the marshals want it. A trebuchet frame goes up crooked twice before a Harrowmark carpenter, cursing fluently in two languages, gets the joint to hold. Nobody has slept properly in three days, and everybody has an opinion about the range on the mangonels."));

c.push(P("Use the Tiered Skill DCs below to resolve however the party wants to help \u2014 hauling material, correcting an engineer\u2019s error, defending a work party from a probing sortie. Success speeds construction and gives Scene 4 a stronger opening position (advantage on the first round, or a wall section already weakened); failure does not lose the siege, only costs time the table can feel."));

// ---------------------------------------------------------------- Scene 3
c.push(H2("Scene 3: The Cavalry Screen"));

c.push(P("Xavier\u2019s own cavalry \u2014 mounted scouts and raiders, Harrowmark-trained \u2014 range beyond the siege lines to cut Vindana\u2019s remaining supply routes inland. The party can ride with them for a fast, mobile scene very different in texture from the siege lines\u2019 grinding patience."));

c.push(BOX("The riders move like people who trust their horses more than their orders, which is not disrespect so much as long habit \u2014 Harrowmark cavalry has always worked this way, fast and loose and hard to pin down. Their captain, without much ceremony, waves the party into formation as though they had always ridden with the unit."));

c.push(H3("Running the Scene"));

c.push(P("Vindana also has an undercity \u2014 a drain and cistern system older than the walls above it, which the coalition would very much like to use and which the garrison stopped guarding two years ago because something else took it over. Perhaps sixty kobolds (SRD, CR 1/8, 25 XP each) hold the tunnels now, and they have had two years and nothing else to do. Every junction is trapped, every trap is signposted in a language nobody in the coalition reads, and they will trade: the coalition wants a route under the walls, and the kobolds want the garrison\u2019s grain stores and a written promise that nobody comes back down afterwards. Run four to eight of them in any fight, in terrain that lets pack tactics and darkvision do the work, and let the negotiation be the real scene."));

c.push(PS([DM("DM Only: "), { t: "the kobold route is the single best reward available for a party that solves Module Six sideways instead of by force, and it pays off directly in Module Seven \u2014 a party that has the tunnels does not need the postern gate and can open the assault from inside the walls. Do not require it. Do reward it enormously if they find it." }]));

c.push(P("A supply column bound for Vindana \u2014 modest, lightly guarded, moving fast to avoid exactly this \u2014 can be intercepted on the road. This is a fast, mobile engagement; let mounted combat, terrain, and speed matter more than raw numbers. Use the Occupation Guard stat block (Module 3) for the column\u2019s escort, three or four of them, easily overwhelmed by a proper cavalry raid \u2014 the point of this scene is momentum and competence, not danger."));

c.push(P("A successful raid here (the column stopped, captured, or turned back) meaningfully weakens Vindana\u2019s position for Scene 4\u2019s assault \u2014 a DM may grant advantage on one relevant roll during the assault, or simply narrate the garrison as visibly shorter on options. This is the module\u2019s clearest chance for the party to feel unambiguously effective before Scene 4\u2019s setback."));

// -------------------------------------------------------------- Skill DCs
c.push(H2("Tiered Skill DCs"));

c.push(P("Easy 10, Moderate 13, Hard 16, matching the tiers used throughout this campaign."));

c.push(table(
  ["Task", "Skill", "DC", "Tier"],
  [44, 26, 10, 20],
  [
    ["Correct a flawed siege engine joint before it fails", "Investigation / a relevant craft", "13", "Moderate"],
    ["Haul and position engine timber efficiently", "Athletics", "10", "Easy"],
    ["Defend an engine work party from a probing sortie", "Combat or relevant skill, DM\u2019s judgment", "13", "Moderate"],
    ["Track and intercept the supply column before it reaches Vindana", "Survival", "13", "Moderate"],
    ["Read Marshal Drell\u2019s dispositions during First Assault", "Investigation / Perception", "16", "Hard"]
  ]
));

// ---------------------------------------------------------------- Scene 4
c.push(H2("Scene 4: First Assault"));

c.push(P("With the engines raised and the supply raid\u2019s results in hand, the coalition tries the walls once, before the real siege settles into its longer rhythm \u2014 not out of impatience, but to test what Drell\u2019s defense actually looks like under pressure. This is meant to fail, or to succeed only at real cost, and the module should let that land honestly."));

c.push(BOX("The first breach in the outer wall holds for less than a minute before Vindana\u2019s garrison closes on it from three directions at once, and the coalition\u2019s own advance stalls in the gap rather than through it. Somewhere to the left, a company that went in confident comes back considerably smaller than it went."));

c.push(H3("Running the Scene"));

c.push(P("This can be run as a large-scale combat with the party at its center (use Marshal Drell \u2014 see Stat Block \u2014 commanding a knot of Occupation Guards at the point of heaviest fighting) or narrated at the level of the assault as a whole, with the party\u2019s actions determining how badly it goes rather than whether it succeeds. Either way, the assault does not take Vindana. A clean tactical win for the party in their own local fight is entirely possible and should not contradict the larger setback \u2014 they can win their corner of a battle the coalition as a whole does not win, which is itself worth letting them feel."));

c.push(H3("Stat Block"));

c.push(...SB({
  name: "Marshal Ossian Drell",
  meta: "Medium humanoid (hobgoblin), lawful neutral \u2014 SRD Veteran, renamed",
  ac: "17 (splint armor)",
  hp: "58 (9d8 + 18)",
  speed: "30 ft.",
  str: 16, dex: 13, con: 14, int: 10, wis: 11, cha: 10,
  skills: "Athletics +5, Perception +2",
  senses: "passive Perception 12",
  langs: "Common",
  cr: "3 (700 XP)",
  actions: [
    { n: "Multiattack", t: "Drell makes two longsword attacks. If he has a shortsword drawn, he can also make a shortsword attack." },
    { n: "Longsword", t: "Melee Weapon Attack: +5 to hit, reach 5 ft., one target. Hit: 7 (1d8 + 3) slashing damage, or 8 (1d10 + 3) slashing damage if used with two hands." },
    { n: "Shortsword", t: "Melee Weapon Attack: +5 to hit, reach 5 ft., one target. Hit: 6 (1d6 + 3) piercing damage." },
    { n: "Heavy Crossbow", t: "Ranged Weapon Attack: +3 to hit, range 100/400 ft., one target. Hit: 6 (1d10 + 1) piercing damage." }
  ]
}));

c.push(P("Drell does not die or fall back easily in this module \u2014 he is meant to survive Scene 4 regardless of how the fight at the point of contact goes, withdrawing his forces in good order once the breach fails rather than pressing an advantage he does not need. Save a real, decisive confrontation with him for Module Seven if the table wants one."));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("Puzzles and Set Pieces"));

c.push(P("Raising the Engines is how that scene is run. The undercity is additive and substantial \u2014 sixty to ninety minutes, and a table that wants a proper dungeon crawl can take considerably longer. It is the best thing in the module and it is optional, which is deliberate: it pays out in Module Seven for the parties that go looking."));

c.push(H2("The Undercity, Keyed"));

c.push(P("A drain and cistern system older than the walls above it, which the coalition would very much like to use and which the garrison stopped patrolling two years ago. It is the single best reward available to a party that solves this module sideways, and it pays out directly in Module Seven."));

c.push(table(
  ["Area", "What is there", "Notes"],
  [22, 40, 38],
  [
    ["A. The seaward outfall", "Entry. Six feet of water at high tide, none at low.", "Timing is the whole approach. Tide tables are buyable in any harbour tavern."],
    ["B. The signposted run", "300 ft. of trapped corridor, every trap marked in Draconic.", "Same principle as the Old Workings. The marks are honest."],
    ["C. The cistern hall", "A vaulted chamber, forty feet across, waist deep.", "Sound carries perfectly. Anything said here is heard at D."],
    ["D. The warren", "Sixty kobolds. Larders, nests, two generations of improvement.", "Run four to eight in any fight. The negotiation is the real scene."],
    ["E. The old sluice", "A gate mechanism, seized, that once drained the cistern.", "Freeing it is DC 16 Strength or an hour with tools. It matters in Module Seven."],
    ["F. The under-wall", "Where the drain passes beneath the inner wall.", "This is the prize. A route under the walls, into the garrison quarter."],
    ["G. The grain undercroft", "The garrison\u2019s reserve store, reachable only from F.", "What the kobolds want. They cannot reach it and the party can."]
  ]
));

c.push(H2("The Puzzle: What the Kobolds Want"));

c.push(P("They have had two years and nothing else to do, and they have prepared a speech. The negotiation is a three-part problem and every part has a real solution."));

c.push(BUL("They want the grain undercroft (G).", "They cannot get into it, because the only route runs under the wall past a garrison post. The party can. This is a straight trade and it is the easy part."));
c.push(BUL("They want a written promise that nobody comes back down.", "In writing, signed, by somebody with a name. This is harder than it sounds, because the party cannot bind the coalition and the kobolds know what a worthless signature looks like. A Norvatch factor can draft one that will actually hold, which is a genuinely good reason to have kept Doria Kell\u2019s goodwill."));
c.push(BUL("They will not hand over the route until both are done.", "No amount of Persuasion moves this. They have been lied to before, by the garrison, in year one, and they will tell the party about it at length."));

c.push(PS([DM("DM Only: "), { t: "do not require this. Do reward it enormously. A party that has the under-wall route does not need the postern gate in Module Seven and can open the assault from inside the walls, which changes that module\u2019s opening from a wall problem into a street fight and saves a great many coalition lives the DM should mention afterward. A party that simply kills sixty kobolds gets the route as well, and gets it four hours later, and gets no help with the sluice at E, which they will want." }]));

c.push(H2("Set Piece: Raising the Engines"));

c.push(P("Nine months of work at Krenholt, four hundred miles of road, and one afternoon of getting them up onto the platform in range of a wall whose garrison would very much prefer they were not. Auberitz has done this before. Auberitz is still going to lose an engine."));

c.push(BOX("\u201CThe chief of the works is four feet tall, has built eleven of these, has never once seen one used the way she intended, and is currently standing on a barrel telling a Harrowmark knight exactly what will happen if he moves his horses across the traverse. He moves his horses across the traverse. Ninety seconds later a stone comes off the Ward Gate and takes the traverse, the horses, and an argument nobody is going to finish.\u201D"));

c.push(B("Three problems, running simultaneously.", "The platform is short of fascines and the nearest brushwood is inside bowshot. The counterweight for the second engine is still two miles back on a wagon with a broken axle. And the garrison\u2019s own engines have the range and are ranging in, three stones at a time, adjusting."));

c.push(B("What the party can do.", "Any of it. Fetch the counterweight, screen the fascine parties, or \u2014 the answer the engineers are hoping somebody thinks of \u2014 go and do something about the garrison\u2019s ranging shots, which are being directed by an observer on the Ward Gate who can be reached, with difficulty, from the seaward side."));

c.push(P("The chief of the works will tell the party exactly what will go wrong an hour before it does, and will be right, and will not say so afterward. She is the module\u2019s best NPC and she has eleven lines in her, all of them dry."));

c.push(H1("NPC Profiles"));

c.push(H2("Marshal Ossian Drell"));

c.push(P("A hobgoblin of perhaps fifty, and a legionary officer of the kind his people produce reliably and other peoples mostly envy: literate, unhurried, personally brave in a way he considers unremarkable, and entirely without any stake in Vale\u2019s war beyond the contract that pays for it. He has held Vindana for three years, competently, and has filed a monthly report about it every one of those months."));
c.push(P("Vindana\u2019s garrison commander, professional rather than cruel, visibly proud of a defense he has had three years to prepare. Speech: economical, respectful of competence in an enemy, entirely without the theatrical menace a table might expect \u2014 he is fighting a siege, not delivering a villain\u2019s speech."));
c.push(P("Open thread: Drell is this campaign\u2019s model of the occupation\u2019s real competence, and a DM can use him again in Module Seven as the officer whose actual weakness the party finds \u2014 or, if the table prefers, as an officer who can be talked into surrender once that weakness is found, rather than one who must be killed."));

// --------------------------------------------------------------- Optional
c.push(H1("Optional Content"));

c.push(H2("The Coalition\u2019s Own Argument"));
c.push(P("If the table wants more camp texture, let Oksitan and Auberitz officers disagree openly, for the first time, about how the siege should be run \u2014 patience versus a faster assault, roughly the tension the module itself just dramatized in Scene 4. No mechanical stakes; this is a chance to deepen the coalition\u2019s internal friction established in Module Four."));

c.push(H2("The Ward\u2019s Read on the City"));
c.push(P("If the Ward was rescued in Module Four, she has real, specific knowledge of Vindana from before the occupation \u2014 a postern gate, a garrison habit, something small but true. This does not need to change the module\u2019s outcome; it is a chance to let her earlier promise (see her NPC profile) pay off in a small way."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("How the engines were raised.", "Smoothly or roughly \u2014 track it, and carry forward any advantage or complication into Module Seven\u2019s assault."));
c.push(BUL("The supply raid\u2019s outcome.", "A clean success meaningfully weakens Vindana for Module Seven; a failed or messy raid leaves Drell\u2019s position stronger than it might otherwise have been."));
c.push(BUL("How First Assault was fought.", "Track whether the party\u2019s own corner of the fight was a clear local win, a costly draw, or a real loss \u2014 this sets the emotional temperature Module Seven opens on."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BUL("Norvatch manifests.", "Found with the intercepted column \u2014 bills of lading, correctly filed, showing that a share of what feeds Vindana comes up through Norvatch bottoms under a standing contract the coalition has no legal instrument to void. Nobody in the column did anything illegal. That is the part that will irritate the party most."));

c.push(BUL("Captured supply.", "Whatever the cavalry took from the intercepted column \u2014 modest, practical, and a genuine (if small) blow to Vindana\u2019s stores rather than treasure."));
c.push(BUL("A garrison officer\u2019s dispatch.", "Recovered from the supply column or from First Assault\u2019s dead \u2014 real intelligence about Vindana\u2019s internal state, useful DM ammunition for Module Seven rather than something that needs to pay off here.", { keepNext: true }));

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
  fs.writeFileSync(stagePath("KC_Module06_VindanaInvestment.docx"), buf);
  console.log("Written.");
});
