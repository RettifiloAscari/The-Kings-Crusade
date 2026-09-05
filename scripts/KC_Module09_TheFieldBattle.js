// KC_Module09_TheFieldBattle.js -- Session Module Nine: The Field Battle.
//
// Canon lives here. corpus/ and documents/ are generated from this file and are
// never edited by hand. See CLAUDE.md for the sign-off rules: anything in the
// "Not yet decided" table must not appear here until it has been approved.
//
// ESCAPE CONVENTION: all prose lives as \uXXXX escapes, never as literal
// typographic characters. tools/build.sh fails the build if that slips. If
// hand-typing an escape, use ONE backslash -- a doubled backslash compiles
// clean and passes the non-ASCII scanner but leaks literal text into the PDF.
// A literal newline inside a double-quoted JS string is invalid syntax -- use
// two separate BOX() or P() calls instead of embedding a line break.
//
// Arsuf-keyed: the open-field battle that proves Vale\u2019s army beatable in a
// fair fight, not just behind a broken siege. Per the module breakdown, this
// module carries either a third rescue thread or a real loss -- written here
// as a real loss (Tam Ondry, introduced in Module Five) by default, with an
// explicit DM-facing alternate for tables that would rather not lose him.
//
// Encounter design note: General Ilyana Voss is the SRD Gladiator (CR 5, 1800
// XP), a deliberate step up from the Veteran used for Thane (2A) and Drell
// (Module Six/Seven) -- Vale\u2019s field army is a different register of threat
// than an occupation garrison, and the boss-tier block should feel like it.

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

// An ordered sequence -- the phases of a set piece, the movements of a battle -- is a
// list, and printing it as unmarked bold-led prose beside a real bulleted list is
// what made those pages read as two idioms doing one job. ORDERED marks it as what
// it is. A document with a second ordered list passes { instance: 1 }, because one
// numbering reference is one running counter.
const ORDERED = (segs, opts = {}) => new Paragraph({
  numbering: { reference: "steps", level: 0 },
  spacing: { after: 120 },
  ...opts,
  children: segs.map(s => new TextRun({ text: s.t, bold: !!s.b, italics: !!s.i, color: s.c }))
});

const B = (lead, rest, opts = {}) => PS([{ t: lead + " ", b: true }, { t: rest }], opts);
const BUL = (lead, rest, opts = {}) => BULLET(lead ? [{ t: lead + " ", b: true }, { t: rest }] : [{ t: rest }], opts);
const ORD = (lead, rest, opts = {}) => ORDERED(lead ? [{ t: lead + " ", b: true }, { t: rest }] : [{ t: rest }], opts);

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
const SB = (d) => { const out = []; out.push(new Paragraph({ spacing: { before: 240, after: 40 }, keepNext: true, children: [new TextRun({ text: d.name, bold: true, size: 26, color: "5B1F1F" })] })); out.push(PS([{ t: d.meta, i: true }], { spacing: { after: 120 }, keepNext: true })); out.push(B("Armor Class:", d.ac, { keepNext: true })); out.push(B("Hit Points:", d.hp, { keepNext: true })); out.push(B("Speed:", d.speed, { keepNext: true })); out.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ new TableRow({ cantSplit: true, tableHeader: true, children: ["STR","DEX","CON","INT","WIS","CHA"].map(h => abCell(h, true)) }), new TableRow({ cantSplit: true, children: [d.str,d.dex,d.con,d.int,d.wis,d.cha].map(v => abCell(v + " (" + mod(v) + ")")) }) ] })); out.push(P("", { spacing: { after: 60 } })); if (d.saves) out.push(B("Saving Throws:", d.saves)); if (d.skills) out.push(B("Skills:", d.skills)); if (d.senses) out.push(B("Senses:", d.senses)); if (d.langs) out.push(B("Languages:", d.langs)); out.push(B("Challenge:", d.cr)); (d.traits||[]).forEach(t => out.push(PS([{ t: t.n + ". ", b: true, i: true }, { t: t.t }]))); if (d.actions && d.actions.length) { out.push(PS([{ t: "ACTIONS", b: true }], { spacing: { before: 80, after: 80 }, keepNext: true })); d.actions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } if (d.reactions && d.reactions.length) { out.push(PS([{ t: "REACTIONS", b: true }], { spacing: { before: 80, after: 80 }, keepNext: true })); d.reactions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } return out; };


// ---------- content ----------
const c = [];

c.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: "The Field Battle", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module Nine", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("The coalition marches on Caer Ysolde, and Vale\u2019s occupation makes its first attempt to stop that march in the open rather than from behind a wall \u2014 a field army under General Ilyana Voss, meaning to catch the coalition\u2019s column strung out and unprepared. It does not work, and this module is the campaign\u2019s proof, stated in blood rather than in a speech, that Vale\u2019s army can be beaten in a fair fight. That proof costs something. Core scenes run four to four and a half hours; Optional Content fills out the rest of a five-hour session and can be cut if the table is short on time."));

c.push(H2("A Battle, Not a Massacre"));

c.push(P("Vindana fell partly through the party\u2019s own cleverness and partly through Xavier\u2019s extraordinary intervention. This battle should not repeat either trick \u2014 it is won by discipline, by a column that holds formation under pressure the way Richard\u2019s did at the historical Arsuf, and by the party\u2019s own competence in a fight that is genuinely dangerous rather than pre-decided. Let the danger be real."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 19, 51],
  [
    ["1. The Column Strung Out", "20\u201330 min", "The ambush begins to reveal itself."],
    ["2. Tam\u2019s Warning", "20\u201330 min", "The module\u2019s cost \u2014 or its alternate. Read carefully before running."],
    ["3. Holding the Line", "60\u201390 min", "The battle itself. DC table and stat block below."],
    ["4. The Rout", "30\u201345 min", "General Voss\u2019s line breaks. Proof, not celebration."],
    ["Optional Content", "30\u201345 min", "Run if the table has time; cut cleanly if not."]
  ]
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("General Voss is a real tactician and her ambush is a genuinely good plan \u2014 she has read the coalition\u2019s march discipline correctly and is exploiting the one moment a column is actually vulnerable, the transition from road formation to battle formation. It very nearly works. What defeats her is not a trick or a twist; it is the coalition holding its nerve and its lines exactly long enough for the party, and Tam Ondry\u2019s warning, to buy the time Xavier\u2019s officers need to form up properly."));

c.push(PS([DM("DM Only: "), { t: "this module defaults to Tam Ondry dying in Scene 2, delivering the warning that saves the column. He is not fridged for shock value \u2014 his death is the direct, earned cost of a warning that genuinely matters, and it should be played with the same weight Module Five gave his introduction. If your table has grown attached to him and you would rather not lose him, the alternate in Scene 2 below turns his warning into the discovery of a third captive member of the royal house instead, opening a rescue thread rather than closing a life. Choose deliberately, not by default." }]));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: The Column Strung Out"));

c.push(P("The coalition\u2019s march north from Vindana is disciplined but long, strung across miles of road through country that stopped being safely held the moment it stopped being inside sight of the city\u2019s walls. The first sign of trouble is not an attack \u2014 it is an absence."));

c.push(BOX("The scouts who should have reported in an hour ago have not, and the ones sent after them have not come back either. Somewhere ahead, past a treeline that should not be able to hide as much as it apparently does, the road goes quiet in a way that has nothing to do with Elduvaine\u2019s habits and everything to do with an army waiting for the right moment."));

c.push(P("Let the party notice the wrongness before the coalition\u2019s officers do \u2014 an Investigation or Survival check (DC 13) reads the terrain as a natural ambush site, or Insight (DC 13) reads the missing scouts as deliberate rather than delayed. This scene\u2019s job is dread, not combat; end it on the certainty that something is about to happen, not on contact."));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: Tam\u2019s Warning"));

c.push(P("A rider comes back down the column at a dead run \u2014 the same shape of scene as Module Five\u2019s Scene 1, deliberately, because it is the same man doing the same job one more time."));

c.push(BOX("Tam Ondry doesn\u2019t bother reining in properly, half-falling out of the saddle in his hurry to reach an officer, anyone, with rank enough to matter. \u201CAmbush \u2014 treeline, both sides, more of them than us if they catch us stretched out like this \u2014 you have to form up now, not in a minute, now \u2014\u201D An arrow that was aimed at someone else entirely finds him mid-sentence."));

c.push(H3("Running the Default Scene"));

c.push(P("Tam\u2019s warning reaches the column\u2019s officers in time regardless of what the party does \u2014 that is not in question, and the party cannot fail to receive it. What is in question is whether they reach him before the end. A DC 16 Medicine check, attempted within the first round after he falls, stabilizes him long enough for a few last words; failing that check, or simply not reaching him in time, means he dies having delivered exactly the warning he came to give, aware that it worked. Either way, this is the module\u2019s emotional cost, and it should be allowed to land before Scene 3\u2019s battle begins."));

c.push(H3("The Alternate: A Third Thread"));

c.push(P("If your table would rather not lose Tam, run this instead: his warning is the same and he survives it, but what he saw scouting ahead of the ambush was not only Voss\u2019s army. It was a small column under Elduvish colours moving the other way, unhurried, barely escorted, with a grey-haired elf riding in the middle of it who was not tied to anything and did not look like a prisoner. That is Emrys Ysolde, the Envoy, being moved somewhere quieter \u2014 and the reason Tam mentions the detail at all is that nobody in the column was guarding him so much as accompanying him."));

c.push(PS([DM("DM Only: "), { t: "if the party pulled Serjeant Hoth\u2019s file out of Sennoch Hall in Module Four, they already have the first hard evidence that Emrys has spent three years talking to Vale. This is where that stops being a document and becomes a man on a road. He is not a traitor, there is no reveal, and he will not defend himself: he took the only job nobody else in his house would take, and he cannot produce a clean accounting of whether it helped. What his sister\u2019s physician, the published levy and two mass reprisals that did not happen actually cost is a question he has had three years to word and still cannot answer. Open the thread here and let it stay open." }]));

c.push(P("Do not run both versions of this scene. Pick one before the session and commit to it."));

// ---------------------------------------------------------------- Scene 3
c.push(H2("Scene 3: Holding the Line"));

c.push(P("The ambush breaks from both treelines at once, and for several very long minutes the outcome is genuinely uncertain. This is the module\u2019s real fight, and it should feel like one."));

c.push(BOX("The coalition line buckles and does not break \u2014 Oksitan spears on the left with a dragonborn house-knight somewhere in the middle of them roaring instructions nobody can hear, Auberitz heavier foot anchoring the center, Harrowmark\u2019s own discipline holding a flank that has no business holding against these numbers. Somewhere in the press, unmistakable even through a helm, is a woman directing Voss\u2019s attack with the calm of someone who still, even now, believes this is winnable."));

c.push(H3("Running the Scene"));

c.push(P("Use Legionaries of the Sixth (Bestiary, CR 1/2) for General Voss\u2019s rank and file \u2014 six of them, engaging the party and the coalition line in waves rather than all at once, with Formation working every time two of them reach the same target \u2014 and General Voss herself (see Stat Block) as the encounter\u2019s real threat, seeking out whoever on the coalition side looks most like a commander to kill or capture. If the party has protected Xavier\u2019s officers or otherwise distinguished themselves, Voss may target them directly, which is a genuine compliment from an enemy tactician and should read as one."));

c.push(H3("Scaling the Fight"));

c.push(P("Voss (1,800 XP) plus six Legionaries of the Sixth (100 XP each, 600) totals 2,400 base XP across seven monsters, which sits in the 7\u201310 band: \u00D72.5 at party sizes three to five, for 6,000 adjusted, and \u00D72 at six or more, for 4,800. The table below is computed at 7th level, because that is where most tables will actually be by Module Nine \u2014 but this campaign prescribes no milestone schedule, so check it against your own party\u2019s real level before you run it."));

c.push(table(
  ["PCs", "Adj. XP", "Hard", "Deadly", "Reads as"],
  [10, 18, 16, 18, 38],
  [
    ["4 at 7th", "6,000", "4,200", "6,800", "Hard, close to Deadly. Run as written."],
    ["5 at 7th", "6,000", "5,250", "8,500", "Hard. This is the calibration the module wants."],
    ["6 at 7th", "4,800", "6,300", "10,200", "Medium. Add two more legionaries, not a second officer."],
    ["4 at 5th", "6,000", "3,000", "4,400", "Well past Deadly. Cut to four legionaries."]
  ]
));

c.push(P("This is deliberately the most dangerous fight in the campaign so far, and unlike every earlier encounter it is meant to read Hard rather than Easy-to-Medium. If it comes out too heavy against your own table, remove legionaries rather than reducing Voss \u2014 she is the fight\u2019s whole point, and a Voss who is not frightening has cost the module its subject."));

c.push(H3("Stat Block"));

c.push(...SB({
  name: "General Ilyana Voss",
  meta: "Medium humanoid (orc), lawful neutral \u2014 SRD Gladiator, renamed",
  ac: "16 (studded leather armor, shield)",
  hp: "112 (15d8 + 45)",
  speed: "30 ft.",
  str: 18, dex: 15, con: 16, int: 10, wis: 12, cha: 15,
  saves: "Str +7, Dex +5, Con +6",
  skills: "Athletics +10, Intimidation +5",
  senses: "passive Perception 11",
  langs: "Common",
  cr: "5 (1,800 XP)",
  traits: [
    { n: "Brave", t: "Voss has advantage on saving throws against being frightened." },
    { n: "Brute", t: "A melee weapon deals one extra die of its damage when Voss hits with it (included in the attacks below)." }
  ],
  actions: [
    { n: "Multiattack", t: "Voss makes three melee attacks or two ranged attacks." },
    { n: "Spear", t: "Melee or Ranged Weapon Attack: +7 to hit, reach 5 ft. or range 20/60 ft., one target. Hit: 11 (2d6 + 4) piercing damage, or 13 (2d8 + 4) piercing damage if used with two hands to make a melee attack." },
    { n: "Shield Bash", t: "Melee Weapon Attack: +7 to hit, reach 5 ft., one creature. Hit: 9 (2d4 + 4) bludgeoning damage. If the target is Medium or smaller, it must succeed on a DC 15 Strength saving throw or be knocked prone." }
  ],
  reactions: [
    { n: "Parry", t: "Voss adds 3 to her AC against one melee attack that would hit her. To do so, she must see the attacker and be wielding a melee weapon." }
  ]
}));

// -------------------------------------------------------------- Skill DCs
c.push(H2("Tiered Skill DCs"));

c.push(P("Easy 10, Moderate 13, Hard 16, matching the tiers used throughout this campaign."));

c.push(table(
  ["Task", "Skill", "DC", "Tier"],
  [44, 26, 10, 20],
  [
    ["Read the treeline as a prepared ambush site", "Investigation / Survival", "13", "Moderate"],
    ["Recognize the missing scouts as deliberate", "Insight", "13", "Moderate"],
    ["Stabilize Tam Ondry after he falls (default Scene 2)", "Medicine", "16", "Hard"],
    ["Hold a coalition position against Voss\u2019s pressure", "Athletics / relevant combat skill", "13", "Moderate"],
    ["Talk Voss into a battlefield surrender once clearly beaten", "Persuasion / Intimidation", "16", "Hard"]
  ]
));

// ---------------------------------------------------------------- Scene 4
c.push(H2("Scene 4: The Rout"));

c.push(P("Once General Voss falls, is captured, or breaks off \u2014 a defeated field commander she genuinely is not the type to flee cleanly, but she will disengage if the fight is unambiguously lost rather than die for nothing \u2014 her line comes apart within minutes. This is proof, not celebration."));

c.push(BOX("It is not a rout in the way stories tell it \u2014 no cheering, no clean lines of fleeing soldiers. It is simply, suddenly, over: Voss\u2019s remaining troops breaking for the treeline they came from, the coalition too battered and too relieved to give proper chase, and a field that will need a great deal of tending before anyone can say anything about who won it that does not sound obscene."));

c.push(P("Let the aftermath be heavy rather than triumphant, especially if Tam died in Scene 2 \u2014 the coalition has proven something real about Vale\u2019s army today, and it cost real people to prove it. End the session on that weight rather than on a victory speech. Hand off directly to Module Ten for the approach to Caer Ysolde."));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("Puzzles and Set Pieces"));

c.push(P("The five movements are the module\u2019s existing scenes, sequenced, with the party put at the hinge of each. Voss\u2019s Field Orders is additive: fifteen minutes, and it changes the casualty list rather than the outcome."));

c.push(H2("The Puzzle: Voss\u2019s Field Orders"));

c.push(P("General Ilyana Voss writes her orders in the Legion\u2019s working cipher, which is not a cipher so much as a professional shorthand a hundred and forty years old. Any captured dispatch is readable and the reading is a genuine puzzle rather than a check."));

c.push(P("The Sixth writes movement as three elements: the company number, the hour on a twelve-mark day, and a bearing given as a clock face from the standard. Nothing is written in words, because words can be read by anybody and numbers can be read by a professional."));

c.push(table(
  ["What is written", "What it means", "How a party can establish it"],
  [22, 38, 40],
  [
    ["IV \u00b7 iii \u00b7 x", "Fourth company, third hour, bearing ten o\u2019clock", "Compare any two dispatches against movements the party watched happen."],
    ["A doubled numeral", "The order is a feint and is to be seen", "The only dispatches ever sent uncovered are doubled. Two examples is enough."],
    ["A struck-through hour", "Hold until countermanded", "A prisoner will confirm this without considering it a betrayal, because it is drill, not intelligence."],
    ["No bearing at all", "Reserve. Committed at the commander\u2019s word only", "This is the one that matters and it is the one the party will find last."]
  ]
));

c.push(P("Put together, three captured dispatches tell the party where Voss\u2019s reserve is and that she has not yet committed it \u2014 which is the single most valuable piece of information available on the field, and which the coalition\u2019s own scouts have failed to get for two days."));

c.push(PS([DM("DM Only: "), { t: "give them a real advantage for solving it. A party that finds the reserve and tells Xavier changes the shape of the battle, and the DM should say so in narration: the coalition\u2019s left does not break, and about four hundred people who would have died do not. Do not make this the difference between victory and defeat \u2014 the field battle is won either way, because the campaign needs it won \u2014 make it the difference in the casualty list, and read part of the casualty list out." }]));

c.push(H2("Set Piece: The Battle, in Five Movements"));

c.push(P("Eight thousand people on each side and five characters somewhere in the middle of it. The party cannot win a field battle and should never be asked to; what they can do is be the hinge at five specific moments, and the module is built to put them at each one."));

c.push(ORD("The Approach.", "Two hours of standing in a line while the enemy does the same thing four hundred yards away. Nothing happens. This is deliberate and should be played straight \u2014 the fear in a set battle is almost all in the waiting, and a DM who skips it has thrown away the only chance to make the rest of it land."));

c.push(ORD("The Left Gives.", "Oksitan spears, already shaken by what happened to their king, bend and do not break. The party is behind them and can see it happening before anyone in command can. A dragonborn house-knight is somewhere in the middle of that line roaring instructions nobody can hear."));

c.push(ORD("The Ironshanks.", "The Sixth\u2019s heavy foot come through the gap in a shield wall and the party is what is in front of them. Four ironshanks and an optio, in the open, with the battle noise making every command a shouted argument. This is the module\u2019s real fight and it should be genuinely frightening."));

c.push(ORD("The Reserve.", "Voss commits, or does not, depending on the puzzle above. If the party found it, the coalition meets it ready. If not, it arrives on the flank and the module gets a sixth movement nobody wanted."));

c.push(ORD("Voss.", "She does not flee and she does not die swinging. When the field is lost she stops, puts her sword point down in the mud, and waits, because she is a professional on a contract and the contract does not require this. What the party does about that is the end of the session."));

c.push(PS([DM("DM Only: "), { t: "a captured Voss is worth more to this campaign than a dead one by a wide margin and the module should not tip the scale. She will not be turned, will not inform, and will not apologise for any of it \u2014 and she will, if asked correctly, explain exactly what the Sixth\u2019s contract says and when it expires, which is information nobody else in Elduvaine will give the party and which matters enormously in Module Eleven." }]));

c.push(H1("NPC Profiles"));

c.push(H2("General Ilyana Voss"));

c.push(P("An orc, and a career soldier of the sort every army wants and few can keep. She came up through the same legion Drell did and outgrew it; she commands Vale\u2019s field army because she is the best available person to command it, and she has never in her life been asked to believe in anything. The campaign makes no comment anywhere about the fact that orcs hunt wyverns for Harrowmark at Greywatch and command armies for Vale in the field. Neither should any NPC. Let the players sit with it."));
c.push(P("A real tactician rather than a fanatic, calm under pressure in a way that reads as competence rather than menace. Speech, in the brief window a table might hear her actually speak (command orders during the fight, or a surrender if the party earns one): direct, economical, entirely without contempt for an enemy who is currently winning."));
c.push(P("Open thread: if captured rather than killed, Voss is a genuine long-term asset for the coalition \u2014 a professional soldier with real knowledge of Vale\u2019s remaining field strength, who a DM can develop as a reluctant, pragmatic informant in later modules rather than a villain who must be disposed of."));

// --------------------------------------------------------------- Optional
c.push(H1("Optional Content"));

c.push(H2("What the Column Says About Tam"));
c.push(P("If Tam died in Scene 2, let the coalition\u2019s reaction be genuine rather than perfunctory \u2014 a name added to a growing list, spoken plainly rather than eulogized, echoing exactly the register the Standing Water gave the second king\u2019s death in Module Five. No mechanical stakes; this is the module giving its cost real weight."));

c.push(H2("Voss\u2019s Own Papers"));
c.push(P("If Voss is captured or her body searched, her own field orders reveal real, specific intelligence about the remaining distance to Caer Ysolde and what defends it \u2014 useful DM ammunition for Module Ten rather than something that needs to resolve here."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("Which Scene 2 was run.", "Tam\u2019s death or the third-thread alternate \u2014 record which, since it changes what Module Ten and beyond can reference."));
c.push(BUL("General Voss\u2019s fate.", "Killed, captured, or fled \u2014 a captured Voss is a real long-term asset; track it the same way Marshal Drell\u2019s fate was tracked in Module Seven."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BUL("Voss\u2019s field orders.", "Real intelligence about Caer Ysolde\u2019s defenses \u2014 see Optional Content."));

c.push(BUL("A cloak of elvenkind.", "Off one of Voss\u2019s scouts, and Elduvish work \u2014 which is to say looted, three years ago, from somebody who is not alive to want it back. A DM who wants that to land can let the party find out whose it was."));
c.push(BUL("Captured field equipment.", "Modest but genuine \u2014 weapons, armor, and supply recovered from Voss\u2019s broken column, worth collecting rather than a windfall.", { keepNext: true }));

// -------------------------------------------------------------- Refrain
c.push(H1("The Refrain"));

c.push(VERSE([
  "By thought, and by word, and by deed,",
  "the king\u2019s own chosen kept their creed.",
  "Far from home, where the quiet land lay,",
  "they held the line, and would not stray."
]));

const doc = new Document({
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 260, hanging: 260 } } } }] },
    // The same measure as the bullets, deliberately: a numbered list and a bulleted
    // one appear on the same page often enough that their text has to hang off one
    // left edge, and "1." is near enough the width of a dot for the gap to match. It
    // holds to "99."; nothing in this campaign counts past five.
    { reference: "steps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 260, hanging: 260 } } } }] }
  ] },
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
  fs.writeFileSync(stagePath("KC_Module09_TheFieldBattle.docx"), buf);
  console.log("Written.");
});
