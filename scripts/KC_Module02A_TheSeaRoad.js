// KC_Module02A_TheSeaRoad.js -- Session Module Two (Sea Road variant).
//
// Canon lives here. corpus/ and documents/ are generated from this file and are
// never edited by hand. See CLAUDE.md for the sign-off rules: anything in the
// "Not yet decided" table must not appear here until it has been approved.
//
// ESCAPE CONVENTION: all prose lives as \uXXXX escapes, never as literal
// typographic characters. tools/build.sh fails the build if that slips.
//
// This is one of two mutually exclusive Module Two variants -- run this one if
// the party chose the sea road at the end of Module One; run 2B if they chose
// the mountain road. Per CLAUDE.md, whichever road they take is written in
// full; the other is met only as aftermath, delivered in Module 5.
//
// Encounter design note: Warden Ivor Thane is the SRD Veteran (CR 3, 700 XP)
// renamed and reflavored, backed by SRD Bandits (CR 1/8, 25 XP each) as his
// household guards -- both pulled from 5e-bits/5e-database rather than from
// memory. The DMG math (below, in Scene 3) reads Easy-to-Medium rather than
// Hard or Deadly, and that is deliberate: this is an opportunistic, largely
// one-sided historical beat (Richard's taking of Cyprus), not the campaign's
// hardest fight, and its real stakes are the captives at risk and the choice
// of what becomes of Calanthe afterward, not raw lethality.

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

const BULLET = (segs) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 120 },
  children: segs.map(s => new TextRun({ text: s.t, bold: !!s.b, italics: !!s.i, color: s.c }))
});

const B = (lead, rest) => PS([{ t: lead + " ", b: true }, { t: rest }]);
const BUL = (lead, rest) => BULLET(lead ? [{ t: lead + " ", b: true }, { t: rest }] : [{ t: rest }]);

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

const { Table, TableRow, TableCell, WidthType, ShadingType } = require('docx');
const cell = (text, opts = {}) => new TableCell({ width: { size: opts.w || 20, type: WidthType.PERCENTAGE }, shading: opts.head ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, margins: { top: 50, bottom: 50, left: 45, right: 45 }, children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text, bold: !!opts.head, size: 18 })] })] });
const row = (cells, opts = {}) => new TableRow({ children: cells, cantSplit: true, ...opts });
const FULLWIDTH = "KCFullWidth";
const table = (headers, widths, rows, opts = {}) => new Table({ ...(opts.full ? { style: FULLWIDTH } : {}), width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ row(headers.map((h, i) => cell(h, { head: true, w: widths[i] })), { tableHeader: true }), ...rows.map(r => row(r.map((v, i) => cell(v, { w: widths[i] })))) ] });

const mod = (v) => { const m = Math.floor((v - 10) / 2); return (m >= 0 ? "+" : "\u2212") + Math.abs(m); };
const abCell = (text, bold) => new TableCell({ width: { size: 16.6, type: WidthType.PERCENTAGE }, shading: bold ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40, before: 40 }, keepNext: !!bold, children: [new TextRun({ text, bold: !!bold, size: 20 })] })] });
const SB = (d) => { const out = []; out.push(new Paragraph({ spacing: { before: 240, after: 40 }, children: [new TextRun({ text: d.name, bold: true, size: 26, color: "5B1F1F" })] })); out.push(PS([{ t: d.meta, i: true }], { spacing: { after: 120 } })); out.push(B("Armor Class:", d.ac)); out.push(B("Hit Points:", d.hp)); out.push(B("Speed:", d.speed)); out.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ new TableRow({ cantSplit: true, tableHeader: true, children: ["STR","DEX","CON","INT","WIS","CHA"].map(h => abCell(h, true)) }), new TableRow({ cantSplit: true, children: [d.str,d.dex,d.con,d.int,d.wis,d.cha].map(v => abCell(v + " (" + mod(v) + ")")) }) ] })); out.push(P("", { spacing: { after: 60 } })); if (d.saves) out.push(B("Saving Throws:", d.saves)); if (d.skills) out.push(B("Skills:", d.skills)); if (d.senses) out.push(B("Senses:", d.senses)); if (d.langs) out.push(B("Languages:", d.langs)); out.push(B("Challenge:", d.cr)); (d.traits||[]).forEach(t => out.push(PS([{ t: t.n + ". ", b: true, i: true }, { t: t.t }]))); if (d.actions && d.actions.length) { out.push(PS([{ t: "ACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.actions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } return out; };


// ---------- content ----------
const c = [];

c.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: "The Sea Road", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module Two, Sea Road Variant", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("Run this module if the party chose the sea road at the end of Module One. If they chose the mountain road, use Module 2B instead \u2014 the two are mutually exclusive, and only one is ever played at a given table. The coalition fleet, three days out from Harrowmark, is scattered by a storm and driven onto the coast of Calanthe, a small independent island that answers to no crown. Its self-styled Warden, Ivor Thane, has spent years quietly ransoming and plundering ships wrecked on his coast, and several coalition vessels \u2014 including one carrying supplies the crusade cannot easily spare \u2014 have just become his newest opportunity. Core scenes run three and a half to four hours; Optional Content fills out the rest of a five-hour session and can be cut cleanly if the table is short on time."));

c.push(H2("The Shape of a Sea-Lane War"));

c.push(P("Play Calanthe as ordinary and opportunistic, not sinister. Warden Thane is not a monster and has no connection to Vale, to Elduvaine, or to anything the party will learn about later \u2014 he is a petty ruler who found a profitable position at a strait every southbound fleet has to pass, and has been quietly working it for longer than anyone in the coalition realized. This module\u2019s tension is not \u201Cwhat is this evil,\u201D the way Elduvaine\u2019s will be. It is a smaller and in some ways harder question: what does a war for someone else\u2019s deliverance do to the small, uninvolved places it happens to pass through \u2014 and the party is about to decide the answer for one of them, in miniature, several modules before they have to decide it for a whole kingdom."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 15, 55],
  [
    ["1. The Storm", "15\u201320 min", "Fast and mostly narrative; ends with the party ashore on Calanthe."],
    ["2. Landfall", "30\u201345 min", "Wreckage, frightened survivors, and Thane\u2019s reputation, secondhand."],
    ["3. Thane\u2019s Hold", "90\u2013120 min", "Negotiation, and combat if it comes to that. DC table and stat blocks below."],
    ["4. Calanthe\u2019s Fate", "30\u201345 min", "Loot, the island\u2019s future, and the road onward \u2014 Branch Ledger entry."],
    ["Optional Content", "30\u201345 min", "Run if the table has time; cut cleanly if not."]
  ],
  { full: true }
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("Ivor Thane is exactly what he appears to be: a minor lord who inherited a rock in the strait, worked out early that wrecked ships are worth more than the fishing rights ever were, and has spent a decade being quietly, unglamorously terrible to whoever the sea hands him. He is not secretly working for anyone. He is not a piece of a larger plot. He genuinely believes he is owed a toll for the use of his coastline, and he is genuinely surprised, every time, that anyone objects strenuously enough to do something about it."));

c.push(PS([DM("DM Only: "), { t: "the point of this module is not Thane himself. It is the choice at the end of Scene 4 \u2014 what the coalition does with a small, real place it now has power over \u2014 and it should feel exactly as weighty as the party lets it. A table that wants to think hard about it should be able to; a table that wants to hand it to Sera Vosk and move on should be able to do that too, without being made to feel they missed something." }]));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: The Storm"));

c.push(P("Three days south of Harrowmark, with the coast long out of sight, the sky does what northern sailors say it always does when a crossing has gone too easily for too long."));

c.push(BOX("The first anyone knows of it is the light going wrong \u2014 a green-grey dusk arriving hours early \u2014 and then the wind, which does not build so much as simply begin, all at once, as if a door somewhere had been thrown open. Somewhere off the bow, another ship\u2019s lantern swings wildly, then is gone."));

c.push(P("Run this fast and mostly in narration. Let each PC make one relevant check if they want to help the crew (Strength (Athletics) to work the lines, Wisdom (Survival) to read the worst of it coming, an appropriate spell to calm water or clear fog) \u2014 success or failure barely matters mechanically and mostly colors how battered the party\u2019s own ship is when the storm breaks. What matters is the ending: the party\u2019s vessel, and at least one other coalition ship, are driven onto an unfamiliar coast as the wind finally drops. By morning, that coast has a name \u2014 Calanthe \u2014 offered by whichever sailor recognizes the headland first, along with a look that suggests the name means something unpleasant."));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: Landfall"));

c.push(P("The beach the party makes for is already occupied \u2014 wreckage from a second coalition ship, planking and barrels and one overturned longboat, and a handful of soaked, shaken survivors doing what shipwrecked people do: counting who is missing."));

c.push(BOX("A woman in a torn Auberitz quartermaster\u2019s coat looks up as the party approaches, calculates something visible behind her eyes, and does not waste time on relief. \u201CIf you\u2019ve a ship that still floats, you\u2019re the first good news I\u2019ve had all morning. If you don\u2019t \u2014 well. Welcome to Calanthe. Try not to look valuable.\u201D"));

c.push(P("This is Sera Vosk (see NPC Profiles). She has enough of the picture to be useful without knowing everything: Calanthe answers to a self-declared Warden named Ivor Thane, whose men have a habit of finding wrecks before anyone official does, and half of her own crew \u2014 including the ship\u2019s master, whom she will not stop mentioning by name \u2014 were taken up the cliff road toward his hold before dawn. She has been debating, alone, whether to go after them unarmed or wait for exactly the kind of help that just walked out of the surf."));

c.push(P("Let the party spend some time here gathering what Calanthe\u2019s few frightened locals will say, if approached carefully: Thane holds captives for ransom, has for years, and nobody on the island loves him for it \u2014 but nobody has ever had a coalition fleet\u2019s worth of reasons to do anything about it either."));

// ---------------------------------------------------------------- Scene 3
c.push(H2("Scene 3: Thane\u2019s Hold"));

c.push(P("The hold is less a fortress than a fortified manor \u2014 old stone, patched with newer and cheaper stone, flying a banner nobody recognizes over a gate that has clearly been rebuilt more than once. Thane receives visitors readily; ransom is, after all, his business, and armed strangers asking after his newest guests are simply customers who have not yet been told the price."));

c.push(BOX("\u201CAh,\u201D says the man in the doorway, unhurried, a ledger genuinely under one arm. \u201CMore of the fleet that thought my strait was free water. I\u2019m Ivor Thane, and I\u2019m told this happens to be my island. Your people are safe, fed, and entirely available \u2014 for a reasonable consideration. I find negotiations go better before anyone has drawn a sword, so let\u2019s try that first, shall we?\u201D"));

c.push(H3("Running the Scene"));

c.push(P("Thane will genuinely negotiate, and a table that wants to buy, threaten, or talk their way to the captives\u2019 release without a fight should be able to \u2014 he names an absurd price, but he is not attached to getting it, only to being paid something, and a successful Charisma (Persuasion or Intimidation) check against DC 15 gets his price down to something a coalition fleet can actually spare, or gets him to agree to release the captives in exchange for safe passage and a promise to leave his toll uncollected this once. He is not brave and is not interested in dying for a business model."));

c.push(P("If the party refuses to pay, threatens him convincingly, is caught trying to free the captives by force, or simply attacks, it goes to violence \u2014 Thane\u2019s guards (use the Bandit stat block) fight to protect the hold and its stores, but break and flee the moment Thane himself surrenders, is captured, or is reduced to half his hit points and makes his own morale call to cut losses. This should read as a real fight, not a formality, but it is not meant to be this module\u2019s hardest moment; see Scaling the Fight, below, for why."));

c.push(H3("Scaling the Fight"));

c.push(P("Ivor Thane is the SRD Veteran (CR 3, 700 XP) renamed and reflavored, unmodified otherwise; his guards are SRD Bandits (CR 1/8, 25 XP each), both pulled from 5e-bits/5e-database. Run with Thane plus four Bandits \u2014 five total monsters, comfortably inside the 3\u20136 monster band, so no table\u2019s party size pushes the count across a multiplier boundary on its own."));

c.push(table(
  ["Party size", "Base XP", "Multiplier", "Adjusted XP", "Medium threshold", "Reads as"],
  [12, 12, 14, 14, 16, 22],
  [
    ["4", "800", "\u00D72", "1,600", "2,000", "Easy\u2013Medium"],
    ["5", "800", "\u00D72", "1,600", "2,500", "Easy\u2013Medium, softer"],
    ["6", "800", "\u00D71.5", "1,200", "3,000", "below Easy"]
  ],
  { full: true }
));

c.push(P("This is deliberately not a Hard or Deadly fight, and the table below is not a problem to fix. Historically, Richard\u2019s taking of Cyprus was closer to a rout than a battle \u2014 Isaac Komnenos\u2019s forces were simply outmatched \u2014 and the module is built the same way on purpose: the real tension is Sera\u2019s crew as captives (see below), not whether the party can win. If a table wants more bite, the cleanest way to add it is to give Thane one Veteran-statted lieutenant instead of two of the Bandits, which raises the ceiling without changing the monster count or crossing a boundary. Do not simply add more Bandits to compensate for a party of six \u2014 crossing from 5 to 7 monsters jumps the multiplier band up a full step and can overcorrect badly, exactly the trap CLAUDE.md warns about."));

c.push(P("If it comes to a fight, the captives are in an outbuilding across the hold\u2019s yard, guarded by one Bandit who will threaten them the moment things turn against Thane \u2014 not out of malice, but as a bargaining chip. A PC who breaks off to reach them (a round\u2019s worth of movement and a DC 13 Athletics or Acrobatics check to cross the yard without being cut off) heads that off cleanly. A table that ignores the captives entirely should still free them once the fight ends \u2014 this is tension, not a fail state."));

// -------------------------------------------------------------- Skill DCs
c.push(H2("Tiered Skill DCs"));

c.push(P("Easy 10, Moderate 13, Hard 16, matching the tiers used throughout this campaign."));

c.push(table(
  ["Task", "Skill", "DC", "Tier"],
  [40, 24, 12, 24],
  [
    ["Help the crew through the worst of the storm", "Athletics / Survival / relevant spell", "13", "Moderate"],
    ["Recognize Calanthe\u2019s coastline before making landfall", "Survival / Nature", "13", "Moderate"],
    ["Get Calanthe\u2019s frightened locals talking about Thane", "Persuasion / Insight", "10", "Easy"],
    ["Talk Thane\u2019s price down without a fight", "Persuasion / Intimidation", "15", "Hard"],
    ["Cross the yard to reach the captives once a fight starts", "Athletics / Acrobatics", "13", "Moderate"],
    ["Read that Thane is bluffing about a price he won\u2019t die for", "Insight", "13", "Moderate"]
  ],
  { full: true }
));

// ---------------------------------------------------------------- Stat Blocks
c.push(H2("Stat Blocks"));

c.push(...SB({
  name: "Warden Ivor Thane",
  meta: "Medium humanoid (human), any non-lawful alignment \u2014 SRD Veteran, renamed",
  ac: "17 (splint armor)",
  hp: "58 (9d8 + 18)",
  speed: "30 ft.",
  str: 16, dex: 13, con: 14, int: 10, wis: 11, cha: 10,
  skills: "Athletics +5, Perception +2",
  senses: "passive Perception 12",
  langs: "Common",
  cr: "3 (700 XP)",
  actions: [
    { n: "Multiattack", t: "Thane makes two longsword attacks. If he has a shortsword drawn, he can also make a shortsword attack." },
    { n: "Longsword", t: "Melee Weapon Attack: +5 to hit, reach 5 ft., one target. Hit: 7 (1d8 + 3) slashing damage, or 8 (1d10 + 3) slashing damage if used with two hands." },
    { n: "Shortsword", t: "Melee Weapon Attack: +5 to hit, reach 5 ft., one target. Hit: 6 (1d6 + 3) piercing damage." },
    { n: "Heavy Crossbow", t: "Ranged Weapon Attack: +3 to hit, range 100/400 ft., one target. Hit: 6 (1d10 + 1) piercing damage." }
  ]
}));

c.push(...SB({
  name: "Thane\u2019s Guard",
  meta: "Medium humanoid (any race), any non-lawful alignment \u2014 SRD Bandit, unmodified",
  ac: "12 (leather armor)",
  hp: "11 (2d8 + 2)",
  speed: "30 ft.",
  str: 11, dex: 12, con: 12, int: 10, wis: 10, cha: 10,
  senses: "passive Perception 10",
  langs: "Common",
  cr: "1/8 (25 XP)",
  actions: [
    { n: "Scimitar", t: "Melee Weapon Attack: +3 to hit, reach 5 ft., one target. Hit: 4 (1d6 + 1) slashing damage." },
    { n: "Light Crossbow", t: "Ranged Weapon Attack: +3 to hit, range 80/320 ft., one target. Hit: 5 (1d8 + 1) piercing damage." }
  ]
}));

// ---------------------------------------------------------------- Scene 4
c.push(H2("Scene 4: Calanthe\u2019s Fate"));

c.push(P("However Thane\u2019s hold was settled \u2014 paid off, talked down, or fought \u2014 the captives are freed, the fleet\u2019s missing stores are recovered from his cellars (along with a fair amount that is unmistakably not the coalition\u2019s, taken from other ships over other years), and the party is, whether they sought it or not, the reason a small independent island now has no one clearly in charge of it."));

c.push(BOX("Sera finds the party before anyone else does, already counting crates. \u201CThere\u2019s enough here to make good the losses and then some,\u201D she says. \u201CThat\u2019s the easy part. The island\u2019s the hard part. Someone\u2019s going to have to decide what happens to it, and I\u2019d rather it wasn\u2019t me.\u201D"));

c.push(P("Let the party decide Calanthe\u2019s fate. A coalition garrison secures the strait for the rest of the crusade\u2019s shipping, at the cost of putting an armed foreign power in charge of people who did not ask for one. Leaving Thane\u2019s household to govern itself, under new terms and a promise of no further tolls, costs the coalition a secure waypoint it may want later. A table can also split the difference \u2014 a light garrison with a local council retaining real authority \u2014 if they think to propose it; do not offer it as a third option unprompted."));

c.push(PS([DM("DM Only: "), { t: "record this choice in the DM Reference Guide as its own Branch Ledger entry, separate from the road choice. It is a small-scale rehearsal of the campaign\u2019s largest question \u2014 what a liberator owes the place it liberates \u2014 and it is worth revisiting, briefly, whenever the party\u2019s coalition allies discuss what Elduvaine should become after the war. Do not force the parallel; if the table draws it themselves, that is the module working as intended." }]));

c.push(P("The fleet resupplies, repairs what the storm damaged, and continues south within the day. Hand off directly to Module 5 for the approach to Vindana, where the mountain road\u2019s own losses \u2014 including the second king\u2019s \u2014 catch up with the party as news from the road not taken."));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("NPC Profiles"));

c.push(H2("Sera Vosk"));
c.push(P("An Auberitz-born quartermaster, mid-thirties, more practical than the situation strictly allows. Speech: fast, dry, allergic to wasted words \u2014 she will assess a room\u2019s exits before its people, and say so if asked. Not a fighter by training, but not helpless either; she has spent a career keeping ships supplied through worse than one bad morning."));
c.push(P("Open thread: Sera is a natural recurring contact for logistics and information for the rest of the campaign \u2014 a DM can bring her back at Vindana, or later, as the person who always seems to know where the supplies actually are. She has no stake in Calanthe\u2019s fate beyond wanting the crates loaded and the fleet moving, and will say so plainly if the party asks her opinion on the island\u2019s future."));

c.push(H2("Warden Ivor Thane"));
c.push(P("Middle-aged, comfortable rather than menacing, carries a ledger more often than a weapon. Speech: unhurried, transactional, genuinely puzzled by moral objections to what he considers a straightforward toll on a strait he happens to control. Not cruel for its own sake \u2014 captives are an asset, and a damaged asset is worth less."));
c.push(P("Open thread: if he survives \u2014 captured, exiled, or simply talked out of his hold \u2014 Thane is exactly the kind of small, self-interested operator who resurfaces later in a campaign, working for whoever will have him. A DM looking for a low-stakes recurring rogue later in the crusade has him available."));

// --------------------------------------------------------------- Optional
c.push(H1("Optional Content"));

c.push(H2("What the Locals Remember"));
c.push(P("Before or after Thane\u2019s hold, let the party spend some time among Calanthe\u2019s ordinary people \u2014 fisherfolk, a handful of farms, an old woman who remembers a Warden before Thane and says the position used to mean something. Play this for texture and moral weight rather than plot: nobody here asked for a war to wash up on their coast, and their opinions about what should happen next are worth hearing even though the party is not obligated to follow them."));

c.push(H2("The Second Ship"));
c.push(P("If the table wants more to do before Thane\u2019s hold, a third coalition vessel can be found wrecked on a further stretch of coast, its survivors already having made their own uneasy peace with a Calanthe fishing village. A short scene of relief and minor logistics \u2014 no combat required \u2014 that adds texture without adding a beat the module needs."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("How Thane\u2019s hold is settled.", "Paid off, talked down, or fought \u2014 track which. A talked-down or paid-off resolution leaves Thane in a position to resurface later (see his NPC profile); a fight that kills or captures him removes that thread but earns Calanthe\u2019s locals more open gratitude in the near term."));
c.push(BUL("Calanthe\u2019s fate.", "Garrisoned by the coalition, left independent under new terms, or a hybrid the table proposes \u2014 record this as its own Branch Ledger entry, separate from the road choice, and note it as the campaign\u2019s first small-scale rehearsal of what a liberator owes the place it liberates."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BUL("The recovered stores.", "Enough of the fleet\u2019s lost supplies, plus a portion of Thane\u2019s accumulated plunder, to be a genuine funding boost for the crusade \u2014 play this narratively as the coalition arriving at Vindana somewhat better supplied than it otherwise would have, rather than a large individual coin payout at 5th level."));
c.push(BUL("Thane\u2019s ledger.", "A genuinely useful item: years of records of ships, cargoes, and (if he is captured or cooperative) contacts along the southern sea-lanes. A DM can use it to justify a Diverging Paths hook or a piece of information the party needs later without it feeling handed to them."));
c.push(BUL("A signet ring.", "Thane\u2019s own, if he is defeated or surrenders it as part of a bargain. Non-magical, but recognizable to anyone in the region who has dealt with Calanthe\u2019s toll \u2014 a DM may let it open doors, or close them, wherever that is true."));

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
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, font: "Georgia", color: "3B2F2F" }, paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: "Georgia", color: "3B2F2F" }, paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 22, bold: true, font: "Georgia", color: "3B2F2F" }, paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } }
    ]
  },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } }, children: c }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(stagePath("KC_Module02A_TheSeaRoad.docx"), buf);
  console.log("Written.");
});
