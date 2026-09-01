// KC_Module01_TheMuster.js -- Session Module One: The Muster.
//
// Canon lives here. corpus/ and documents/ are generated from this file and are
// never edited by hand. See CLAUDE.md for the sign-off rules: anything in the
// "Not yet decided" table must not appear here until it has been approved.
//
// ESCAPE CONVENTION: all prose lives as \uXXXX escapes, never as literal
// typographic characters. tools/build.sh fails the build if that slips.
//
// Module shape per CLAUDE.md: overview + pacing budget, What Is Actually
// Happening (DM Only), numbered scenes with boxed read-aloud, tiered skill
// DCs, full stat blocks, NPC profiles, Optional Content, Diverging Paths
// (DM Only), loot, and the Refrain -- printed identically, never varied,
// except in the campaign's final module.
//
// Encounter design note: the wyvern below is the SRD monster unmodified (CR 6,
// 110 hp, +7 to hit, stinger DC 15 Con save for 7d6 poison), pulled from
// 5e-bits/5e-database rather than from memory, per Mechanical Validation.
// Adjusted-XP math against DMG thresholds for 4/5/6 characters is computed and
// stated in Scene 3 rather than asserted.

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
  children: [new TextRun({ text: "The Muster", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module One", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("The party is summoned to Duncarrow, seat of Xavier III of Harrowmark, and chosen personally for the crusade he has called. Before the muster can march, Harrowmark asks one more thing of them: help see off a wyvern that has turned on the hold of Greywatch. The session ends with the coalition\u2019s ships or column setting out, and the table making the campaign\u2019s first real choice \u2014 the sea road or the mountain road. Core scenes run three and a half to four hours; Optional Content is built to fill out the rest of a five-hour session, not to be squeezed into it, and can be skipped entirely on a slow night without costing the module anything it needs later."));

c.push(P("If the party does not already know one another, run Scene 1 as their introduction rather than as exposition \u2014 they can meet on the road, all answering the same rider. If they do already know each other, skip straight to the summons and let the scene move."));

c.push(H2("The Weight of Harrowmark"));

c.push(P("Play the texture of the place before anything else happens in it. Harrowmark in the last cold weeks before winter is grey stone under a greyer sky, woodsmoke that never quite clears the low valleys, and light that arrives late and leaves early. Nothing glows, nothing hums, nothing remembers a kindness done to it \u2014 water is only water here, and a mile is exactly as long as it looks. That flatness is not absence; it is a fact about the place, stated plainly, and it is what will make Elduvaine\u2019s magic land as wonder rather than as a rules explanation when the party finally crosses into it."));

c.push(P("What Harrowmark has instead is competence. Doors are barred against weather, not spirits. Watch fires burn all night on every hold\u2019s high ground, not from superstition but because something with wings sometimes comes out of the crags, and has for longer than anyone can date. People here do not gasp at danger; they reach for the nearest rope, or pike, or child, in that order, and get on with it. Voice every NPC in this module \u2014 Xavier included \u2014 from that same flat, competent register, and save the wonder for later."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 15, 55],
  [
    ["1. The Summons", "20 min", "Optional introductions fold in here if needed."],
    ["2. Audience with the King", "45\u201360 min", "The Promise, the crusade, and the party\u2019s place in it."],
    ["3. The Wyvern at Greywatch", "90\u2013120 min", "The session\u2019s combat. See DC table and stat block below."],
    ["4. The Muster\u2019s End", "30\u201345 min", "Departure, and the road choice \u2014 Branch Ledger entry 1."],
    ["Optional Content", "30\u201345 min", "Run if the table has time; cut cleanly if not."]
  ],
  { full: true }
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("Nothing in this module is a trick. The wyvern at Greywatch is exactly what Brenna Vane says it is: a young, hungry animal pushed out of its usual range by a harder winter than most, doing what wyverns do. It has no connection to Vale, to Elduvaine, or to anything the party will learn about later. Its only job is to be the first thing this campaign asks the party to be brave about, and to show them \u2014 and the table \u2014 what \u201Cunimpressed\u201D actually looks like from the inside."));

c.push(PS([DM("DM Only: "), { t: "Xavier is not yet the Wyvernheart, and nothing in this module may suggest that he will be. He does not fly, does not fight the wyvern personally if he is present at Greywatch (he should not be \u2014 keep him at Duncarrow), and no NPC may use the name early, even as a joke or a slip. See CLAUDE.md under Canon and Sources of Truth if you are unsure why this matters this much." }]));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: The Summons"));

c.push(P("Wherever the party is when this begins, make it somewhere ordinary \u2014 a roadside inn, a garrison bunkhouse, a farmhold at the edge of the crags. The rider is easy to hear before he is seen: hooves on frozen ruts, faster than anyone rides without a reason, and by the time he is in view his horse is lathered white at the neck despite the cold."));

c.push(BOX("The rider comes at a hard trot, mud and ice to the knee of his horse, and does not slow at the gate. \u201CDuncarrow,\u201D he says, to whoever is closest, breath fogging between words. \u201CThe king summons you by name. Today, if you can manage it.\u201D He says each name once, plainly, the way a man reads a list he has already checked twice, and does not wait to be thanked before he turns his horse back the way he came."));

c.push(P("Play this scene short. Its only job is to get the party moving toward Duncarrow with a name attached to the summons \u2014 their own. If the party has not met, this is where they cross paths: on the same road, answering the same rider, sizing each other up over the two or three days\u2019 ride it takes to reach the seat, trading the little each of them knows about why a king would want them by name. Nobody in Harrowmark is summoned by name for a small reason, and the party should arrive at Duncarrow already half-certain of that, before anyone tells them so."));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: Audience with the King"));

c.push(P("Duncarrow is a working fortress, not a palace \u2014 grey stone, low ceilings, arrow-slit windows that let in more cold than light. The great hall smells of woodsmoke, wet dog, and oiled steel rather than incense; the tapestries on its walls are march-maps and hide charts, not heraldry. Guards at the door look tired rather than ceremonial, and nod the party through without announcing them \u2014 Duncarrow has had a great many visitors this season, and does not have the staff to make an occasion of one more."));

c.push(P("Xavier receives the party standing, at a steward\u2019s table crowded with maps and weighted at the corners with what look like ordinary stones. He is not what the songs will eventually make of him: greying at the temples earlier than he should be, a soldier\u2019s build going a little soft from a year of councils rather than campaigns, dressed plainly enough that a stranger might take him for one of his own captains. Only when he looks up does the room organize itself around him \u2014 not because anyone commands it to, but because everyone in it has clearly done this before."));

c.push(BOX("He does not make them kneel. \u201CI know your names,\u201D he says, before anyone has given one, and it is not a boast \u2014 he simply has. \u201CA kingdom I have never seen has been taken by a man I have never met, and I am asking you to help me take it back. I will not pretend that is a small thing to ask of anyone."));

c.push(P("Let Xavier speak plainly about the Promise \u2014 a share in Elduvaine\u2019s resident magic, and the run of the Ysolde Archive, for those who answer the call \u2014 without dwelling on it; the sourcebook covers its substance, and this scene\u2019s job is the man, not the policy. He is direct, economical, uncomfortable with the theatre of being knelt to, and visibly more at ease discussing march order and supply than titles. If a player asks him something personal, he answers once, briefly, and changes the subject; he did not survive his own youth by talking about it."));

c.push(PS([DM("DM Only: "), { t: "if a player pushes on why he chose them specifically, he has a real answer, not a flattering one: he asked his officers for the people other people trusted under pressure, not the people other people were impressed by. He believes this is the same thing Harrowmark has always valued in its own wyvern-hunters, and it is not a coincidence that Greywatch is where he is about to send them next." }]));

c.push(P("He closes the audience by asking one thing of them before the muster marches: Greywatch, three days out, has sent word that a wyvern has taken to raiding its herds and, twice now, its people. He would rather send the crusade\u2019s own chosen than levy more of Greywatch\u2019s own dead."));

c.push(P("He does not walk them out. A steward does that, through a yard full of the ordinary business of a fortress readying for war \u2014 barrels counted, horses shod, a smith\u2019s hammer going somewhere out of sight \u2014 and the party\u2019s last sight of Xavier, glanced back through the hall door, is of a man already bent over his maps again before they have finished crossing the threshold."));

// ---------------------------------------------------------------- Scene 3
c.push(H2("Scene 3: The Wyvern at Greywatch"));

c.push(P("Greywatch announces itself before the road does: watchtowers on every high point, each with a horn-bearer standing a longer shift than seems reasonable, and a palisade studded here and there with a wyvern claw or a length of jaw nailed up as a ward that everyone insists is decoration and no one will take down. It is a hold built for exactly this, by people who have buried others to wyverns before and, without much ceremony about it, expect to again. Huntmaster Brenna Vane meets the party at the gate, sizes them up without much comment, and puts them to work."));

c.push(BOX("\u201CIt took two lambs and a shepherd\u2019s boy\u2019s arm Tuesday last,\u201D Brenna says, already walking, not waiting to see if they follow. \u201CYoung, hungry, pushed down off the high crags by the cold \u2014 not clever, not old enough to be clever. We\u2019ll rope it if we can and put it back where it came from. We\u2019ll kill it if we have to. Either way, nobody\u2019s grandmother is losing a leg on my watch. You keep up or you keep out of the way.\u201D"));

c.push(P("Give the party the rest of the daylight to see Greywatch properly before dusk brings the fight: pens reinforced with iron banding rather than wood, a bone yard behind the smithy where old kills are rendered down for nothing gets wasted here, children who know the horn-calls \u2014 one long for sighted, three short for struck, one long again for down \u2014 the way other children know a nursery rhyme. Nobody here is grim about it. This is simply what a season looks like."));

c.push(H3("Running the Encounter"));

c.push(P("The wyvern comes at dusk, out of a sky gone the color of a bruise, and is heard before it is seen \u2014 a dry, heavy beat of wings, then a cry that every dog in the hold answers at once. It strikes from height, going first for whatever looks least defended: a penned animal, a straggler, a party member who has wandered from the group. Brenna and two or three Greywatch hunters (use commoner or scout-tier statistics as needed; they are support, not combatants, and should not be expected to survive a direct hit) fight alongside the party with ropes and long pikes rather than swords, calling positions to each other in the flat, practiced shorthand of people who have done this before."));

c.push(P("Harrowmark does not fight wyverns the way most parties expect to fight a dragon. A PC adjacent to the wyvern, or within its reach, may spend their action fixing a hunting rope to a wing or leg (Strength (Athletics) or Dexterity (Sleight of Hand), DC 13 \u2014 Brenna will call out which, and why, the first time it comes up). Three successful ropes from three different characters ground the wyvern: it loses its fly speed until it breaks free (its action, a DC 13 Athletics contest against the ropes) or the encounter ends. A grounded wyvern fighting on the ground, hemmed in by pikes, is a substantially safer fight than one still in the air, and this is deliberate \u2014 it is the module\u2019s built-in partial resolution, not a trick or a trap. Play the moment it goes down for everything it is worth: the ropes snapping taut, the hold-folk shouting the count, the sudden and enormous silence of something huge no longer in the air."));

c.push(P("The stinger is the fight\u2019s real danger and should read as one \u2014 not a number on a sheet but a visible, spreading wrongness in whoever it hits: the wound going numb, then cold, then someone else shouting for the antitoxin before the poison finishes its work. The wyvern is unaligned and not suicidal. If it drops to 25 hit points or fewer, or is grounded and clearly losing, it breaks off and flees at its next opportunity rather than fighting to the death. A fleeing wyvern that escapes counts as a full success at the table \u2014 Greywatch does not expect a corpse, only a hold that is safe again."));

c.push(H3("Scaling the Fight"));

c.push(P("The wyvern below is the SRD statistics unmodified, run against Dungeon Master\u2019s Guide adjusted-XP thresholds for 5th-level characters (Easy 250 / Medium 500 / Hard 750 / Deadly 1100 per character). A single monster carries a \u00D71 multiplier at party sizes of 3\u20135 and \u00D70.5 at 6 or more \u2014 the party-of-six discount CLAUDE.md warns about, and it lands hard here."));

c.push(table(
  ["Party size", "Adjusted XP", "Deadly threshold", "Reads as", "Compensate with"],
  [14, 16, 16, 18, 36],
  [
    ["4", "2,300", "4,400", "Medium\u2013Hard", "Run as written."],
    ["5", "2,300", "5,500", "just under Medium", "Run as written; the rope tactic and morale give it teeth the raw math understates."],
    ["6", "1,150", "6,600", "below Easy", "Do not add a second wyvern \u2014 crossing the 1\u21922 monster boundary roughly doubles the multiplier and overcorrects badly. Instead add one complication: a burning hayrick blocking the herd pens, or a trapped child in the loft, that raises the practical stakes without touching the monster count."]
  ],
  { full: true }
));

// -------------------------------------------------------------- Skill DCs
c.push(H2("Tiered Skill DCs"));

c.push(P("Easy 10, Moderate 13, Hard 16, matching the tiers used throughout this campaign."));

c.push(table(
  ["Task", "Skill", "DC", "Tier"],
  [40, 24, 12, 24],
  [
    ["Notice the summons rider is telling the truth about Duncarrow", "Insight", "10", "Easy"],
    ["Read that Xavier is uncomfortable with pageantry, not with them", "Insight", "13", "Moderate"],
    ["Track the wyvern\u2019s likely strike point before dusk", "Survival", "13", "Moderate"],
    ["Fix a hunting rope to the wyvern (per attempt; three needed)", "Athletics / Sleight of Hand", "13", "Moderate"],
    ["Steady panicking livestock or hold-folk mid-attack", "Animal Handling / Persuasion", "13", "Moderate"],
    ["Spot the stinger wind-up in time to call a warning", "Perception", "16", "Hard"],
    ["Earn Brenna\u2019s open respect rather than her polite tolerance", "Persuasion (or demonstrated competence)", "16", "Hard"]
  ],
  { full: true }
));

// ---------------------------------------------------------------- Stat Block
c.push(H2("Stat Block"));

c.push(...SB({
  name: "Wyvern",
  meta: "Large dragon, unaligned",
  ac: "13 (natural armor)",
  hp: "110 (13d10 + 39)",
  speed: "20 ft., fly 80 ft.",
  str: 19, dex: 10, con: 16, int: 5, wis: 12, cha: 6,
  skills: "Perception +4",
  senses: "darkvision 60 ft., passive Perception 14",
  langs: "\u2014",
  cr: "6 (2,300 XP)",
  actions: [
    { n: "Multiattack", t: "The wyvern makes two attacks: one with its bite and one with its stinger. While flying, it can use its claws in place of one other attack." },
    { n: "Bite", t: "Melee Weapon Attack: +7 to hit, reach 10 ft., one creature. Hit: 11 (2d6 + 4) piercing damage." },
    { n: "Claws", t: "Melee Weapon Attack: +7 to hit, reach 5 ft., one target. Hit: 13 (2d8 + 4) slashing damage." },
    { n: "Stinger", t: "Melee Weapon Attack: +7 to hit, reach 10 ft., one creature. Hit: 11 (2d6 + 4) piercing damage. The target must make a DC 15 Constitution saving throw, taking 24 (7d6) poison damage on a failed save, or half as much damage on a successful one." }
  ]
}));

// ---------------------------------------------------------------- Scene 4
c.push(H2("Scene 4: The Muster\u2019s End"));

c.push(P("Whatever became of the wyvern, the party returns to Duncarrow to find the muster nearly ready to move. The yard that was ordinary business three days ago is now a camp too large for it \u2014 wagons loaded and lashed, Harrowmark levies falling into column by the sound of a horn rather than a shouted order, and, past the edge of the Harrowmark tents, the first coalition banners: colors and cuts of armor the party has not seen before, voices in accents that mark out where in the muster Oksitan and Auberitz have pitched their own camps. Nobody introduces the party to any of it yet \u2014 that is a later module\u2019s work \u2014 but it should be visibly, unmistakably there, the first sign that this war is bigger than Harrowmark."));

c.push(P("Xavier meets them once more, briefly, in the noise of it rather than the quiet of his hall, to say the thing he did not say at the audience."));

c.push(BOX("\u201CGreywatch will remember this longer than I will be able to thank you for it,\u201D he says, close enough to be heard over the camp without raising his voice. \u201CThat is worth more than anything I can pin to your coats. But I mean to pin something anyway.\u201D"));

c.push(P("This is where the road is chosen \u2014 sea or mountain. Let the party decide it in character, weighing whatever they have heard in camp (speed and a side-conquest against the sea, or a longer and harder overland march against the mountains); do not let Xavier decide it for them. Whichever they choose, hand off directly to that module."));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("NPC Profiles"));

c.push(H2("Xavier III of Harrowmark"));
c.push(P("Speech: short sentences, plain words, and a habit of answering the question actually asked rather than the one people expect. Never raises his voice; the room quiets for him anyway."));
c.push(P("Open thread: he has told no one in this party, and will tell no one for some time, that leaving Harrowmark under a regency he does not fully trust is the decision from this crusade that frightens him most \u2014 not Vale, not the road, that. A DM looking for a private, human moment with him later in the campaign has this to draw on."));

c.push(H2("Huntmaster Brenna Vane"));
c.push(P("Late fifties, missing two fingers on her left hand to a stinger she still calls \u201Ca fair trade.\u201D Runs Greywatch\u2019s wyvern-watch and, unofficially, a hold-wide betting ledger on every kill \u2014 which limb, which method, how long. Speech: dry, economical, allergic to being thanked directly; deflects gratitude by immediately assigning a chore."));
c.push(P("Open thread: if the party impresses her, she offers what Greywatch actually has to give \u2014 not gold, but competence: a standing invitation to send word if Harrowmark-trained hands are ever needed again, which a DM can call in during Module 8 or later as a recurring, grounded ally rather than a one-scene NPC."));

// --------------------------------------------------------------- Optional
c.push(H1("Optional Content"));

c.push(H2("The Betting Ledger"));
c.push(P("Before or after the wyvern, Brenna\u2019s ledger makes the rounds of Greywatch\u2019s hall \u2014 a battered book of wagers going back decades, with a fresh page opened the moment the party arrives. Play this for pure levity: hold-folk arguing odds, someone trying to get the party to wager on themselves, an old wound reopened as an anecdote rather than a warning. No mechanical stakes; this is the relief valve CLAUDE.md asks every module to plan for rather than hope for."));

c.push(H2("Fortifying Greywatch"));
c.push(P("If the table wants a prep phase before the wyvern strikes, run a short group effort to reinforce the hold\u2019s pens and towers \u2014 two or three successful checks (Athletics, Investigation, or anything else a player can justify) buy the party an advantageous position for the fight: advantage on the first round\u2019s initiative, or the wyvern\u2019s opening attack targeting a decoy instead of a person. Skip this cleanly if the table is eager to get to the fight itself."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("The wyvern\u2019s fate.", "Killed, driven off, or (rarely) roped and released elsewhere \u2014 track which. Greywatch remembers a killing as competence and a driving-off as mercy; both earn Brenna\u2019s respect, and only a table that badly mishandles the fight (fleeing, refusing to help, or getting a Greywatch hunter killed through carelessness) should lose it."));
c.push(BUL("The road choice \u2014 Branch Ledger entry 1.", "Sea road or mountain road, decided here and carried into Module 2A or 2B. This is the campaign\u2019s first tracked divergence; record it in the DM Reference Guide the moment it is made, including who argued for which and why, since that reasoning is worth revisiting when the drowned king\u2019s aftermath reaches the party in Module 5."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BUL("A Greywatch hunting pike.", "Masterwork but not magical \u2014 a reach weapon balanced by people who use them for a living. Reflavor as a longspear or glaive per the wielder\u2019s preference; treat as a normal weapon of its type, finely made."));
c.push(BUL("A vial of Greywatch antitoxin.", "A single-use draught brewed from the hold\u2019s long experience with stinger wounds. Grants advantage on the next saving throw against poison damage or the poisoned condition within the next 24 hours."));
c.push(BUL("Xavier\u2019s token.", "A plain iron badge, the king\u2019s own mark, given without ceremony. It carries no mechanical benefit and opens no doors by itself \u2014 but it is recognizable to anyone who has served under Xavier, and a DM may let it matter socially wherever that is true."));

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
  fs.writeFileSync(stagePath("KC_Module01_TheMuster.docx"), buf);
  console.log("Written.");
});
