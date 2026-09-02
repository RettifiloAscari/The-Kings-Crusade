// KC_Bestiary.js -- creatures of Elduvaine, the occupation, and the road between.
//
// Every block is calibrated against real SRD creatures at the same and neighbouring
// CR, never against the DMG Monster Statistics by CR table -- see CLAUDE.md.
//
// Canon lives here alongside KC_Sourcebook.js. corpus/ and documents/ are
// generated from this file and are never edited by hand. See CLAUDE.md for the
// sign-off rules: anything in the "Not yet decided" table must not appear here
// until it has been approved.
//
// ESCAPE CONVENTION: all prose lives as \uXXXX escapes, never as literal
// typographic characters. tools/build.sh fails the build if that slips.
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

// paragraph from segments: [{t: "text", b: bool, i: bool, c: "RRGGBB"}]
const PS = (segs, opts = {}) => new Paragraph({
  spacing: { after: 200 },
  ...opts,
  children: segs.map(s => new TextRun({ text: s.t, bold: !!s.b, italics: !!s.i, color: s.c }))
});

const DM = (t) => ({ t, b: true, c: "5B1F1F" });   // DM-only marker: bold book-red
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

// boxed read-aloud text
const BOX = (text) => new Paragraph({
  spacing: { before: 120, after: 160 },
  shading: { type: "clear", fill: "F3EFE4" },
  indent: { left: 220, right: 220, firstLine: 0 },   // template default firstLine=180 otherwise leaks in
  keepLines: true,   // never let a boxed passage tear across a page break
  children: [new TextRun({ text, italics: true })]
});

// Verse in a read-aloud box: keeps its line breaks instead of running together.
const VERSE = (lines) => new Paragraph({
  spacing: { before: 120, after: 160 },
  shading: { type: "clear", fill: "F3EFE4" },
  indent: { left: 220, right: 220, firstLine: 0 },   // same fix as BOX -- see its comment
  keepLines: true,   // the refrain must never split across a page break
  children: lines.map((l, i) => new TextRun({ text: l, italics: true, ...(i ? { break: 1 } : {}) }))
});

// Artwork from images/. Width and height are points at 72/inch, so 288 is four
// inches. `mdPath` is the shim-only field the Markdown corpus links with; it is
// relative to corpus/, where the generated .md lives. The real docx ignores it.
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
const cell = (text, opts = {}) => new TableCell({ width: { size: opts.w || 20, type: WidthType.PERCENTAGE }, shading: opts.head ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, margins: { top: 60, bottom: 60, left: 110, right: 110 }, children: [new Paragraph({ spacing: { after: 0 }, indent: { firstLine: 0 }, children: [new TextRun({ text, bold: !!opts.head, size: 18 })] })] });
// cantSplit keeps a row\u2019s cells from being torn across a column or page break;
// tableHeader repeats the header row when a long table does span a break.
const row = (cells, opts = {}) => new TableRow({ children: cells, cantSplit: true, ...opts });
const FULLWIDTH = "KCFullWidth";   // marker only; transplant.py acts on it and strips it
// docx-js emits <w:tblGrid> only when given columnWidths in DXA. Without a grid
// LibreOffice ignores the per-cell percentages and distributes columns evenly, so a
// d6 column holding one digit took a third of the table. Only the ratios matter.
const GRID = 9360;
const table = (headers, widths, rows, opts = {}) => new Table({ ...(opts.full ? { style: FULLWIDTH } : {}), layout: TableLayoutType.FIXED, columnWidths: widths.map(w => Math.round(w / 100 * GRID)), width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ row(headers.map((h, i) => cell(h, { head: true, w: widths[i] })), { tableHeader: true }), ...rows.map(r => row(r.map((v, i) => cell(v, { w: widths[i] })))) ] });

const mod = (v) => { const m = Math.floor((v - 10) / 2); return (m >= 0 ? "+" : "\u2212") + Math.abs(m); };
const abCell = (text, bold) => new TableCell({ width: { size: 16.6, type: WidthType.PERCENTAGE }, shading: bold ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40, before: 40 }, indent: { firstLine: 0 }, keepNext: !!bold, children: [new TextRun({ text, bold: !!bold, size: 20 })] })] });
const SB = (d) => { const out = []; out.push(new Paragraph({ spacing: { before: 240, after: 40 }, children: [new TextRun({ text: d.name, bold: true, size: 26, color: "5B1F1F" })] })); out.push(PS([{ t: d.meta, i: true }], { spacing: { after: 120 } })); out.push(B("Armor Class:", d.ac)); out.push(B("Hit Points:", d.hp)); out.push(B("Speed:", d.speed)); out.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ new TableRow({ cantSplit: true, tableHeader: true, children: ["STR","DEX","CON","INT","WIS","CHA"].map(h => abCell(h, true)) }), new TableRow({ cantSplit: true, children: [d.str,d.dex,d.con,d.int,d.wis,d.cha].map(v => abCell(v + " (" + mod(v) + ")")) }) ] })); out.push(P("", { spacing: { after: 60 } })); if (d.saves) out.push(B("Saving Throws:", d.saves)); if (d.skills) out.push(B("Skills:", d.skills)); if (d.senses) out.push(B("Senses:", d.senses)); if (d.langs) out.push(B("Languages:", d.langs)); out.push(B("Challenge:", d.cr)); (d.traits||[]).forEach(t => out.push(PS([{ t: t.n + ". ", b: true, i: true }, { t: t.t }]))); if (d.actions && d.actions.length) { out.push(PS([{ t: "ACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.actions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } if (d.reactions && d.reactions.length) { out.push(PS([{ t: "REACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.reactions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } return out; };

// ---------- content ----------
const c = [];

c.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [new TextRun({ text: "The King\u2019s Crusade", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "A Bestiary of Elduvaine, the Occupation, and the Road Between", i: true }],
  { alignment: AlignmentType.CENTER }));

c.push(H1("Using This Bestiary"));

c.push(P("Every block here was built against real SRD creatures at the same and neighbouring challenge ratings, not against the DMG\u2019s Monster Statistics by Challenge Rating table. Measured against that table every humanoid in this book looks badly under-tuned, and so does every humanoid in the SRD: a CR 3 Veteran has 58 hit points where the table asks for 101 to 115. The table describes monsters. These are mostly people, and people are squishier than the arithmetic wants them to be."));

c.push(P("Where a creature is deliberately off-baseline it says so in its own entry. Support casters, pack fighters and control-focused adversaries are meant to be individually weak and collectively appalling, and correcting them toward the curve would break the encounters they were built for."));

c.push(H2("What Resident Magic Produces"));

c.push(P("The creatures in the first section are not monsters that happen to know magic. They are habits with appetites \u2014 the same kind of thing as the Willing Road or the Listening Water, given shape because a shape was useful. This has one consequence that runs through every entry: as its region is drained, a resident-magic creature visibly suffers, changes, or goes wrong. None of them are immune to what is being done to Elduvaine, and several of them are the most direct way a party can see it happening."));

c.push(PS([DM("DM Only: "), { t: "the single best use of this section is not combat. Three of the six creatures below have a nonviolent resolution that is more rewarding than the fight, and one of them cannot meaningfully be fought at all. A party that reaches for initiative every time will get a serviceable campaign. A party that occasionally does not will get this one." }]));

c.push(H1("Creatures of the Living Realm"));

c.push(...SB({
  name: "Waystone Warden",
  meta: "Large construct, unaligned",
  ac: "16 (natural armor)", hp: "105 (14d10 + 28)", speed: "30 ft.",
  str: 20, dex: 10, con: 15, int: 6, wis: 16, cha: 8,
  skills: "Perception +6", senses: "Truesight 60 ft., passive Perception 16",
  langs: "understands Elduvish and Sylvan but does not speak", cr: "5 (1,800 XP)",
  traits: [
    { n: "Unbarred Road", t: "The warden cannot attack, and has no reaction against, any creature that has not offered violence on the Willing Road within the last hour. This is not a moral judgement and cannot be argued with, invoked, or appealed to. It is simply how the thing works, and a party that discovers it has discovered the single most useful fact in this entry." },
    { n: "Stonebound", t: "The warden cannot move more than 120 feet from a waystone. There are three hundred and eleven waystones and it has never been established how the warden chooses between them." },
    { n: "Failing", t: "In a drained region the warden has disadvantage on all attack rolls and its speed is halved. Two have been found stopped entirely, upright, in the road, and nobody has been able to move them." }
  ],
  actions: [
    { n: "Multiattack", t: "The warden makes two slam attacks." },
    { n: "Slam", t: "Melee Weapon Attack: +8 to hit, reach 10 ft., one target. Hit: 14 (2d8 + 5) bludgeoning damage." },
    { n: "Set Aside (Recharge 5-6)", t: "The warden lifts one Large or smaller creature it can reach and places it 30 feet away on the road, gently. The target must succeed on a DC 15 Strength saving throw or be moved and knocked prone. This deals no damage. The warden much prefers it to fighting." }
  ]
}));

c.push(PS([DM("DM Only: "), { t: "the warden is a puzzle wearing a stat block. Its first action in almost every encounter should be Set Aside, repeatedly, on whoever is in the way \u2014 which reads to players as a fight going strangely and then, about three rounds in, as something else entirely. Do not explain Unbarred Road. Let them work out that the party member who has not attacked is being ignored." }]));

c.push(...SB({
  name: "The Withering",
  meta: "Medium fey, neutral",
  ac: "13", hp: "60 (11d8 + 11)", speed: "30 ft.",
  str: 10, dex: 16, con: 13, int: 14, wis: 15, cha: 18,
  skills: "Perception +4, Stealth +5", senses: "Darkvision 60 ft., passive Perception 14",
  langs: "Elvish, Sylvan, Elduvish", cr: "3 (700 XP)",
  traits: [
    { n: "Was a Dryad", t: "The Withering is what a dryad becomes when the wood it was bound to is drained rather than felled \u2014 still alive, still bound, and bound now to something that is no longer there. It is not undead and it is not evil. It is bereaved, and it has been bereaved continuously for somewhere between four months and two years." },
    { n: "Grieving", t: "The Withering has advantage on saving throws against being charmed or frightened, and disadvantage on any check made to leave the boundary of what its wood used to be." },
    { n: "Speak with Beasts and Plants", t: "The Withering can communicate with beasts and plants as if they shared a language. Almost nothing answers any more." }
  ],
  actions: [
    { n: "Multiattack", t: "The Withering makes two claw attacks." },
    { n: "Claw", t: "Melee Weapon Attack: +5 to hit, reach 5 ft., one target. Hit: 9 (2d6 + 3) slashing damage." },
    { n: "Bitter Bough (Recharge 5-6)", t: "Dead wood erupts in a 20-foot radius. Each creature in the area must make a DC 14 Dexterity saving throw, taking 21 (6d6) piercing damage on a failure, or half as much on a success. The area becomes difficult terrain." }
  ]
}));

c.push(PS([DM("DM Only: "), { t: "the Withering will stop if anybody tells it, plainly and without softening, what happened to its wood and who did it. Not a Persuasion check \u2014 an actual answer, delivered by a player, out loud. It has spent months unable to work out what it did wrong. It is not attacking the party because it is hostile; it is attacking them because they are the first thing to come through and it has no other way to ask. Run the fight, and let somebody talk, and end it the moment they do." }]));

c.push(...SB({
  name: "Light-Hollow",
  meta: "Medium construct, unaligned",
  ac: "17 (natural armor)", hp: "85 (10d8 + 40)", speed: "20 ft.",
  str: 16, dex: 8, con: 18, int: 3, wis: 10, cha: 1,
  senses: "Blindsight 60 ft. (blind beyond this radius), passive Perception 10",
  langs: "\u2014", cr: "4 (1,100 XP)",
  traits: [
    { n: "Emptied Stone", t: "A Light-Hollow is a piece of Standing Light masonry that was drained past what it could survive. Six centuries of held afternoons came out of it in an afternoon, and what was left got up. There are perhaps forty of them in Elduvaine and every one was somebody\u2019s doorstep, lintel or garden wall." },
    { n: "Thirst", t: "Any nonmagical light source within 30 feet of the Light-Hollow is extinguished at the start of the creature\u2019s turn. A magical light of 3rd level or lower is suppressed while it remains in range; the caster may make a DC 15 Constitution check at the end of each of their turns to restore it." },
    { n: "Immutable Form", t: "The Light-Hollow is immune to any spell or effect that would alter its form." }
  ],
  actions: [
    { n: "Multiattack", t: "The Light-Hollow makes two slam attacks." },
    { n: "Slam", t: "Melee Weapon Attack: +6 to hit, reach 5 ft., one target. Hit: 11 (2d8 + 3) bludgeoning damage." },
    { n: "Draw Down (Recharge 6)", t: "The Light-Hollow pulls at every light in a 30-foot radius, its own included. Each creature in the area must succeed on a DC 15 Constitution saving throw or take 18 (4d8) necrotic damage and be blinded until the end of its next turn. The area is plunged into magical darkness until the start of the Light-Hollow\u2019s next turn." }
  ]
}));

c.push(PS([DM("DM Only: "), { t: "fight this in the dark. That is the entire design: Thirst removes the party\u2019s torches on round one, Draw Down removes everything else on round three or four, and the encounter becomes a problem about position and memory rather than about damage. A party that thinks to bring a light-stone lamp \u2014 which holds daylight rather than making it \u2014 finds it works, once, and is worth every copper they paid for it." }]));

c.push(...SB({
  name: "Season-Bound Stag",
  meta: "Large fey, neutral good",
  ac: "13", hp: "45 (6d10 + 12)", speed: "60 ft.",
  str: 16, dex: 16, con: 14, int: 8, wis: 16, cha: 14,
  skills: "Perception +5, Stealth +5", senses: "passive Perception 15",
  langs: "understands Sylvan and Elduvish but does not speak", cr: "2 (450 XP)",
  traits: [
    { n: "Of Its Wood", t: "The stag carries the season of the Kept Season wood that bore it, visibly, wherever it goes \u2014 frost on its flanks in high summer, or blossom caught in its antlers in the depth of winter, always the wrong one, always four days or nine weeks into it. It is the most photographed thing in Elduvaine, in the sense that every visiting scholar for four centuries has tried to draw one." },
    { n: "Season\u2019s Passage", t: "The stag ignores difficult terrain and cannot be tracked by nonmagical means." },
    { n: "Skittish", t: "The stag will not initiate combat and flees at the first opportunity. It has no morale to break; it simply leaves." }
  ],
  actions: [
    { n: "Ram", t: "Melee Weapon Attack: +5 to hit, reach 5 ft., one target. Hit: 12 (2d8 + 3) bludgeoning damage, and the target must succeed on a DC 13 Strength saving throw or be knocked prone." },
    { n: "Turn of the Year (1/Day)", t: "The stag stamps once. Every creature within 30 feet that can see it regains 9 (2d8) hit points and is cured of one level of exhaustion. It does this for reasons nobody has established and to whoever happens to be standing there, including, on two recorded occasions, an occupation patrol." }
  ]
}));

c.push(...SB({
  name: "Echo of the Listening Water",
  meta: "Medium elemental, unaligned",
  ac: "12", hp: "27 (6d8)", speed: "0 ft., swim 40 ft.",
  str: 6, dex: 14, con: 10, int: 5, wis: 14, cha: 16,
  senses: "passive Perception 12", langs: "speaks only what it has heard", cr: "1 (200 XP)",
  traits: [
    { n: "Not a Monster", t: "An Echo forms where a great deal was said at a water\u2019s edge over a long time and the water has more of it than it can hold. It has no hostility and no plan. It surfaces, it repeats, and it goes." },
    { n: "In the Speaker\u2019s Voice", t: "Everything the Echo says is something a real person said at that water, in that person\u2019s own voice, exactly as they said it. It does not paraphrase, cannot lie, and has no idea what any of it means." },
    { n: "Watery Form", t: "The Echo can move through a space as narrow as one inch wide without squeezing, and has resistance to bludgeoning, piercing and slashing damage from nonmagical attacks." }
  ],
  actions: [
    { n: "Drench", t: "Melee Weapon Attack: +4 to hit, reach 5 ft., one target. Hit: 7 (2d4 + 2) cold damage. The Echo does this only if struck first, and does it apologetically, in somebody else\u2019s voice." },
    { n: "Give Back", t: "The Echo repeats up to a minute of speech spoken at its water at any point in the last thirty years, chosen by the DM. This is the reason the creature is in this book." }
  ]
}));

c.push(PS([DM("DM Only: "), { t: "the Echo is the Four Voices technique with a stat block attached, for tables that want the delivery mechanism to be a thing in the room rather than a riverbank. Use it when the party needs to hear something that happened somewhere they were not, and give them the court\u2019s version, the commons' version, the proclamation and the private grief, in four different voices, none of them lying and none of them complete. Do not let anybody fight it. If they attack it, it leaves." }]));

c.push(...SB({
  name: "Draining Engine",
  meta: "Large construct, unaligned",
  ac: "17 (plated)", hp: "115 (11d10 + 55)", speed: "0 ft. (cart-mounted)",
  str: 18, dex: 6, con: 20, int: 1, wis: 8, cha: 1,
  senses: "Blindsight 30 ft., passive Perception 9", langs: "\u2014", cr: "6 (2,300 XP)",
  traits: [
    { n: "What It Is For", t: "An Engine does not fight. An Engine drains, in a slow radius, into cut light-stone that is then carted to Vindana and sold. There are perhaps thirty of them in Elduvaine and between them they are the reason the kingdom is dying. Everything below is what an Engine does when interrupted, which is not what it was built for and is quite bad enough." },
    { n: "Drawing", t: "While the Engine is running, every creature within 60 feet that is native to Elduvaine \u2014 including fey, resident-magic creatures, and any character with an Elduvish habit-based feature \u2014 has disadvantage on saving throws. The effect is not targeted and cannot be resisted; it is simply happening to the whole field." },
    { n: "Warded Plate", t: "The Engine has resistance to bludgeoning, piercing and slashing damage from nonmagical attacks, and is immune to poison and psychic damage." }
  ],
  actions: [
    { n: "Discharge", t: "Ranged Spell Attack: +7 to hit, range 60 ft., one target. Hit: 22 (4d10) force damage." },
    { n: "Vent (Recharge 5-6)", t: "The Engine dumps what it has taken in a 30-foot radius. Each creature in the area must make a DC 16 Constitution saving throw, taking 27 (6d8) force damage on a failure, or half as much on a success. Every plant in the area dies. This is what it looks like when Elduvaine is spent, and it should be described that way." }
  ]
}));

c.push(PS([DM("DM Only: "), { t: "an Engine is an objective, not an opponent, and the encounter around it should be built that way: Legion troops defending it while a clock runs, and the Engine itself as terrain that hurts. Breaking one is a real and permanent win \u2014 there are only about thirty \u2014 and the party should be told the number so that they can feel the arithmetic. It is also the campaign\u2019s cleanest answer to a table that wants to hurt Vale before Module Eleven." }]));
c.push(H1("The Sixth Free Legion"));

c.push(P("Vale did not raise a horde. He hired an army that was already good at this, on a written contract, with a pay schedule and a reputation to protect \u2014 a hobgoblin professional core with orc, human and goblin auxiliaries and a stratum of Elduvish collaborator clerks doing the paperwork. The Sixth has been in business for a hundred and forty years, has fought for six different employers, and has never once broken a contract, which is the only reason anybody hires it."));

c.push(P("The blocks below are the Legion as an institution: drilled, disciplined, and individually unremarkable. Their strength is that they do not break, they do not chase, and they do not do anything stupid. Every one of them has explicit morale, because a professional army has explicit morale, and a party that gives the Sixth a reason to withdraw will find it withdraws in good order and takes its wounded."));

c.push(...SB({
  name: "Legionary of the Sixth",
  meta: "Medium humanoid (any race), lawful neutral",
  ac: "18 (chain mail, shield)", hp: "22 (4d8 + 4)", speed: "30 ft.",
  str: 15, dex: 12, con: 13, int: 10, wis: 11, cha: 9,
  skills: "Athletics +4, Perception +2", senses: "Darkvision 60 ft., passive Perception 12",
  langs: "Common, Goblin", cr: "1/2 (100 XP)",
  traits: [
    { n: "Formation", t: "The legionary has advantage on attack rolls against a creature if at least one other legionary is within 5 feet of that creature and is not incapacitated. This is not flanking; it is drill, and it works whether or not the table uses flanking rules." },
    { n: "Contracted", t: "The Sixth withdraws in good order when a fight has clearly been lost, when its officer orders it, or when its pay is in doubt. It does not fight to the last man for anybody, and Vale is not an exception." }
  ],
  actions: [
    { n: "Legion Pike", t: "Melee Weapon Attack: +4 to hit, reach 10 ft., one target. Hit: 7 (1d10 + 2) piercing damage." },
    { n: "Shortsword", t: "Melee Weapon Attack: +4 to hit, reach 5 ft., one target. Hit: 5 (1d6 + 2) piercing damage." },
    { n: "Light Crossbow", t: "Ranged Weapon Attack: +3 to hit, range 80/320 ft., one target. Hit: 5 (1d8 + 1) piercing damage." }
  ]
}));

c.push(...SB({
  name: "Legion Optio",
  meta: "Medium humanoid (any race), lawful neutral",
  ac: "18 (chain mail, shield)", hp: "44 (8d8 + 8)", speed: "30 ft.",
  str: 16, dex: 13, con: 13, int: 12, wis: 14, cha: 13,
  saves: "Wis +4", skills: "Athletics +5, Perception +4",
  senses: "Darkvision 60 ft., passive Perception 14",
  langs: "Common, Goblin", cr: "2 (450 XP)",
  traits: [
    { n: "Deliberately Off-Baseline", t: "An Optio is a force multiplier and is tuned as one: the armour class is high for the challenge rating and the damage is not. Do not correct this. An Optio alone is a soft CR 2; an Optio with six legionaries is the reason the encounter works." },
    { n: "Formation", t: "As the legionary." }
  ],
  actions: [
    { n: "Multiattack", t: "The optio makes two longsword attacks." },
    { n: "Longsword", t: "Melee Weapon Attack: +5 to hit, reach 5 ft., one target. Hit: 8 (1d10 + 3) slashing damage." },
    { n: "Dress the Line (Recharge 5-6)", t: "Up to six legionaries within 30 feet that can hear the optio may each immediately move up to half their speed without provoking opportunity attacks, and gain 5 temporary hit points." }
  ],
  reactions: [
    { n: "Hold", t: "When a creature within 30 feet that can hear the optio fails a saving throw against being frightened, that creature may reroll the save. The optio is not inspiring anybody. The optio is shouting a number." }
  ]
}));

c.push(...SB({
  name: "Legion Ironshank",
  meta: "Medium humanoid (any race), lawful neutral",
  ac: "19 (splint, shield)", hp: "60 (8d8 + 24)", speed: "25 ft.",
  str: 18, dex: 10, con: 16, int: 9, wis: 12, cha: 10,
  saves: "Str +6, Con +5", skills: "Athletics +6",
  senses: "Darkvision 60 ft., passive Perception 11",
  langs: "Common, Goblin", cr: "3 (700 XP)",
  traits: [
    { n: "Immovable", t: "The ironshank has advantage on saving throws and ability checks made to resist being moved, knocked prone, or shoved, and its speed cannot be reduced below 10 feet." },
    { n: "Shield Wall", t: "While within 5 feet of another ironshank, both have three-quarters cover against ranged attacks." }
  ],
  actions: [
    { n: "Multiattack", t: "The ironshank makes two maul attacks, or one maul attack and one shield bash." },
    { n: "Maul", t: "Melee Weapon Attack: +6 to hit, reach 5 ft., one target. Hit: 11 (2d6 + 4) bludgeoning damage." },
    { n: "Shield Bash", t: "Melee Weapon Attack: +6 to hit, reach 5 ft., one target. Hit: 7 (1d6 + 4) bludgeoning damage, and the target must succeed on a DC 14 Strength saving throw or be pushed 10 feet and knocked prone." }
  ],
  reactions: [
    { n: "Parry", t: "The ironshank adds 3 to its AC against one melee attack that would hit it. To do so, it must see the attacker and be wielding a melee weapon." }
  ]
}));

c.push(...SB({
  name: "Legion Battle-Mage",
  meta: "Medium humanoid (any race), lawful neutral",
  ac: "15 (bracers)", hp: "66 (12d8 + 12)", speed: "30 ft.",
  str: 9, dex: 14, con: 13, int: 17, wis: 12, cha: 11,
  saves: "Int +7, Wis +5", skills: "Arcana +7, History +7",
  senses: "passive Perception 11", langs: "Common, Goblin, Draconic", cr: "6 (2,300 XP)",
  traits: [
    { n: "Deliberately Off-Baseline", t: "Low hit points for the challenge rating, exactly as the SRD Mage is at CR 6 with 40. A battle-mage is a threat because of what it does on round one and a corpse if it is reached on round two, and the whole encounter is about which of those happens." },
    { n: "Spellcasting", t: "The battle-mage is a 9th-level spellcaster. Its spellcasting ability is Intelligence (spell save DC 15, +7 to hit with spell attacks). It has the following wizard spells prepared: Cantrips (at will): fire bolt, light, mage hand, prestidigitation; 1st level (4 slots): magic missile, shield, thunderwave; 2nd level (3 slots): misty step, scorching ray, web; 3rd level (3 slots): counterspell, fireball, fly; 4th level (3 slots): ice storm, wall of fire; 5th level (1 slot): cone of cold." },
    { n: "Hired, Not Devoted", t: "The battle-mage is on the same contract as everybody else and has the same withdrawal terms. It will surrender if the alternative is dying for an employer whose kingdom is visibly running out, and several have." }
  ],
  actions: [
    { n: "Dagger", t: "Melee or Ranged Weapon Attack: +5 to hit, reach 5 ft. or range 20/60 ft., one target. Hit: 4 (1d4 + 2) piercing damage." }
  ]
}));

c.push(H1("Named Figures"));

c.push(P("Blocks for people the party is likely to fight beside rather than against. All of them are built as people rather than as monsters, which means they sit below the DMG\u2019s curve in exactly the way the SRD\u2019s own humanoids do."));

c.push(...SB({
  name: "Xavier III of Harrowmark",
  meta: "Medium humanoid (human), lawful good",
  ac: "20 (plate, shield)", hp: "130 (that is 17d8 + 51, and he has been at this a long time)", speed: "30 ft.",
  str: 18, dex: 12, con: 16, int: 13, wis: 15, cha: 18,
  saves: "Str +8, Con +7, Wis +6", skills: "Athletics +8, Insight +6, Persuasion +8, Survival +6",
  senses: "passive Perception 12", langs: "Common, Marchspeak, Ninefold Cant", cr: "8 (3,900 XP)",
  traits: [
    { n: "Excellent Commander, Indifferent King", t: "Xavier is at his best in exactly the situation this stat block describes and at his worst in a council chamber. Play the difference. He is decisive, physically brave, personally warm to soldiers, and visibly relieved to be doing something he is good at." },
    { n: "Brave", t: "Xavier has advantage on saving throws against being frightened." },
    { n: "Chose Them Himself", t: "Xavier picked this party personally and will say so, in front of people, at moments when it is politically expensive for him to do it." }
  ],
  actions: [
    { n: "Multiattack", t: "Xavier makes three longsword attacks." },
    { n: "Longsword", t: "Melee Weapon Attack: +9 to hit, reach 5 ft., one target. Hit: 12 (1d10 + 7) slashing damage, or 13 (1d12 + 7) when used with two hands." },
    { n: "Rally (1/Day)", t: "Every friendly creature within 60 feet that can hear Xavier gains 20 temporary hit points and advantage on attack rolls until the end of their next turn." }
  ],
  reactions: [
    { n: "Parry", t: "Xavier adds 3 to his AC against one melee attack that would hit him. To do so, he must see the attacker and be wielding a melee weapon." }
  ]
}));

c.push(PS([DM("DM Only: "), { t: "Xavier earns the name Wyvernheart airborne, mid-siege, at Vindana, in a battle that is going badly. No document may use the name in narration set before that moment and the Player Guide does not hint that a name is coming. His stat block does not change when he earns it. That is the point: nothing about him is different afterwards except what he is called." }]));

c.push(...SB({
  name: "Aveline Ysolde, the Regent",
  meta: "Medium humanoid (human), lawful good",
  ac: "15 (studded leather)", hp: "66 (12d8 + 12)", speed: "30 ft.",
  str: 10, dex: 17, con: 12, int: 15, wis: 16, cha: 16,
  saves: "Dex +6, Int +5", skills: "Deception +6, Insight +6, Perception +6, Stealth +9, Survival +6",
  senses: "passive Perception 16", langs: "Common, Elduvish, Marchspeak, Writ-tongue", cr: "4 (1,100 XP)",
  traits: [
    { n: "Unremarkable", t: "Aveline has advantage on Charisma (Deception) checks made to pass as an ordinary person of any station, and on Dexterity (Stealth) checks made in an inhabited place. Three years uncaught in a kingdom where everyone else is known." },
    { n: "Cunning Action", t: "On each of her turns, Aveline can use a bonus action to take the Dash, Disengage, or Hide action." },
    { n: "Sneak Attack (1/Turn)", t: "Aveline deals an extra 14 (4d6) damage when she hits with a finesse or ranged weapon attack and has advantage, or when another enemy of the target is within 5 feet of it." },
    { n: "Will Not Be Evacuated", t: "She has refused in writing, twice, to two different coalition commanders who put it to her as a kindness. She will refuse the party as well, and the refusal is not negotiable and is not a puzzle to be solved." }
  ],
  actions: [
    { n: "Multiattack", t: "Aveline makes two attacks with her shortsword." },
    { n: "Shortsword", t: "Melee Weapon Attack: +6 to hit, reach 5 ft., one target. Hit: 6 (1d6 + 3) piercing damage." },
    { n: "Hand Crossbow", t: "Ranged Weapon Attack: +6 to hit, range 30/120 ft., one target. Hit: 6 (1d6 + 3) piercing damage." }
  ]
}));

c.push(...SB({
  name: "Huntmaster Brenna Vane",
  meta: "Medium humanoid (dwarf), neutral good",
  ac: "16 (breastplate)", hp: "58 (9d8 + 18)", speed: "25 ft.",
  str: 16, dex: 14, con: 15, int: 11, wis: 16, cha: 13,
  saves: "Con +5, Wis +6", skills: "Animal Handling +6, Athletics +6, Perception +6, Survival +8",
  senses: "Darkvision 60 ft., passive Perception 16",
  langs: "Common, Dwarvish, Marchspeak", cr: "3 (700 XP)",
  traits: [
    { n: "Two Centuries of This", t: "Brenna has advantage on attack rolls against any creature with a flying speed, and on saving throws against being frightened by one. She has been killing wyverns since before Xavier\u2019s grandfather was born and finds the whole business tedious rather than heroic." },
    { n: "Pike Drill", t: "Brenna and any friendly creature within 5 feet of her have advantage on attack rolls of opportunity, and their reach weapons deal an extra 3 damage against a creature that entered their reach this turn." }
  ],
  actions: [
    { n: "Multiattack", t: "Brenna makes two attacks with her wyvern pike." },
    { n: "Wyvern Pike", t: "Melee Weapon Attack: +6 to hit, reach 10 ft., one target. Hit: 10 (1d12 + 3) piercing damage. On a hit against a creature with a flying speed, the target must succeed on a DC 14 Strength saving throw or be pulled 10 feet toward Brenna and knocked prone." },
    { n: "Handaxe", t: "Melee or Ranged Weapon Attack: +6 to hit, reach 5 ft. or range 20/60 ft., one target. Hit: 6 (1d6 + 3) slashing damage." }
  ]
}));

c.push(...SB({
  name: "Doria Kell",
  meta: "Medium humanoid (tiefling), neutral",
  ac: "14 (fine leather, under considerably finer clothes)", hp: "45 (10d8)", speed: "30 ft.",
  str: 9, dex: 15, con: 11, int: 17, wis: 15, cha: 18,
  saves: "Int +6, Cha +7", skills: "Deception +7, History +6, Insight +8, Investigation +9, Persuasion +10",
  senses: "Darkvision 60 ft., passive Perception 12",
  langs: "Common, Infernal, Writ-tongue, Elduvish, Goblin", cr: "2 (450 XP)",
  traits: [
    { n: "Not a Combatant", t: "Doria has a stat block because parties are unpredictable, not because this encounter is intended. She will not fight. She will leave, and the house will note it, and every price the party is ever quoted afterward will reflect the note." },
    { n: "The Letter of It", t: "Doria has advantage on any check made to draft, read, or find the flaw in a written agreement, and she cannot be deceived about the contents of one she has read." },
    { n: "Genuinely Takes No Side", t: "This is not a pose and is not a secret allegiance waiting to be uncovered. Norvatch priced a situation it did not create and has traded inside that price for three years. She will say so to anybody who asks and will not be embarrassed by it." }
  ],
  actions: [
    { n: "Dagger", t: "Melee or Ranged Weapon Attack: +5 to hit, reach 5 ft. or range 20/60 ft., one target. Hit: 4 (1d4 + 2) piercing damage." },
    { n: "Innate Spellcasting", t: "Doria can cast thaumaturgy at will, and hellish rebuke (2nd level) and darkness once each per day. Charisma is her spellcasting ability (save DC 15). She regards all three as vulgar and uses them only when leaving." }
  ]
}));

c.push(H2("Maelis Ysolde"));

c.push(P("There is no stat block for the Sovereign of Elduvaine, and the omission is deliberate."));

c.push(P("She has not stood unaided since the second winter of the occupation. She cannot fight, cannot flee, cannot be usefully protected, and cannot be carried out of Caer Ysolde without a question nobody can answer about what leaving would do to the habits bound to her. Anything a party does in a room with Maelis Ysolde is a conversation, and a DM who wants numbers for her should use her Intelligence, which is 20, and her Wisdom, which is 20, and remember that she has had three years of confinement with nothing whatsoever to do except think about the man holding her."));

c.push(PS([DM("DM Only: "), { t: "if a table insists on a mechanical handle: she has passive Insight 20, she cannot be lied to about Elduvaine, and she knows more about Maedoc Vale than anyone alive including Vale. That is the whole of it. Do not give her hit points. The moment she has hit points, some table somewhere will find out what happens when they run out, and this campaign is better without that scene in it." }]));

c.push(H1("Creatures by Region"));

c.push(P("SRD creatures used by name in the modules and the gazetteer, without a reprinted block, because they are unmodified and a DM can look them up in seconds."));

c.push(table(
  ["Where", "What lives there"],
  [24, 76],
  [
    ["Harrowmark", "Wyvern (CR 6), bandit (CR 1/8), wolf (CR 1/4), hill giant (CR 5), guard (CR 1/8)"],
    ["The sea road", "Merrow (CR 2), giant octopus (CR 1), sahuagin (CR 1/2)"],
    ["The mountain road", "Kobold (CR 1/8), ogre (CR 2), griffon (CR 2), troll (CR 5), giant eagle (CR 1)"],
    ["The Vaunt", "Sprite (CR 1/4), ghoul (CR 1), scout (CR 1/2), giant crab (CR 1/8)"],
    ["The Braid", "Will-o'-wisp (CR 2), spy (CR 1), veteran (CR 3), mage (CR 6)"],
    ["The Orchard Marches", "Dryad (CR 1), winter wolf (CR 3), troll (CR 5), satyr (CR 1/2), awakened tree (CR 2)"],
    ["The Standing Marches", "Gargoyle (CR 2), earth elemental (CR 5), stone giant (CR 7)"],
    ["Vindana\u2019s undercity", "Kobold (CR 1/8), giant rat (CR 1/8), grey ooze (CR 1/2), otyugh (CR 5)"]
  ], { full: true }
));

c.push(P("Coverage here is SRD-only. There is no Volo\u2019s and no Mordenkainen\u2019s in this book, and anything a DM imports from those is unvalidated against the encounter maths the modules were built with."));

c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300, after: 100 }, children: [new TextRun({ text: "~", size: 24 })] }));
c.push(PS([{ t: "\u201CIt set me down in the road. Twice. Very carefully, like a man moving a chair. And then it went back and stood by the stone, and it did not look at me again, and I have thought about it every day since.\u201D", i: true }], { alignment: AlignmentType.CENTER }));
c.push(PS([{ t: "\u2014 a carter of Lisswater, deposition to the Magistrate", i: true }], { alignment: AlignmentType.CENTER }));
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
  fs.writeFileSync(stagePath("KC_Bestiary.docx"), buf);
  console.log("Written.");
});
