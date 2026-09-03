// KC_Module04_TheCoalition.js -- Session Module Four: The Coalition.
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
// GATED CONTENT NOTICE: two things this module needs are still open in
// CLAUDE.md and are deliberately NOT invented here. First, what Oksitan and
// Auberitz each actually want beyond the Promise -- this module gives named,
// individual officers personal opinions and friction, never a stated national
// policy for either realm. Second, the royal family\u2019s names -- the captive
// rescued in Scene 4 is written as a full, usable NPC with no proper name,
// referred to as "the Ward" throughout, with a DM-only note on where to drop
// a name in once CLAUDE.md\u2019s open item is signed off. Do not fill either gap
// in without sign-off, including in future revisions of this file.

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
  children: [new TextRun({ text: "The Coalition", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module Four", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("The party reaches the coalition\u2019s forward camp above Caerwyn and meets, for the first time at real length, the allies marching under Xavier\u2019s call: Oksitan and Auberitz troops in numbers, and \u2014 approached privately, rather than in camp \u2014 a representative of unaligned Norvatch with something to sell. By the module\u2019s end, the party has learned exactly what the Promise is worth to the people who marched for it, and has the chance to rescue the first of Elduvaine\u2019s captive royal house: a young member of it, held at a lightly-guarded hall not far from where the coalition now stands. Core scenes run three and a half to four hours; Optional Content fills out the rest of a five-hour session and can be cut cleanly if the table is short on time."));

c.push(H2("Allies, Not Friends"));

c.push(P("Play Oksitan and Auberitz as genuinely allied and genuinely self-interested at once \u2014 neither cynical nor noble, the way real coalitions actually are. No officer from either realm should ever speak for their crown\u2019s official policy in this module; what each realm actually wants beyond the Promise is not yet settled, and this module deliberately gives individual named people opinions, ambitions, and doubts instead. That is more useful at the table besides: a party remembers a captain who said something honest over a cookfire long after they have forgotten a stated national position."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 19, 51],
  [
    ["1. The Muster Camp", "30\u201340 min", "First real look at Oksitan and Auberitz, in numbers."],
    ["2. What the Promise Is Worth", "45\u201360 min", "The coalition\u2019s tension, in individual voices, not policy."],
    ["3. Doria Kell\u2019s Offer", "30\u201345 min", "Norvatch approaches privately. Leverage, not alliance."],
    ["4. Sennoch Hall", "75\u201390 min", "The first rescue. DC table and stat block below."],
    ["Optional Content", "30\u201345 min", "Run if the table has time; cut cleanly if not."]
  ]
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("The coalition is exactly as fragile as a coalition built on the Promise should be. Oksitan and Auberitz both marched for their own reasons, which this module does not state, and both are already privately uneasy about how much of Elduvaine\u2019s resident magic will actually be theirs to hold once Vale is gone \u2014 a doubt none of their officers would say aloud to a superior, and several will say plainly to the party, who are outsiders to the coalition\u2019s own politics and therefore safe to be honest with."));

c.push(PS([DM("DM Only: "), { t: "the Ward, held at Sennoch Hall, is a real, specific person with no proper name yet \u2014 this campaign\u2019s royal family is deliberately unnamed. Play them fully: young, sharp, has spent the occupation being quietly furious rather than afraid, and deeply unimpressed by the idea that their rescue is a political opportunity for whoever performs it. Do not invent a name for them at the table if it can be avoided; if a player asks directly, have the Ward deflect (\u201Cnames are for people with the leisure to be found by them\u201D) rather than supply one. Once the royal family\u2019s names are settled, insert the name here and everywhere else the Ward appears." }]));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: The Muster Camp"));

c.push(P("The camp above Caerwyn is the first time the party has seen the coalition as a whole rather than as a rumor at the edge of the Harrowmark muster: Oksitan tents in orderly rows, their banners a color the party has not seen before; Auberitz\u2019s heavier cavalry picketed apart, grooms working in a language half the camp does not speak; Harrowmark\u2019s own levies looking, for the first time, like the smallest contingent present."));

c.push(BOX("The camp smells of a dozen different cooking fires and sounds like three armies that have not yet decided how loudly they are allowed to dislike each other. An Oksitan sergeant argues good-naturedly with a Harrowmark quartermaster over a supply wagon; an Auberitz officer watches both of them with the particular patience of someone counting days until the argument becomes his problem."));

c.push(P("Let the party walk the camp and meet it at their own pace. If Module 2A was played, Sera Vosk (see her NPC profile in Module 2A) is a natural, familiar face to have turn up here, now properly attached to the Auberitz contingent\u2019s logistics; if Module 2B was played, use a new, unnamed Auberitz quartermaster instead, or introduce Garrick Hollow\u2019s opposite number for texture. Either way, someone in camp should treat the party as already known \u2014 the story of whichever Module Two variant they played has traveled ahead of them."));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: What the Promise Is Worth"));

c.push(P("Around a cookfire, a council table, or wherever the party finds themselves drawn into camp conversation, individual officers from Oksitan and Auberitz will talk \u2014 carefully at first, then plainly, if the party listens rather than interrogates \u2014 about what they actually hope the crusade gets them."));

c.push(BOX("\u201CMy grandmother always said Oksitan never got the better end of a bargain it didn\u2019t write itself,\u201D says one officer, watching the fire rather than the party. \u201CI don\u2019t know what my king actually wants out of this. I know what I want. A place where my sons don\u2019t have to be soldiers. I don\u2019t much care whose crown is over it.\u201D"));

c.push(P("Use two or three named minor officers (invent freely \u2014 this campaign\u2019s Creative Latitude covers minor NPCs without limit) to carry this scene, and let each want something small, personal, and plausible rather than grand: land for a family, a debt paid, simple relief at being somewhere the war has not yet reached their own home. None of them should articulate a coherent national policy, because none of them has the standing to state one, and neither does this module."));

c.push(PS([DM("DM Only: "), { t: "if a player asks a pointed question about what Oksitan or Auberitz officially wants, the honest in-fiction answer is that nobody at this table \u2014 including the officers themselves \u2014 actually knows, or agrees. Let that stand as an answer. It is true, and it is more interesting than a policy statement would be." }]));

// ---------------------------------------------------------------- Scene 3
c.push(H3("What the Camp Looks Like"));

c.push(P("Three armies, and they do not look alike. Harrowmark\u2019s lines are humans and dwarves and orcs in mixed companies, unglamorous and quiet, cooking badly. Oksitan\u2019s quarter is horse country transplanted \u2014 human riders and a scattering of dragonborn house-knights whose armour is better than anybody else\u2019s and who are perfectly aware of it. Auberitz has pitched furthest from the latrines on purpose and got it right: gnome engineers arguing over a half-assembled hoist, halfling quartermasters running a supply system that visibly works, and a general air of a duchy that considers this whole war a logistics problem being mishandled by people who enjoy shouting."));

c.push(P("Let the party walk through all three. The point is not information; it is that the coalition is a real coalition, with three ideas of how a war should be run and three cuisines, and that the party is about to spend months inside it."));

c.push(H2("Scene 3: Doria Kell\u2019s Offer"));

c.push(P("Away from camp \u2014 at a well, on a supply run, wherever the party can be approached without an audience \u2014 a woman in plain, well-made traveling clothes falls into step beside them, unhurried and entirely unbothered by the coalition\u2019s presence."));

c.push(BOX("\u201CDoria Kell, factor of Norvatch,\u201D she says, as though the introduction were a formality rather than an offer. \u201CI don\u2019t march with your coalition and I don\u2019t intend to. I trade with whoever\u2019s buying, on either side of this war, and just now I\u2019m buying goodwill. I know where one of Elduvaine\u2019s royal wards is being kept. I\u2019ll tell you, no charge, no favor owed \u2014 because it costs me nothing and it might cost the man holding her something later, and I find that arrangement agreeable.\u201D"));

c.push(P("Doria is not an ally and does not pretend to be one. She trades information the way she trades goods \u2014 because it is profitable, not because she is choosing a side \u2014 and this scene should make that legible without making her untrustworthy; what she says checks out exactly as she describes it. If the party wants to press her for more (payment for better information, a standing arrangement, anything beyond this one free lead), she will discuss terms plainly, and a DM can use this as a recurring thread: Norvatch as a source the party can return to, always for a price, never as a friend."));

c.push(H3("What She Is, If Anybody Asks"));

c.push(P("A party that asks Doria directly what Norvatch actually does in this war gets a straight answer, delivered without a flicker of embarrassment, because she does not consider it embarrassing. Norvatch buys. It has been buying out of Elduvaine for three years, under a written contract with the occupation, and the goods are the kingdom: light-stone by the cartload, Archive material by the crate, worked timber out of woods that have stopped holding their season. Vale extracts. Norvatch moves it. She will say so in as many words if the question is put plainly."));

c.push(P("She will also point out, if pressed and sometimes if not, that Norvatch did not open the wards, would have advised against it on commercial grounds, and has broken no agreement with anybody. All of which is true, and none of which is a defence, and she knows both of those things and says it anyway."));

c.push(PS([DM("DM Only: "), { t: "this is the campaign\u2019s quietest gut-punch and it should be played completely flat. There is no villainy in the scene. There is a professional explaining her trade to people who had not thought about where a drained kingdom physically goes, and the answer is that somebody has been buying it, promptly, for three years, and filing the paperwork correctly. Do not let any NPC editorialise. Let the party do it." }]));

c.push(PS([DM("DM Only: "), { t: "she does not offer the ledgers here and should not. What this scene plants is that they exist \u2014 three years of purchase records, the only complete account anywhere of how much of Elduvaine has already left it. A sharp party will work out on their own what that would be worth to a coalition that does not know how long it has. When they come back for them, the price is in Module Ten." }]));

// ---------------------------------------------------------------- Scene 4
c.push(H2("Scene 4: Sennoch Hall"));

c.push(P("Doria\u2019s information is accurate: a half-day\u2019s ride from camp, Sennoch Hall \u2014 a minor Elduvish manor requisitioned early in the occupation \u2014 holds one member of the royal house, guarded lightly precisely because nobody expected a coalition to reach this far this fast. This is the campaign\u2019s first rescue, and it is meant to be winnable."));

c.push(BOX("The Hall is quiet in the particular way of a place trying not to be noticed. Two guards at the gate, bored rather than alert; a third somewhere inside, audible before visible. Through a high window, briefly, a figure crosses a lit room without hurrying \u2014 someone who has stopped expecting rescue and has not stopped watching for it anyway."));

c.push(H3("Running the Scene"));

c.push(P("This can be played as stealth, as a direct assault, or as a mix of both \u2014 the guard presence is genuinely light, and the module should not force a fight. A DC 13 Stealth (Dexterity) check gets past the gate guards unseen; a DC 13 Investigation or Perception check finds the Ward\u2019s room without alerting the interior guard. The three guards are hobgoblin legionaries \u2014 professionals on a wage, bored, and a long way down anybody\u2019s list of priorities. If it comes to violence, use the Occupation Guard stat block (as in Module 3) for all three \u2014 none of them are heroes, and all three will surrender or flee once it is clear the fight is lost."));

c.push(P("The Ward (see NPC Profiles) is sharp, unharmed, and immediately useful rather than a burden once freed \u2014 she knows the Hall\u2019s layout, the guards\u2019 habits, and enough about the wider occupation\u2019s posture in this region to be a genuine asset on the ride back to camp, not just a rescued NPC to be delivered and forgotten."));

c.push(H3("Stat Block"));

c.push(...SB({
  name: "Occupation Guard",
  meta: "Medium humanoid (hobgoblin, or any race), any alignment \u2014 SRD Guard, unmodified",
  ac: "16 (chain shirt, shield)",
  hp: "11 (2d8 + 2)",
  speed: "30 ft.",
  str: 13, dex: 12, con: 12, int: 10, wis: 11, cha: 10,
  skills: "Perception +2",
  senses: "passive Perception 10",
  langs: "Common",
  cr: "1/8 (25 XP)",
  actions: [
    { n: "Spear", t: "Melee or Ranged Weapon Attack: +3 to hit, reach 5 ft. or range 20/60 ft., one target. Hit: 4 (1d6 + 1) piercing damage, or 5 (1d8 + 1) piercing damage if used with two hands to make a melee attack." }
  ]
}));

// -------------------------------------------------------------- Skill DCs
c.push(H2("Tiered Skill DCs"));

c.push(P("Easy 10, Moderate 13, Hard 16, matching the tiers used throughout this campaign."));

c.push(table(
  ["Task", "Skill", "DC", "Tier"],
  [44, 26, 10, 20],
  [
    ["Get camp officers talking honestly rather than diplomatically", "Persuasion / Insight", "13", "Moderate"],
    ["Verify Doria Kell\u2019s information is genuine before acting on it", "Insight / Investigation", "13", "Moderate"],
    ["Slip past Sennoch Hall\u2019s gate guards unseen", "Stealth", "13", "Moderate"],
    ["Find the Ward\u2019s room without alerting the interior guard", "Investigation / Perception", "13", "Moderate"],
    ["Negotiate an ongoing arrangement with Doria Kell", "Persuasion", "16", "Hard"]
  ]
));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("Puzzles and Set Pieces"));

c.push(P("All three of these are the Sennoch Hall rescue, given a map, a way in and a way out. They replace describing the Hall freehand and add no time to the module; the keyed table is there so the DM is not inventing the building while the party is inside it."));

c.push(H2("Sennoch Hall, Keyed"));

c.push(P("The Hall is the campaign\u2019s first proper dungeon and it is a country house, which means the map is the puzzle. It was never built to hold anybody and the occupation has adapted it badly: a garrison of eighteen in a building with eleven exterior doors, four of which are glazed."));

c.push(table(
  ["Area", "What is there", "What a party can use"],
  [22, 40, 38],
  [
    ["1. The ornamental water", "A shallow formal canal round three sides. Elduvish. Listening.", "It has heard three years of guard-changes. See the puzzle below."],
    ["2. The south front", "Four tall glazed doors onto the terrace. Shuttered at dusk, latched from inside.", "Silent entry, DC 14 Sleight of Hand. Loud entry, no check at all."],
    ["3. The kitchen court", "Staff who are not permitted to leave. Six of them, none guarded.", "Every one of them knows the rotation and four will say so."],
    ["4. The long library", "Ninian\u2019s days are spent here. One guard at the door, bored.", "The party\u2019s likely first contact. Also: eleven years of estate maps."],
    ["5. The garrison range", "Converted stables. Twelve off-duty legionaries, unarmoured.", "Taking this room ends the fight before it starts, and is loud."],
    ["6. The commander\u2019s study", "Serjeant Hoth, and three years of correct paperwork.", "He would like somebody to notice he has been decent. See below."],
    ["7. The north tower", "Ninian\u2019s rooms. One door, one window, forty feet up.", "The window is the extraction route and it needs rope."],
    ["8. The gate range", "Six on watch, alert, with a bell.", "The bell is the whole problem. Cutting its rope is a 20-ft. climb."]
  ]
));

c.push(H2("The Puzzle: What the Water Heard"));

c.push(P("The ornamental canal at Sennoch Hall is Elduvish water, and Elduvish water listens. Three years of guards have walked its edge four times a day saying the same twelve words to each other, and it has all of it."));

c.push(P("A character who spends ten minutes at the water\u2019s edge \u2014 with the listening water spell, the Listening-Trained feat, a flask, or simply by knowing the trick and asking, which any Elduvish NPC can teach in one sentence \u2014 gets this, in four different voices, none of them lying:"));

c.push(BOX("\u201CSecond watch, all quiet.\u201D \u2014 \u201CSecond watch relieved, the tower door is fast.\u201D \u2014 \u201CGate to tower, gate to tower, and the bell rope wants seeing to.\u201D \u2014 \u201CThird watch, and if that dog is out again I am shooting it.\u201D"));

c.push(P("Four true facts fall out of it, and a party should have to assemble them rather than be handed them. The watch changes three times a night, not two. The tower door is checked at every change, so the extraction window is the twenty minutes between. The gate range and the tower speak to each other by shouted call-and-response, so an unanswered call is an alarm. And the bell rope has wanted seeing to for some time, which means it is frayed, which means it can be cut from below with a thrown weapon at DC 16."));

c.push(PS([DM("DM Only: "), { t: "this is the first time the party gets to use Elduvaine\u2019s magic as a tool rather than admire it as scenery, and it should land that way. Do not summarise the four voices \u2014 perform them, in four different registers, and let the players do the assembling. If nobody thinks to try the water, Ninian mentions it afterward, drily, and the party feels the loss without being punished for it." }]));

c.push(H2("Set Piece: Serjeant Hoth\u2019s Surrender"));

c.push(P("The garrison commander at Sennoch Hall is a hobgoblin of the Sixth who has held a royal hostage in comfortable confinement for three years, has not once permitted a hand to be laid on her, has filed a monthly report every one of those months, and is entirely aware of how this ends for him."));

c.push(BOX("\u201CI have eighteen. You have got past the water and the bell and I have not heard from the tower in some time, so let us both save an hour. My terms are these: my people walk out with their arms and their wounded, and I will give you the last three years of the file, which is complete, and which is the only account anybody has of who came here and what was asked. You will want it. I would like it written down that the Hall was kept correctly.\u201D"));

c.push(P("He means every word and there is no trick in it. The file is real and is genuinely valuable \u2014 it names the collaborator clerks in the Braid, it establishes what the occupation asked the royal house and when, and it is the campaign\u2019s first hard evidence that Emrys Ysolde has been talking to Vale."));

c.push(P("A party that fights him instead wins. Eighteen legionaries and an optio against five characters at 5th level is a real fight and not a close one if the party has taken the range or the gate. They also lose the file, which Hoth burns, correctly, per his standing orders, in the four minutes it takes them to reach the study."));

c.push(H1("NPC Profiles"));

c.push(H2("Doria Kell"));
c.push(P("A tiefling of middle years and Norvatch to the bone, dressed for travel rather than for a court, with the unhurried manner of somebody who has never once been the most anxious person in a negotiation. Speech: unhurried, transactional, entirely without malice \u2014 she genuinely does not care who wins this war, only that she is positioned well when it ends. Not a spy in any dramatic sense; simply a trader whose stock is information."));
c.push(P("Open thread: Doria is a standing resource for the rest of the campaign \u2014 a DM can bring her back whenever the party needs information no ally would risk giving them, always for a price, and her price is never coin alone."));

c.push(H2("The Ward"));
c.push(P("A young half-elf of Elduvaine\u2019s royal house \u2014 which is, like most old Elduvish families, a mixed one, and does not think of itself as anything else \u2014 with no proper name yet assigned (see the DM-Only note in What Is Actually Happening). Speech: quick, dry, allergic to being treated as fragile \u2014 three years of captivity have made her precise rather than broken. Speaks of Vale, when she speaks of him at all, with contempt rather than fear."));
c.push(P("Open thread: once freed, the Ward is a genuine resource and a genuine person with her own opinions about what Elduvaine should become after the war \u2014 opinions that may or may not match the rest of her family\u2019s, per the royal house\u2019s canon division on that question. A DM can bring her back throughout the rest of the campaign as a recurring ally with real standing to speak for what the liberated actually want."));

// --------------------------------------------------------------- Optional
c.push(H1("Optional Content"));

c.push(H2("The Quartermasters\u2019 War"));
c.push(P("If the table wants more camp texture before Sennoch Hall, a supply dispute between the coalition\u2019s quartermasters \u2014 Harrowmark, Oksitan, and Auberitz all convinced their own contingent is shorted \u2014 gives the party a chance to mediate or take a side, with no stakes beyond camp goodwill and a good scene."));

c.push(H2("What Doria Doesn\u2019t Say"));
c.push(P("A perceptive party (DC 16 Insight) can notice that Doria Kell\u2019s free information came a little too easily \u2014 not a trap, but a deliberate investment. If asked directly, she admits it plainly: goodwill now is cheaper than goodwill bought later, once the party is worth more to her. This does not need to go anywhere; it is simply true, and worth letting the party notice on their own."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("How Sennoch Hall was resolved.", "Stealth, violence, or a mix \u2014 track which, and whether the guards were killed or allowed to flee. A clean, quiet rescue keeps the occupation from learning quickly that Sennoch Hall was compromised; a loud one does not."));
c.push(BUL("The party\u2019s standing with Doria Kell.", "Whether they took her offer at face value, pressed for more, or treated her with suspicion \u2014 this sets the tone for every future Norvatch interaction, and is worth tracking as its own thread rather than folding into the road-choice or Calanthe entries."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BUL("The Ward herself.", "Not loot in the ordinary sense, but the module\u2019s actual reward: a freed, capable ally with real knowledge of the region and a real stake in how this war ends."));
c.push(BUL("Sennoch Hall\u2019s stores.", "Modest \u2014 the occupation did not garrison this Hall richly. Enough coin and supply to be worth collecting, not enough to be a windfall at 5th level."));
c.push(BUL("A sealed letter.", "Found among the Hall\u2019s papers, addressed to an occupation official the party has not yet met and has not yet had reason to trust or distrust \u2014 a DM\u2019s hook for later, not something that needs to pay off in this module.", { keepNext: true }));

// -------------------------------------------------------------- Refrain
c.push(H1("The Refrain"));

c.push(VERSE([
  "By thought, and by word, and by deed,",
  "the king\u2019s own chosen kept their creed.",
  "Far from home, where the quiet land lay,",
  "they held the line, and would not stray."
]));

const doc = new Document({
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 260, hanging: 260 } } } }] }] },
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
  fs.writeFileSync(stagePath("KC_Module04_TheCoalition.docx"), buf);
  console.log("Written.");
});
