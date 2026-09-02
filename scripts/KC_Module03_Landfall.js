// KC_Module03_Landfall.js -- Session Module Three: Landfall.
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
// This is the campaign's tonal pivot: the party's first landfall in Elduvaine.
// CLAUDE.md's own instruction for Running Elduvaine is followed directly here
// -- "describe the habit before you describe the damage" -- so this module
// shows a working Standing Light before it shows the Dead Mile, and closes
// with the Four Voices device (also specified in CLAUDE.md) delivering the
// last three years in four registers: a court account, a commons account, the
// occupation's own proclamation, and a private grief. None of the four lies;
// none is complete.
//
// Encounter design note: this module does not require combat. The checkpoint
// in Scene 3 can turn violent if the party forces it, and a light stat block
// (the SRD Guard, unmodified) is provided for that case, but the scene is
// built to reward patience and reading the situation over force -- consistent
// with "peril is the point, but so is agency."

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

const { Table, TableRow, TableCell, WidthType, ShadingType } = require('docx');
const cell = (text, opts = {}) => new TableCell({ width: { size: opts.w || 20, type: WidthType.PERCENTAGE }, shading: opts.head ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, margins: { top: 60, bottom: 60, left: 110, right: 110 }, children: [new Paragraph({ spacing: { after: 0 }, indent: { firstLine: 0 }, children: [new TextRun({ text, bold: !!opts.head, size: 18 })] })] });
const row = (cells, opts = {}) => new TableRow({ children: cells, cantSplit: true, ...opts });
const FULLWIDTH = "KCFullWidth";
const table = (headers, widths, rows, opts = {}) => new Table({ ...(opts.full ? { style: FULLWIDTH } : {}), width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ row(headers.map((h, i) => cell(h, { head: true, w: widths[i] })), { tableHeader: true }), ...rows.map(r => row(r.map((v, i) => cell(v, { w: widths[i] })))) ] });

const mod = (v) => { const m = Math.floor((v - 10) / 2); return (m >= 0 ? "+" : "\u2212") + Math.abs(m); };
const abCell = (text, bold) => new TableCell({ width: { size: 16.6, type: WidthType.PERCENTAGE }, shading: bold ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40, before: 40 }, indent: { firstLine: 0 }, keepNext: !!bold, children: [new TextRun({ text, bold: !!bold, size: 20 })] })] });
const SB = (d) => { const out = []; out.push(new Paragraph({ spacing: { before: 240, after: 40 }, children: [new TextRun({ text: d.name, bold: true, size: 26, color: "5B1F1F" })] })); out.push(PS([{ t: d.meta, i: true }], { spacing: { after: 120 } })); out.push(B("Armor Class:", d.ac)); out.push(B("Hit Points:", d.hp)); out.push(B("Speed:", d.speed)); out.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ new TableRow({ cantSplit: true, tableHeader: true, children: ["STR","DEX","CON","INT","WIS","CHA"].map(h => abCell(h, true)) }), new TableRow({ cantSplit: true, children: [d.str,d.dex,d.con,d.int,d.wis,d.cha].map(v => abCell(v + " (" + mod(v) + ")")) }) ] })); out.push(P("", { spacing: { after: 60 } })); if (d.saves) out.push(B("Saving Throws:", d.saves)); if (d.skills) out.push(B("Skills:", d.skills)); if (d.senses) out.push(B("Senses:", d.senses)); if (d.langs) out.push(B("Languages:", d.langs)); out.push(B("Challenge:", d.cr)); (d.traits||[]).forEach(t => out.push(PS([{ t: t.n + ". ", b: true, i: true }, { t: t.t }]))); if (d.actions && d.actions.length) { out.push(PS([{ t: "ACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.actions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } if (d.reactions && d.reactions.length) { out.push(PS([{ t: "REACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.reactions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } return out; };


// ---------- content ----------
const c = [];

c.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: "Landfall", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module Three", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("Whichever road brought them here, the party makes landfall in Elduvaine for the first time \u2014 at Caerwyn, a small coastal town the coalition has chosen as a forward landing, still far enough from Caer Ysolde and Vindana that the occupation\u2019s grip on it is thin rather than absolute. This module\u2019s job is tonal before it is anything else: it is the party\u2019s first sight of a land that is actually magical, the first mundane machinery of Vale\u2019s occupation, and the first time the Listening Water hands them the last three years directly. Core scenes run three and a half to four hours; Optional Content fills out the rest of a five-hour session and can be cut cleanly if the table is short on time."));

c.push(H2("Wonder Before Damage"));

c.push(P("Run Scene 1 before Scene 2, in that order, without exception. The sourcebook\u2019s own instruction for this campaign is to describe a habit working before describing one failing, and this module is built entirely around that discipline: the party should feel Elduvaine as marvellous for one full scene before they feel it as wounded. If a table\u2019s pacing forces a cut somewhere in this module, cut from the checkpoint or Optional Content \u2014 never from Scene 1."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 15, 55],
  [
    ["1. First Light of Elduvaine", "30\u201340 min", "A working habit. Play this for wonder, at length."],
    ["2. The Dead Mile", "20\u201330 min", "A failing habit. Short, quiet, and the module\u2019s gut-punch."],
    ["3. The Caerwyn Checkpoint", "45\u201375 min", "The occupation, in bureaucratic miniature. Rarely needs to be a fight."],
    ["4. The Riverbank", "60\u201390 min", "The Four Voices. See below \u2014 this is the module\u2019s centerpiece."],
    ["Optional Content", "30\u201345 min", "Run if the table has time; cut cleanly if not."]
  ],
  { full: true }
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("Caerwyn is halflings and humans in about equal measure, with a gnome family running the one stone-yard and an elderly elf who keeps the orchard above the town and has kept it, in the same four days of spring, since before anybody now living was born. It has survived the occupation better than most of Elduvaine, mostly through insignificance \u2014 it is small, it is far from the capital, and its habits are correspondingly slower to drain. Its Standing Light still glows most nights, faintly. That will not last, and nothing in this module should promise the party that it will; the point of showing them a working habit here is to let them feel what is actually being lost everywhere else, not to suggest Caerwyn is safe."));

c.push(PS([DM("DM Only: "), { t: "Vale himself never appears in this module and should not be referenced by name at the checkpoint or by any occupation soldier the party meets \u2014 the occupation\u2019s day-to-day machinery runs on permits and levies issued by a distant authority, not on any personal presence, and that facelessness is deliberate. The Four Voices in Scene 4 is where the party first hears his name spoken by people who lived through what he did, and it should land as the first time he becomes real to them." }]));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: First Light of Elduvaine"));

c.push(P("The final approach to Caerwyn goes faster than the coalition\u2019s own maps predicted, and nobody aboard can quite say why \u2014 a current that ran the right way, a headland that seemed to arrive sooner than the charts allowed. It is the Willing Road\u2019s cousin at work, though nobody in the party has cause to name it yet."));

c.push(BOX("Caerwyn rises out of the dusk already lit \u2014 not by torches, not by hearth-fires banked for evening, but from within: pale stone walls holding a soft, steady glow, as though the town had swallowed the last of the daylight and was giving it back slowly, all night, for anyone who needed it. A child on the harbor wall waves at the ship without any particular urgency, the way you wave at something ordinary."));

c.push(P("Give this scene real time. Let the party walk Caerwyn\u2019s streets before anything else happens: stone that is warm to the touch after dark, a well that a local swears will repeat back a lullaby sung to it forty years ago if you ask it kindly, a baker who insists \u2014 matter-of-factly, not proudly \u2014 that the ovens here have never needed lighting twice. None of this is remarkable to the people living in it. That is exactly what should make it remarkable to the party."));

c.push(PS([DM("DM Only: "), { t: "resist any urge to explain the mechanism. Elduvaine\u2019s magic is a habit, not a spell, and no NPC in this scene should describe it in rules language. A fisherman does not know why the light holds; he only knows it does, the way a Harrowmark hunter knows a wyvern comes down off the crags in a hard winter \u2014 as a fact about the world, not a system to be understood." }]));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: The Dead Mile"));

c.push(P("The road inland from Caerwyn runs true for a while, and then, without any visible seam, it stops being true. There is no marker. There is no wall. A local guide, if the party has one, simply slows and says the name \u2014 the Dead Mile \u2014 the way a person names a grave they walk past often."));

c.push(BOX("It looks like an ordinary road. That is the entire horror of it. No stone glows. No well answers. The party\u2019s own breath sounds too loud, and it takes a long, uncomfortable moment for anyone to place why: nothing here is listening. For one mile, Elduvaine is exactly as ordinary as Harrowmark, and every local who crosses it does the same thing without discussing it \u2014 walks a little faster, and does not talk until the far side."));

c.push(P("Play this scene short and quiet. There is nothing to fight and nothing to solve; its only job is to make the drainage real and specific after Scene 1 made the habits real and specific. If a spellcaster wants to detect magic or otherwise investigate the Mile directly, let them find exactly nothing \u2014 not a residue, not a trace, not evidence of a spell having been cast or removed. The Mile is not damaged the way a wound is damaged. It is simply, completely absent, and that absence is the point."));

// ---------------------------------------------------------------- Scene 3
c.push(H2("Scene 3: The Caerwyn Checkpoint"));

c.push(P("Past the Dead Mile, the road reaches a checkpoint \u2014 a barrier, a brazier, and a folding table with a ledger, staffed by two or three occupation soldiers and one Elduvish clerk who does the actual talking. This is Vale\u2019s war in its least dramatic and most revealing form: permits, a grain levy receipt, a question about where the party is bound and why."));

c.push(BOX("The clerk \u2014 a woman in a coat too well-made for the checkpoint she is standing at \u2014 does not look up immediately. \u201CPapers, or a reason,\u201D she says, pen already uncapped. \u201CI don\u2019t enjoy this part any more than you will. Let\u2019s get through it quickly.\u201D"));

c.push(P("This is Wyn Alder (see NPC Profiles). She has no papers to offer that would satisfy a party arriving in warships, and both she and the party know it; her real job in this scene is to demonstrate that the occupation is staffed by ordinary, frightened, complicit people rather than villains, and to give the party a first, small, human-scale decision about how they intend to fight this war \u2014 through the people caught inside it, or around them."));

c.push(H3("Running the Scene"));

c.push(P("A checkpoint this small cannot stop a coalition landing party and everyone present knows it. Wyn will not fight, will not raise an alarm she cannot win, and will answer honest questions honestly if approached without threats \u2014 she has no loyalty to Vale, only a family in Caerwyn she cannot afford to have marked as uncooperative. A Charisma (Persuasion) or Wisdom (Insight) check against DC 13 gets her talking plainly about levy quotas, patrol patterns, and how thin the occupation\u2019s actual presence is this far from the capital; nothing requires a check if the party simply treats her like a person rather than an obstacle."));

c.push(P("If the party threatens or attacks the checkpoint, the soldiers (use the Occupation Guard stat block) fight briefly and then flee or surrender \u2014 they are a levy garrison, not a real defense, and know it. Wyn does not fight under any circumstances and should not be placed in danger by this module regardless of how the scene resolves; if combat breaks out, she ducks for cover and stays there. A violent resolution here has a real cost the table should feel: word of an attacked checkpoint reaches Caer Ysolde days sooner than it otherwise would have, which a DM can pay off later as tightened security somewhere that mattered."));

c.push(H3("Stat Block (If Needed)"));

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
c.push(H3("The First Strange Thing"));

c.push(P("Somewhere in the first hour ashore, before any of the plot happens, let the party see one thing that does not work the way the world works. The strong recommendation is the smallest available: a sprite in the orchard hedge, four inches of extremely rude opinion with wings, entirely uninterested in the war and deeply interested in whether anyone has brought food. The elf who keeps the orchard treats it the way a Harrowmark farmer treats a difficult neighbour \u2014 politely, with a running grievance, and without the faintest suggestion that anything remarkable is occurring."));

c.push(P("That reaction is the point of the scene, not the sprite. The party has just crossed into a country where this is Tuesday, and the fastest way to establish it is to have a local be bored by it."));

c.push(H2("Tiered Skill DCs"));

c.push(P("Easy 10, Moderate 13, Hard 16, matching the tiers used throughout this campaign."));

c.push(table(
  ["Task", "Skill", "DC", "Tier"],
  [46, 26, 8, 20],
  [
    ["Get a Caerwyn local talking about the Standing Light", "Persuasion / Insight", "10", "Easy"],
    ["Notice the exact point where the Dead Mile begins", "Perception", "13", "Moderate"],
    ["Get Wyn Alder talking plainly about the occupation", "Persuasion / Insight", "13", "Moderate"],
    ["Talk the checkpoint down without papers to show", "Persuasion", "16", "Hard"],
    ["Find the Listening Water locals use at Caerwyn", "Survival / Investigation", "10", "Easy"]
  ],
  { full: true }
));

// ---------------------------------------------------------------- Scene 4
c.push(H2("Scene 4: The Riverbank"));

c.push(P("A short walk from Caerwyn, a slow, wide stream runs beneath old willows \u2014 a Listening Water, well used, its banks worn smooth by generations of people who came here to be heard by something. A local, if asked, will explain the custom plainly: speak at the water\u2019s edge, and it gives your words back, in your own voice, to whoever listens after you. It has done this for as long as anyone remembers. It has not stopped."));

c.push(BOX("The water is unremarkable to look at \u2014 brown-green, unhurried, catching the last of the light the way any river does. Then a voice rises out of it, not the party\u2019s own, mid-sentence, as though the bank had been listening for three years and was only now getting a chance to answer."));

c.push(H3("Running the Scene"));

c.push(P("Let the party sit at the water\u2019s edge and listen. Read each of the four voices below as its own boxed passage, in order, with a pause between each \u2014 do not summarize or paraphrase them together. None of the four is lying. None of them is complete. The party\u2019s job, and the table\u2019s pleasure, is assembling what actually happened from four honest, incomplete, differently motivated accounts, the same technique named directly in the sourcebook under Running Elduvaine."));

c.push(H3("The Court\u2019s Account"));
c.push(BOX("\u201C...and so the Keeper, whom the realm trusted with its oldest law, unmade that trust in a single night. Let it be recorded that no siege took this kingdom, no army broke its walls \u2014 a man we ourselves elevated turned the key we ourselves gave him. History will not forgive us the choosing, whatever it makes of the choosing itself...\u201D"));

c.push(H3("The Commons\u2019 Account"));
c.push(BOX("\u201C...didn\u2019t know it were happening till the well went quiet on a Tuesday and stayed quiet. Levy man come round a fortnight after, very polite about it, asking after grain like he weren\u2019t the reason half the granary walked off nobody knows where. We buried the old well-keeper in spring. Water never did give him back his own voice, after...\u201D"));

c.push(H3("The Occupation\u2019s Own Proclamation"));
c.push(BOX("\u201CBy order, the following is established: all holdings shall be assessed and levied on the published schedule; all persons shall carry proof of standing upon request; the wards of the realm, once the Keeper\u2019s charge, are now the Keeper\u2019s alone, and no other authority may contest their disposition. Compliance ensures continuity. This is not punishment. This is administration.\u201D"));

c.push(H3("A Private Grief"));
c.push(BOX("\u201C...I know you can\u2019t give him back to me, I\u2019m not simple, I know that\u2019s not how you work. I just wanted somebody besides me to have heard his name out loud today. That\u2019s all. That\u2019s all I came here for.\u201D"));

c.push(P("Let the party sit with this before moving on. If a player wants to speak at the water themselves \u2014 a question, a message, simply their own name \u2014 let the water take it and give it back once, faithfully, in their own voice, and then let the scene end there. Nothing further needs to happen; the module\u2019s job is complete the moment the party has heard all four voices."));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("NPC Profiles"));

c.push(H2("Wyn Alder"));

c.push(P("A gnome in her sixties, which for a gnome is a working prime, and a clerk by trade rather than by conviction. Gnomes keep most of Elduvaine\u2019s ledgers \u2014 it is the same profession that staffs the Archive, several rungs down \u2014 and the occupation found the local records already competently kept and simply carried on employing the people keeping them."));
c.push(P("She wears a coat better made than her posting suggests she can afford. Speech: brisk, professional, and visibly rehearsed \u2014 she has said \u201Cpapers, or a reason\u201D enough times that it no longer costs her anything to say. Not a believer in the occupation; a person with a family in Caerwyn and no illusions about what happens to families of the uncooperative."));
c.push(P("Open thread: Wyn is this campaign\u2019s first answer to the question of the occupation\u2019s ordinary, complicit people \u2014 not a villain, not a secret ally, just someone getting through it. A DM can bring her back later, changed by how the party treated her here: grateful and useful if treated well, gone or worse if not."));

// --------------------------------------------------------------- Optional
c.push(H1("Optional Content"));

c.push(H2("The Baker\u2019s Ovens"));
c.push(P("If the table wants more of Caerwyn before moving on, its baker \u2014 matter-of-fact about her never-cooling ovens, per Scene 1 \u2014 is a good source of purely human texture: gossip, a minor local favor asked of the party, or simply an excuse to let the table enjoy the town a while longer before the checkpoint\u2019s tension arrives. No mechanical stakes."));

c.push(H2("What Wyn Knows"));
c.push(P("If the party earns Wyn Alder\u2019s trust in Scene 3 rather than merely getting past her, she has real, specific intelligence about patrol schedules and levy routes inland \u2014 useful DM ammunition for a later module rather than anything this one needs. Do not have her volunteer this unprompted; it should feel earned."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("How the party treated Wyn Alder.", "Threatened, ignored, or treated as a person worth understanding \u2014 track which. This is the campaign\u2019s first real answer to the question it keeps asking about the occupation\u2019s ordinary, complicit people, and it is worth revisiting whenever the party meets another collaborator later in the crusade."));
c.push(BUL("Whether the checkpoint turned violent.", "If it did, word reaches Caer Ysolde days early \u2014 a DM may pay this off as tightened security, an alerted garrison, or a suspicious captor somewhere the party needed an easier path later."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BUL("A standing-stone lamp.", "If the party helped Caerwyn\u2019s baker, she gives them one \u2014 gnome-cut, fist-sized, and full of a morning three weeks ago. Leave it in daylight and it lights a room for hours. It is worth roughly nothing in Caerwyn and will be the strangest object anyone in Harrowmark has ever held."));

c.push(BUL("Otherwise, nothing material.", "This module is deliberately not about treasure. If the party helped the baker or earned Wyn\u2019s trust, the payoff is information and goodwill (see Optional Content and Diverging Paths), not coin or equipment \u2014 keep it that way; Caerwyn has little to spare, and it should not feel plundered by the people meant to be delivering it.", { keepNext: true }));

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
  fs.writeFileSync(stagePath("KC_Module03_Landfall.docx"), buf);
  console.log("Written.");
});
