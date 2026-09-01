// KC_Module11_TheDecisionAtTheGates.js -- Session Module Eleven: The Decision at the Gates.
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
// THE FINAL MODULE. Maedoc Vale appears in person for the first and only time
// in the whole campaign -- every prior module kept him deliberately absent.
// Both endings (hold Elduvaine, or turn back within sight of it) are written
// in full; CLAUDE.md is explicit that neither is favored and a table decides
// at the table. The Refrain's last line changes here, and only here -- see
// the very end of this file, and do not let it change anywhere else.
//
// PRESERVED AMBIGUITIES -- do not resolve these, here or in any future
// revision, without explicit sign-off:
//   - Whether Vale is still human. CLAUDE.md: "deliberately ambiguous in v1."
//   - What the deepest vaults actually contain. CLAUDE.md: the sourcebook
//     "states his motive; it does not state his destination."
// This module gives the party a real, dangerous, in-person confrontation
// with Vale without answering either question. Read Scene 2 and Scene 3
// carefully before running or revising them.

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
  children: [new TextRun({ text: "The Decision at the Gates", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module Eleven (Final)", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("The coalition takes Caer Ysolde. The party descends into the Ysolde Archive and meets Maedoc Vale in person for the first and only time in this campaign. And then \u2014 with the wizard stopped and the capital theirs \u2014 the table makes the choice this entire crusade has been walking toward without ever being told which answer is correct: hold Elduvaine, and everything the Promise was owed, or set it down, having proven it could be taken, and go home with nothing but that proof. Both endings are written here in full. Core scenes run four and a half to five hours; this module is built to take the whole session, and should not be rushed to fit a shorter one."));

c.push(H2("What This Module Will Not Do"));

c.push(P("It will not tell you whether Vale is still human. It will not tell you what the deepest vaults actually contain. It will not tell you which ending is the right one. These are load-bearing silences, kept on purpose across every module that came before this one, and breaking any of them now \u2014 to make the climax feel more resolved \u2014 would cost the campaign more than it would gain. Run this module trusting the silence. It has held this long because it works."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 15, 55],
  [
    ["1. The Gates of Caer Ysolde", "30\u201345 min", "The capital falls; brief resistance, not another siege."],
    ["2. Into the Archive", "30\u201345 min", "Descent. Atmosphere over combat. See the note below."],
    ["3. Maedoc Vale", "75\u201390 min", "The confrontation. Stat block below. Read before running."],
    ["4. The Decision at the Gates", "45\u201360 min", "The table chooses. Both endings follow, written in full."]
  ],
  { full: true }
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("Vale is exactly what the sourcebook has always said he is: a man who wanted to finish reading, priced the world against that want, and got the sum right by his own arithmetic. He is not surprised the coalition reached him. He is mildly, genuinely inconvenienced by it \u2014 the way a scholar is inconvenienced by a library closing before he finishes a chapter, not the way a tyrant is inconvenienced by losing a war. Play him without raising his voice, ever, in this entire module."));

c.push(PS([DM("DM Only: "), { t: "when the party finally sees him, describe him precisely and describe nothing that settles whether he is still human. He looks like a man who has not slept properly in three years and does not appear to have noticed. His hands, when he gestures, do exactly what a person\u2019s hands do. Leave it there. If a player casts a spell or uses an ability that would definitively answer the question \u2014 true seeing, a paladin\u2019s Divine Sense reading him as something other than a living humanoid, anything similarly conclusive \u2014 it is your table\u2019s call whether to let that mechanic resolve the ambiguity or to rule that whatever he has become resists exactly that kind of certainty. Either choice is valid; picking one is a bigger decision than this module makes for you on purpose." }]));

c.push(P("Likewise, whatever the party glimpses of the vaults he has been reading toward \u2014 and they should glimpse something, briefly, on the way to him \u2014 describe it as overwhelming and specific in texture (light that behaves wrong, shelving that recedes further than the room containing it should allow, a smell like a held breath) without ever stating what any of it actually is or means. The sourcebook does not answer this question. This module should not either."));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: The Gates of Caer Ysolde"));

c.push(P("Caer Ysolde\u2019s garrison, already thin after Vindana and the field battle, does not mount a real defense \u2014 this is not another Module Six or Module Seven, and should not be run at that length or intensity. A short, sharp engagement at the outer gate (use the Occupation Guard stat block, Module 3, four to six of them) is enough; let the capital fall quickly, because the module\u2019s actual content is what waits underneath it, not another wall to break."));

c.push(BOX("The city opens ahead of the coalition almost too easily, streets emptying rather than resisting, until it becomes clear this is not fear of the coalition \u2014 it is the particular quiet of a place whose attention has been somewhere else entirely for three years. Nobody in Caer Ysolde seems especially surprised that today is the day it ends."));

c.push(P("Let the party move quickly through the city toward the Archive. This scene\u2019s job is momentum, not a second siege."));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: Into the Archive"));

c.push(P("The Ysolde Archive, even three years drained, is the largest collection of magical knowledge in the world, and the party descends through galleries that were built for scholarship rather than defense \u2014 which does not make them safe, only strange."));

c.push(BOX("The stacks go down further than the building above them should allow, and further than anyone escorting the party is willing to explain. Somewhere below, faintly, is a sound like a held breath, and a light that does not come from any torch or window doing something none of the party\u2019s own light sources do."));

c.push(P("Run this as atmosphere and tension rather than a dungeon crawl \u2014 a handful of Tiered Skill DC checks (below) to navigate safely, and perhaps one non-lethal hazard (a ward triggered by curiosity rather than malice, a section of stacks that briefly does not behave like ordinary space) rather than a string of fights. The Archive\u2019s remaining wards are not trying to kill the party; they are simply still doing their job, and their job was never intruder defense in the ordinary sense."));

// -------------------------------------------------------------- Skill DCs
c.push(H2("Tiered Skill DCs"));

c.push(P("Easy 10, Moderate 13, Hard 16, matching the tiers used throughout this campaign."));

c.push(table(
  ["Task", "Skill", "DC", "Tier"],
  [40, 24, 12, 24],
  [
    ["Navigate the Archive\u2019s stacks without becoming lost", "Investigation / Survival", "13", "Moderate"],
    ["Recognize a ward as passive rather than hostile", "Arcana / Investigation", "13", "Moderate"],
    ["Avoid triggering a curiosity-ward while examining the stacks", "Investigation / Wisdom (DM\u2019s judgment)", "13", "Moderate"],
    ["Read Vale\u2019s intentions correctly before he speaks", "Insight", "16", "Hard"]
  ],
  { full: true }
));

// ---------------------------------------------------------------- Scene 3
c.push(H2("Scene 3: Maedoc Vale"));

c.push(P("At the bottom of the Archive, past wards that have stopped noticing anything but the specific work they were left doing, the party finds him \u2014 surrounded by open books rather than guards, in a room that does not feel like the climax of anything until he actually looks up."));

c.push(BOX("He does not reach for a weapon. He closes the book nearest to him \u2014 carefully, with a strip of cloth marking his place \u2014 before he says anything at all. \u201CI rather thought it would be you, eventually. I want to say I\u2019m sorry for what this cost everyone, and I find that I mean it and it changes nothing about what I\u2019m going to keep doing until someone makes me stop. So. I suppose that\u2019s you, now.\u201D"));

c.push(H3("Running the Scene"));

c.push(P("Let him talk before he fights. He will answer honest questions honestly and without defensiveness \u2014 about the wards, about the deepest vaults (without ever actually naming what is in them; he will describe his want, never his destination), about the royal house, about the coalition\u2019s Promise, which he finds mildly, academically interesting as an example of exactly the kind of arithmetic he understands very well. He does not beg, does not monologue, and does not offer a redemption the party can accept \u2014 this is not that kind of scene. When the conversation ends (the party attacks, or he simply decides the conversation is over and returns to his books in a way that forces the issue), the fight is real and dangerous."));

c.push(H3("Stat Block"));

c.push(...SB({
  name: "Maedoc Vale",
  meta: "Medium humanoid, any alignment \u2014 SRD Archmage, renamed",
  ac: "15 (mage armor)",
  hp: "99 (18d8 + 18)",
  speed: "30 ft.",
  str: 10, dex: 14, con: 12, int: 20, wis: 15, cha: 16,
  saves: "Int +9, Wis +6",
  skills: "Arcana +13, History +13",
  senses: "passive Perception 12",
  langs: "any six languages",
  cr: "12 (8,400 XP)",
  traits: [
    { n: "Magic Resistance", t: "Vale has advantage on saving throws against spells and other magical effects." },
    { n: "Spellcasting", t: "Vale is an 18th-level spellcaster (spell save DC 17, +9 to hit with spell attacks; Intelligence). At will: disguise self, invisibility. Cantrips: fire bolt, light, mage hand, prestidigitation, shocking grasp. He has spell slots for 1st through 9th level and, before combat, casts mage armor and stoneskin on himself. His prepared spells include detect magic, identify, and magic missile (1st); detect thoughts, mirror image, and misty step (2nd); counterspell, fly, and lightning bolt (3rd); banishment and fire shield (4th); cone of cold, scrying, and wall of force (5th); globe of invulnerability (6th); teleport (7th); mind blank (8th, cast on himself before combat); and time stop (9th). Full spell mechanics follow the SRD Archmage entry exactly; only the name and flavor have changed." },
    { n: "A Warder\u2019s Instinct", t: "Vale treats every spell as a rule to be applied rather than a force to be unleashed \u2014 flavor his spellcasting as precise and procedural, closer to a lock turning than an explosion, even when the effect is devastating." }
  ],
  actions: [
    { n: "Fire Bolt", t: "Ranged Spell Attack: +9 to hit, range 120 ft., one target. Hit: 22 (4d10) fire damage. (Cantrip; his default action when not spending a higher-level slot.)" }
  ]
}));

c.push(P("Vale fights to disable and delay, not to kill quickly \u2014 counterspell, banishment, and wall of force are all more in character for him than raw damage, and a DM should feel free to lean on control effects over pure lethality. He does not surrender. He also does not fight to the death in the ordinary heroic sense: at the DM\u2019s discretion, a Vale reduced to under 20 hit points may attempt to teleport away rather than be finished off, particularly if the table would prefer the ambiguity of his fate to his confirmed death \u2014 either resolution (killed outright, or escaping to an unknown end) is fully compatible with everything else this module does."));

// ---------------------------------------------------------------- Scene 4
c.push(H2("Scene 4: The Decision at the Gates"));

c.push(P("However Vale\u2019s confrontation ended, Elduvaine\u2019s occupation is, in every practical sense, over. The coalition holds Caer Ysolde. The royal house \u2014 however much of it the party has freed \u2014 is free to reclaim its own capital. And the Promise, made all the way back at Duncarrow, is finally payable, if the coalition chooses to collect it."));

c.push(BOX("Xavier finds the party one last time, in a city that is theirs now in whatever sense anyone can agree on. He does not tell them what to decide. He has clearly been thinking about the question as long as they have, and has not arrived anywhere he is willing to call an answer. \u201CI made a promise to get allies to march,\u201D he says. \u201CI never actually decided, myself, what I thought we should do once we got here. I don\u2019t know that it\u2019s mine to decide. I\u2019m not sure it\u2019s only yours, either. But I think it has to be somebody\u2019s, today.\u201D"));

c.push(P("Let the table talk this through as themselves, not only as their characters \u2014 this is the campaign\u2019s actual final decision, per CLAUDE.md, and it is deliberately not adjudicated by any NPC, any mechanic, or this module. Weigh whatever the campaign has given them to weigh: the Promise and what the coalition is owed; the royal house\u2019s own divided opinions (the Ward, the Magistrate, the Regent, if met); what Elduvaine\u2019s ordinary people \u2014 Wyn Alder, Caerwyn\u2019s baker, the resistance the Regent led \u2014 actually seem to want; and what holding a kingdom whose magic responds to intention would even mean, held by people who marched here for a share of it."));

c.push(P("When the table decides, run the matching ending below. Do not run both. Do not hint, before this moment, which one the campaign favors \u2014 it does not favor either."));

// ============================================================ ENDING ONE
c.push(H1("Ending One: The Crown Held"));

c.push(P("The coalition claims Elduvaine. Xavier\u2019s champions stand at the center of it \u2014 this is the ending in the tradition of a war won and a promise kept, and it should be played with real, earned triumph."));

c.push(BOX("The banners go up over Caer Ysolde by evening \u2014 Harrowmark\u2019s, Oksitan\u2019s, Auberitz\u2019s, and, restored to its proper place above them, Elduvaine\u2019s own. The Standing Light does not come back all at once; that will take years, the royal house\u2019s own scholars say, now that Vale\u2019s draining has stopped. But somewhere in the city tonight, for the first time in three years, one single street lights itself without anyone striking a flame, and the crowd that gathers to watch it is not entirely sure whether they are permitted to cheer. They do anyway."));

c.push(P("The Promise is paid. The coalition\u2019s land grants are drawn up within weeks, and the Ysolde Archive opens its doors \u2014 carefully, under terms the surviving royal house insists on negotiating rather than simply granting \u2014 to scholars from every allied realm. It is not a clean ending. Whichever member of the royal house the party has come to know best has real, complicated feelings about a foreign coalition now holding a legal claim on part of their own kingdom, and says so, and is not wrong to. Elduvaine is saved, and it now belongs, in part, to people who were not born in it. Both of those things are true at once, and the campaign does not ask you to resolve the tension between them \u2014 only to have earned the right to feel it."));

c.push(P("Xavier is remembered as the king who answered the call and, more than that, as the Wyvernheart \u2014 the name is fully his own by the time the histories are written, whatever he privately still thinks of it. The party is remembered by whatever name the table gave them, in whatever register a Listening Water somewhere is already returning it in."));

// ============================================================ ENDING TWO
c.push(H1("Ending Two: The Crown Set Down"));

c.push(P("The coalition proves the crusade could be won, and then, deliberately, does not collect on it. This ending should be played as the harder-won triumph, not the sad one \u2014 Balian gave up Jerusalem having traded it for every life inside; the Skaldic Bard\u2019s Richard never took the holy city and the song does not treat that as failure. Play this the same way."));

c.push(BOX("There is no ceremony for turning a kingdom down. Xavier simply gives the order, and the coalition \u2014 grumbling, in places relieved, in places genuinely angry \u2014 begins the work of leaving rather than the work of staying. The royal house is left exactly what it was always owed: its own capital, its own choices, and no foreign banner over any of it."));

c.push(P("This costs something real, and the module should let it. The Promise goes unpaid \u2014 Oksitan and Auberitz march home with nothing but the proof that Vale could be beaten, which some of their own people will not consider sufficient payment for months of war and real losses. The party may need to answer for this decision to allies who feel betrayed by it, and that reckoning is not this module\u2019s to resolve; it is the shape of whatever the table imagines happens next. What Elduvaine gets instead is what it was actually owed from the start: itself, undiminished by a debt to its own rescuers."));

c.push(P("Whichever member of the royal house the party has come to know best has their own reaction to this, and it should not be uncomplicated gratitude \u2014 someone who wanted to hold and someone who wanted to give have both had a say in what almost happened, and not everyone in that divided house will agree the party chose correctly. Xavier\u2019s own place in this ending is quieter than in the other one; the songs about him will still be written, but they will have to reckon with a king who marched an army to the gates of a kingdom and then, deliberately, did not take it. That is a harder story to tell well, and this campaign trusts its table to have earned the right to tell it anyway."));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("NPC Profiles"));

c.push(H2("Maedoc Vale"));
c.push(P("See the sourcebook for his full established character. In this module specifically: unhurried, precise, genuinely regretful about the cost without that regret changing anything about his intentions. Never raises his voice. Closes his books carefully even when the room is about to become a battlefield."));
c.push(P("Open thread: whether he dies, escapes, or is captured is this module\u2019s to decide at the table (see Scene 3\u2019s stat block note) \u2014 a DM keeping the campaign\u2019s deliberate ambiguity about his humanity intact may prefer to leave his ultimate fate just as unresolved as that question, whatever the mechanical outcome of the fight."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("The campaign\u2019s final choice.", "Held or set down \u2014 this is the Branch Ledger\u2019s last and largest entry, and everything else tracked across the campaign (the road choice, Calanthe or Ashgate, every rescue, every NPC\u2019s fate) is worth revisiting here as context for how the table actually arrived at this decision."));
c.push(BUL("Vale\u2019s fate.", "Killed, escaped, or captured \u2014 record which, and whether the table wanted the ambiguity preserved or resolved."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BUL("Whatever the Archive itself now represents.", "Not itemized loot \u2014 the campaign\u2019s actual final reward is the ending the table chose, and this module does not attach a treasure list to that choice on purpose."));

// -------------------------------------------------------------- Refrain
c.push(H1("The Refrain"));

c.push(P("This is the only place in the entire campaign where the following changes. Read it exactly as written below, once, and end the session."));

c.push(VERSE([
  "By thought, and by word, and by deed,",
  "the king\u2019s own chosen kept their creed.",
  "Far from home, where the quiet land lay,",
  "and the water kept their names that day."
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
  fs.writeFileSync(stagePath("KC_Module11_TheDecisionAtTheGates.docx"), buf);
  console.log("Written.");
});
