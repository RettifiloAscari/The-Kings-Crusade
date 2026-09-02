// KC_Module05_TheRoadToVindana.js -- Session Module Five: The Road to Vindana.
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
// This module resolves Branch Ledger entry 1 (the two roads) regardless of
// which variant of Module Two the table played: the second king, lost on
// whichever road the party did not take, reaches them here as news, via the
// Listening Water -- the second consecutive module to use the Four Voices
// device, deliberately, per CLAUDE.md's instruction to use it as a recurring
// technique. The Kyffh\u00E4user refrain from the Barbarossa touchstone (see
// CLAUDE.md's Touchstones) pays off directly in one of the four voices here.
//
// GATED CONTENT NOTICE: the second king's realm is NOT named or confirmed as
// Oksitan's here, even though CLAUDE.md notes that inference as plausible --
// it is explicitly flagged there as not signed off, and this module does not
// invent past that gate. He is referred to only as "the second king."

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
  children: [new TextRun({ text: "The Road to Vindana", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module Five", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("The coalition, the Ward beside it now, marches the final stretch toward Vindana \u2014 and partway there, news catches up with them from the road they did not take. The second king, who led the other column, is dead: lost early, in water, his army come apart behind him. This module delivers that news and lets the party, and the coalition around them, sit with it before Vindana\u2019s siege begins in Module 6. It is shorter and quieter than the modules before it by design. Core scenes run two and a half to three hours; Optional Content fills out the rest of a five-hour session and can be cut cleanly if the table is short on time."));

c.push(H2("Playing the News"));

c.push(P("This module resolves Branch Ledger entry 1 \u2014 the road choice from Module One \u2014 regardless of which variant of Module Two the table played. If the party took the sea road (Module 2A), the second king took the mountain road and was lost fording a swollen river, much as the Ashgate crossing nearly cost the party something in that variant. If the party took the mountain road (Module 2B), the second king took the sea road and was lost in a storm, much as the party\u2019s own crossing was in Calanthe\u2019s telling. Either way, the parallel is deliberate: whatever the party\u2019s own road cost them, the other road cost more, and it is worth letting a player notice that without the module explaining it."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 15, 55],
  [
    ["1. Word on the Road", "15\u201320 min", "A rider arrives first, before the full story does."],
    ["2. The Standing Water", "60\u201390 min", "The Four Voices, again \u2014 deliberately. The module\u2019s centerpiece."],
    ["3. Vindana in Sight", "30\u201345 min", "The march ends in view of the siege to come."],
    ["Optional Content", "30\u201345 min", "Run if the table has time; cut cleanly if not."]
  ],
  { full: true }
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("The second king\u2019s death was exactly what it appears to be: an accident of weather and bad ground, not an ambush and not Vale\u2019s doing \u2014 he has never heard of this man and would not consider him worth the effort if he had. That plainness is the point. The coalition loses its second crown to a river or a storm, the same way real armies have, and the loss is not made more dramatic than it actually was. What the module is actually about is not the death itself but what the survivors do with it \u2014 and the old soldiers\u2019 belief that a king of that stamp does not die but sleeps in some hill, waiting until he is needed again, is offered here as one live option among several for how the coalition receives it, not the only one."));

c.push(PS([DM("DM Only: "), { t: "the second king\u2019s realm is deliberately not named in this module. Oksitan is the plausible candidate by elimination \u2014 Auberitz is a grand duchy, Norvatch does not march \u2014 but that has never actually been settled, and neither this module nor any NPC in it should confirm it. If a player asks directly which realm lost its king, the honest answer any NPC can give is that word has not settled that far down the column yet, which is true and also convenient." }]));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: Word on the Road"));

c.push(P("The news arrives the way bad news on a march always does: badly, in pieces, ahead of anyone who actually knows the whole of it. A lone rider reaches the column first, half a day ahead of any official word."));

c.push(BOX("Tam Ondry rides in on a horse that has clearly been ridden too hard for too long, and does not wait for permission before finding the nearest officer. \u201CThe other column,\u201D he says, and then has to stop and start again. \u201CThe king. He\u2019s not \u2014 he didn\u2019t make the crossing. Nobody\u2019s saying it plain yet, but I was there for the part before I wasn\u2019t, and I\u2019m saying it plain. He\u2019s gone.\u201D"));

c.push(P("Tam is a Harrowmark rider, attached to the other column as a courier, and genuinely shaken rather than performing grief \u2014 let him answer the party\u2019s immediate questions plainly and incompletely, the way someone actually present for a disaster usually can. He does not have the full story; nobody does yet. That is what Scene 2 is for."));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: The Standing Water"));

c.push(P("By evening, the column has reached a Listening Water of its own \u2014 a broad, still pool fed by a spring, the kind of feature every Elduvish village seems to have one of within an hour\u2019s walk. Word has spread that this is where mourning is properly done in this land, and the coalition, uncertain of its own customs for a loss like this, borrows Elduvaine\u2019s without much discussion about whether it is entitled to."));

c.push(BOX("Nobody organizes it. People simply arrive at the water\u2019s edge through the evening, in ones and twos, and say what they have to say to it. By full dark there have been enough voices that the water, when it finally answers, has a great deal to give back."));

c.push(H3("Running the Scene"));

c.push(P("As in Module Three, read each voice below as its own boxed passage, with a pause between each. None of the four is lying. None of them is complete, and this time that incompleteness is doing real emotional work \u2014 the coalition genuinely does not agree yet on what happened or what it means, and the party is hearing that disagreement unfold in real time rather than as settled history."));

c.push(H3("A Soldier\u2019s Account"));
c.push(BOX("\u201C...ford looked no worse than a dozen we\u2019d already crossed. He went in ahead, the way he always did, wouldn\u2019t let anyone else take the lead line first. Current took the horse out from under him before anyone downstream could reach a rope to him. That\u2019s the whole of it. That\u2019s all it ever is, in the end \u2014 a bad step and cold water, no matter whose head wore the crown...\u201D"));

c.push(H3("An Officer\u2019s Account"));
c.push(BOX("\u201C...the column\u2019s command structure held, credit where it\u2019s due \u2014 his marshals had the sense to keep the line moving rather than let the army come apart on the riverbank grieving. But a coalition is a fragile thing to lose a crown out of, and I will not pretend to you that this doesn\u2019t change the arithmetic of everything that comes after it...\u201D"));

c.push(H3("What the Soldiers Are Already Saying"));
c.push(BOX("\u201C...my sergeant swears he\u2019s not dead, not really \u2014 swears a king like that doesn\u2019t just drown in a ditch, says he\u2019s sleeping somewhere under the water or under a hill same as the old stories tell it, and he\u2019ll wake up and come back for us when it matters most. I don\u2019t know if I believe it. I know I\u2019d rather believe it than the alternative, and I\u2019m not the only one...\u201D"));

c.push(H3("A Private Grief"));
c.push(BOX("\u201C...I served his household eleven years and I don\u2019t even know what I\u2019m meant to call this. Not treason, staying loyal to a dead man. Not foolishness either. I just wanted to say his name somewhere it would be heard, because nobody\u2019s asked me how I am, and I don\u2019t expect anybody will.\u201D"));

c.push(P("Let the table sit with this. If a player wants to speak at the water themselves, let it take their words and give them back once, faithfully, exactly as in Module Three. The scene\u2019s job is complete once all four voices have been heard; do not rush it to make room for Scene 3."));

// ---------------------------------------------------------------- Scene 3
c.push(H2("Scene 3: Vindana in Sight"));

c.push(P("The column crests a final ridge the following day, subdued in a way it was not before Scene 2, and Vindana comes into view for the first time \u2014 walls, harbor, and the specific, particular dread of a siege that has not yet begun but is now unmistakably close."));

c.push(BOX("The city sits where the coast folds around a natural harbor, walls pale even at this distance, and for a long moment nobody in the column says anything at all. Then, somewhere behind the party, someone starts walking again, and the rest of the column follows, because that is what a column does."));

c.push(P("End the session here, on the sight of the city rather than on grief \u2014 the tone should turn forward, not linger. Hand off directly to Module 6 for the investment of Vindana."));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("NPC Profiles"));

c.push(H2("Tam Ondry"));
c.push(P("A young Harrowmark rider, courier-attached to the other column, badly shaken by what he witnessed and not yet good at hiding it. Speech: plain, halting when the subject is the crossing itself, otherwise ordinary and even easygoing \u2014 grief has not replaced his personality, just interrupted it."));
c.push(P("Open thread: Tam is a natural recurring minor NPC for the rest of the campaign \u2014 a DM can use him as a courier again whenever a module needs news to arrive from elsewhere, now already established as reliable and already known to the party."));

// --------------------------------------------------------------- Optional
c.push(H1("Optional Content"));

c.push(H2("What the Coalition Believes Now"));
c.push(P("If the table wants more of the coalition\u2019s reaction before Vindana, let the party spend time in camp hearing how differently Oksitan, Auberitz, and Harrowmark soldiers are each processing the loss \u2014 without confirming which realm the crown belonged to (see the DM-Only note above). Play this for texture: an argument about what kind of memorial, if any, is appropriate; a quiet, un-discussed uptick in soldiers touching still water on the march since. No mechanical stakes."));

c.push(H2("The Held Winter"));

c.push(P("The road to Vindana runs within half a day of a birch wood that was planted four days into spring three hundred years ago and has been nine weeks into a winter it was never sown in for most of a year. A party that goes to look finds the most legible single image of the draining the campaign has: a wood in the wrong season, dying of it, with an ecology that has moved in behind the change."));

c.push(P("What lives there now came down out of the hills after the cold did. A pair of winter wolves (SRD, CR 3, 700 XP each) hunt the wood\u2019s edges and will shadow a small party for an hour before committing. Deeper in, a troll (SRD, CR 5, 1,800 XP) has taken the old orchard-keeper\u2019s cottage and is the reason the last two woodsmen did not come back \u2014 a straightforward, dangerous fight for a party of this level, and one that rewards anybody who remembers what regeneration does and does not survive. A DM running the wolves and the troll together should check the total against their table before committing; either alone is a full encounter."));

c.push(P("The wood also has a dryad (SRD, CR 1), and she is the actual content of the scene. She is not an encounter. She has been the spirit of a spring wood for three centuries and is now the spirit of a winter one, and she is neither hostile nor grateful nor able to leave. She will talk, at length, and what she wants is for somebody to explain to her what has been done and why, which nobody in the party can do. A table that fights her has misread the scene; a table that sits down with her gets the campaign\u2019s grief in one conversation, from something that is not a person and is nonetheless quite plainly grieving."));

c.push(PS([DM("DM Only: "), { t: "if the party asks whether killing Vale fixes her wood, the honest answer is that nobody knows and she does not expect it to. Do not have her forgive anyone, and do not have her curse anyone. She is a wood in the wrong season, and the wrong season is not going to end because a war did." }]));

c.push(H2("A Skirmish on the Ridge"));
c.push(P("If the table wants combat before this module ends, a small occupation patrol \u2014 scouting Vindana\u2019s approaches, not expecting a coalition column this size \u2014 can be spotted and engaged on the final ridge before Scene 3. Use the Occupation Guard stat block from Module 3 (three or four is sufficient); they flee to warn the city the moment the fight turns against them, which is itself useful information for Module 6 rather than a failure state."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("Branch Ledger entry 1, resolved.", "Record here how the party\u2019s own road compared to the second king\u2019s \u2014 what their crossing cost against what his cost him. This closes the loop opened in Module One."));
c.push(BUL("How the party engaged with the mourning at the Standing Water.", "Whether they spoke at the water themselves, listened only, or kept apart from it entirely \u2014 worth a line in the ledger, since it is the party\u2019s first real exposure to how this coalition grieves, and it will not be its last."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BUL("A flask of Listening Water.", "Drawn at the Standing Water while the mourning was going on, by whoever thought to. It holds what was said at its mouth and gives it back once, in the speaker\u2019s own voice, and then it is only water. Whatever the party chose to put in it is the point; a DM should ask, and should write down the answer."));

c.push(BULLET([{ t: "Otherwise, nothing material. ", b: true }, { t: "This module is not about treasure any more than Module Three was. Its currency is information and tone \u2014 the Branch Ledger entry above is the actual payoff." }], { keepNext: true }));

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
  fs.writeFileSync(stagePath("KC_Module05_TheRoadToVindana.docx"), buf);
  console.log("Written.");
});
