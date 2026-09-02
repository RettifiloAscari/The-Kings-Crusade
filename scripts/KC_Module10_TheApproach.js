// KC_Module10_TheApproach.js -- Session Module Ten: The Approach.
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
// The last quiet module before the finale. It does not force the turn-back-
// or-hold decision -- that belongs to Module Eleven -- it only makes sure the
// table has actually weighed it before Module Eleven asks. The resistance
// leader (the royal house member who escaped and refuses evacuation, per
// CLAUDE.md) appears in person for the first time here, referred to as "the
// Regent" for the same reason the Ward and the Magistrate have no proper
// names -- the royal family's names are still open in CLAUDE.md.

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
  children: [new TextRun({ text: "The Approach", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module Ten", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("The coalition\u2019s final march brings it within sight of Caer Ysolde \u2014 and for the first time since Landfall, the party sees what the deepest drainage actually looks like, at the scale of a capital rather than a mile of road. This module also introduces the resistance leader in person: the member of the royal house who escaped the fall of the kingdom and has refused every offer of evacuation since. Nothing here forces the campaign\u2019s central choice. It only makes sure the table has actually sat with it before Module Eleven asks. Core scenes run three to three and a half hours; Optional Content fills out the rest of a five-hour session and can be cut cleanly if the table is short on time."));

c.push(H2("A Quiet Module on Purpose"));

c.push(P("Resist adding combat to pad this module\u2019s length. Its entire job is weight and reflection after Module Nine\u2019s real cost, and a fight here would undercut that rather than serve it. If a table wants more to do, use Optional Content rather than inventing an encounter \u2014 this is one of the few modules in the campaign built to run under its target time without anything having gone wrong."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 15, 55],
  [
    ["1. The Last Miles", "20\u201330 min", "The march\u2019s final stretch; accumulated weight, not new danger."],
    ["2. Caer Ysolde, Dark", "30\u201340 min", "First sight of the capital. The Standing Light has failed here."],
    ["3. The Regent", "60\u201375 min", "The resistance leader, met in person for the first time."],
    ["4. What Waits at the Gates", "30\u201345 min", "The turn-back-or-hold question, weighed narratively."],
    ["Optional Content", "30\u201345 min", "Run if the table has time; cut cleanly if not."]
  ],
  { full: true }
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("The Regent has spent three years running a resistance that could never win, on purpose \u2014 not to defeat the occupation, which was never possible with the forces available, but to make sure Elduvaine never became a place that had simply accepted Vale. They know a great deal about the deepest state of the drainage, about the Archive\u2019s vaults from the outside, and about what the rest of the royal house actually wants for the kingdom afterward, which is not something they all agree on. None of this resolves the campaign\u2019s central open question \u2014 what Elduvaine should become after the war \u2014 and this module should not resolve it either."));

c.push(PS([DM("DM Only: "), { t: "the Regent is a real, developed person, not an exposition device \u2014 give them their own opinions, including ones the party might disagree with. Do not use them to state a \u201Ccorrect\u201D answer to the turn-back-or-hold question; per CLAUDE.md, that question is deliberately left to the table, and a later session must not reweight it, including through this NPC\u2019s own persuasive certainty." }]));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: The Last Miles"));

c.push(P("The road from the field battle to Caer Ysolde is the campaign\u2019s longest quiet stretch, and it should feel like one. Let the party notice how much has actually changed since Duncarrow: who they have lost, who they have freed, and how little any of it resembles the war they thought they were marching into."));

c.push(BOX("Nobody talks much on the last day\u2019s march. It isn\u2019t grim exactly \u2014 more that everyone in the column seems to be doing the same private arithmetic, adding up what the last several months have actually cost against what they came here for, and none of them are quite ready to say the sum out loud."));

c.push(P("This scene has no mechanical content. Let the party talk to whichever recurring NPCs the campaign has given them \u2014 the Ward, the Magistrate, Brenna Vane, Sera Vosk or Garrick Hollow, a captured Drell or Voss if either survived \u2014 about anything at all. This is the campaign giving its own cast room to be people rather than plot before the finale asks something of all of them."));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: Caer Ysolde, Dark"));

c.push(P("The column crests a final rise, and Caer Ysolde is below them \u2014 the capital where three rivers braid together, exactly as the sourcebook describes it, except for the one thing no description quite prepares a party for."));

c.push(BOX("It is dark. Not ruined, not besieged-looking from this distance \u2014 simply dark, in a way a city its size has apparently never been before tonight, its pale stone holding no more light than any ordinary quarry\u2019s would. Somewhere down there is the Ysolde Archive, the largest collection of magical knowledge in the world, and the party is looking at the exact place the sourcebook already told them the draining hits first and hardest."));

c.push(P("Let this land as the wonder-into-grief beat it is meant to be, deliberately paired against Module Three\u2019s Caerwyn glowing at dusk. If a player remarks on the contrast, that is the scene working exactly as intended \u2014 do not have an NPC explain it; let the silence around the column do that."));

// ---------------------------------------------------------------- Scene 3
c.push(H2("Scene 3: The Regent"));

c.push(P("A resistance contact \u2014 cautious, testing the party before committing to anything \u2014 leads them to a meeting the coalition\u2019s own officers have been trying to arrange for weeks: the Regent, in person, for the first time."));

c.push(BOX("They are older than the party expects, and plainer \u2014 no crown, no court dress, a coat as travel-worn as anything Doria Kell wears. \u201CThree years I\u2019ve been told, by people who mean well, that I should let myself be taken somewhere safe,\u201D they say, by way of greeting. \u201CI have found that \u2018safe\u2019 and \u2018useful\u2019 are very rarely the same offer. You, I take it, are neither.\u201D"));

c.push(H3("Running the Scene"));

c.push(P("The Regent will speak plainly about what the resistance actually knows: the drainage\u2019s current state, Vale\u2019s general pattern of movement within the city (he is rarely seen, and never predictably), and, if asked directly, their own view on what should happen to Elduvaine after the war \u2014 a view the sourcebook establishes the royal house does not universally share. Let this be a real conversation with a real person who has earned the right to an opinion, not a briefing. No skill check gates any of this information; the Regent gives it freely to people they have decided, correctly, are worth trusting."));

c.push(PS([DM("DM Only: "), { t: "if the party has freed the Ward and the Magistrate, this is a natural moment for the Regent to ask after them directly and personally rather than as political assets \u2014 let that land as real family feeling, complicated by three years apart and by disagreement about what comes next." }]));

// ---------------------------------------------------------------- Scene 4
c.push(H2("Scene 4: What Waits at the Gates"));

c.push(P("Before the final push, let the coalition\u2019s leadership \u2014 and the party themselves \u2014 sit with the question the whole campaign has been building toward without answering: take Elduvaine, or, having proven it can be done, choose not to hold it."));

c.push(BOX("Nobody puts it that bluntly. It comes out instead as smaller questions, asked by different people for different reasons \u2014 an Oksitan officer asking what happens to the Promise if the coalition marches home early; the Magistrate asking, very precisely, what \u201Cvictory\u201D is actually meant to mean here; Xavier, quieter than usual, asking the party directly whether they think this is a kingdom that wants saving on the coalition\u2019s terms."));

c.push(P("Do not resolve this scene with a decision. Its entire purpose is to make sure the table has heard the question asked seriously, by people with real stakes in the answer, before Module Eleven makes them answer it. End the module on the coalition making final preparations \u2014 supply, formation, the ordinary business of an army about to do something enormous \u2014 and hand off directly to Module Eleven."));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("NPC Profiles"));

c.push(H2("The Regent"));
c.push(P("No proper name yet assigned, for the same reason as the Ward and the Magistrate (see the DM-Only note in What Is Actually Happening). Speech: measured, unsentimental, entirely unbothered by three years of danger that would have broken most people. Has genuinely considered leaving Elduvaine and genuinely chosen not to, more than once."));
c.push(P("Open thread: the Regent\u2019s own view on Elduvaine\u2019s future is theirs to hold and the campaign\u2019s not to adjudicate \u2014 a DM developing the endgame should treat their opinion as one real voice among several in the royal house, not the deciding one."));

// --------------------------------------------------------------- Optional
c.push(H1("Optional Content"));

c.push(H2("The Resistance, Briefly"));
c.push(P("If the table wants more of Elduvaine\u2019s ordinary resistance before the finale, let them meet a handful of the Regent\u2019s people directly \u2014 farmers, former Archive clerks, people who have been doing small, unglamorous sabotage for three years rather than anything a story would call heroic. No mechanical stakes; pure texture, and a chance to make the liberated feel like people rather than a cause."));

c.push(H2("Letters Home"));
c.push(P("If the table wants a quiet character scene, let each PC write (or dictate, or simply compose in their head) a letter home before the finale \u2014 to Harrowmark, to wherever they came from, to whoever they left behind. No mechanical content; this is purely for the players who want a moment of reflection before the campaign\u2019s climax."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("What the party told the Regent, and what the Regent told them.", "Not mechanically trackable in the usual sense, but worth a DM\u2019s private note \u2014 this conversation is the table\u2019s clearest signal yet of which way they are leaning on the campaign\u2019s central question, and it is worth remembering going into Module Eleven."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BULLET([{ t: "Nothing material. ", b: true }, { t: "This module, like Module Three and Module Five before it, is not about treasure. Its currency is entirely the conversation in Scenes 3 and 4." }], { keepNext: true }));

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
  fs.writeFileSync(stagePath("KC_Module10_TheApproach.docx"), buf);
  console.log("Written.");
});
