// KC_DM_Reference_Guide.js -- The DM Reference Guide.
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
// SINGLE-COLUMN by design (SINGLE_COL_MATCH in tools/build.sh already points
// at this document's basename): wide scannable tables, a stat block index,
// and the Branch Ledger -- every Diverging Paths entry from all eleven module
// slots, compiled here with a blank column for what actually happened at a
// given table. Regenerate this file whenever a new module adds its own
// Diverging Paths section; do not let the ledger drift from the modules it
// summarizes.

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
  children: [new TextRun({ text: "The DM Reference Guide", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("This document is the table\u2019s working reference, not a fifth telling of the setting or the story. It gathers three things a DM needs at hand and does not want to hunt for across twelve other files: a one-page map of the whole campaign\u2019s shape, an index of every stat block already built so a name can be found without a search, and the Branch Ledger \u2014 the compiled record of every tracked divergence, module by module, with a blank column for what actually happened at this table. Everything here is drawn from the sourcebook and the eleven module slots; nothing here is new canon, and nothing in the \u201cNot yet decided\u201d table in CLAUDE.md has been resolved to make this document tidier."));

c.push(P("Use it during play as a quick-lookup: which module a name belongs to, what CR a recurring stat block sits at, what got decided last session that this session might reference. Use the Branch Ledger after every session, not just at the end of the campaign \u2014 an entry filled in while it is fresh is worth more than one reconstructed from memory two modules later."));

// ------------------------------------------------------ Campaign at a Glance
c.push(H1("Campaign at a Glance"));

c.push(P("Eleven module slots (twelve files \u2014 Module Two forks by road and only one half is played). Starting level 5, milestone advancement, built for a party of 4\u20136. Each module\u2019s own Overview carries a scene-by-scene pacing table against its five-hour budget; this table is the one-line shape of the whole march."));

c.push(table(
  ["Module", "Title", "Core Content"],
  [12, 30, 58],
  [
    ["1", "The Muster", "Duncarrow; the wyvern at Greywatch; the road choice \u2014 sea or mountain."],
    ["2A", "The Sea Road", "Warden Thane\u2019s hold; the port of Calanthe and its fate."],
    ["2B", "The Mountain Road", "The Ashgate Ford; Baron Vell\u2019s toll-keep."],
    ["3", "Landfall", "First ground in occupied Elduvaine; the checkpoint; Wyn Alder."],
    ["4", "The Coalition", "Sennoch Hall rescue; the Ward freed; Doria Kell and Norvatch."],
    ["5", "The Road to Vindana", "The second king\u2019s drowning resolved; mourning at the Standing Water."],
    ["6", "The Siege of Vindana: Investment", "Engines raised; the supply raid; First Assault."],
    ["7", "The Siege of Vindana: The Breaking", "Marshal Drell\u2019s fate; the city breaks; Xavier becomes the Wyvernheart."],
    ["8", "Held Ground", "Securing Vindana; the Magistrate freed."],
    ["9", "The Field Battle", "General Voss\u2019s army met in the open and broken; Tam Ondry."],
    ["10", "The Approach", "The march on Caer Ysolde; counsel with the Regent."],
    ["11", "The Decision at the Gates (Final)", "Caer Ysolde falls; Maedoc Vale confronted; hold or set down."]
  ]
));

// ---------------------------------------------------------------- Stat Blocks
c.push(H1("Stat Block Index"));

c.push(P("Every named stat block built so far, in order of first appearance. \u201cSRD base\u201d names the unmodified SRD creature a block was built from, per CLAUDE.md\u2019s Mechanical Validation rules \u2014 renamed blocks keep the base creature\u2019s numbers exactly and change only name, flavor text, and (where noted in the module) minor cosmetic description. The Occupation Guard reappears by direct reference in Modules 4, 6, 7, 9, and 11 rather than being redefined each time; look for it under Module 3."));

c.push(table(
  ["Name", "SRD Base", "CR (XP)", "AC", "HP", "Module"],
  [20, 20, 14, 10, 10, 26],
  [
    ["Wyvern", "Wyvern, unmodified", "6 (2,300)", "13", "110", "1 \u2014 The Muster"],
    ["Warden Ivor Thane", "Veteran, renamed", "3 (700)", "17", "58", "2A \u2014 The Sea Road"],
    ["Thane\u2019s Guard", "Bandit, unmodified", "1/8 (25)", "12", "11", "2A \u2014 The Sea Road"],
    ["Baron Osgar Vell", "Bandit Captain, renamed", "2 (450)", "15", "65", "2B \u2014 The Mountain Road"],
    ["Ashgate Scout", "Scout, unmodified", "1/2 (100)", "13", "16", "2B \u2014 The Mountain Road"],
    ["Occupation Guard", "Guard, unmodified", "1/8 (25)", "16", "11", "3 \u2014 Landfall (reused: 4, 6, 7, 9, 11)"],
    ["Marshal Ossian Drell", "Veteran, renamed", "3 (700)", "17", "58", "6 \u2014 Vindana: Investment (returns: 7)"],
    ["General Ilyana Voss", "Gladiator, renamed", "5 (1,800)", "16", "112", "9 \u2014 The Field Battle"],
    ["Maedoc Vale", "Archmage, renamed", "12 (8,400)", "15", "99", "11 \u2014 The Decision at the Gates"]
  ]
));

c.push(P("A stat block\u2019s full traits and actions live only in its own module\u2019s script \u2014 this index is for finding the right file fast, not for running the creature from this page alone."));

// ---------------------------------------------------------------- Recurring NPCs
c.push(H1("Recurring NPCs"));

c.push(P("Named NPCs the campaign\u2019s own modules flag as \u201copen threads\u201d \u2014 built to be brought back, not to be used once and discarded. Three (the Ward, the Magistrate, the Regent) are members of Elduvaine\u2019s royal house and deliberately carry no proper name yet; see CLAUDE.md\u2019s \u201cNot yet decided\u201d table before assigning one."));

c.push(table(
  ["Name", "Introduced", "Role"],
  [22, 14, 64],
  [
    ["Xavier III of Harrowmark", "1", "The calling king. Leads in person; not yet the Wyvernheart until Module 7."],
    ["Huntmaster Brenna Vane", "1", "Greywatch\u2019s huntmaster; runs the wyvern hunt; a recurring source of camp levity."],
    ["Sera Vosk", "2A", "Auberitz quartermaster; a standing logistics and information contact."],
    ["Garrick Hollow", "2B", "Hired mountain guide; a standing source for overland routes and rumor."],
    ["Wyn Alder", "3", "An occupation-era collaborator at the Landfall checkpoint; the campaign\u2019s first answer to what an ordinary complicit person looks like."],
    ["Doria Kell", "4", "Norvatch factor. Trades information for a price; genuinely takes no side."],
    ["The Ward", "4", "Royal house, freed at Sennoch Hall. Precise rather than broken; a standing voice for what the liberated want."],
    ["Tam Ondry", "5", "Harrowmark courier attached to the lost column; a standing source for news arriving from elsewhere."],
    ["Marshal Ossian Drell", "6", "Commands Vindana\u2019s garrison; a competent officer, not a fanatic. Fate resolved in Module 7."],
    ["The Magistrate", "8", "Royal house, freed in Vindana. Fought her captivity with bureaucratic warfare and won."],
    ["General Ilyana Voss", "9", "Commands Vale\u2019s field army. Fate resolved this module; a captured Voss is a real long-term asset."],
    ["The Regent", "10", "Royal house, at large; runs what resistance survives and refuses evacuation. See the sourcebook."],
    ["Maedoc Vale", "11", "The wizard. Appears in person only here. Whether he is still human is deliberately unresolved."]
  ]
));

// ---------------------------------------------------------------- Branch Ledger
c.push(H1("The Branch Ledger"));

c.push(P("Every tracked divergence across the campaign, compiled from each module\u2019s own Diverging Paths (DM Only) section. Fill in the last column at the table, as it happens \u2014 this is the record that makes the campaign replayable rather than a railroad with scenery, and it is the reason this document is single-column: these rows need the width."));

c.push(table(
  ["Module", "Divergence", "Options", "What Actually Happened"],
  [10, 26, 40, 24],
  [
    ["1", "The wyvern\u2019s fate", "Killed \u00b7 driven off \u00b7 roped and released", ""],
    ["1", "Branch Ledger entry 1: the road choice", "Sea road (2A) \u00b7 Mountain road (2B)", ""],
    ["2A", "How Thane\u2019s hold is settled", "Paid off \u00b7 talked down \u00b7 fought", ""],
    ["2A", "Calanthe\u2019s fate", "Garrisoned \u00b7 independent \u00b7 hybrid", ""],
    ["2B", "What the ford cost", "A wagon \u00b7 a delay \u00b7 a near-miss \u00b7 an actual loss", ""],
    ["2B", "How Vell\u2019s toll was settled", "Paid \u00b7 talked down \u00b7 bypassed \u00b7 fought", ""],
    ["3", "How the party treated Wyn Alder", "Threatened \u00b7 ignored \u00b7 understood", ""],
    ["3", "Whether the checkpoint turned violent", "Yes (word reaches Caer Ysolde early) \u00b7 No", ""],
    ["4", "How Sennoch Hall was resolved", "Stealth \u00b7 violence \u00b7 a mix", ""],
    ["4", "The party\u2019s standing with Doria Kell", "Face value \u00b7 pressed for more \u00b7 suspicion", ""],
    ["5", "Branch Ledger entry 1, resolved", "The party\u2019s own crossing, set against the second king\u2019s", ""],
    ["5", "How the party engaged with the mourning at the Standing Water", "Spoke at the water \u00b7 listened only \u00b7 kept apart", ""],
    ["6", "How the engines were raised", "Smoothly \u00b7 roughly", ""],
    ["6", "The supply raid\u2019s outcome", "Clean success \u00b7 failed or messy", ""],
    ["6", "How First Assault was fought", "Clear local win \u00b7 costly draw \u00b7 real loss", ""],
    ["7", "How Marshal Drell\u2019s fate was resolved", "Surrendered \u00b7 killed \u00b7 fled", ""],
    ["7", "Scene 3: as written, or a direct dragon fight", "As written \u00b7 direct dragon fight", ""],
    ["8", "How the Magistrate was freed", "Caper \u00b7 force", ""],
    ["8", "Whether the party noticed the pressed clerk", "Yes \u00b7 No", ""],
    ["9", "Which Scene 2 was run", "Tam\u2019s death \u00b7 the third-thread alternate", ""],
    ["9", "General Voss\u2019s fate", "Killed \u00b7 captured \u00b7 fled", ""],
    ["10", "What the party told the Regent, and what the Regent told them", "(Private DM note \u2014 not mechanically tracked)", ""],
    ["11", "The campaign\u2019s final choice", "Held \u00b7 set down", ""],
    ["11", "Vale\u2019s fate", "Killed \u00b7 escaped \u00b7 captured", ""]
  ]
));

c.push(P("The final choice draws on everything above it \u2014 before running Module 11\u2019s Scene 4, a DM may find it worth reading back across this whole ledger with the table."));

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
  fs.writeFileSync(stagePath("KC_DM_Reference_Guide.docx"), buf);
  console.log("Written.");
});
