// smoke.js -- THROWAWAY. Delete once the real generators build clean.
//
// A one-page document that exercises every rendering path the pipeline supports:
// headings, body prose, the bold book-red DM marker, bullets, a table, boxed
// read-aloud, and a full stat block. Its only job is to prove the chain works --
// docx-js -> transplant.py -> LibreOffice -> Ghostscript -> normalize_pdf.py --
// before a word of canon is written. Debugging fonts and PDF reproducibility is
// far easier on a page nobody cares about.
//
// Run: tools/build.sh   (GENERATORS=(smoke) in tools/build.sh)
// Expect: corpus/Smoke_Test.md, documents/Smoke_Test.pdf, "verified clean",
//         and a byte-identical file on a second build.
//
// NOTE THE ESCAPE CONVENTION: all prose lives here as \uXXXX escapes, never as
// literal typographic characters. Verify with:
//     grep -Pc '[^\x00-\x7F]' scripts/*.js        # MUST be 0
//     pdftotext documents/Smoke_Test.pdf - | grep -c '\\u'   # MUST be 0
// The doubled backslash in the second grep is load-bearing.

const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, LevelFormat } = require('docx');
const fs = require('fs');

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

const BULLET = (segs) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 120 },
  children: segs.map(s => new TextRun({ text: s.t, bold: !!s.b, italics: !!s.i, color: s.c }))
});

const B = (lead, rest) => PS([{ t: lead + " ", b: true }, { t: rest }]);
const BUL = (lead, rest) => BULLET(lead ? [{ t: lead + " ", b: true }, { t: rest }] : [{ t: rest }]);

// boxed read-aloud text
const BOX = (text) => new Paragraph({
  spacing: { before: 120, after: 160 },
  shading: { type: "clear", fill: "F3EFE4" },
  indent: { left: 220, right: 220 },
  children: [new TextRun({ text, italics: true })]
});

const { Table, TableRow, TableCell, WidthType, ShadingType } = require('docx');
const cell = (text, opts = {}) => new TableCell({ width: { size: opts.w || 20, type: WidthType.PERCENTAGE }, shading: opts.head ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, margins: { top: 50, bottom: 50, left: 45, right: 45 }, children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text, bold: !!opts.head, size: 18 })] })] });
const row = (cells) => new TableRow({ children: cells });
const table = (headers, widths, rows) => new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ row(headers.map((h, i) => cell(h, { head: true, w: widths[i] }))), ...rows.map(r => row(r.map((v, i) => cell(v, { w: widths[i] })))) ] });

const mod = (v) => { const m = Math.floor((v - 10) / 2); return (m >= 0 ? "+" : "\u2212") + Math.abs(m); };
const abCell = (text, bold) => new TableCell({ width: { size: 16.6, type: WidthType.PERCENTAGE }, shading: bold ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40, before: 40 }, children: [new TextRun({ text, bold: !!bold, size: 20 })] })] });
const SB = (d) => { const out = []; out.push(new Paragraph({ spacing: { before: 240, after: 40 }, children: [new TextRun({ text: d.name, bold: true, size: 26, color: "5B1F1F" })] })); out.push(PS([{ t: d.meta, i: true }], { spacing: { after: 120 } })); out.push(B("Armor Class:", d.ac)); out.push(B("Hit Points:", d.hp)); out.push(B("Speed:", d.speed)); out.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ new TableRow({ children: ["STR","DEX","CON","INT","WIS","CHA"].map(h => abCell(h, true)) }), new TableRow({ children: [d.str,d.dex,d.con,d.int,d.wis,d.cha].map(v => abCell(v + " (" + mod(v) + ")")) }) ] })); out.push(P("", { spacing: { after: 60 } })); if (d.saves) out.push(B("Saving Throws:", d.saves)); if (d.skills) out.push(B("Skills:", d.skills)); if (d.senses) out.push(B("Senses:", d.senses)); if (d.langs) out.push(B("Languages:", d.langs)); out.push(B("Challenge:", d.cr)); (d.traits||[]).forEach(t => out.push(PS([{ t: t.n + ". ", b: true, i: true }, { t: t.t }]))); if (d.actions && d.actions.length) { out.push(PS([{ t: "ACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.actions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } return out; };

// ---------- content ----------
const c = [];

c.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: "Pipeline Smoke Test", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "Delete this document, and scripts/smoke.js, once the real generators build clean.", i: true }],
  { alignment: AlignmentType.CENTER }));

c.push(H1("What This Page Proves"));
c.push(P("If this document rendered, the whole chain works: docx-js produced a .docx, transplant.py applied the visual template, LibreOffice converted it to PDF, Ghostscript rewrote it, and normalize_pdf.py stripped the per-run randomness. If tools/build.sh also reported \u201call documents verified clean,\u201d then no escape sequences leaked into the output and no template font was silently substituted."));
c.push(P("Build it twice and compare the bytes. They must be identical \u2014 that is what makes the git history stay clean when a document has not actually changed."));

c.push(H2("Every Rendering Path, Exercised"));
c.push(BUL("Bullets:", "this is one, produced by the BUL helper."));
c.push(BUL("Bold leads:", "the lead is bold, the rest is roman, and the helper puts the space in the right place."));
c.push(BUL(null, "A bullet with no lead, for continuation points."));

c.push(H3("The DM Marker Convention"));
c.push(PS([DM("DM Only: "), { t: "the marker is bold and book-red; the prose after it stays roman. Colour is preattentive \u2014 a DM spots red without reading \u2014 and leaving the body roman matters because these notes run long. Colour the marker, not the prose. Italic is reserved for read-aloud, quotations, and epigraphs, and overloading it makes both signals ambiguous. Note also that the marker segment carries its own trailing space and the segment after it never begins with one, because the Markdown shim appends a space after every bold run." }]));

c.push(H2("Read-Aloud"));
c.push(BOX("Boxed text renders in italic on a tinted ground, indented from both margins. This is where the players hear the scene, and it is the one place italic is doing its intended job."));

c.push(H2("A Table"));
c.push(table(["Check", "What it catches", "Must be"], [26, 52, 22], [
  ["grep -Pc non-ASCII", "Literal typographic characters in the generator source", "0"],
  ["pdftotext | grep -c escape", "Escape sequences leaking into the rendered output", "0"],
  ["pdffonts | grep -c DejaVu", "A template font missing and silently substituted", "0"],
  ["Second build, cmp", "Per-run randomness the normalizer failed to strip", "identical"]
]));

c.push(H2("A Stat Block"));
c.push(...SB({
  name: "Test Construct",
  meta: "Medium construct, unaligned \u2014 exists only to prove the SB helper renders",
  ac: "15 (plated)", hp: "27 (5d8 + 5)", speed: "30 ft.",
  str: 14, dex: 12, con: 13, int: 3, wis: 10, cha: 1,
  senses: "blindsight 30 ft., passive Perception 10",
  langs: "understands its instructions; cannot speak",
  cr: "1 (200 XP)",
  traits: [{ n: "Immutable Form", t: "The construct is immune to any spell or effect that would alter its form." }],
  actions: [{ n: "Slam", t: "Melee Weapon Attack: +4 to hit, reach 5 ft., one target. Hit: 6 (1d8 + 2) bludgeoning damage." }]
}));

c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300, after: 100 }, children: [new TextRun({ text: "~", size: 24 })] }));
c.push(PS([{ t: "Pipeline proved. Now go and decide what the campaign is about.", i: true }], { alignment: AlignmentType.CENTER }));

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
  fs.writeFileSync("/home/claude/Smoke_Test.docx", buf);
  console.log("Written.");
});
