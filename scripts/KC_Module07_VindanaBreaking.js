// KC_Module07_VindanaBreaking.js -- Session Module Seven: The Siege of Vindana, The Breaking.
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
// This is the campaign's structural centre: the siege breaks, Vale's one
// bound dragon is deployed, and Xavier earns "the Wyvernheart" -- all now
// signed off in CLAUDE.md and drafts/DRACONIC-LAYER.RESOLVED.md.
//
// DESIGN NOTE ON THE DRAGON: it is not statted as a monster the party fights
// directly, on purpose. Canon says Vale does not tame, he coerces -- the
// dragon is a hostage, not a pet -- and a straightforward kill would betray
// that. Xavier's actual heroism in Scene 3 is freeing it, not defeating it;
// the party's own combat in this module is the ground assault through
// Vindana's found weakness, run in parallel to Xavier's airborne arc rather
// than against the dragon directly. A DM whose table wants a direct dragon
// fight instead has an explicit, lightweight opt-in in Scene 3 rather than a
// full stat block -- see the note there.
//
// Xavier is not called "the Wyvernheart" until the exact line marked below.
// No document, and no scene before that line, may use the name.

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
  children: [new TextRun({ text: "The Siege of Vindana: The Breaking", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module Seven", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("The siege turns. The party finds the weakness Vindana\u2019s defense has been hiding since Module Six, leads the assault that exploits it, and \u2014 mid-breach, with the ground fight going their way \u2014 watches it go badly all at once: Marshal Drell, out of better options, unleashes something the coalition has never seen used against it before. What happens in the sky over Vindana in the next several minutes is the reason this king will be called something else for the rest of his life. Core scenes run four to four and a half hours; this module is written to run long and should not be trimmed for time if the table is engaged \u2014 it is the campaign\u2019s centre, and it can afford to take the whole session."));

c.push(H2("Two Fights, One Battle"));

c.push(P("Run this module as two threads happening at once: the party\u2019s own ground assault, which they control directly and which should feel like a real, winnable fight with real stakes, and Xavier\u2019s airborne confrontation, which they witness rather than pilot. Do not hand Xavier\u2019s scene to the party to resolve mechanically \u2014 this is the one moment in the campaign that belongs to him, the way Module One belonged to Brenna Vane\u2019s expertise and Module Four belonged to the Ward\u2019s freedom. The party\u2019s job is to be magnificent in their own fight while his happens over their heads."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 15, 55],
  [
    ["1. Finding the Weakness", "45\u201360 min", "Investigation and payoff from Module Six\u2019s threads."],
    ["2. The Breach", "60\u201375 min", "The party\u2019s own fight. DC table and stat blocks below."],
    ["3. The Dragon", "20\u201330 min", "A hazard, not a monster fight. Read carefully before running."],
    ["4. The Wyvernheart", "20\u201330 min", "Read-aloud. Do not let the party resolve this scene mechanically."],
    ["5. Vindana Falls", "30\u201345 min", "The city secured; the legend begins; hand off to Module Eight."]
  ],
  { full: true }
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("Drell\u2019s position was never actually unbreakable \u2014 Module Six was honest about that \u2014 and the party\u2019s own work in Scenes 1 and 2 is what breaks it. What Drell does in response is not a plan he has held in reserve out of patience; it is the last option of a competent officer who has genuinely run out of better ones. Somewhere above Vindana, under conditions this module deliberately does not over-explain, an occupation officer with the authority to do so releases the one dragon Vale has ever bound. It is not loyal. It is compelled, the same way the wards of Elduvaine were compelled \u2014 by a rule, not by force \u2014 and it is in genuine pain the entire time it is airborne."));

c.push(PS([DM("DM Only: "), { t: "Xavier is not yet the Wyvernheart when this scene begins, and every line of narration before Scene 4\u2019s marked moment must hold that line. When he takes to a wyvern \u2014 grabbing the mount of whichever Harrowmark rider is nearest, not a beast prepared in advance for him \u2014 it should read as improvised and reckless, not as a plan. He does not defeat the dragon. He recognizes what it is: a captive, coerced by the same kind of rule Vale used on Elduvaine itself, and what he actually does in Scene 4 is free it, at real personal risk, rather than kill it. That choice \u2014 mercy over a kill he could plausibly have taken instead \u2014 is the actual content of the legend, not raw combat prowess. Do not let the table anticipate this; play Scenes 1 through 3 as though a straightforward fight were coming." }]));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: Finding the Weakness"));

c.push(P("Whatever the party learned in Module Six \u2014 the Ward\u2019s knowledge of a postern gate, a garrison officer\u2019s captured dispatch, or simply patient scouting since \u2014 now pays off. Vindana\u2019s defense has a real seam, and finding it precisely is this scene\u2019s job."));

c.push(BOX("The postern gate is exactly where the Ward said it would be, half-hidden behind a repair to the sea wall that was never quite finished \u2014 a door built for a peacetime harbor, not a siege, and evidently forgotten by a garrison that has had three years to fortify everything it remembered to fortify."));

c.push(P("If the Ward was rescued in Module Four and her thread was used in Module Six\u2019s Optional Content, this scene is a clean payoff \u2014 let it be found quickly and specifically. If not, run a short investigation (Investigation or Survival, DC 13) locating the same seam through ordinary reconnaissance instead; the weakness is real either way, just differently earned."));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: The Breach"));

c.push(P("The party leads, or joins, a strike through the found weakness while the main siege lines press Vindana\u2019s attention elsewhere. This is a real fight, the party\u2019s own, and should feel decisive and earned."));

c.push(BOX("The postern door gives on the third strike, and for a long moment nothing happens \u2014 then the alarm goes up all at once, and the fight that follows is close, loud, and entirely theirs to win."));

c.push(H3("Running the Scene"));

c.push(P("The garrison that meets them in the breach is the legion doing its job: hobgoblin line troops falling back in order, an orc file-closer swearing at them to keep the order, and not one person on that side of the fight who believes in anything except the wage and the drill. Play it as competence under pressure, not as fury. Use the Occupation Guard stat block (Module 3) for the garrison response \u2014 four to six of them, reinforced by Marshal Drell himself (see Module 6\u2019s stat block) if the table wants a direct confrontation with him. Per his NPC profile, Drell can be talked into surrender once his position is genuinely lost (Persuasion or Intimidation, DC 16, once he is clearly beaten) rather than killed outright \u2014 both resolutions are equally valid endings for him, and both let the module continue identically into Scene 3."));

// -------------------------------------------------------------- Skill DCs
c.push(H2("Tiered Skill DCs"));

c.push(P("Easy 10, Moderate 13, Hard 16, matching the tiers used throughout this campaign."));

c.push(table(
  ["Task", "Skill", "DC", "Tier"],
  [46, 26, 8, 20],
  [
    ["Locate Vindana\u2019s weakness without the Ward\u2019s knowledge", "Investigation / Survival", "13", "Moderate"],
    ["Force the postern door quietly", "Athletics / Thieves\u2019 Tools", "13", "Moderate"],
    ["Talk Marshal Drell into surrender once beaten", "Persuasion / Intimidation", "16", "Hard"],
    ["Keep a routing ally from panicking during Scene 3", "Persuasion / Animal Handling", "13", "Moderate"]
  ],
  { full: true }
));

// ---------------------------------------------------------------- Scene 3
c.push(H2("Scene 3: The Dragon"));

c.push(P("Mid-breach, with the fight turning decisively toward the coalition, everything stops mattering for a moment. Something enormous crosses the sun."));

c.push(BOX("It does not roar. That is the detail that will stay with everyone who is there \u2014 it comes over the walls in near silence, wrongly quiet for something that size, and the sound that finally does come out of it, when the first coalition line breaks and scatters beneath it, is closer to a scream than anything a story would have prepared them for."));

c.push(H3("Running the Scene"));

c.push(P("Play this as a hazard the party survives and responds to, not a monster they fight to the death. Passes overhead force a Dexterity saving throw (DC 15) or the target is knocked prone and takes 11 (2d10) bludgeoning damage from the wind of its wings alone; it does not land near the party or engage them directly \u2014 its attention, and Drell\u2019s remaining forces\u2019 attention, is on the wider battle line. The party\u2019s job here is damage control: steadying routing coalition soldiers (see Tiered Skill DCs), protecting anyone downed by a pass, and \u2014 this is the important part \u2014 noticing Xavier."));

c.push(BOX("Above the chaos, unmistakably, a single wyvern climbs hard toward the thing crossing the sky, and even at this distance there is no doubt at all whose banner-colors are on its rider."));

c.push(PS([DM("DM Only: "), { t: "if your table specifically wants a direct fight with the dragon instead of this hazard framing, the cleanest fix is not to write one into this module \u2014 reskin a young dragon of appropriate challenge from the SRD (checked against the party\u2019s actual level) as an alternate Scene 3, with Xavier\u2019s freeing of it in Scene 4 becoming the thing that ends that fight rather than a separate beat. This is a real departure from how the scene is designed above, and changes the module\u2019s emotional shape from witnessed legend to shared victory \u2014 make it deliberately, not by default." }]));

// ---------------------------------------------------------------- Scene 4
c.push(H2("Scene 4: The Wyvernheart"));

c.push(P("Read this scene aloud, slowly, and do not take mechanical input from the table during it beyond what they are already doing in Scene 3\u2019s hazard below them. This is the campaign\u2019s title-earning moment, and it belongs to Xavier."));

c.push(BOX("He reaches it not with a killing stroke but with both hands empty, which nobody watching from the ground understands until it is already happening \u2014 and there, at the base of one enormous, straining wing, is the thing holding the whole horror together: a mark, a binding, something small enough to miss and exactly the kind of rule Vale has always preferred to a chain. Xavier tears it free with his bare hands, at a height that should kill him for the attempt alone, and for one full second nothing in the sky over Vindana moves at all."));

c.push(BOX("Then the dragon screams again \u2014 differently, this time \u2014 and the battle changes shape beneath it."));

c.push(BOX("It is Ondry\u2019s sergeant, three ranks back and still bleeding from the first pass, who says it first, and it is not a title yet, not really \u2014 just a soldier saying what he saw. \u201CDid you see that? Did anyone else see that? The king \u2014 the king just freed it. He went up there with nothing in his hands and he freed it.\u201D By evening, up and down the whole coalition line, men who were not there are already telling each other that Xavier the Wyvernheart went into the sky over Vindana and came back with a dragon at his back instead of a kill to his name. Nobody corrects them. Nobody will, for the rest of this war."));

c.push(P("The freed dragon does not become an ally in any mechanical sense \u2014 it is not tamed, and this campaign should never suggest it has been. What it does, once freed, is leave the fight: a single pass low over the garrison\u2019s own lines, close enough to break what remains of their nerve, and then it is gone, climbing away from Vindana entirely. That alone is enough. Drell\u2019s remaining defense, already broken by Scene 2 and now watching its last desperate measure fail this completely, does not hold."));

// ---------------------------------------------------------------- Scene 5
c.push(H2("Scene 5: Vindana Falls"));

c.push(P("The city\u2019s defense comes apart within the hour. This scene is aftermath, not more combat \u2014 let the party walk through a city changing hands in real time: garrison soldiers surrendering in ones and twos, Vindana\u2019s own people emerging cautiously from wherever they sheltered, and a coalition that has just watched something it will be telling stories about for the rest of its life."));

c.push(BOX("Xavier finds the party before the day is out, still faintly singed, entirely unbothered by it, exactly as plain-spoken as he was at Duncarrow. \u201CI don\u2019t entirely recommend it,\u201D he says, of the flight, and does not elaborate further, and does not yet seem to have noticed what the men behind him have started calling him."));

c.push(P("Do not have Xavier react to the new name in this scene \u2014 per the DM-Only note above, let it exist in the ranks before it ever reaches him directly; that gap is worth preserving into Module Eight. End the session on Vindana secured and hand off directly to Module Eight, the campaign\u2019s deliberate relief-valve module after its largest set piece."));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("NPC Profiles"));

c.push(H2("Xavier III of Harrowmark, called the Wyvernheart (as of this module)"));
c.push(P("See his existing profile in Module One for baseline speech and bearing, unchanged by this module \u2014 the point of how he is written here is that nothing about his personality changes, only what people now call him. He does not perform the moment, does not narrate his own heroism, and visibly has not yet processed what he did any more than the party has."));
c.push(P("Open thread: how Xavier himself feels about the name \u2014 earned by an act of mercy he may not think of as heroic at all \u2014 is deliberately not resolved in this module. A DM can develop this in any later scene where he and the party have a quiet moment together."));

// --------------------------------------------------------------- Optional
c.push(H1("Optional Content"));

c.push(H2("What Vindana Remembers"));
c.push(P("If the table has time after Vindana falls, let them walk the city and meet its people directly for the first time \u2014 the same drainage-and-endurance texture established in Landfall, now at the scale of a major port rather than a small coastal town. No mechanical stakes; this is pure world texture, and a chance to let Vindana feel like a real place before later modules use it as a staging ground."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("How Marshal Drell\u2019s fate was resolved.", "Surrendered, killed, or fled \u2014 track which. A surrendered Drell is a genuine long-term thread (an occupation officer who chose to stop, worth more alive than as a corpse); a killed or fled Drell closes that thread but changes nothing else about this module\u2019s outcome."));
c.push(BUL("Whether the table ran Scene 3 as written or opted into a direct dragon fight.", "Record which \u2014 it changes the emotional register of everything that follows this module, and a DM should know which version of this legend their table actually witnessed."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BUL("Vindana itself.", "Not loot in the ordinary sense, but the module\u2019s actual prize: a major port, secured, with everything that implies for the rest of the campaign\u2019s logistics and momentum."));
c.push(BUL("Drell\u2019s campaign sword.", "A +1 longsword of legion pattern, plain, superbly maintained, with three years of Vindana service filed into the guard as regulation notches. If Drell surrendered rather than died, he hands it over correctly and asks for a receipt, and means it."));

c.push(BUL("A wand out of the harbour-mage\u2019s quarters.", "A wand of magic missiles (SRD), taken from the rooms of an occupation battle-mage who did not stay for the ending. The first genuinely significant magical reward of the campaign, and the party has earned it."));

c.push(BUL("The garrison\u2019s stores and armory.", "Substantial \u2014 Vindana was well-supplied for a long siege it did not get to fight. The first genuinely significant material reward of the campaign; a DM may introduce one uncommon magic item here without it feeling out of place for the first time."));
c.push(BUL("The binding-mark.", "What Xavier tore free from the dragon, if he keeps it rather than discarding it \u2014 a small, cold object that means nothing to anyone who examines it and everything to Vale, who will know exactly what its absence means. A DM\u2019s hook for a later module, not something that needs to resolve here.", { keepNext: true }));

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
  fs.writeFileSync(stagePath("KC_Module07_VindanaBreaking.docx"), buf);
  console.log("Written.");
});
