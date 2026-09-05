// KC_Module08_HeldGround.js -- Session Module Eight: Held Ground.
//
// Canon lives here. corpus/ and documents/ are generated from this file and are
// never edited by hand. See CLAUDE.md for the sign-off rules: anything in the
// "Not yet decided" table must not appear here until it has been approved.
//
// ESCAPE CONVENTION: all prose lives as \uXXXX escapes, never as literal
// typographic characters. tools/build.sh fails the build if that slips. If
// hand-typing an escape, use ONE backslash -- a doubled backslash compiles
// clean and passes the non-ASCII scanner but leaks literal text into the PDF.
// A literal newline inside a double-quoted JS string is also invalid syntax --
// use two separate BOX() or P() calls instead of embedding a line break.
//
// This module carries no required plot beat, by design -- CLAUDE.md asks for
// levity planned in, not accidental, and this is where that design shows up
// structurally rather than as scattered asides between set pieces. It follows
// the campaign\u2019s largest set piece on purpose. The second rescue thread is
// real and resolves in this module, but it is written to carry the same comic
// register as everything around it rather than reading as inserted homework.
//
// GATE LIFTED: the second captive is Ottoline Vahn, the Magistrate -- gnome,
// past two hundred, a magistrate of the Braid for a hundred and sixty years of
// it, and Ninian\u2019s great-aunt by a marriage nobody now living attended. The
// House of Ysolde is named and peopled in the sourcebook. Use the name.

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

// An ordered sequence -- the phases of a set piece, the movements of a battle -- is a
// list, and printing it as unmarked bold-led prose beside a real bulleted list is
// what made those pages read as two idioms doing one job. ORDERED marks it as what
// it is. A document with a second ordered list passes { instance: 1 }, because one
// numbering reference is one running counter.
const ORDERED = (segs, opts = {}) => new Paragraph({
  numbering: { reference: "steps", level: 0 },
  spacing: { after: 120 },
  ...opts,
  children: segs.map(s => new TextRun({ text: s.t, bold: !!s.b, italics: !!s.i, color: s.c }))
});

const B = (lead, rest, opts = {}) => PS([{ t: lead + " ", b: true }, { t: rest }], opts);
const BUL = (lead, rest, opts = {}) => BULLET(lead ? [{ t: lead + " ", b: true }, { t: rest }] : [{ t: rest }], opts);
const ORD = (lead, rest, opts = {}) => ORDERED(lead ? [{ t: lead + " ", b: true }, { t: rest }] : [{ t: rest }], opts);

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
const SB = (d) => { const out = []; out.push(new Paragraph({ spacing: { before: 240, after: 40 }, keepNext: true, children: [new TextRun({ text: d.name, bold: true, size: 26, color: "5B1F1F" })] })); out.push(PS([{ t: d.meta, i: true }], { spacing: { after: 120 }, keepNext: true })); out.push(B("Armor Class:", d.ac, { keepNext: true })); out.push(B("Hit Points:", d.hp, { keepNext: true })); out.push(B("Speed:", d.speed, { keepNext: true })); out.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ new TableRow({ cantSplit: true, tableHeader: true, children: ["STR","DEX","CON","INT","WIS","CHA"].map(h => abCell(h, true)) }), new TableRow({ cantSplit: true, children: [d.str,d.dex,d.con,d.int,d.wis,d.cha].map(v => abCell(v + " (" + mod(v) + ")")) }) ] })); out.push(P("", { spacing: { after: 60 } })); if (d.saves) out.push(B("Saving Throws:", d.saves)); if (d.skills) out.push(B("Skills:", d.skills)); if (d.senses) out.push(B("Senses:", d.senses)); if (d.langs) out.push(B("Languages:", d.langs)); out.push(B("Challenge:", d.cr)); (d.traits||[]).forEach(t => out.push(PS([{ t: t.n + ". ", b: true, i: true }, { t: t.t }]))); if (d.actions && d.actions.length) { out.push(PS([{ t: "ACTIONS", b: true }], { spacing: { before: 80, after: 80 }, keepNext: true })); d.actions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } if (d.reactions && d.reactions.length) { out.push(PS([{ t: "REACTIONS", b: true }], { spacing: { before: 80, after: 80 }, keepNext: true })); d.reactions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } return out; };


// ---------- content ----------
const c = [];

c.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: "Held Ground", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade \u2014 Module Eight", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Overview
c.push(H1("Overview"));

c.push(P("Vindana is held, the coalition is exhausted in the specific, good way an army gets after a fight it won, and this module\u2019s entire job is to let that feeling breathe before the campaign asks anything hard of the party again. There is no required plot beat here. A second member of the royal house is found and freed along the way, but it happens inside the same relaxed, comic register as everything else in this module rather than as inserted business. Core scenes run three to three and a half hours; this is one of the few modules in the campaign where running short is not a problem \u2014 let the table linger."));

c.push(H2("Levity, Planned In"));

c.push(P("This campaign\u2019s relief valves are designed rather than accidental, and this module is where that design shows up structurally. Play every scene here for warmth first: reunions, absurd logistics, recurring characters getting real room. If a DM feels the urge to raise the stakes, resist it \u2014 Module Nine will do that job. This one\u2019s job is rest."));

c.push(table(
  ["Scene", "Target time", "Notes"],
  [30, 19, 51],
  [
    ["1. The Long Exhale", "45\u201360 min", "Vindana at rest. Recurring characters welcome."],
    ["2. The Grain Ledger War", "30\u201345 min", "Absurd logistics. Comedy, not combat."],
    ["3. The Magistrate", "60\u201375 min", "The second rescue, played as a caper. DC table below."],
    ["4. What the Name Costs", "20\u201330 min", "A quiet scene with Xavier. Closes the module."],
    ["Optional Content", "30\u201345 min", "Run if the table has time; cut cleanly if not."]
  ]
));

// --------------------------------------------- What Is Actually Happening
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("Nothing in this module is a trap, a test, or a delayed complication. Vindana is genuinely secure, the coalition is genuinely at ease for the first time since Module One, and the Magistrate\u2019s rescue in Scene 3 is genuinely as easy as it plays \u2014 the occupation\u2019s administration in a freshly-fallen city is in real disarray, and a party willing to use its own paperwork against it will find that disarray works entirely in their favor. If a DM feels this module is too easy, that is correct and intentional."));

c.push(PS([DM("DM Only: "), { t: "Xavier has still not directly addressed being called the Wyvernheart as of this module\u2019s start \u2014 the gap held open in Module Seven continues into Scene 4, which is the one place this module allows real feeling. Do not let the earlier scenes anticipate it." }]));

// ---------------------------------------------------------------- Scene 1
c.push(H2("Scene 1: The Long Exhale"));

c.push(P("Vindana\u2019s taken harbor district has become, within days, the loudest and most cheerful place the coalition has occupied since Harrowmark. Let the party spend real time here."));

c.push(BOX("Someone has chalked a running tally on the side of a captured warehouse \u2014 days since the siege, casks emptied, and, unofficially, a tally of how many times a passing soldier has tried and failed to get a straight answer out of Huntmaster Brenna Vane about whether she actually rode south for the war or just to see the fuss for herself."));

c.push(P("If Module One was played, Brenna Vane has arrived with a handful of Greywatch hands, drawn south by rumor of the Wyvernheart with an expression that suggests she will never admit that is the reason. If Module 2A or 2B was played, Sera Vosk or Garrick Hollow can likewise turn up here, already embedded in the coalition\u2019s logistics or scouting corps. Use whichever recurring faces the table has earned; this scene\u2019s only job is reunion and warmth."));

// ---------------------------------------------------------------- Scene 2
c.push(H2("Scene 2: The Grain Ledger War"));

c.push(P("Vindana\u2019s occupation-era grain levy records \u2014 thousands of receipts, quotas, and permits, exactly the bureaucracy Module Three showed the party in miniature \u2014 have fallen into coalition hands whole, and nobody sent to sort them was remotely prepared for the scale of it."));

c.push(BOX("Three different coalition quartermasters are shouting at each other over a single warehouse of requisitioned grain, each one certain their own contingent\u2019s paperwork takes precedence, while an increasingly desperate Elduvish clerk \u2014 pressed into the coalition\u2019s service rather than freed from it, which nobody seems to have noticed yet \u2014 tries to explain that the numbers do not actually add up for anyone."));

c.push(P("Let the party resolve this however they find funniest: mediate, cut through it with authority, or simply steal a wagon of grain themselves while the argument continues. No mechanical stakes and no wrong answer. If a perceptive party notices the pressed clerk\u2019s situation and does something about it \u2014 a small kindness, not a rescue \u2014 let that pay forward into how Elduvaine\u2019s ordinary people regard the coalition later."));

// ---------------------------------------------------------------- Scene 3
c.push(H2("Scene 3: The Magistrate"));

c.push(P("Among the captured administration\u2019s own records is a name the party will recognise if they have been paying attention, and it is not filed under prisoners. Ottoline Vahn \u2014 magistrate of the Braid for a hundred and sixty years, the Ward\u2019s great-aunt by a marriage nobody now living attended, and a gnome somewhere past two hundred \u2014 has been held under house arrest in a Vindana townhouse she insists on calling her chambers, kept comfortable and thoroughly bored by an occupation that was too disorganised before the siege to know what to do with her and is considerably more disorganised now."));

c.push(BOX("She receives the party at a desk, in what is unmistakably an office, with the weary patience of somebody who has spent three years winning arguments against clerks a tenth her age. \u201CYou\u2019ll want papers,\u201D she says, before anyone has explained themselves. \u201CEverybody wants papers. Fortunately I have spent three years becoming extremely good at papers. Sit down. Do not move the third pile.\u201D"));

c.push(H3("Running the Scene"));

c.push(P("This is a caper, not a fight. The Magistrate\u2019s household is guarded by a skeleton crew of occupation clerks and one or two soldiers, none of them expecting trouble in a freshly-fallen city, and she herself is entirely capable of walking out the front door if the party can produce (or forge) plausible transfer papers \u2014 Vindana\u2019s administrative chaos, established in Scene 2, means nobody is checking closely. A DC 13 Deception or Forgery-adjacent check (use whatever tool proficiency a character has) produces convincing papers; a DC 10 Persuasion check talks a guard into not looking too hard at them regardless."));

c.push(P("If the party would rather simply walk in and take her by force, that works too and is considerably less interesting \u2014 a DM should let the Magistrate herself gently steer the party toward the caper option if asked, since she has clearly been planning her own exit for some time and has opinions about the tidiest way to manage it."));

// -------------------------------------------------------------- Skill DCs
c.push(H2("Tiered Skill DCs"));

c.push(P("Easy 10, Moderate 13, Hard 16, matching the tiers used throughout this campaign."));

c.push(table(
  ["Task", "Skill", "DC", "Tier"],
  [44, 26, 10, 20],
  [
    ["Resolve the quartermasters\u2019 dispute without violence", "Persuasion / Insight", "10", "Easy"],
    ["Forge or produce convincing transfer papers", "Deception / an appropriate tool proficiency", "13", "Moderate"],
    ["Talk a guard out of scrutinizing the papers", "Persuasion", "10", "Easy"],
    ["Notice the pressed Elduvish clerk\u2019s actual situation", "Insight / Perception", "13", "Moderate"]
  ]
));

// ---------------------------------------------------------------- Scene 4
c.push(H2("Scene 4: What the Name Costs"));

c.push(P("Late, with the day\u2019s business done, Xavier finds the party somewhere quiet \u2014 away from the harbor\u2019s noise, away from anyone keeping official record of the conversation."));

c.push(BOX("\u201CThey\u2019ve started calling me something,\u201D he says, not quite a question. \u201CI keep waiting for someone to explain it to me properly and nobody will. I remember the dragon. I remember being fairly sure I was about to die. I don\u2019t remember deciding to be brave about it \u2014 I remember being too frightened to think of anything else to do.\u201D He is quiet for a moment. \u201CIs that what the songs are going to say happened? Because I\u2019d rather they didn\u2019t, if it\u2019s all the same.\u201D"));

c.push(P("Let this be a real conversation rather than a scene to resolve. Xavier is not fishing for reassurance and does not need the party to tell him he is a hero; he genuinely does not know yet how he feels about the name, and the party\u2019s honest reaction \u2014 whatever it is \u2014 matters more here than any mechanical outcome. This scene needs no check and produces no loot. It is the module\u2019s actual point, arriving last, exactly where a session built on deliberate levity has been heading since Scene 1."));

// ------------------------------------------------------------ NPC Profiles
c.push(H1("Puzzles and Set Pieces"));

c.push(P("Three Years of Filing expands the ledger war and adds twenty minutes. Putting the Light Back is additive, takes ten, and should not be cut \u2014 it is the beat the whole module is built around."));

c.push(H2("The Puzzle: Three Years of Filing"));

c.push(P("Ottoline Vahn fought her captivity by filing, and won, and the consequence is that the occupation of Vindana produced the most complete administrative record in Elduvaine and then had to rely on it. The party needs one document out of it. She will help, at her own pace, and her pace is the puzzle."));

c.push(BOX("\u201CYou want the carriage authorisations. Everybody wants the carriage authorisations. They are not filed under carriage, they are not filed under authorisation, and they are not filed under Norvatch, and if you tell me why you want them I will tell you where they are, and if you do not, we can spend a pleasant afternoon and you can find them yourself.\u201D"));

c.push(P("The registry is filed by the issuing office, not by subject, which is standard Elduvish practice and is baffling to everyone else. Four offices issued paper in Vindana and the party can work out which from any docket they already hold:"));

c.push(table(
  ["Office", "Issued", "Docket mark"],
  [26, 44, 30],
  [
    ["The Levy", "Grain, fodder, and the published quarterly rate", "A single stroke, top right"],
    ["The Harbour", "Everything that moved by water, in or out", "Two strokes and a date"],
    ["The Garrison", "Permits for persons, and passes", "A stamp, always smudged"],
    ["The Keeper\u2019s Office", "Anything touching the Archive or the quarries", "No mark at all, which is itself the mark"]
  ]
));

c.push(P("The carriage authorisations went out by water, so they are Harbour, and they are filed by date, and the date the party wants is the week Norvatch\u2019s volume tripled \u2014 which they can get from Morgarth\u2019s harbourmaster, or from the counting-house manifests in this module, or by asking Ottoline the question she is waiting to be asked."));

c.push(PS([DM("DM Only: "), { t: "the unmarked fourth office is the real find and the party should stumble on it while looking for something else. Anything issued by the Keeper\u2019s Office carries no docket mark, because the Keeper\u2019s Office was never part of the civil registry and never needed one \u2014 which means every quarry order, every Archive requisition and every light-stone consignment for three years is sitting in a drawer that nobody has thought to look in, unmarked, because it always was. Ottoline knows. She has been waiting three years for somebody to notice, and she will not point at it, because pointing at it is not how a magistrate establishes anything." }]));

c.push(H2("Set Piece: Putting the Light Back"));

c.push(P("The city has been dark for fourteen months and the coalition holds it, and on the fourth night somebody works out that the wall will still take a charge."));

c.push(BOX("\u201CIt is a gnome from Cairn Ithel and two Auberitz sappers and a Concord priest of Aurine who has no business being on that scaffold, and between them they pour about four hours of lamplight into eleven feet of the inner wall. It holds. It comes up slow, the colour of late afternoon, and it spreads maybe thirty feet along the course and stops. Eleven feet of a city that used to do this from the ground up on every clear night. Somebody on the harbour steps starts crying and does not stop and nobody says anything about it.\u201D"));

c.push(P("It is not a victory and everyone present knows the arithmetic: four hours of lamp for eleven feet of wall, against a city that held six centuries of afternoons for free. The stone is not repaired. The stone is being manually filled, one small stretch at a time, by people who will be dead long before the second street is done."));

c.push(P("They do it anyway. They do it every night for the rest of the occupation of the city, and by the time the party leaves Vindana about two hundred feet of the inner wall is lit, and the party will be able to see it from the road."));

c.push(PS([DM("DM Only: "), { t: "this is the levity module and this is not a levity scene, and it should sit in the middle of an evening that is otherwise warm and loud and full of terrible singing. Do not build to it. Let somebody mention it in passing, let the party wander over, and let them stand there. Then go back to the singing. The contrast is the whole design of the module and this is the beat it is built around." }]));

c.push(H1("NPC Profiles"));

c.push(H2("Ottoline Vahn, the Magistrate"));
c.push(P("A gnome somewhere past two hundred, dry as a ledger, and entirely unbroken by three years of house arrest \u2014 if anything sharpened by it. Ninian\u2019s great-aunt by a marriage nobody now living attended, which is the sort of thing old Elduvish houses are full of. Speech: precise, faintly amused, allergic to being underestimated by people who had every opportunity to know better. She fought her captivity with the only weapon she has ever needed and won: three years of correctly formatted objections, appeals and requests for clarification, every one of them lawful under the occupation\u2019s own published code, which have cost the administration of Vindana an estimated four hundred clerk-days and produced a written record of the occupation so complete that the occupation came to rely on it."));
c.push(P("Open thread: freed, she is a formidable and very funny recurring NPC \u2014 a DM can use her as a source of administrative chaos turned against the occupation in any later module, or simply as comic relief who happens to also be dangerous in exactly the way nobody expects from an old woman with a ledger."));

// --------------------------------------------------------------- Optional
c.push(H1("Optional Content"));

c.push(H2("Downtime, Properly, for the First Time"));

c.push(P("This is the first stretch of the campaign long enough and safe enough for downtime, and the sourcebook\u2019s six options under Downtime on a March are all available in a held city. Each is meant to produce a scene rather than a die roll, and a DM running this module should offer them by name rather than waiting to be asked."));

c.push(P("Stand a watch \u2014 four hours on a picket line with one other person, chosen by the player, and the DM asks what they talked about. It is the single most productive downtime action in the campaign and it costs nothing. Work the column, and earn the goodwill of Auberitz\u2019s people, which converts directly into supply and being told things early. Drink with the levy, for one rumour that is true and one that is not, with no indication which. Keep the observance, at a Ninefold House or a wayside shrine, and be seen doing it, which in this coalition is a political act whichever one is chosen. Train with the watch, if Brenna Vane is here, and take the advantage it grants against the next large flying thing. Or copy, if anybody has Archive training, at about a page a week of something Vale would rather nobody had."));

c.push(PS([DM("DM Only: "), { t: "keep the observance is the one worth pushing in this module specifically. Vindana is a freshly-taken Elduvish city with a Concord army sitting in it and a Tenth Work chapter house being framed in a street that had an Observance on it last year. A player character who chooses one of those two on a given evening has made a statement, and somebody will comment, and it should cost nothing and be remembered." }]));

c.push(H2("The Counting-House"));

c.push(P("Vindana is a port, and Norvatch trades through ports. Three streets back from the harbour, a Norvatch counting-house has been open through the entire occupation and is open now, under the coalition, with the shutters down and a clerk at the desk and absolutely nothing to apologise for. This is the one cold scene in an otherwise warm module, and the module can carry it."));

c.push(BOX("The factor does not stand up. \u201CWe are open,\u201D he says, in the tone of a man confirming the weather. \u201CWe were open last month as well. If your king would like to discuss terms, my principals will hear him \u2014 and if he would rather not, the arrangement we have runs another nine months and I would ask you not to interfere with lawful carriage.\u201D"));

c.push(P("Everything he says is true and none of it is deniable. Norvatch has broken no agreement with the coalition because it never made one. The goods moving through that house are Elduvaine, by weight, sold under a contract with the occupation that predates the crusade\u2019s arrival by three years. A party that wants to seize the ledgers by force can \u2014 and will find them impeccably kept, entirely legal, and worth rather less taken than bought, because what they actually need is the house\u2019s cooperation in reading them."));

c.push(PS([DM("DM Only: "), { t: "the ledgers are the campaign\u2019s clock, and this is where the party first sees the physical object. Do not let them buy it here \u2014 the factor is not authorised and says so. What he will do is send word to Doria Kell, which is what sets up Module Ten. If the party takes them by force anyway, they get numbers they cannot interpret and a Norvatch house that will deal with them thereafter strictly, correctly, and never generously again. That is a real consequence and not a punishment; play it exactly that flat." }]));

c.push(H2("The Betting Ledger Returns"));
c.push(P("If Module One was played and Brenna Vane is present per Scene 1, her old betting ledger from Greywatch makes a reappearance, now taking wagers on considerably higher-stakes nonsense \u2014 how long the peace will hold, whether the Magistrate and the Ward will get along, anything the table finds funny. Pure levity, no stakes."));

c.push(H2("A Quiet Word With the Ward"));
c.push(P("If the Ward was rescued in Module Four, let her and the Magistrate meet for the first time here \u2014 two members of a divided royal house with different ideas about what should follow the war, meeting in a moment that has no pressure on it yet. Play their first exchange for warmth and a little friction, not conflict; save the real disagreement for later in the campaign, per the royal house\u2019s established canon division on the question."));

// -------------------------------------------------------------- Diverging
c.push(H1("Diverging Paths (DM Only)"));

c.push(BUL("How the Magistrate was freed.", "Caper or force \u2014 track which. A caper leaves the occupation\u2019s administration none the wiser for a while longer, which a DM can use later; a forced rescue is noisier and gives the coalition\u2019s presence in Vindana a slightly harder edge sooner."));
c.push(BUL("Whether the party noticed the pressed clerk in Scene 2.", "A small thread, but worth tracking \u2014 it is the campaign\u2019s recurring question about the occupation\u2019s ordinary people, in miniature, again."));

// ---------------------------------------------------------------- Loot
c.push(H1("Loot"));

c.push(BUL("The Magistrate herself.", "As with the Ward in Module Four, the module\u2019s actual reward \u2014 a second freed royal, formidable in her own register, and a second real voice in whatever the royal house eventually decides about Elduvaine\u2019s future."));
c.push(BUL("A warehouse\u2019s worth of grain.", "However the Grain Ledger War was resolved, the coalition ends this module better supplied than it started it \u2014 modest, practical, not a coin windfall.", { keepNext: true }));

// -------------------------------------------------------------- Refrain
c.push(H1("The Refrain"));

c.push(VERSE([
  "By thought, and by word, and by deed,",
  "the king\u2019s own chosen kept their creed.",
  "Far from home, where the quiet land lay,",
  "they held the line, and would not stray."
]));

const doc = new Document({
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 260, hanging: 260 } } } }] },
    // The same measure as the bullets, deliberately: a numbered list and a bulleted
    // one appear on the same page often enough that their text has to hang off one
    // left edge, and "1." is near enough the width of a dot for the gap to match. It
    // holds to "99."; nothing in this campaign counts past five.
    { reference: "steps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 260, hanging: 260 } } } }] }
  ] },
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
  fs.writeFileSync(stagePath("KC_Module08_HeldGround.docx"), buf);
  console.log("Written.");
});
