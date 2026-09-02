// KC_Character_Options.js -- backgrounds, feats, subclasses, spells and items.
//
// 2014 rules, SRD 5.1. Race not species; no weapon masteries, Bastions, Epic
// Boons or Heroic Inspiration; 2014 exhaustion and grapple. See CLAUDE.md.
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

const { Table, TableRow, TableCell, WidthType, ShadingType } = require('docx');
const cell = (text, opts = {}) => new TableCell({ width: { size: opts.w || 20, type: WidthType.PERCENTAGE }, shading: opts.head ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, margins: { top: 60, bottom: 60, left: 110, right: 110 }, children: [new Paragraph({ spacing: { after: 0 }, indent: { firstLine: 0 }, children: [new TextRun({ text, bold: !!opts.head, size: 18 })] })] });
// cantSplit keeps a row's cells from being torn across a column or page break;
// tableHeader repeats the header row when a long table does span a break.
const row = (cells, opts = {}) => new TableRow({ children: cells, cantSplit: true, ...opts });
const FULLWIDTH = "KCFullWidth";   // marker only; transplant.py acts on it and strips it
const table = (headers, widths, rows, opts = {}) => new Table({ ...(opts.full ? { style: FULLWIDTH } : {}), width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ row(headers.map((h, i) => cell(h, { head: true, w: widths[i] })), { tableHeader: true }), ...rows.map(r => row(r.map((v, i) => cell(v, { w: widths[i] })))) ] });

const mod = (v) => { const m = Math.floor((v - 10) / 2); return (m >= 0 ? "+" : "\u2212") + Math.abs(m); };
const abCell = (text, bold) => new TableCell({ width: { size: 16.6, type: WidthType.PERCENTAGE }, shading: bold ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40, before: 40 }, indent: { firstLine: 0 }, keepNext: !!bold, children: [new TextRun({ text, bold: !!bold, size: 20 })] })] });
const SB = (d) => { const out = []; out.push(new Paragraph({ spacing: { before: 240, after: 40 }, children: [new TextRun({ text: d.name, bold: true, size: 26, color: "5B1F1F" })] })); out.push(PS([{ t: d.meta, i: true }], { spacing: { after: 120 } })); out.push(B("Armor Class:", d.ac)); out.push(B("Hit Points:", d.hp)); out.push(B("Speed:", d.speed)); out.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ new TableRow({ cantSplit: true, tableHeader: true, children: ["STR","DEX","CON","INT","WIS","CHA"].map(h => abCell(h, true)) }), new TableRow({ cantSplit: true, children: [d.str,d.dex,d.con,d.int,d.wis,d.cha].map(v => abCell(v + " (" + mod(v) + ")")) }) ] })); out.push(P("", { spacing: { after: 60 } })); if (d.saves) out.push(B("Saving Throws:", d.saves)); if (d.skills) out.push(B("Skills:", d.skills)); if (d.senses) out.push(B("Senses:", d.senses)); if (d.langs) out.push(B("Languages:", d.langs)); out.push(B("Challenge:", d.cr)); (d.traits||[]).forEach(t => out.push(PS([{ t: t.n + ". ", b: true, i: true }, { t: t.t }]))); if (d.actions && d.actions.length) { out.push(PS([{ t: "ACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.actions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } if (d.reactions && d.reactions.length) { out.push(PS([{ t: "REACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.reactions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } return out; };

// ---------- content ----------
const c = [];

c.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [new TextRun({ text: "The King\u2019s Crusade", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "Character Options for the Crusade and the Living Realm", i: true }],
  { alignment: AlignmentType.CENTER }));

c.push(H1("Using These Options"));

c.push(P("Everything in this book is written for the 2014 rules, SRD 5.1. That is a deliberate choice for this campaign and not a default: race rather than species, no weapon masteries, the 2014 exhaustion track, and monster maths from the 2014 DMG. If your table plays the 2024 rules, most of what follows will convert, and none of it has been tested that way."));

c.push(P("The campaign starts at 5th level, so a background's feature matters less than usual and its bonds matter more. Take one anyway. Every background here is a claim about where you stood on the night the wards opened, and the campaign will ask about it."));

c.push(P("Nothing in this chapter is required. A party built entirely from the SRD will play this campaign perfectly well, and the setting guidance in the Player Guide covers how to reskin standard options into Elduvaine and Harrowmark without touching a single mechanic."));

c.push(H1("Backgrounds"));

c.push(H2("Wyvern-Watch"));

c.push(P("You stood on a pike line eleven hundred feet up, or you went over the edge on a rope, and you did it more than once. Harrowmark has kept the watch for nine centuries and it has never had to conscript for it, which outsiders find baffling and nobody from Greywatch finds worth explaining."));

c.push(B("Skill Proficiencies:", "Animal Handling, Athletics"));
c.push(B("Tool Proficiencies:", "One type of artisan's tools of your choice, usually rope-work or leatherworking"));
c.push(B("Equipment:", "A wyvern pike or a fifty-foot coil of hempen rope, a set of climber's tools, a hooked knife, traveller's clothes, and a belt pouch containing 10 gp"));

c.push(PS([{ t: "Feature: Names on the Wall. ", b: true }, { t: "Every wyvern-watch in Harrowmark keeps a roster of its dead and does not cross them off. You can read those rosters, and any watch in the kingdom will house, feed and equip you on the strength of yours \u2014 and will expect you to take a shift while you are there. Watches talk to each other. Word of what you did at one will reach the next before you do." }]));

c.push(P("The wyvern-watch also produces a specific personality, which is not courage. It is the flat refusal to be impressed by anything, up to and including a dragon, and the campaign will give you several opportunities to demonstrate it."));

c.push(table(
  ["d6", "Bond"],
  [8, 92],
  [
    ["1", "Somebody went over the edge on my rope. I was holding it. I would like to be sure that never happens again."],
    ["2", "My name is already on a wall. There was a mistake about a body. I have never asked them to correct it."],
    ["3", "I taught a dozen people to hold a line and eleven of them are still alive."],
    ["4", "The huntmaster vouched for me when nobody else would. That debt is not settled and will not be."],
    ["5", "I have killed nine wyverns and I remember all of them, and I did not enjoy any of it."],
    ["6", "There is a village that has livestock because of me. They do not know my name. That is correct."]
  ], { full: true }
));

c.push(H2("Archive Clerk"));

c.push(P("You worked in the largest collection of magical knowledge in the world, at a desk, fetching what the rule permitted for people who were entitled to it. It was the best job in Elduvaine and everybody who had it knew that."));

c.push(B("Skill Proficiencies:", "Arcana, History"));
c.push(B("Languages:", "Two of your choice, one of which is Elduvish or Ninefold Cant"));
c.push(B("Equipment:", "A set of calligrapher's supplies, a requisition seal that is no longer valid, a book of hands and ciphers, common clothes, and a belt pouch containing 15 gp"));

c.push(PS([{ t: "Feature: Access by Rule. ", b: true }, { t: "You know how the Archive is organised, which is a rarer skill than it sounds and does not stop being useful because the Archive is in enemy hands. Given an hour in any substantial collection of books or records, you can determine whether it contains what you are looking for and roughly where, without reading it. You also know, precisely, what a person of any given station was and was not permitted to read \u2014 and you know that the Keeper was permitted less than most people assume." }]));

c.push(table(
  ["d6", "Bond"],
  [8, 92],
  [
    ["1", "I filed requisitions for Maedoc Vale for four years. He was unfailingly courteous and I liked him."],
    ["2", "There is a shelf I was never cleared for and I have thought about it every day for three years."],
    ["3", "I got eleven volumes out on the second night. They are buried and I am the only one who knows where."],
    ["4", "The clerk at the next desk stayed on and is being paid. I have not decided what I think about that."],
    ["5", "Somebody has to be able to put it back in order afterward. It might as well be me."],
    ["6", "I refused an order to burn a catalogue. It is the only brave thing I have ever done."]
  ], { full: true }
));

c.push(H2("Norvatch Factor"));

c.push(P("You were trained to draft, to read, and above all to notice \u2014 in a realm whose entire standing in the world rests on the proposition that its word, once written, is kept to the letter no matter what the letter turns out to have meant."));

c.push(B("Skill Proficiencies:", "Insight, Persuasion"));
c.push(B("Tool Proficiencies:", "One gaming set"));
c.push(B("Languages:", "Writ-tongue"));
c.push(B("Equipment:", "A ledger, a sealed and current letter of credit, fine clothes that have been carefully kept, a set of scales, and a belt pouch containing 25 gp"));

c.push(PS([{ t: "Feature: The House Stands Behind It. ", b: true }, { t: "Any Norvatch counting-house will honour your letter of credit, extend you reasonable hospitality, and forward a message anywhere Norvatch trades, which is everywhere. In return the house expects an accurate account of what you have seen, and it will notice if you shade one. You may also draft an agreement that will hold up in the Writ House, which in four kingdoms means it will simply hold up." }]));

c.push(table(
  ["d6", "Bond"],
  [8, 92],
  [
    ["1", "I signed something three years ago that is still being honoured and should not be."],
    ["2", "The house made me. Everything I am is theirs and I have never once resented it, which worries me."],
    ["3", "Doria Kell taught me to read a manifest. I would like her to be wrong about something, once."],
    ["4", "I have never broken a written word and I intend to die with that intact."],
    ["5", "A contract I drafted killed people. Every clause in it was correct."],
    ["6", "Somebody is going to have to write the settlement at the end of this. It should be somebody competent."]
  ], { full: true }
));

c.push(H2("Season-Keeper"));

c.push(P("You tended a wood that holds the season it was planted in. It is not farming and it is not druidry; it is a four-hundred-year institution with an apprenticeship, a body of practice, and extremely strong opinions about where the next planting goes."));

c.push(B("Skill Proficiencies:", "Nature, Survival"));
c.push(B("Tool Proficiencies:", "Herbalism kit"));
c.push(B("Languages:", "Sylvan"));
c.push(B("Equipment:", "A herbalism kit, a pruning hook, a sealed packet of seeds from a wood you tended, traveller's clothes, and a belt pouch containing 10 gp"));

c.push(PS([{ t: "Feature: Read the Wood. ", b: true }, { t: "You can tell at a glance what season a stand of trees is holding, how long it has held it, and \u2014 the part that matters in this campaign \u2014 whether it is failing and roughly how fast. Any Elduvish rural community will shelter you on that skill alone. Fey in Elduvaine will generally hear you out before deciding anything, which is not the same as helping and is a great deal better than nothing." }]));

c.push(table(
  ["d6", "Bond"],
  [8, 92],
  [
    ["1", "My wood turned last spring. I was there. I could not do anything and I did not leave."],
    ["2", "The dryad of my orchard knew my name and used it."],
    ["3", "My grandmother chose where the last planting went. It will still be standing in three hundred years, or it will not, and that is now a real question."],
    ["4", "I have seeds in my pack that have not been sown. I am waiting to know they will hold."],
    ["5", "The Keepers went on working for the man killing the orchards. I was one of them."],
    ["6", "Somebody married under my wood every spring for sixty years."]
  ], { full: true }
));

c.push(H2("Crusade Levy"));

c.push(P("You are one of the eight thousand. You answered a summons read in a language you do not speak, promising a place in a country you could not find on a map, and you are walking to it. This is the most common background in the entire coalition and the campaign takes it seriously."));

c.push(B("Skill Proficiencies:", "Athletics, and either Insight or Survival"));
c.push(B("Tool Proficiencies:", "One type of artisan's tools from whatever you did before"));
c.push(B("Equipment:", "A soldier's kit, a token from home, a Concord medal of the Call in cheap tin, common clothes, and a belt pouch containing 8 gp"));

c.push(PS([{ t: "Feature: One of Eight Thousand. ", b: true }, { t: "You can find a bed, a meal and an honest answer anywhere in the coalition camp, from anybody, at any hour. Levymen know each other across companies, kingdoms and languages, and they will cover for you with an officer without being asked and without expecting anything. It also means you cannot move through the camp unnoticed. Somebody always knows where you went." }]));

c.push(table(
  ["d6", "Bond"],
  [8, 92],
  [
    ["1", "I am carrying a letter for somebody in the other column. I have not been able to deliver it."],
    ["2", "The priest promised me a place at the end of this. I have started wondering what that means in practice."],
    ["3", "My whole village sent four. I am the one who came back last time, so I went again."],
    ["4", "I have never in my life been chosen for anything before the king chose me."],
    ["5", "There is a farm and it is thirty acres and it is not going to work itself."],
    ["6", "I do not care about Elduvaine. I care about the eleven people in my file."]
  ], { full: true }
));

c.push(H2("Concord Devotee"));

c.push(P("You were raised, taught, or taken in by a Ninefold House \u2014 chapter, school, hospital and court of appeal at once, and in a bad year the granary as well. You may be clergy. You may simply be somebody the Concord fed and educated and who has never quite stopped reporting in."));

c.push(B("Skill Proficiencies:", "Religion, and either Medicine or History"));
c.push(B("Languages:", "Ninefold Cant"));
c.push(B("Equipment:", "A Ninefold token of your Work, a book of the liturgy, vestments or a scholar's robe, common clothes, and a belt pouch containing 12 gp"));

c.push(PS([{ t: "Feature: The House Will Take You In. ", b: true }, { t: "Any Ninefold House in four kingdoms will give you shelter, food and care, and will treat your word about what you have seen as evidence. The Concord's houses keep the best records outside Elduvaine and will let you read them. This works in every realm that took the Call and does not work in Elduvaine, where there are no Ninefold Houses at all \u2014 except the new ones the Tenth Work is building, which will also take you in, and will want to talk." }]));

c.push(table(
  ["d6", "Bond"],
  [8, 92],
  [
    ["1", "The House fed me for nine years and never once mentioned it. I mention it constantly."],
    ["2", "I read the Call aloud in a town square. I have wondered ever since how many of them died on the strength of it."],
    ["3", "My teacher joined the Tenth Work. I have not been able to write back."],
    ["4", "Sennet is a Work of the written word and I have never broken one."],
    ["5", "I want to see Elduvaine because I do not believe a word of what we teach about it."],
    ["6", "Ossuar keeps the dead properly. Somebody has to do that out here and it is going to be me."]
  ], { full: true }
));

c.push(H2("Occupation Survivor"));

c.push(P("You were in Elduvaine for the three years. Not in the resistance, necessarily, and not a collaborator either \u2014 just there, alive, getting a permit stamped, paying a levy that was published in advance, and watching the light go out of the stone quarter by quarter over fourteen months."));

c.push(B("Skill Proficiencies:", "Deception, Stealth"));
c.push(B("Tool Proficiencies:", "Forgery kit"));
c.push(B("Languages:", "Elduvish"));
c.push(B("Equipment:", "A forgery kit, three permits in three different names of which one is genuine, worn common clothes, a light-stone pebble that no longer holds anything, and a belt pouch containing 5 gp"));

c.push(PS([{ t: "Feature: Paper and Patience. ", b: true }, { t: "You know how the occupation's bureaucracy actually functions \u2014 which permits are checked, which are waved through, which clerk to approach and at what hour, and what a levy docket is supposed to look like. You can usually get one other person through a routine checkpoint alongside you. You also know, by name or by face, a startling number of the Elduvish clerks working for the administration, and about a third of them will not raise the alarm." }]));

c.push(table(
  ["d6", "Bond"],
  [8, 92],
  [
    ["1", "I hid four people for a winter and one of them informed on the other three."],
    ["2", "I paid the levy every quarter, on time, for three years. Nobody made me."],
    ["3", "I know exactly which of my neighbours took the work, and I am not certain I would have refused it."],
    ["4", "I have a child who does not remember the light being on."],
    ["5", "The garrison sergeant on my street was decent to me for two years. I do not know what to want for him."],
    ["6", "Somebody has to tell the coalition what it was actually like, because they have got it wrong in both directions."]
  ], { full: true }
));
c.push(H1("Feats"));

c.push(P("Six feats, all of them tied to something the setting actually does. As always, feats are an optional rule; if your table does not use them, nothing here is load-bearing."));

c.push(H2("Wyvern-Handler"));
c.push(BULLET([{ t: "Increase your Strength or Wisdom score by 1, to a maximum of 20." }]));
c.push(BULLET([{ t: "You have advantage on Wisdom (Animal Handling) and Intelligence (Nature) checks concerning beasts, dragons and monstrosities with a flying speed." }]));
c.push(BULLET([{ t: "When a creature with a flying speed enters your reach, you may use your reaction to make one melee weapon attack against it with a reach weapon. On a hit, its flying speed is 0 until the end of its next turn, and if it is airborne it falls." }]));
c.push(BULLET([{ t: "You are never frightened by a creature solely because of its size or its ability to fly. Greywatch does not consider this a virtue. Greywatch considers it the minimum." }]));

c.push(H2("Waystone-Walked"));
c.push(BULLET([{ t: "Increase your Wisdom score by 1, to a maximum of 20." }]));
c.push(BULLET([{ t: "You always know whether you are standing on the Willing Road, and you can find the nearest waystone without a check." }]));
c.push(BULLET([{ t: "When you and up to five companions travel together with a purpose you would be willing to state aloud, none of you suffer exhaustion from a forced march that day." }]));
c.push(BULLET([{ t: "You may not, cannot and will never be able to explain any of this to anybody, and four hundred years of Archive scholarship is on your side about that." }]));

c.push(H2("Contract-Bound"));
c.push(BULLET([{ t: "Increase your Intelligence or Charisma score by 1, to a maximum of 20." }]));
c.push(BULLET([{ t: "You gain proficiency in Writ-tongue, and you cannot be deceived about the contents of a written agreement you have read." }]));
c.push(BULLET([{ t: "Once per long rest, when you make a specific promise aloud and in front of a witness, you gain a d6 you may add to any one ability check, attack roll or saving throw made in direct service of keeping it. You lose the die, and cannot use this feature again for a week, if you act against the promise." }]));

c.push(H2("Light-Pourer"));
c.push(BULLET([{ t: "Increase your Intelligence or Charisma score by 1, to a maximum of 20." }]));
c.push(BULLET([{ t: "You learn the light cantrip if you do not already know it." }]));
c.push(BULLET([{ t: "You can spend ten minutes pouring light into a piece of worked stone you are touching. It sheds bright light in a 20-foot radius and dim light for a further 20 feet, for eight hours, and can be muffled or unmuffled with a touch. Only stone from the Standing Marches holds it well; ordinary stone holds it for one hour. You can maintain a number of pieces equal to your proficiency bonus." }]));

c.push(H2("Legion Drilled"));
c.push(BULLET([{ t: "Increase your Strength, Dexterity or Constitution score by 1, to a maximum of 20." }]));
c.push(BULLET([{ t: "While at least one willing ally is within 5 feet of you, you have advantage on attack rolls against any creature that is also within 5 feet of that ally." }]));
c.push(BULLET([{ t: "You may use your reaction to move up to half your speed toward an ally within 30 feet who has been reduced to 0 hit points, without provoking opportunity attacks." }]));

c.push(H2("Listening-Trained"));
c.push(BULLET([{ t: "Increase your Wisdom score by 1, to a maximum of 20." }]));
c.push(BULLET([{ t: "You have advantage on Wisdom (Insight) checks made to determine whether somebody is repeating words they do not understand." }]));
c.push(BULLET([{ t: "Once per long rest, you can spend ten minutes at the edge of any Elduvish water and hear up to one minute of what was spoken there, in the speaker's own voice, chosen by the DM. You do not get to choose. Nobody has ever got to choose." }]));

c.push(H1("Subclasses"));

c.push(H2("Cleric: Domain of the Kept"));

c.push(P("Elduvaine has no church, so it has no clerics in the Concord's sense. What it has are Keepers, who tend a habit the way a Concord priest tends a congregation \u2014 and a few of whom, over four hundred years, have found that the habit tends them back. This is the one Elduvish divine tradition, it is not organised, and it has no doctrine to speak of."));

c.push(B("Domain Spells.", "1st: goodberry, sanctuary. 3rd: pass without trace, warding bond. 5th: plant growth, sending. 7th: guardian of faith, stone shape. 9th: commune with nature, hallow."));

c.push(PS([{ t: "Keeper's Tending (1st level). ", b: true }, { t: "You gain proficiency with herbalism kits and mason's tools, and in the Nature skill. When you finish a long rest in a place that holds a habit \u2014 a Kept Season wood, a Willing Road waystone, a Listening Water, or Standing Light stone \u2014 you regain one expended spell slot of 3rd level or lower." }]));

c.push(PS([{ t: "Channel Divinity: Hold the Season (2nd level). ", b: true }, { t: "As an action, you fix a 30-foot radius in the state it is currently in for one minute. Within the area, no creature's hit point maximum can be reduced, no ongoing damage or condition worsens, burning fires do not spread, and no plant dies. Creatures may still be damaged, healed and killed normally; what stops is deterioration. Concentration is not required." }]));

c.push(PS([{ t: "Bound to the Place (6th level). ", b: true }, { t: "Choose one location no larger than a village, a wood, or a mile of road. While within it, you have advantage on saving throws against being charmed or frightened, and you cannot be surprised. You may change the location by tending a new one for a month, and most Keepers never do." }]));

c.push(PS([{ t: "Potent Spellcasting (8th level). ", b: true }, { t: "You add your Wisdom modifier to the damage you deal with any cleric cantrip." }]));

c.push(PS([{ t: "The Habit Answers (17th level). ", b: true }, { t: "Once per long rest, as an action within your bound location, you may ask the place for one thing: a road to be shorter, water to give something back, a wood to hold, or stone to give up its light. The DM decides what happens. It is always something, it is never precisely what was asked for, and it has never in four hundred years been explained." }]));

c.push(H2("Fighter: Pikewatch"));

c.push(P("Nine centuries of killing things that fly, reduced to a drill any competent person can be taught in a season and nobody masters in less than twenty years. It is not a chivalric tradition. There is no code, no oath and no title \u2014 just a rope, a very long stick, and an institution that has never had to conscript."));

c.push(PS([{ t: "Bracing Set (3rd level). ", b: true }, { t: "You gain proficiency with rope-work and climber's tools. When you take the Ready action to attack a creature that enters your reach, and you are wielding a weapon with the reach or heavy property, the attack deals an extra 1d8 damage and the target must succeed on a Strength saving throw (DC 8 + your proficiency bonus + your Strength modifier) or be knocked prone." }]));

c.push(PS([{ t: "Rope and Pike (3rd level). ", b: true }, { t: "You can use a bonus action to make a special attack with a rope, hook or net against a Large or smaller creature within 20 feet. The target must succeed on a Dexterity saving throw against the same DC or have its speed reduced to 0 until the end of its next turn. Against a flying creature, a failure means it falls." }]));

c.push(PS([{ t: "Reach of the Watch (7th level). ", b: true }, { t: "Your reach with a reach weapon increases by 5 feet, and you may make opportunity attacks against creatures that enter your reach as well as those that leave it. You may make one such attack per creature per turn." }]));

c.push(PS([{ t: "Bring It Down (10th level). ", b: true }, { t: "When you hit a creature with a flying speed with a melee weapon attack, its flying speed is reduced by 20 feet until the end of its next turn. If this reduces it to 0, it falls." }]));

c.push(PS([{ t: "Unimpressed (15th level). ", b: true }, { t: "You are immune to being frightened. If an effect would frighten you, you may instead impose disadvantage on the source's next attack roll against you, because you looked at it and it noticed." }]));

c.push(PS([{ t: "Nine Centuries of Practice (18th level). ", b: true }, { t: "Once per turn, when a Large or larger creature within your reach is prone or has a speed of 0, your first attack against it is an automatic critical hit." }]));

c.push(H2("Warlock: The Sleeping Archive"));

c.push(P("Something older than the wards sleeps near or under the Ysolde Archive. It answers to neither Maedoc Vale nor anybody else, it has not woken in living memory, and it has, on a very small number of occasions across four centuries, made an arrangement with a reader."));

c.push(PS([DM("DM Only: "), { t: "this patron does not answer what is in the deepest vaults, and a DM should be careful not to let a player conclude that it does. The dragon is a single approved exception that sleeps under the building. It is not a key to the building's contents, it does not know what is on the sealed shelves, and its interest in the campaign is entirely its own. If a player asks, the honest answer is that it has never said." }]));

c.push(B("Expanded Spell List.", "1st: comprehend languages, sleep. 2nd: detect thoughts, locate object. 3rd: clairvoyance, tongues. 4th: divination, arcane eye. 5th: legend lore, dream."));

c.push(PS([{ t: "Reader's Privilege (1st level). ", b: true }, { t: "You can read any written language, though not necessarily understand what is being discussed. Additionally, when you spend at least an hour with a written work, you learn one true and specific fact about the person who wrote it that they did not intend to record." }]));

c.push(PS([{ t: "What the Vaults Hold Back (6th level). ", b: true }, { t: "As a reaction when you fail a saving throw against a spell of 5th level or lower, you may succeed instead. Something enormous shifts in its sleep and the spell simply does not apply to you. You cannot use this feature again until you finish a long rest, and you dream about it." }]));

c.push(PS([{ t: "Deep Sleeper's Ward (10th level). ", b: true }, { t: "You have resistance to psychic damage, and you cannot be put to sleep by magic. You do sleep. You sleep extremely well, and for slightly longer than you intend, and this has never once been convenient." }]));

c.push(PS([{ t: "It Turns Over (14th level). ", b: true }, { t: "As an action, choose a creature you can see within 60 feet. It must succeed on a Wisdom saving throw against your warlock spell save DC or be frightened of you and incapacitated for 1 minute, repeating the save at the end of each of its turns. Nothing visible happens. Everyone present, including your allies, is briefly and privately certain that something very large has just noticed them. Once per long rest." }]));
c.push(H1("Spells"));

c.push(P("Six spells drawn from Elduvish practice, and one taken off a dead battle-mage. Elduvish worked magic is unusual in that most of it is an attempt to do deliberately and briefly what the land does permanently and without being asked, which every Elduvish caster will tell you is embarrassing and does it anyway."));

c.push(H2("Pour Light"));
c.push(PS([{ t: "Transmutation cantrip", i: true }]));
c.push(B("Casting Time:", "1 action"));
c.push(B("Range:", "Touch"));
c.push(B("Components:", "V, S"));
c.push(B("Duration:", "8 hours"));
c.push(P("You touch a piece of worked stone no larger than a helmet and pour light into it. It sheds bright light in a 15-foot radius and dim light for a further 15 feet, and can be muffled or unmuffled by any creature touching it and speaking a word you choose at casting. Stone quarried in the Standing Marches holds it for the full duration; other stone holds it for one hour. If you cast this on stone that has been drained, nothing happens at all, and it is the single most reliable test for whether a place has been worked over."));
c.push(B("Available to:", "Artificer, bard, cleric, druid, sorcerer, warlock, wizard"));

c.push(H2("Listening Water"));
c.push(PS([{ t: "2nd-level divination (ritual)", i: true }]));
c.push(B("Casting Time:", "10 minutes"));
c.push(B("Range:", "Touch"));
c.push(B("Components:", "V, S, M (a mouthful of water from the place you are asking about)"));
c.push(B("Duration:", "Instantaneous"));
c.push(P("You touch a body of standing or slow-moving water and it gives back up to one minute of speech that was spoken at its edge within the last thirty days, in the speaker's own voice. You do not choose which minute. The DM does, and should choose what is most interesting rather than what is most useful."));
c.push(P("In Elduvaine, where the water does this anyway, the spell reaches back thirty years rather than thirty days. Outside Elduvaine it works on any water at all, which is the only reason it was ever written down, and Elduvish casters regard using it at home as roughly equivalent to shouting at somebody who was going to answer you anyway."));
c.push(B("Available to:", "Bard, cleric, druid, warlock, wizard"));

c.push(H2("Willing Step"));
c.push(PS([{ t: "2nd-level conjuration", i: true }]));
c.push(B("Casting Time:", "1 action"));
c.push(B("Range:", "Self"));
c.push(B("Components:", "V"));
c.push(B("Duration:", "1 hour"));
c.push(P("For the duration, your speed increases by 10 feet and you ignore difficult terrain, provided you are moving toward a destination you named aloud when you cast the spell and have not since changed. If you change your destination, or move away from it for more than one round, the spell ends."));
c.push(P("The spell has no way of establishing whether you meant what you said and makes no attempt to. This is considered its principal shortcoming and also the only honest thing about it."));
c.push(B("Available to:", "Bard, druid, ranger, sorcerer, wizard"));

c.push(H2("Kept Season"));
c.push(PS([{ t: "3rd-level transmutation", i: true }]));
c.push(B("Casting Time:", "1 minute"));
c.push(B("Range:", "60 feet"));
c.push(B("Components:", "V, S, M (a seed from the season you are calling)"));
c.push(B("Duration:", "24 hours"));
c.push(P("Choose a point within range. A 40-foot-radius sphere centred there takes on a season of your choosing, in full: the light, the temperature, the state of every plant within it, and the smell. Creatures in the area are comfortable in it regardless of the weather outside, gain no benefit against magical cold or heat, and can forage as though in that season."));
c.push(P("Cast at 5th level or higher, the duration becomes seven days. Cast at 7th level, it becomes a year and a day, and a wood so treated will occasionally simply keep it, which is how the Orchard Marches began and is not repeatable on purpose."));
c.push(B("Available to:", "Druid, ranger, warlock, wizard"));

c.push(H2("Waystone's Refusal"));
c.push(PS([{ t: "4th-level abjuration", i: true }]));
c.push(B("Casting Time:", "1 action"));
c.push(B("Range:", "30 feet"));
c.push(B("Components:", "V, S, M (a chip of unlettered stone)"));
c.push(B("Duration:", "Concentration, up to 10 minutes"));
c.push(P("You set the stone down and name a purpose aloud. For the duration, any creature that attempts to enter a 20-foot radius around it must succeed on a Wisdom saving throw or be unable to do so, and must spend its movement going around. A creature that intends the named purpose no harm automatically succeeds and passes freely."));
c.push(P("The spell determines intent by a method its inventors were never able to describe and which four hundred years of subsequent scholarship has not improved on. It can be fooled. Nobody has established how, but it has been, twice, and both cases are in the Archive under a rule that permits almost nobody to read them."));
c.push(B("Available to:", "Cleric, druid, paladin, wizard"));

c.push(H2("Draw Down"));
c.push(PS([{ t: "5th-level necromancy", i: true }]));
c.push(B("Casting Time:", "1 action"));
c.push(B("Range:", "60 feet"));
c.push(B("Components:", "V, S, M (a piece of cut light-stone, which is consumed)"));
c.push(B("Duration:", "Instantaneous"));
c.push(P("You pull the life out of a 30-foot-radius sphere. Each creature in the area must make a Constitution saving throw, taking 8d8 necrotic damage on a failed save, or half as much on a success. Every ordinary plant in the area dies, the ground is barren for a year, and any resident-magic effect in the area is suppressed for 24 hours."));
c.push(PS([DM("DM Only: "), { t: "this is what the occupation's battle-mages use and it is on the list so that a party can recognise it when it is cast at them, and can understand exactly what they are looking at the first time somebody does it to a field. A player character can learn it. Nothing in the rules prevents that and nothing in this book will stop them. What will happen is that an Elduvish NPC will see them cast it, and the campaign should absolutely make that a scene." }]));
c.push(B("Available to:", "Sorcerer, warlock, wizard"));

c.push(H1("Magic Items"));

c.push(P("Elduvaine's great treasures are habits made portable, and their defining quality is that they are ordinary at home and astonishing anywhere else. A light-stone lamp is a doorstep in Caer Ysolde and a wonder in Duncarrow, and the campaign gets a great deal of mileage out of party members from different countries disagreeing about whether something is remarkable."));

c.push(table(
  ["Item", "Rarity", "Attunement", "Where it first appears"],
  [30, 18, 18, 34],
  [
    ["Standing-stone lamp", "Common", "No", "Module 3, from Caerwyn's baker"],
    ["Flask of Listening Water", "Uncommon", "No", "Module 5, at the Standing Water"],
    ["Road-token", "Uncommon", "Yes", "DM's discretion"],
    ["Kept Season seeds", "Uncommon", "No", "DM's discretion"],
    ["Legion pattern blade (+1)", "Uncommon", "No", "Module 7, Drell's campaign sword"],
    ["Wand of magic missiles", "Uncommon", "No", "Module 7, the harbour-mage's quarters"],
    ["Cloak of elvenkind", "Uncommon", "Yes", "Module 9, off one of Voss's scouts"],
    ["Keeper's pruning hook", "Rare", "Yes", "The Orchard Marches"],
    ["Ysolde reading-glass", "Rare", "Yes", "The Archive, and not easily"],
    ["Sovereign's veil", "Very rare", "Yes", "Not lootable. See its entry."]
  ], { full: true }
));

c.push(H2("Standing-Stone Lamp"));
c.push(PS([{ t: "Wondrous item, common", i: true }]));
c.push(P("A gnome-cut block of Standing Marches stone about the size of a fist, on a leather thong. Left in daylight for an hour it holds it, and gives it back for eight hours as bright light in a 20-foot radius and dim light for a further 20 feet. Muffle or unmuffle it as an action by closing your hand around it."));
c.push(P("It holds daylight rather than making it, which is a distinction that matters exactly once in this campaign, against a Light-Hollow, and is worth every copper on that occasion."));

c.push(H2("Flask of Listening Water"));
c.push(PS([{ t: "Wondrous item, uncommon", i: true }]));
c.push(P("A stoppered flask of water drawn from an Elduvish river while something was being said over it. Unstopper it and it gives back what was spoken at its mouth, once, in the speaker's own voice, for up to one minute. Then it is only water."));
c.push(P("A flask can be filled again at any Listening Water in Elduvaine. The campaign's better use is the other one: a party can deliberately speak into it and carry the words somewhere. What they choose to put in it is the point, and a DM should ask, and should write the answer down."));

c.push(H2("Road-Token"));
c.push(PS([{ t: "Wondrous item, uncommon (requires attunement)", i: true }]));
c.push(P("A palm-sized chip cut from a Willing Road waystone, which is technically a serious crime and has been done perhaps forty times. While attuned, you always know which direction leads most directly to a destination you have named aloud, and you and up to five companions travelling with you ignore exhaustion from forced march on any day you began by naming it."));
c.push(P("The token explains itself no more than the road does. It works less well as you get further from Elduvaine and nobody has established the rate."));

c.push(H2("Kept Season Seeds"));
c.push(PS([{ t: "Wondrous item, uncommon", i: true }]));
c.push(P("A sealed packet of a dozen seeds taken from a Kept Season wood. Sown together in a plot of at least twenty feet square and left for a year, they produce a stand that holds the season they were sown in, permanently, in the way the Orchard Marches do."));
c.push(P("This is not an adventuring item and is very deliberately not one. It is a thing a party can plant, somewhere they choose, and come back to in an epilogue. Give it to them early. Say nothing about it afterward."));

c.push(H2("Keeper's Pruning Hook"));
c.push(PS([{ t: "Weapon (sickle), rare (requires attunement by a druid, ranger, or cleric)", i: true }]));
c.push(P("You gain a +1 bonus to attack and damage rolls with this weapon. While holding it, you can cast the kept season spell once per long rest without expending a spell slot or material components, at 3rd level."));
c.push(P("In addition, as an action you can touch a dying plant, tree or wood and stop it dying for one week. This does not heal it, cure it, or address whatever is killing it. It simply stops, for a week, and then resumes. Six of these were made and four are accounted for."));

c.push(H2("Ysolde Reading-Glass"));
c.push(PS([{ t: "Wondrous item, rare (requires attunement)", i: true }]));
c.push(P("A hand lens in a plain gnome-cut frame, issued to Archive clerks of the fourth grade and above. Looking through it, you can read any written language, and you can tell at a glance whether a document is a forgery, a copy, or an original, and roughly how old it is."));
c.push(P("It also does the thing it was actually made for, which is that it will not let you read anything you are not entitled to read. Text beyond your permission is simply blank through the glass. Nobody has ever worked out how it decides, the Archive did not consider this a problem, and Maedoc Vale used one for nineteen years."));

c.push(H2("Sovereign's Veil"));
c.push(PS([{ t: "Wondrous item, very rare (requires attunement by the reigning sovereign of Elduvaine)", i: true }]));
c.push(P("Grey silk, unremarkable, and the reason nobody has seen Maelis Ysolde's face in four years. While attuned and worn, the wearer cannot be scried, read, charmed or compelled, cannot be lied to about the state of Elduvaine's habits, and knows at all times, precisely, how much of the Living Realm remains."));
c.push(P("The last of those is not a benefit and was never intended as one. It was made so that a sovereign would always know the condition of the thing they were bound to, in a kingdom where that binding was a nine-hundred-year source of long life and good weather. She has worn it through three years of the draining."));

c.push(PS([DM("DM Only: "), { t: "the veil is in this book so that a DM knows it exists and knows exactly what it has been doing to her, and not because it is loot. It cannot be taken, attunes to nobody else, and does nothing for anyone who is not bound to the habits. If a party asks her what the number is, she will tell them, and it is the single most demoralising piece of information available anywhere in the campaign, and she will give it to them without softening it because she has never once softened anything." }]));

c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300, after: 100 }, children: [new TextRun({ text: "~", size: 24 })] }));
c.push(PS([{ t: "\u201CIn Caer Ysolde it is a doorstep. Here it is the most remarkable thing anyone in this village has ever seen. Both of those are correct and I have stopped trying to reconcile them.\u201D", i: true }], { alignment: AlignmentType.CENTER }));
c.push(PS([{ t: "\u2014 a Harrowmark levyman, on being given a light-stone", i: true }], { alignment: AlignmentType.CENTER }));
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
  fs.writeFileSync(stagePath("KC_Character_Options.docx"), buf);
  console.log("Written.");
});
