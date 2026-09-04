// KC_Sourcebook.js -- setting canon for The King\u2019s Crusade.
//
// Canon lives here. corpus/ and documents/ are generated from this file and are
// never edited by hand. See CLAUDE.md for the sign-off rules: anything in the
// "Not yet decided" table must not appear here until it has been approved.
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

const { Table, TableRow, TableCell, WidthType, ShadingType, TableLayoutType } = require('docx');
const cell = (text, opts = {}) => new TableCell({ width: { size: opts.w || 20, type: WidthType.PERCENTAGE }, shading: opts.head ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, margins: { top: 55, bottom: 55, left: 60, right: 60 }, children: [new Paragraph({ spacing: { after: 0 }, indent: { firstLine: 0 }, children: [new TextRun({ text, bold: !!opts.head, size: 18 })] })] });
// cantSplit keeps a row\u2019s cells from being torn across a column or page break;
// tableHeader repeats the header row when a long table does span a break.
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
  children: [new TextRun({ text: "The King\u2019s Crusade", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "A Sourcebook of Elduvaine, and of the War for Its Deliverance", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- The Call
c.push(H1("The Call"));

c.push(BOX("Three years ago the Living Realm went quiet. The couriers stopped. The roads that used to carry a letter from Caer Ysolde to the coast in a day began to take nine, and then to take travellers nowhere at all. What came out of Elduvaine after that came on foot, and did not want to talk about it."));

c.push(P("Elduvaine has fallen. Not to an empire, not to a rival crown, but to one man who was already inside it \u2014 Maedoc Vale, once Keeper of the Ysolde Archive, who opened the wards he had spent his life maintaining and let his army in through them. The royal house is taken. The capital is his. And the land itself, which in Elduvaine is not scenery but a participant, is being drained of the thing that made it worth living in."));

c.push(P("Xavier III of Harrowmark has called a crusade. He is going himself, which is not what kings usually do, and he is taking the better part of his realm with him. Two allied powers march under the same call. A third party has been paid, or has decided, to be interested, and answers to neither."));

c.push(P("You are not the army. You are what the king sends ahead of it \u2014 four to six people chosen because a column of eight thousand cannot go where you can, cannot move as quickly as you can, and cannot be denied as plausibly if it is caught somewhere it should not be. The march will take months. Elduvaine is a long way from Harrowmark, and everything between the two has an opinion about armies crossing it."));

c.push(PS([DM("DM Only: "), { t: "the campaign opens on the road, not at the gates. Resist the urge to get the party to Elduvaine quickly. The road is half the campaign and the reason the arrival lands at all \u2014 players who have spent four sessions getting there will care about the place in a way that players who started there will not." }]));

// ------------------------------------------------- Elduvaine Before the Fall
c.push(H1("The Peoples of the Crusade"));

c.push(P("Before any of the realms below are described, one thing about all of them: none of them belongs to a single people. Every kingdom in this war is a mixed country, and what distinguishes them is not who lives there but what that place has made of them \u2014 its institutions, its trades, and the particular things it asks its people to be good at."));

c.push(P("The clearest way to see it is the dwarves. In Harrowmark they are crag-hold folk who hunt wyverns with pikes and a great deal of rope, and they are famous for it. In Norvatch they are factors and contract-lawyers whose written word is the most trusted instrument in the trading world. Same people. Two institutions. A Harrowmark dwarf and a Norvatch dwarf standing in the same room have markedly less in common with one another than either has with the neighbours they grew up beside \u2014 and both of them will tell you so, at length, if you make the mistake of assuming otherwise."));

c.push(PS([DM("DM Only: "), { t: "this is the note that keeps the setting from collapsing into a race-per-nation chart, and it is worth landing early and once. Every realm is mixed. Nobody in this campaign is their species. When the party meets an orc, the interesting question is never what an orc is like \u2014 it is which country raised this one, and which side of this war is paying them." }]));

c.push(H2("Harrowmark"));

c.push(P("Humans in plurality, farming thin soil behind hill-fort walls. Dwarves in the crag holds, where the stone is worth working and the winters are worth surviving indoors. And orcs in the high country, for as many generations as Harrowmark has bothered counting, which is to say all of them \u2014 the notion that they came from somewhere else is a thing foreigners believe and Harrowmark finds tiresome to correct. Half-orcs and half-elves throughout, unremarked, because Harrowmark does not have the energy to remark on things."));

c.push(P("Greywatch is the picture of it in miniature: a hold that is roughly half dwarves and half orcs with a scatter of everyone else, whose wyvern-watch has worked the same way for nine centuries. Dwarves take the pikes, because a dwarf braced properly does not move. Orcs take the ropes, because when there is a wyvern on the far end of a line somebody has to be strongest and it is usually them. Everyone else does the shouting and the counting. Nobody in Greywatch considers this arrangement remarkable, or has ever explained it to an outsider without visible impatience."));

c.push(P("Because Harrowmark\u2019s land is mundane and its people are not, every caster in the realm is a worked-magic professional: a dwarven runesmith who cuts her spells into a shield because a shield stays where you put it, an orc war-priest whose prayers are answered as reliably as anyone\u2019s anywhere, the king\u2019s own battle-mages drawing pay. What Harrowmark lacks is a land that joins in. It has never lacked people who can do magic."));

c.push(H2("Elduvaine"));

c.push(P("Elves are the oldest inhabitants and the Kept Season is essentially their institution \u2014 a wood planted deliberately into a season and held there, some of them tended by the same family since before the wards existed. An elven orchard-keeper who has stood four days into spring for three centuries has opinions about the calendar that no other people in the world share."));

c.push(P("Gnomes are the Archive\u2019s clerks and the Standing Light\u2019s masons, and the two trades are closer than they look: gnome-cut stone holds poured light longer and more evenly than anything else anyone has managed, which is why Caer Ysolde glowed and why the drainage there shows first. Halflings hold the river parishes, and the entire Elduvish etiquette of living politely beside a landscape that eavesdrops is halfling manners exported upward into a kingdom. Humans throughout, in numbers, doing everything."));

c.push(P("And Elduvaine\u2019s magic is populated as well as ambient. There are dryads in the older Kept Season woods \u2014 a dryad here is not a visitor to the wood, she is the wood, with opinions \u2014 and sprites and satyrs in parishes that have been strange for long enough to be relaxed about it, and things in the deeper Listening Waters that are not fish and do not care to be described. None of this is exotic to an Elduvish farmer. All of it is dying."));

c.push(H2("Oksitan, Auberitz, Norvatch"));

c.push(P("Oksitan is a kingdom of river country and horse country, human in the main, with an old landed nobility that is substantially dragonborn \u2014 houses that have held the same fords and the same crossings since before anyone was writing it down, and that consider a river a thing you are responsible for rather than a thing you cross. Half-elves are common at court and commoner in its diplomacy."));

c.push(P("Auberitz is the grand duchy that builds things: humans, gnomes and halflings, mercantile and engineering rather than martial, and privately of the view that this entire war is a logistics problem being mishandled by people who enjoy shouting. The coalition\u2019s siege train is Auberitz work, gnome-designed and halfling-quartermastered, and its artisans will tell anybody who stands still long enough exactly how it should be used."));

c.push(P("Norvatch is guild-law country \u2014 dwarves and tieflings, humans throughout \u2014 where a bargain is written, witnessed, filed, and binding, and where the realm\u2019s entire standing in the world rests on the fact that it honours the letter of one no matter what the letter turns out to have meant. Norvatch does not break contracts. It writes them very carefully first."));

c.push(H2("The Occupation"));

c.push(P("Vale\u2019s army is a legion, and somebody paid for it."));

c.push(P("Its professional core is hobgoblin: disciplined, literate, drilled, and entirely capable of filing correct paperwork about something terrible. Orc and human companies serve alongside them under the same pay-chest, with goblin auxiliaries doing the work nobody senior wants and a stratum of Elduvish clerks \u2014 collaborators, or simply people with children and a job that still exists \u2014 running the permits and the levies underneath all of it."));

c.push(P("This explains the occupation better than anything else about it. A horde does not issue permits. A horde does not collect a grain levy on a published schedule, or keep officers who negotiate, or contain no fanatics at all. What Vale did three years ago was not raise a mob; he hired a professional army that was already good at exactly this and had been waiting for the work. They are not his believers. They are his employees, and several of them are extremely good at their jobs."));

c.push(PS([DM("DM Only: "), { t: "run the occupation as competent people doing an ugly job for money, and never as a species. The campaign puts orcs on both sides on purpose \u2014 Greywatch\u2019s ropers and Vale\u2019s field general are the same people in different pay \u2014 and it never once remarks on it in narration. Let the players notice. If a player asks an occupation soldier why they are here, the honest answer is a wage and a three-year contract, and it should be delivered without a shred of self-justification." }]));

c.push(H1("Faith in the Nine Works"));

c.push(P("Harrowmark\u2019s land does nothing. Its roads are the length they look, its rivers keep nothing, its stone is dark at night, and nine centuries of people have lived there anyway by building every single thing they needed with their hands. The religion that grew out of that country is a religion of the made thing, and it is now the established faith of four kingdoms."));

c.push(P("The Concord teaches that the powers who made the world worked, and then withdrew \u2014 deliberately, and while the work was still unfinished, so that there would be something left for hands to do. They are not called gods. Calling them gods is a provincialism the Concord\u2019s clergy will correct you on once, politely, and thereafter by simply not using the word. They are called the Works, there are nine of them, and every one is a thing being done rather than a person doing it."));

c.push(table(
  ["The Work", "Sphere", "Domains"],
  [26, 46, 28],
  [
    ["Ashet the Anvil", "Craft, making, the honest tool, the thing built well enough to outlive its builder", "Knowledge, War"],
    ["Voran of the Long Road", "Travel, messengers, guest-right, the stranger fed without question", "Life, Trickery"],
    ["Sennet the Witness", "Oath, contract, law, testimony, the word that holds", "Knowledge, Trickery"],
    ["Halevin the Hearth-Kept", "Home, harvest, healing, the household that eats", "Life, Nature"],
    ["Aurine the Unshuttered", "Light, truth, dawn, courage of the ordinary kind", "Light"],
    ["Duran Ninefold", "War as discipline rather than glory; the line that holds", "War"],
    ["Threnn Greywater", "Sea, storm, river, and everyone the water keeps", "Tempest"],
    ["Ossuar the Quiet Warden", "Death, the grave, remembrance, endings done properly", "Death"],
    ["Saveth of the Green Verge", "Wilds, beasts, the turning of the seasons", "Nature"]
  ]
));

c.push(P("A Concord priest does not ask a Work for anything. A Concord priest reports. The liturgy is closer to a guild inspection than a prayer \u2014 here is what was built this season, here is what failed, here is what we intend next \u2014 and the answering miracle, when it comes, is understood as a tool handed down rather than a favour granted. This is why Harrowmark\u2019s clergy are so difficult to impress and so hard to frighten. They were never in the business of being awed."));

c.push(H2("The Concord"));

c.push(P("The institution that carries all this is old, wealthy, literate, and spread across Harrowmark, Oksitan, Auberitz and Norvatch without belonging to any of them. It keeps the only archive in the mundane world that rivals Elduvaine\u2019s, teaches most of the letters anybody in those kingdoms can read, and holds land in every one of the four. Its Ninefold Houses are chapter, school, hospital and court of appeal at once, and in a bad year they are also the granary."));

c.push(P("It was the Concord that put the Call into words. Xavier called the crusade; the Concord published it, in the old liturgical tongue, in every Ninefold House in four kingdoms on the same morning \u2014 and it was the Concord, not the king, that attached the promise of remission and the promise of a place at the end. A great many people marching in this army answered a summons that arrived in a priest\u2019s voice."));

c.push(H2("Elduvaine Has No Church"));

c.push(P("You do not build a temple to ask for an answer in a country where the river gives you one directly. Elduvaine acknowledges the Works \u2014 it is not ignorant, and it trades with four kingdoms that hold them \u2014 but it has never addressed them, because it has never needed an intermediary and does not entirely understand why anyone else does."));

c.push(P("What Elduvaine has instead are the Observances: local, seasonal, unwritten, and different in every parish. An Observance is not worship. It is closer to good manners practised at enormous scale \u2014 the things you do so that a road stays willing, a wood stays kept, and water that has heard you does not have cause to repeat the worst of it. Nobody supervises this. Nobody collects from it. Two river parishes eleven miles apart will observe entirely different things and each will find the other slightly eccentric."));

c.push(P("Elduvaine\u2019s religious professionals are called Keepers, and a Keeper tends a habit the way a Concord priest tends a congregation \u2014 a Road-keeper, a Water-keeper, a Season-keeper, a Light-keeper. The office is practical, hereditary as often as not, and carries no authority whatsoever outside the thing kept."));

c.push(PS([DM("DM Only: "), { t: "the title Keeper of the Ysolde Archive means exactly what it sounds like it means, and the campaign has been using it in plain sight since the first page. Maedoc Vale was a religious officer of a faith with no gods, charged with tending the largest single accumulation of Elduvaine\u2019s magic in the world. Do not underline this. Let a player notice it." }]));

c.push(H2("The Tenth Work"));

c.push(P("Not every Concord theologian was content to file Elduvaine under folk custom. Roughly two centuries ago a reading of the doctrine emerged which took the founding claim entirely seriously and followed it one step further: if the Works withdrew leaving the world unfinished, and if there is a country where something clearly unfinished is still lying about in the open, then that country is not a heresy. It is the remainder. It is the tenth and last Work, left undone, and completing it is the highest act available to hands."));

c.push(P("The Order of the Tenth Work has marched with this crusade from the day it was called, at its own expense, in disciplined and well-supplied numbers. Its people are personally brave, scrupulously honest, generous to the poor of any kingdom, and entirely prepared to consecrate a Kept Season orchard by burning it to the roots and raising a Ninefold House on the ash. They do not consider this a contradiction. They consider it the job."));

c.push(PS([DM("DM Only: "), { t: "the Tenth Work is the campaign\u2019s answer to the question of where its complexity lives, and the answer is: on the party\u2019s own side of the line. They are allies. They will hold a wall for the party and lose people doing it. They also intend to do something to the liberated kingdom that a great many Elduvish would consider a second occupation with better manners, and they will do it lawfully, at the Concord\u2019s direction, with the Promise as their instrument. Do not make them hypocrites and do not give them a secret evil plan. They are exactly what they say they are, and that is the problem." }]));

c.push(H2("Clerics, Paladins, and Everyone Else"));

c.push(P("Divine magic works. That is the entire extent of what this campaign asserts about the question. A Concord cleric channels the Works and a Keeper channels an Observance, and both of them heal the same wound at the same speed, and neither side\u2019s spells have ever once failed in a way that settled the argument. Players should be allowed to hold any position on this they like, including the position that the whole business is obviously one thing wearing two hats, which is a respectable and entirely unprovable view held by most Norvatch dwarves."));

c.push(P("A player character cleric or paladin from Harrowmark, Oksitan, Auberitz or Norvatch takes a Work and its domains from the table above. One from Elduvaine takes a habit instead \u2014 the Kept Season maps to Nature or Life, the Standing Light to Light, the Willing Road to Trickery or Life, and the Listening Water to Knowledge \u2014 and answers awkward questions about it for the entire march."));

c.push(H1("Elduvaine Before the Fall"));

c.push(H2("The Living Realm"));

c.push(P("Magic elsewhere is worked. It is studied, cast, bound, and spent, and when the spell ends the world goes back to what it was. Elduvaine does not work that way and never has. Here the magic is resident. It lives in the country the way weather lives in a sky, and it is not commanded so much as accommodated."));

c.push(P("What this means in practice is that the land has habits. Not rules \u2014 habits, which is a harder thing to plan around and a much better thing to travel through. Elduvaine\u2019s roads are shorter for a traveller who means well by the people at the other end, and no one has ever established exactly what the road is measuring. Its rivers keep what was said on their banks and will give it back, in the speaker\u2019s own voice, to anyone patient enough to sit and listen for it. Its forests hold the season they were planted in, so that a wood sown in autumn is in autumn permanently, and a stand of birches three centuries old has never once been anything but four days into spring."));

c.push(P("Elduvaine is therefore the only country in the world where the phrase \u201Cthe land will provide\u201D is a statement of fact rather than a piety, and where it is considered gravely rude to say something you would not want repeated near water."));

c.push(P("The habits are not uniform and not entirely mapped. Every parish has a local one, and a good deal of Elduvaine\u2019s culture is the accumulated etiquette of living politely alongside a landscape that notices."));

c.push(P("Below are the habits a party will meet first. They are the ones a traveller can rely on \u2014 which is what makes their failure, later, legible without a word of explanation."));

c.push(table(
  ["The habit", "What it does", "How it fails now"],
  [24, 44, 32],
  [
    ["The Willing Road", "A road shortens for travellers who mean well by whoever waits at the far end. Journeys take a fraction of their measured distance.", "Roads have gone honest. Nine days where there was one, and no shortening for anyone, however good their errand."],
    ["The Listening Water", "Rivers, wells and standing pools keep speech spoken at their edge and return it in the original voice.", "Some waters have forgotten entirely. Others return speech that nobody in living memory said."],
    ["The Kept Season", "A wood holds the season it was planted in, permanently and regardless of the calendar.", "Woods are slipping. A spring stand has gone to a winter it was never sown in and cannot leave."],
    ["The Standing Light", "Worked stone in Elduvaine holds light poured into it, so cities need no lamps and cellars are never dark.", "Draining first and fastest. Caer Ysolde is dark at night for the first time in its history."]
  ]
));

c.push(P("Habits are not the only thing a resident magic produces. Where Harrowmark\u2019s crags breed wyverns \u2014 mundane, appalling, ordinary animals \u2014 Elduvaine\u2019s resident magic breeds something at the same slot in the world that is not an animal at all. A dragon here is not a beast that happens to know magic; it is closer kin to the Willing Road or the Listening Water, a habit that grew teeth and appetite and left the ground. They are rare, and a land in the process of being drained is not a land producing many of anything. What the party is more likely to meet is the damage: a dragon visibly suffering, changed, or grown strange as its home region empties, the same way a wood does in the Held Winter."));

c.push(P("Nor are dragons the only thing living in a country whose magic is resident. Elduvaine has always had fey in it the way other countries have deer \u2014 dryads in the Kept Season woods, sprites in the hedgerows, satyrs in parishes remote enough not to be embarrassed about it. An Elduvish farmer treats a dryad roughly the way a Harrowmark farmer treats a bad-tempered neighbour with a legitimate grievance: carefully, politely, and without any sense that something extraordinary is happening. The draining is killing them too, and a dryad whose wood has gone from spring to a winter it was never planted in is the most direct grief this campaign has available. Use it once, properly, rather than often."));

c.push(H2("Caer Ysolde and the Archive"));

c.push(P("The capital sits where three rivers braid together, which was chosen deliberately: three Listening Waters meeting is the closest thing Elduvaine has to a public record. Caer Ysolde is a city of pale stone that glows faintly after dark from the light poured into it during the day, of observatories on every high roof, and of vaults."));

c.push(P("The vaults are the point. The Ysolde Archive is not a library in the ordinary sense \u2014 it is the working memory of a country whose landscape remembers things, and much of what it holds was deposited there by the land rather than by any person. It is the largest single collection of magical knowledge in the world, and Elduvaine has never been especially careful about saying so."));

c.push(P("Access to it was governed by rule rather than by power. A Keeper of the Archive could read almost anything. Almost."));

c.push(PS([DM("DM Only: "), { t: "the word \u201Calmost\u201D is the whole campaign. Do not explain it early. A party that works out on their own why a man with everything would burn a kingdom over a reading privilege has had a better evening than one that was told." }]));

c.push(H1("The Reckoning of Years"));

c.push(P("Elduvaine dates nothing consistently, for reasons covered under Tongues and Years below, so the table is given the way the coalition\u2019s own clerks now give it: relative to the night the wards opened. Everyone in four kingdoms knows which night that was."));

c.push(table(
  ["When", "What happened"],
  [22, 78],
  [
    ["c. 900 years before", "The Braid is settled where three rivers meet. House Ysolde is already the name of the family doing the settling; the city takes its name from them and not the other way about."],
    ["c. 700 years before", "The first Kept Season orchards are planted at Bryn Aeling, four days into a spring that has not ended there since."],
    ["c. 600 years before", "Light-stone is first quarried at Cairn Ithel. Caer Ysolde is lit, and stops needing lamps."],
    ["c. 480 years before", "The Ysolde Archive is founded, and \u2014 the decision that matters \u2014 its access is governed by rule rather than by power. Anyone may read what their rule permits. Nobody may read past it, including the person holding the keys."],
    ["c. 400 years before", "The wards are raised over the kingdom, in layers, by many hands across three generations."],
    ["c. 200 years before", "The doctrine of the Tenth Work is first argued in the Concord and is not condemned, which its opponents have been complaining about ever since."],
    ["c. 60 years before", "Norvatch codifies guild-law and signs its first carriage compact with Elduvaine. The arrangement outlives the kingdom that signed it."],
    ["22 years before", "Xavier of Harrowmark, then a young man, makes war on his father alongside his brothers. He takes three of Harrowmark\u2019s own hill forts, loses the war, and kneels for a mercy he is given."],
    ["19 years before", "Maedoc Vale is appointed Keeper of the Ysolde Archive."],
    ["11 years before", "Maelis Ysolde falls ill. It is nobody\u2019s fault and nobody\u2019s doing, and her physicians expect another century out of her."],
    ["5 years before", "Vale petitions, correctly and in writing, for access to the deepest vaults. He is refused, correctly and in writing, by the rule he is himself charged with enforcing. He does not appeal."],
    ["The night itself", "Vale opens every ward in Elduvaine from the inside, in one night, because he is the person entitled to. An army walks in. The kingdom falls to a key rather than a siege."],
    ["Year One", "Vindana falls in eleven days. The grain levy is published. Permits are issued. A bureaucracy assembles itself out of Elduvish clerks who would like to keep eating."],
    ["Year Two", "The draining becomes measurable to people who are not looking for it. Roads lengthen. Woods go wrong. Norvatch\u2019s carriage contract triples in volume."],
    ["Year Three", "The Call. Four kingdoms hear it read aloud on the same morning. Harrowmark musters at Duncarrow, and the campaign begins."]
  ]
));

// ---------------------------------------------------------------- The Fall
c.push(H1("The Fall"));

c.push(H2("Maedoc Vale"));

c.push(P("Vale was Keeper of the Ysolde Archive for nineteen years, and by every account that survives he was good at it. He designed a third of the wards that ringed Elduvaine and personally maintained the rest. He was not a foreigner, not a usurper with a claim, and not, so far as anyone knew, ambitious. He was the man you would have asked to check the locks."));

c.push(P("Three years ago he opened them. All of them, in one night, from the inside, in the correct order, with the correct words, because he was the person entitled to say them. The army that came through was waiting and had been for some time. Elduvaine did not fall to a siege. It fell to a key."));

c.push(P("This is why the kingdom\u2019s collapse looks so much worse from outside than it deserves to. Nothing was wrong with Elduvaine\u2019s defences. They were excellent. They were simply operated, on the night in question, by the enemy."));

c.push(PS([DM("DM Only: "), { t: "Vale wanted to finish reading. That is the entire motive, and it should be delivered flat, without any attempt to make it grand. The deepest vaults of the Archive were closed to the Keeper by rule \u2014 not by any power he lacked, and not by a lock he could not have picked, but by a rule he had sworn to and a council that would not vote to lift it. He unmade a country over a reading privilege. Play him as a man who considers this proportionate and is faintly puzzled that anyone disagrees. He is not mad. He is not tragic. He is a man who wanted a thing and priced the world correctly against it, and got the sum right by his own arithmetic." }]));

c.push(H2("What the Wards Were For"));

c.push(P("Worth stating plainly, because the party will meet the remains of them: Elduvaine\u2019s wards were never a wall. A country whose magic is resident cannot be walled without smothering it. The wards were a filter \u2014 they governed what could enter and, more importantly, what could be taken out. Elduvaine was not afraid of invasion. It was afraid of extraction."));

c.push(P("Vale has been extracting for three years."));

// ------------------------------------------------------------ The Occupation
c.push(H1("The Occupation"));

c.push(H2("Three Years In"));

c.push(P("Long enough to have a bureaucracy. That is the fact a party arriving from Harrowmark finds hardest to absorb: the occupation issues permits. It collects a grain levy on a published schedule. It employs Elduvish clerks, some of whom are collaborators and most of whom simply have children to feed and a job that still exists. There are people in the border parishes now who were four when the wards opened, and for whom a road that shortens is a story an old woman tells."));

c.push(P("Long enough, also, that everyone above the age of ten remembers before. The occupation is thoroughly established and not remotely accepted, and those two facts sit on top of each other everywhere the party goes."));

c.push(H2("The Draining"));

c.push(P("The visible damage is not ruin in the ordinary sense. Towns still stand. Fields are still worked, under levy. What has gone is the resident magic, and it goes in patches, which is far stranger to walk through than uniform devastation would be."));

c.push(BUL("The Dead Mile.", "A stretch of the old west road where nothing has any habit at all. It is an ordinary mile of ordinary dirt, and it is the single most frightening place in Elduvaine to anyone born there."));
c.push(BUL("The Forgetting.", "A river between two market towns that has lost eleven years of speech. The towns have taken to writing things down, which they have never had to do, and are not good at."));
c.push(BUL("The Held Winter.", "A birch wood four days into spring for three hundred years, now nine weeks into a winter it was never planted in, and getting deeper."));

c.push(P("The pattern of the drainage is not random, and a party that maps it will learn something the crusade\u2019s commanders have not."));

c.push(H2("What Survives"));

c.push(P("A resistance exists. It is not large, it is not well armed, and it is run \u2014 to the considerable inconvenience of everyone trying to rescue them \u2014 by a member of the royal house who got out and has flatly refused to leave since."));

c.push(P("The rest of the family are held, in two or three separate places, under conditions that vary from comfortable to markedly less so. Vale has been careful with them. He has not harmed them, has never threatened to, and appears to regard the question of what to do with them as somebody else\u2019s administrative problem that he will get to eventually."));

c.push(PS([DM("DM Only: "), { t: "the royals are united on the war and divided on what follows it. Every one of them wants Vale gone. What they do not agree on is what Elduvaine should be afterwards \u2014 the one at large has spent three years watching the country survive without a court, and has come to conclusions about that which the ones in captivity have not. Play this as colour that a table can ignore entirely. It becomes an endgame decision only if the players reach for it." }]));

c.push(H1("The House of Ysolde"));

c.push(P("An elven house, and after two centuries of marrying whoever it liked, an elven house containing a half-elf, a gnome and a human without anybody in Elduvaine finding that worth a remark. It has held the Braid since before the wards, it gave its name to the capital rather than taking one from it, and it governed the way Elduvaine does everything else: by long custom, minimal machinery, and the assumption that people will mostly behave."));

c.push(P("That assumption is what Maedoc Vale spent nineteen years standing inside."));

c.push(table(
  ["Name", "Style", "Where they are"],
  [30, 26, 44],
  [
    ["Maelis Ysolde (elf)", "The Veiled Sovereign", "Held in Caer Ysolde, in her own apartments, under guard"],
    ["Aveline Ysolde (human)", "The Regent", "At large. Runs what resistance survives and refuses evacuation"],
    ["Ninian Ysolde (half-elf)", "The Ward", "Held at Sennoch Hall until the party reach it"],
    ["Ottoline Vahn (gnome)", "The Magistrate", "Held in Vindana, in what she insists on calling her chambers"],
    ["Emrys Ysolde (elf)", "The Envoy", "Held separately, and not entirely as a prisoner"]
  ]
));

c.push(H2("Maelis Ysolde, the Veiled Sovereign"));

c.push(P("She has been dying for eleven years, and she was dying before any of this started, which is the only mercy in it. A sovereign of Elduvaine is bound into the habits at accession \u2014 not ceremonially, actually \u2014 and what the habits have, the sovereign has. For nine hundred years that arrangement ran the other way and every monarch of the Braid was uncommonly long-lived, uncommonly hard to poison, and uncommonly aware of what the weather was going to do."));

c.push(P("What began eleven years ago was ordinary, personal and slow, and her physicians expected another century out of her. Three years ago it changed character. The draining started, the binding held, and a wasting that had been hers alone became a second thing wearing the first one\u2019s symptoms. She is going the way the land is going, at the same rate, for the same reason. She wears a veil because the wasting reached her face four years ago and she declines to have it discussed. She has not stood unaided since the second winter of the occupation. She is, without any competition at all, the most intelligent person in this campaign, and three years of confinement have given her nothing whatsoever to do except think about the man holding her."));

c.push(BOX("\u201CHe has not been cruel to me. He has been considerate, and prompt, and he sends a physician who is genuinely skilled, and none of it costs him anything, and all of it is true. I want you to understand that before you meet him. He is not pretending. That is the difficulty.\u201D"));

c.push(PS([DM("DM Only: "), { t: "Maelis is the campaign\u2019s second clock, and unlike Norvatch\u2019s ledgers she cannot be bought, bargained with, or read faster. Anyone who has seen her can tell how far the draining has gone by looking at her, which makes her the only honest instrument in the setting \u2014 and it means every scene she is in is a status report on the entire war. Use her sparingly and never for exposition she would find beneath her. She will not be evacuated either, for the same reason her cousin will not: the habits are bound to her and she is not certain what leaving would do to them. Nobody knows. That is the whole problem with Elduvaine." }]));

c.push(H2("Aveline Ysolde, the Regent"));

c.push(P("Human, forty-one, and the only member of the house who was outside a wall on the night the wards opened \u2014 a fact she has never once described as luck. Her great-grandmother married into the Braid from a Harrowmark trading family, which makes her the closest thing Elduvaine\u2019s royal house has to a foreigner, and which is precisely why three years of occupation have not caught her: she knows how to be unremarkable in a country where everyone else is known."));

c.push(P("She runs the resistance, such as it is, and she is clear-eyed about what it is: a miller who miscounts, two orchard-keepers who hide people, a clerk who forges permits in the same hand she once used to file requisitions. Not an army. A conspiracy of small, patient, deniable treason. She refuses evacuation and has refused it in writing, twice, to two different coalition commanders who put it to her as a kindness."));

c.push(H2("Ninian Ysolde, the Ward"));

c.push(P("Half-elf, twenty-six, heir presumptive, and precise rather than broken \u2014 three years of house arrest at Sennoch Hall have produced somebody who keeps lists, notices things, and has thought very carefully about the difference between being rescued and being useful. Of the whole house, Ninian is the one who has done the arithmetic on the Promise, and the conclusion is not comfortable: a coalition owed a kingdom will collect, and an Elduvaine that cannot pay has simply bought itself a second war with better-armed creditors."));

c.push(H2("Ottoline Vahn, the Magistrate"));

c.push(P("A gnome somewhere past two hundred, the Ward\u2019s great-aunt by a marriage nobody now living attended, and a magistrate of the Braid for a hundred and sixty years of it. She fought her captivity with the only weapon she had ever needed and won: she filed. Three years of correctly formatted objections, appeals, and requests for clarification, every one of them lawful under the occupation\u2019s own published code, have cost the administration of Vindana an estimated four hundred clerk-days and produced a written record of the occupation so complete that the occupation itself now relies on it."));

c.push(H2("Emrys Ysolde, the Envoy"));

c.push(P("Elf, Maelis\u2019s younger brother, and the member of the family the other four do not discuss with strangers. He has been talking to Maedoc Vale for three years. Not under duress and not for advantage \u2014 he simply concluded, in the first week, that somebody in the house had to be in the room with the man, and that nobody else would do it."));

c.push(P("What that has bought is real: his sister\u2019s physician, the hostages held unharmed rather than usefully, the grain levy published rather than arbitrary, and at least two mass reprisals that were proposed and did not happen. What it has cost is also real, and he will tell you it without being asked, because he has had three years to arrive at the wording. He has answered questions. Some of them were about the wards. He does not know which of his answers mattered and he has stopped pretending the not-knowing is a defence."));

c.push(PS([DM("DM Only: "), { t: "Emrys is complexity in the cost, not in the cause, and the distinction is the whole campaign. He is not a traitor and he is not secretly the villain and there is no reveal. He is a man who took the only job nobody else would take, did it for three years, and cannot now produce a clean accounting of whether it helped. Do not let the party resolve him. If they forgive him he will decline the forgiveness on procedural grounds, and if they condemn him he will agree with them and continue doing exactly what he has been doing." }]));

c.push(H2("What They Cannot Agree On"));

c.push(P("They are united on the war. Every one of the five wants Vale stopped, wants the occupation ended, and has been prepared to be personally spent to get it. They are divided \u2014 privately, bitterly, and along lines that have nothing to do with the fighting \u2014 on what Elduvaine should be afterward."));

c.push(BUL("Maelis.", "The habits are restored first and everything else is arranged around that. She is aware this may not be survivable for her and considers the point irrelevant."));
c.push(BUL("Aveline.", "Elduvaine free, and the Ysolde Archive closed. Permanently, and by law. A kingdom that can be opened with a key should not keep the key."));
c.push(BUL("Ninian.", "The coalition is paid what it was promised, in full and quickly, because the alternative is a second war against the people who won the first one."));
c.push(BUL("Ottoline.", "The courts sit again within the month. Institutions first; everything anyone else on this list wants is downstream of a functioning registry."));
c.push(BUL("Emrys.", "Whatever keeps the most people alive. He has stopped having opinions with more content than that and does not expect to get them back."));

c.push(PS([DM("DM Only: "), { t: "the campaign does not adjudicate this and neither should the table. There is no correct position among the five and no mechanism anywhere in the eleven modules for settling which one wins \u2014 that is deliberate, and a table that never engages with the argument at all has lost nothing. What the party can do is be present for it, and be asked, and discover that being asked is worse than not being asked." }]));

// ---------------------------------------------------------------- The Crusade
c.push(H1("The Crusade"));

c.push(H2("Harrowmark and Its King"));

c.push(P("Harrowmark is a cold, stony realm of hill forts and hard winters, and nobody has ever called it marvellous. This matters, because it is where the party comes from, and Elduvaine will not be ordinary to them either."));

c.push(P("Elduvish scholars call Harrowmark unmagical. It is a useful phrase and it is not true, and the difference is worth getting right before the party crosses a border. Harrowmark has wizards. It has priests whose prayers are answered, hedge-witches in the fishing towns, war-mages on the king\u2019s payroll, and a small, cold, well-regarded college that has been teaching evocation for two hundred years. Magic is worked in Harrowmark exactly as it is worked everywhere else in the world: studied, cast, bound and spent, carried in a person and gone when that person stops."));

c.push(P("What Harrowmark does not have is a land that joins in. Its roads are the length they look. Its rivers keep nothing that is said to them. Its stone is dark at night and always has been. A Harrowmark battle-mage can throw fire the length of a hall, and has never in her life walked a road that got shorter because her errand was a kind one, and the first time it happens to her she will stop dead in the mud and refuse to go on until somebody explains."));

c.push(PS([DM("DM Only: "), { t: "keep this distinction crisp, because it is the one that makes the setting legible. Worked magic is normal and portable and belongs to people; a character from Harrowmark can be any caster in the game and should never be told otherwise. Resident magic belongs to Elduvaine and to nowhere else. \u201cUnmagical Harrowmark\u201d is a thing Elduvish characters say, sometimes rudely. It is not narration." }]));

c.push(P("What Harrowmark has instead of wonder is wyverns. They nest in the high crags and have as long as anyone has counted, and dealing with them is not an adventure in Harrowmark but a season of the year, conducted with long pikes and longer ropes and a number of funerals that the rest of the world finds startling. A wyvern is a beast. It is not magic, it cannot be reasoned with, and it can be killed by determined people with the right equipment and an acceptance of losses."));

c.push(P("This produces a particular kind of person. Harrowmark folk are not brave, exactly \u2014 they are unimpressed. Something enormous with wings is a problem their grandmother had. Whatever else the party carries out of their homeland, they carry that, and it will serve them better than they expect and worse than they hope."));

c.push(P("Xavier III is going in person. He is a genuinely excellent field commander and an indifferent king, and he is aware of both facts in a way that makes him better company than it makes him a ruler. Harrowmark will be governed in his absence by people he does not entirely trust, and he has decided that this is a price worth paying, and he is probably wrong."));

c.push(P("He was not always this. At nineteen he made war on his own father alongside his brothers, took three of Harrowmark\u2019s own hill forts, and lost. What is remembered about it in Harrowmark is not the war but the end of it: Xavier knelt in the yard of a castle he had besieged eight weeks earlier and asked mercy of the man inside, and got it, and neither of them ever spoke about it again. He has been an excellent commander and an uneasy king from that day, and he is under no illusion about which of the two he is good at."));

c.push(P("He chose this party himself. He does that sort of thing personally and remembers the names."));

c.push(PS([DM("DM Only: "), { t: "he is Xavier III and nothing else yet. Before this war is over he will be Xavier the Wyvernheart, earning it at the siege of Vindana, in the air, on the back of one of the things his country has spent nine centuries killing, in a battle that is going badly at the time. A handful of Harrowmark wyvern-riders travel with the army for exactly this reason, and should not be introduced with any fanfare before that moment. Do not use the name in any read-aloud before that scene. Do not let an NPC use it early as a slip. Do not foreshadow it. A campaign in which the players watch a king acquire his title is worth considerably more than one in which he arrives holding it." }]));

c.push(H2("The Coalition"));

c.push(P("Two other powers have taken the call. The Kingdom of Oksitan marches, and the Grand Duchy of Auberitz marches, each with its own reasons and its own idea of what a finished war looks like. Neither is doing this for Elduvaine\u2019s sake and neither pretends to be."));

c.push(P("The Kingdom of Norvatch has not taken the call. Norvatch trades with the occupied kingdom, has an understanding with the men holding it, and has made no promises to anybody marching. It will deal with either side. It is worth saying plainly that Norvatch keeps the bargains it makes, which is more than can be said for one of the powers that did take the call."));

c.push(PS([DM("DM Only: "), { t: "what Oksitan and Auberitz actually want is still not decided and must not be invented in passing \u2014 run those two by role. Norvatch is different: it has been settled, and it is written below." }]));

c.push(H3("What Norvatch Wants"));

c.push(P("Norvatch does not want Elduvaine. It wants to remain the market for what comes out of it."));

c.push(P("Specifically: the exclusive, standing, written right to purchase whatever leaves the Living Realm \u2014 guaranteed by whoever is holding the place when the war ends. Norvatch has held precisely that arrangement with the occupation for three years. Vale extracts; Norvatch moves it; and the reason a man can spend a kingdom by the cartload without the cartloads simply piling up in a yard somewhere is that there has been a buyer the entire time, paying promptly, asking nothing, and filing the paperwork correctly."));

c.push(P("This is the understanding the crusade has heard rumours of, and it is less sinister and considerably worse than the rumours suggest. There is no alliance. Norvatch did not want the wards opened and would have advised against it on commercial grounds. It simply priced the situation once the situation existed, the way it prices everything, and it has been trading inside that price ever since."));

c.push(P("Two things follow, and the campaign runs on both. The first is that Norvatch will sell to the coalition exactly as readily as it sells to Vale, honour the sale completely, and go on buying from Vale right up until the day the coalition outbids him \u2014 none of which it considers a contradiction, because none of it is. The second is the price. Norvatch wants its trading rights written into whatever settlement follows the war, which is to say it wants a share of the same kingdom the Promise has already sold to Oksitan and to Auberitz. There is not enough Elduvaine to pay everybody, and there never was."));

c.push(PS([DM("DM Only: "), { t: "Norvatch\u2019s ledgers are the single most valuable object in this campaign that is not the Archive, and the party can buy them. Three years of purchase records are the only complete account anywhere of how much of Elduvaine has already left it \u2014 which is to say, the only way to know how much is left, and therefore how close Vale is to finishing. That is the campaign\u2019s clock, for sale, from someone with no reason to lie about it and every reason to charge. What she wants for it is a signature from Xavier guaranteeing Norvatch\u2019s trade in a liberated Elduvaine. He has no authority to give it, over a kingdom that is not his \u2014 exactly as he had no authority to make the Promise, and made it anyway. Whether the party lets him is theirs. Do not put a thumb on it." }]));

c.push(PS([DM("DM Only: "), { t: "one thing about Norvatch stays open and should not be closed at the table: whether the arrangement with Vale was ever more than commercial. Doria Kell will say it was only ever trade. She is almost certainly telling the truth. \u201cAlmost certainly\u201d is the correct amount, and a DM who resolves it in either direction has spent something they cannot get back." }]));

c.push(H3("What Oksitan Wants"));

c.push(P("Oksitan is ford country. Every one of its great houses holds a crossing, most of them have held the same crossing since before there were records to hold it in, and the whole political economy of the kingdom is the control of movement \u2014 who may pass, at what price, and how fast. An Oksitan lord does not think about land the way an Auberitz one does. He thinks about the distance between two places and who sets it."));

c.push(P("Which is why Oksitan does not primarily want an estate in Elduvaine and does not much care about the Archive. Oksitan wants the Willing Road. A road that shortens itself for travellers who mean well is, to the people who have spent nine centuries taxing the alternative, the single most valuable object in the world, and Oksitan intends to claim its waystones and learn to cut more of them."));

c.push(PS([DM("DM Only: "), { t: "they cannot. Nobody has ever established what the Willing Road measures, the campaign never establishes it, and no amount of Oksitan surveying, purchase or scholarship is going to change that \u2014 which means the second-largest crown in this coalition is spending its army on a thing it is structurally guaranteed never to have. Do not tip this to the players as futility. Play the Oksitan officers as competent, confident and entirely sincere. The tragedy only works if it is not signposted." }]));

c.push(H3("Raimon V of Oksitan"));

c.push(P("Seventy-one years old, grey, small, and in the saddle before most of his staff are awake. He came to the throne at nineteen, spent four decades making Oksitan the arbiter of every crossing between three kingdoms, and answered the Call within eleven days of it reaching him \u2014 not because a priest asked him to, and not for the remission, but because he intends to walk the Willing Road himself before he dies and can think of no other way to arrange it."));

c.push(P("Nobody has been able to talk him out of going in person. It has been tried by his marshals, his heir, his physicians and the Concord, and his answer to all four was the same and is now quoted across the coalition with a good deal of affection: that he has been old for twenty years and it has not inconvenienced him yet."));

c.push(PS([DM("DM Only: "), { t: "Raimon takes the road the party does not take, and a river kills him early. Not a battle \u2014 a ford, in armour, at the head of his own column, in water that was not even particularly deep. The king of the ford-holders, drowned on his way to claim a road. Play him warm and play him competent in whatever scene or two he gets before the crossing, and do not foreshadow the water. His army comes apart behind him and the survivors reach the party later as a rumour, then as refugees, then as a problem." }]));

c.push(P("A detail worth keeping. Some part of the coalition will refuse to believe he drowned. Threnn Greywater keeps what the water takes, and the men of a broken column would very much rather believe their king is kept than dead \u2014 so the story that goes around the camps is that Raimon is sleeping under the Vaskren and will come up when Oksitan needs him. It is not true. It is also not worth arguing with, and a party that tries will find they are arguing with grief rather than with a claim."));

c.push(H2("The Promise"));

c.push(P("A crusade is not answered out of sympathy, and Xavier did not ask anyone to answer one. He offered terms. Those who take the call are promised two things: a share in Elduvaine\u2019s magic \u2014 land held inside the Living Realm, where the habits work \u2014 and access to the Ysolde Archive."));

c.push(P("Both are extraordinary offers, and neither is Xavier\u2019s to make. That has not slowed anybody down."));

c.push(P("It is worth being clear about what this means for the war. Nobody marching under the call is doing so to be kind. The two allied powers want land where the roads shorten and stone holds light, because there is no such land anywhere else and there never has been. They want the Archive because it is the largest collection of magical knowledge in the world and it is currently unattended. These are not shameful motives. They are simply not the same motive as deliverance, and they will not survive contact with an Elduvish farmer who wants to know who exactly promised away the river his grandmother is in."));

c.push(PS([DM("DM Only: "), { t: "the promise is the coalition\u2019s engine and its fault line, and it is only payable on a kingdom that is taken and held. Three things follow. Nobody asked Elduvaine, and the party will meet people who have worked that out. The coalition wants precisely what Maedoc Vale wants, for entirely different reasons, which makes some allied conversations very uncomfortable if the party is paying attention. And Vale is spending the exact asset that was promised \u2014 every month he reads, the thing the coalition marched for is worth less. Do not state any of this. Let them assemble it." }]));

c.push(H2("The Two Roads"));

c.push(P("The coalition could not agree on a route and did not try very hard. The armies split."));

c.push(BUL("The sea road.", "Ship out, coast down, make landfall within reach of Vindana, the great Elduvish port the war will turn on. Fast, episodic, and dependent on weather, harbours, and whoever currently owns the islands along the way."));
c.push(BUL("The mountain road.", "Overland, through passes and through realms that are nominally friendly and practically not. Slower, harder, and the reason armies have historically arrived at Elduvaine at half strength or not at all."));

c.push(P("The party chooses. Whichever road they take is written in full; the other is met later, in the form of what became of the people who took it."));

c.push(PS([DM("DM Only: "), { t: "this is the campaign\u2019s first Branch Ledger entry, and it is a genuine fork rather than a cosmetic one. A second crown marches by the road the party does not take, and that king does not arrive \u2014 he is lost early, in water, and his army comes apart on the road behind him. Whether the party witnesses that or walks into its aftermath months later is the whole difference between the two routes, and both versions are worth running." }]));

c.push(H1("Powers and Factions"));

c.push(P("Six bodies in this war will notice a party that deals with them repeatedly, and all six keep some kind of account. The standing given below is not a mechanic the players should be shown as a bar to fill; it is a reminder that these organisations remember, and a rough guide to what remembering buys."));

c.push(table(
  ["Faction", "Standing, in three steps", "What the top step actually gets you"],
  [22, 40, 38],
  [
    ["The Crusade", "Sworn \u00b7 Lance \u00b7 Banner of the Call", "A hearing with Xavier without an appointment, and the standing to ask the coalition for something it does not want to give."],
    ["The Order of the Tenth Work", "Postulant \u00b7 Hand \u00b7 Warden of the Work", "Disciplined troops who will die where you put them, healing without price, and an expectation you will not like."],
    ["The Ysolde Remnant", "Known \u00b7 Trusted \u00b7 Named", "Safe houses across the Braid, forged permits in a magistrate\u2019s own hand, and Aveline Ysolde answering a message the same day."],
    ["The Sixth Free Legion", "Noted \u00b7 Respected \u00b7 Owed", "Professionals do not hate you. Parley honoured, prisoners exchanged, and at the very top, an officer who will take a bribe and stay bought."],
    ["House Kell of Norvatch", "Client \u00b7 Factor\u2019s Guest \u00b7 Signatory", "Anything that can be bought, delivered on time, plus three years of ledgers and the truth about the war\u2019s arithmetic."],
    ["The Unbound Clerks", "Enquirer \u00b7 Reader \u00b7 Keeper\u2019s Friend", "Archive scholars in hiding who know what is on which shelf, and are the only people alive who can tell you what Vale has already read."]
  ]
));

c.push(PS([DM("DM Only: "), { t: "the Legion is the one on this list players will assume is not available to them, and it is the most interesting. The Sixth is a hired professional army with a contract, a pay schedule and a reputation to protect, and none of that is compatible with atrocity or with fighting to the last man for an employer whose kingdom is visibly running out. A party that treats the Legion as an institution rather than a monster gets an enormous amount out of it, and the campaign is built to reward exactly that." }]));

c.push(H1("Tongues, Years, and Names"));

c.push(H2("Languages"));

c.push(table(
  ["Tongue", "Who speaks it", "Notes"],
  [22, 30, 48],
  [
    ["Common", "Everyone, everywhere", "The trade tongue of four kingdoms and the reason this war can be argued about at all."],
    ["Marchspeak", "Harrowmark", "A blunt dialect of Common with a small vocabulary and a large number of words for weather. Mutually intelligible with Common if both parties are patient."],
    ["Elduvish", "All of Elduvaine", "A dialect of Elvish spoken by every race in the kingdom, dwarves and halflings included. Being unable to speak it marks you as foreign far more reliably than your face does."],
    ["Ninefold Cant", "The Concord\u2019s clergy", "Dead, written, and identical in every Ninefold House in four kingdoms. The Call was published in it. Counts as an exotic language."],
    ["Writ-tongue", "Norvatch", "A deliberately impoverished contract dialect with no idioms, no metaphors and one meaning per word. Ugly to hear and impossible to misread, which is the point."],
    ["Goblin", "The Sixth Free Legion", "The Legion\u2019s working language of command, whatever a given company\u2019s people happen to be. Orders, drill, and a very large body of extremely obscene marching song."],
    ["Draconic", "Oksitan\u2019s houses", "Court and heraldic use among the dragonborn nobility. An Oksitan commoner knows perhaps thirty words of it, all of them titles."],
    ["Sylvan", "Elduvaine\u2019s fey", "Dryads, sprites and satyrs. A great many Elduvish farmers have functional kitchen Sylvan and no idea it is a separate language."]
  ]
));

c.push(H2("The Year, and Why Elduvaine Cannot Agree On It"));

c.push(P("The Concord keeps a Ninefold calendar of nine months named for the nine Works, and it is used without argument in Harrowmark, Oksitan, Auberitz and Norvatch, because those are countries where a season is a season."));

c.push(P("Elduvaine cannot do this. A Kept Season wood holds the season it was planted in, and there are hundreds of them, and a parish that lies between a wood nine weeks into spring and a wood four days into autumn does not experience a year in any sense the Concord\u2019s calendar would recognise. So Elduvish parishes date by local observance instead \u2014 the year the near orchard flowered late, the third cutting after the water went quiet \u2014 and two villages an hour apart routinely disagree about what month it is by a margin of weeks, without either of them being wrong or finding the situation remarkable."));

c.push(P("This is very funny to everyone in the coalition for approximately two weeks, and then becomes a serious logistical problem for the quartermasters, and Auberitz has been quietly imposing the Ninefold calendar on every district it supplies since the landing. Nobody asked Elduvaine about that either."));

c.push(H2("Names"));

c.push(table(
  ["People or place", "Sample names"],
  [24, 76],
  [
    ["Harrowmark", "Brenna, Duncan, Hald, Ivor, Maud, Osric, Ren, Tam, Wat \u00b7 Vane, Ondry, Thane, Marrow, Kettle, Stannock, Crowe"],
    ["Elduvaine", "Aveline, Caerwyn, Emrys, Maelis, Ninian, Ottoline, Rhosyn, Tegan, Wyn \u00b7 Ysolde, Alder, Vahn, Sennoch, Nantcorrow, Aeling"],
    ["Oksitan", "Aimeric, Bertran, Guilhem, Peire, Raimon, Sicart \u00b7 and among the dragonborn houses: Vashkar, Ourrez, Sarrelan, Kaldiss"],
    ["Auberitz", "Anneke, Corvin, Hessel, Maartje, Sera, Wilm \u00b7 Vosk, Brandhoek, Kreyn, Aalder"],
    ["Norvatch", "Doria, Halvard, Karessa, Morvyn, Torvald, Zeruth \u00b7 Kell, Ashvane, Grimmond, Sallow"],
    ["The Sixth Free Legion", "Drell, Voss, Grask, Hoth, Nazira, Ruk, Ossian, Sekh \u00b7 companies are numbered, never named"]
  ]
));

// ------------------------------------------ What Is Actually Happening (DM Only)
c.push(H1("What Is Actually Happening (DM Only)"));

c.push(P("Maedoc Vale opened Elduvaine because the deepest vaults of the Ysolde Archive were closed to him by a rule he could not get lifted. Everything else \u2014 the army, the occupation, the permits, the grain levy, the captured royal house \u2014 is administrative overhead on that one objective. He is not building an empire. He is running a very large, very expensive reading room, and the country is the electricity bill."));

c.push(P("The drainage is not vandalism and not cruelty. Elduvaine\u2019s resident magic is what the deepest vaults are keyed to, and the amount required to open them is roughly the amount a country has. He is spending the land to read the book. He would tell you so if you asked, and would not understand the objection."));

c.push(P("The clock is real: when he has read enough, he no longer needs Elduvaine, and what happens to the kingdom at that point is not something he has given thought to. The party is not racing a doomsday weapon. They are racing a man\u2019s finishing a task and losing interest."));

c.push(P("The occupation\u2019s competence, its officers, and its willingness to negotiate all follow from this. Nothing in Vale\u2019s war is personal, which is exactly what makes him hard to fight and unpleasant to meet."));

c.push(P("One more thing follows, and it is the campaign\u2019s quietest problem. The coalition was promised a share of Elduvaine\u2019s magic and the run of its Archive. Vale is burning the first to open the second. Whatever the crusade arrives to liberate is smaller every month, and the men who marched for a share of it do not yet know that they are racing a subtraction rather than a thief."));

// ------------------------------------------------------------- Running It
c.push(H1("Running Elduvaine"));

c.push(P("Three things keep this campaign in the register it was built for."));

c.push(BUL("Describe the habit before you describe the damage.", "A road that shortens is a delight; a road that has stopped shortening is a bereavement, and only if the players felt the first one. Give them at least one working habit before you take any away."));
c.push(BUL("The villain is a villain.", "There is no argument to be had about whether Vale should be stopped. Put the difficulty somewhere else \u2014 in what the crusade costs, in allies who want their own things, and in what the Elduvish actually want once somebody asks them."));
c.push(BUL("Plan the relief valves.", "A long grim march needs taverns, absurd logistics, and recurring comic figures by design. Elduvaine\u2019s etiquette of living politely alongside a landscape that eavesdrops is a comedy engine as much as a wonder one; use it as both."));

c.push(H2("Wonders Worth Carrying"));

c.push(P("The best treasure in this campaign is not a sword. It is a habit in a container."));

c.push(P("Elduvish craft has spent centuries working out how to make the land\u2019s habits portable, and the results are ordinary inside Elduvaine and astonishing everywhere else. A party that has spent four sessions on the road and finally crosses the border should be able to buy, be given, or take these \u2014 and a Harrowmark character holding one for the first time is one of the better scenes available to a DM who wants the wonder to land on a person rather than on a landscape."));

c.push(BUL("A standing-stone lamp.", "Gnome-cut stone the size of a fist, which holds light poured into it and gives it back. Leave it in daylight and it lights a room for hours. In Caer Ysolde before the fall these were street furniture. They are now, increasingly, the only working ones left, and Elduvish people notice when an outsider carries one carelessly."));

c.push(BUL("A flask of Listening Water.", "Drawn properly from a bank that still remembers. It holds what is said at its mouth and gives it back once, in the speaker\u2019s own voice, and then it is only water. An investigation tool, a courier that cannot be bribed, and the only way in the world to carry a dead friend\u2019s last sentence as an object."));

c.push(BUL("A road-token.", "Cut from a waystone on a Willing Road. It does not explain what the road measures any more than the road does, and a party that tries to work it out by experiment will get exactly as far as nine centuries of Elduvish scholarship has. Keep it that way."));

c.push(BUL("Kept Season seeds.", "Plant them and a small stand holds whatever season they were sown in. Mostly a wonder. Occasionally, in the hands of a party that thinks about it, a tactical problem solved sideways."));

c.push(PS([DM("DM Only: "), { t: "these are also the campaign\u2019s clearest way to make the draining hurt. A standing-stone lamp that has stopped taking light, a flask that gives back a voice nobody recognises, seeds that come up in the wrong season \u2014 each of those is a paragraph of exposition the party feels instead of hearing. Break one at the right moment rather than explaining anything." }]));

c.push(H2("Travelling the Willing Road"));

c.push(P("The road shortens for travellers who mean well. That is the whole of what is known, it has been the whole of what is known for four hundred years, and Elduvaine finds the question of why roughly as interesting as being asked why water is wet."));

c.push(P("At the table, run it like this. When the party travels a stretch of Willing Road, decide privately how much of the expected time it takes \u2014 anywhere from all of it to a little under half \u2014 and then narrate the arrival rather than the arithmetic. They set out at dawn expecting to camp on the road, and the towers are in sight before the light goes. Never announce a percentage. Never roll where they can see it."));

c.push(PS([DM("DM Only: "), { t: "the essential rule, and the one this campaign will hold to the last page: do not invent a criterion. Not privately, not in your notes, not as a rule of thumb you apply consistently. The moment the road is measuring something the DM can name, it stops being the Willing Road and becomes an alignment detector with a travel-time output, and the single strangest thing in the setting is gone. Vary it. Contradict yourself. Let it shorten spectacularly for a party that has just done something appalling and not at all for one that has just been heroic, and never explain either. Players will build elaborate theories. Every one of them should be wrong, including the flattering ones." }]));

c.push(P("Off the Willing Road, travel in Elduvaine is ordinary and increasingly bad. The occupation\u2019s permits apply on every metalled road. The draining has made some stretches longer than the map says, which is the Willing Road running in reverse and is the most frightening thing most Elduvish have ever experienced. Assume standard travel pace, and apply the Dead Mile and its cousins as fixed hazards rather than as random encounters \u2014 they are places, not events."));

c.push(H2("Downtime on a March"));

c.push(P("An army moves for months and then sits for weeks, and the sitting is where a campaign like this either builds its characters or wastes the time. Between modules, each character may take one of these. All of them are meant to produce a scene, not a die roll."));

c.push(BUL("Stand a watch.", "Four hours on a picket line with one other person, chosen by the player. The DM asks what they talked about. This is the single most productive downtime action in the campaign and it costs nothing."));
c.push(BUL("Work the column.", "Smithing, mending, driving, doctoring, or hauling for the quartermasters. Earns the goodwill of Auberitz\u2019s people, which converts directly into supply, information and being told things early."));
c.push(BUL("Drink with the levy.", "One rumour from the module\u2019s table, one piece of camp gossip that is true, and one that is not. The DM should not indicate which is which."));
c.push(BUL("Keep the observance.", "Report to a Ninefold House, or tend a habit at a wayside shrine \u2014 and be seen doing it. In this coalition that is a political act whichever one you choose, and somebody will comment."));
c.push(BUL("Train with the watch.", "Harrowmark\u2019s wyvern-hunters will teach anyone who asks and mock anyone who does not last. Grants advantage on the next check made to handle, drive off, or avoid a large flying predator."));
c.push(BUL("Copy.", "Available only to a character with Archive training or an Unbound Clerk contact. Produces one page of something Vale would rather nobody had, chosen by the DM, at a rate of roughly one page per week of downtime."));

c.push(H2("The Four Voices"));

c.push(P("The Listening Water is the best exposition device this setting has, and it should be used as a habit rather than as a trick. When the party needs to learn what happened somewhere they were not, sit them on a bank and give them the same events four ways: the account a court would give, the account a farming village would give, the summons that set the whole thing moving, and somebody\u2019s private grief said aloud to running water because there was nobody else to say it to. None of the four is lying. None of them is complete. The party assembles the truth themselves, and they do it by listening rather than by rolling."));

c.push(H2("The Refrain"));

c.push(P("Every session module ends on the same four lines, after the loot, printed identically every time:"));

c.push(VERSE([
  "By thought, and by word, and by deed,",
  "the king\u2019s own chosen kept their creed.",
  "Far from home, where the quiet land lay,",
  "they held the line, and would not stray."
]));

c.push(PS([DM("DM Only: "), { t: "read it and end the session. Do not vary it, do not trim it to fit, do not remark on it, and do not let anyone at the table treat it as a joke by the fourth session \u2014 it stops being one around the seventh. In the final module of the campaign, and nowhere else, the last line changes to " }, { t: "and the water kept their names that day.", i: true }, { t: " Nothing else about the verse ever changes, which is the whole reason that lands. Never foreshadow it." }]));

c.push(P("And keep the peril honest. Vale has already won a great deal, and he has three years of practice at holding it. The wonder in this campaign is only worth anything because the danger under it is real."));

c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300, after: 100 }, children: [new TextRun({ text: "~", size: 24 })] }));
c.push(PS([{ t: "\u201CIt was nine days to the coast. It had never once been nine days to the coast.\u201D", i: true }], { alignment: AlignmentType.CENTER }));
c.push(PS([{ t: "\u2014 a courier of Caer Ysolde, on the first morning", i: true }], { alignment: AlignmentType.CENTER }));

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
  fs.writeFileSync(stagePath("KC_Sourcebook.docx"), buf);
  console.log("Written.");
});
