// KC_Sourcebook.js -- setting canon for The King's Crusade.
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
  ],
  { full: true }
));

c.push(P("Habits are not the only thing a resident magic produces. Where Harrowmark\u2019s crags breed wyverns \u2014 mundane, appalling, ordinary animals \u2014 Elduvaine\u2019s resident magic breeds something at the same slot in the world that is not an animal at all. A dragon here is not a beast that happens to know magic; it is closer kin to the Willing Road or the Listening Water, a habit that grew teeth and appetite and left the ground. They are rare, and a land in the process of being drained is not a land producing many of anything. What the party is more likely to meet is the damage: a dragon visibly suffering, changed, or grown strange as its home region empties, the same way a wood does in the Held Winter."));

c.push(P("Nor are dragons the only thing living in a country whose magic is resident. Elduvaine has always had fey in it the way other countries have deer \u2014 dryads in the Kept Season woods, sprites in the hedgerows, satyrs in parishes remote enough not to be embarrassed about it. An Elduvish farmer treats a dryad roughly the way a Harrowmark farmer treats a bad-tempered neighbour with a legitimate grievance: carefully, politely, and without any sense that something extraordinary is happening. The draining is killing them too, and a dryad whose wood has gone from spring to a winter it was never planted in is the most direct grief this campaign has available. Use it once, properly, rather than often."));

c.push(H2("Caer Ysolde and the Archive"));

c.push(P("The capital sits where three rivers braid together, which was chosen deliberately: three Listening Waters meeting is the closest thing Elduvaine has to a public record. Caer Ysolde is a city of pale stone that glows faintly after dark from the light poured into it during the day, of observatories on every high roof, and of vaults."));

c.push(P("The vaults are the point. The Ysolde Archive is not a library in the ordinary sense \u2014 it is the working memory of a country whose landscape remembers things, and much of what it holds was deposited there by the land rather than by any person. It is the largest single collection of magical knowledge in the world, and Elduvaine has never been especially careful about saying so."));

c.push(P("Access to it was governed by rule rather than by power. A Keeper of the Archive could read almost anything. Almost."));

c.push(PS([DM("DM Only: "), { t: "the word \u201Calmost\u201D is the whole campaign. Do not explain it early. A party that works out on their own why a man with everything would burn a kingdom over a reading privilege has had a better evening than one that was told." }]));

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
  fs.writeFileSync(stagePath("KC_Sourcebook.docx"), buf);
  console.log("Written.");
});
