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
  indent: { left: 220, right: 220, firstLine: 0 },   // template default firstLine=180 otherwise leaks in
  children: [new TextRun({ text, italics: true })]
});

// Verse in a read-aloud box: keeps its line breaks instead of running together.
const VERSE = (lines) => new Paragraph({
  spacing: { before: 120, after: 160 },
  shading: { type: "clear", fill: "F3EFE4" },
  indent: { left: 220, right: 220, firstLine: 0 },   // same fix as BOX -- see its comment
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
const cell = (text, opts = {}) => new TableCell({ width: { size: opts.w || 20, type: WidthType.PERCENTAGE }, shading: opts.head ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, margins: { top: 50, bottom: 50, left: 45, right: 45 }, children: [new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text, bold: !!opts.head, size: 18 })] })] });
// cantSplit keeps a row's cells from being torn across a column or page break;
// tableHeader repeats the header row when a long table does span a break.
const row = (cells, opts = {}) => new TableRow({ children: cells, cantSplit: true, ...opts });
const FULLWIDTH = "KCFullWidth";   // marker only; transplant.py acts on it and strips it
const table = (headers, widths, rows, opts = {}) => new Table({ ...(opts.full ? { style: FULLWIDTH } : {}), width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ row(headers.map((h, i) => cell(h, { head: true, w: widths[i] })), { tableHeader: true }), ...rows.map(r => row(r.map((v, i) => cell(v, { w: widths[i] })))) ] });

const mod = (v) => { const m = Math.floor((v - 10) / 2); return (m >= 0 ? "+" : "\u2212") + Math.abs(m); };
const abCell = (text, bold) => new TableCell({ width: { size: 16.6, type: WidthType.PERCENTAGE }, shading: bold ? { type: ShadingType.CLEAR, fill: "E4DCCB" } : undefined, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40, before: 40 }, keepNext: !!bold, children: [new TextRun({ text, bold: !!bold, size: 20 })] })] });
const SB = (d) => { const out = []; out.push(new Paragraph({ spacing: { before: 240, after: 40 }, children: [new TextRun({ text: d.name, bold: true, size: 26, color: "5B1F1F" })] })); out.push(PS([{ t: d.meta, i: true }], { spacing: { after: 120 } })); out.push(B("Armor Class:", d.ac)); out.push(B("Hit Points:", d.hp)); out.push(B("Speed:", d.speed)); out.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ new TableRow({ cantSplit: true, tableHeader: true, children: ["STR","DEX","CON","INT","WIS","CHA"].map(h => abCell(h, true)) }), new TableRow({ cantSplit: true, children: [d.str,d.dex,d.con,d.int,d.wis,d.cha].map(v => abCell(v + " (" + mod(v) + ")")) }) ] })); out.push(P("", { spacing: { after: 60 } })); if (d.saves) out.push(B("Saving Throws:", d.saves)); if (d.skills) out.push(B("Skills:", d.skills)); if (d.senses) out.push(B("Senses:", d.senses)); if (d.langs) out.push(B("Languages:", d.langs)); out.push(B("Challenge:", d.cr)); (d.traits||[]).forEach(t => out.push(PS([{ t: t.n + ". ", b: true, i: true }, { t: t.t }]))); if (d.actions && d.actions.length) { out.push(PS([{ t: "ACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.actions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } return out; };

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

c.push(PS([DM("DM Only: "), { t: "he is Xavier III and nothing else yet. Before this war is over he will be Xavier the Wyvernheart, and he will earn it in the air, on the back of one of the things his country has spent nine centuries killing, in a battle that is going badly at the time. Do not use the name in any read-aloud before that scene. Do not let an NPC use it early as a slip. Do not foreshadow it. A campaign in which the players watch a king acquire his title is worth considerably more than one in which he arrives holding it." }]));

c.push(H2("The Coalition"));

c.push(P("Two other powers have taken the call. The Kingdom of Oksitan marches, and the Grand Duchy of Auberitz marches, each with its own reasons and its own idea of what a finished war looks like. Neither is doing this for Elduvaine\u2019s sake and neither pretends to be."));

c.push(P("The Kingdom of Norvatch has not taken the call. Norvatch trades with the occupied kingdom, has an understanding of some kind with the men holding it, and has made no promises to anybody marching. It will deal with either side. It is worth saying plainly that Norvatch keeps the bargains it makes, which is more than can be said for one of the powers that did take the call."));

c.push(PS([DM("DM Only: "), { t: "the three realms are named; what each of them actually wants is not decided and must not be invented in passing. Run them by role. Norvatch is the useful one for a party that likes leverage, precisely because it is not on anybody\u2019s side and can be relied on once bought \u2014 and because buying it costs something the coalition would rather not spend." }]));

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

c.push(H2("The Four Voices"));

c.push(P("The Listening Water is the best exposition device this setting has, and it should be used as a habit rather than as a trick. When the party needs to learn what happened somewhere they were not, sit them on a bank and give them the same events four ways: the account a court would give, the account a farming village would give, the summons that set the whole thing moving, and somebody\u2019s private grief said aloud to running water because there was nobody else to say it to. None of the four is lying. None of them is complete. The party assembles the truth themselves, and they do it by listening rather than by rolling."));

c.push(H2("The Refrain"));

c.push(P("Every session module ends on the same four lines, after the loot, printed identically every time:"));

c.push(VERSE([
  "By thought, and by word, and by deed,",
  "the king\u2019s chosen kept the road.",
  "Far from home, under a borrowed sky,",
  "they stood against the quiet."
]));

c.push(PS([DM("DM Only: "), { t: "read it and end the session. Do not vary it, do not trim it to fit, do not remark on it, and do not let anyone at the table treat it as a joke by the fourth session \u2014 it stops being one around the seventh. In the final module of the campaign, and nowhere else, the last line changes to " }, { t: "and the water kept their names.", i: true }, { t: " Nothing else about the verse ever changes, which is the whole reason that lands. Never foreshadow it." }]));

c.push(P("And keep the peril honest. Vale has already won a great deal, and he has three years of practice at holding it. The wonder in this campaign is only worth anything because the danger under it is real."));

c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300, after: 100 }, children: [new TextRun({ text: "~", size: 24 })] }));
c.push(PS([{ t: "\u201CIt was nine days to the coast. It had never once been nine days to the coast.\u201D", i: true }], { alignment: AlignmentType.CENTER }));
c.push(PS([{ t: "\u2014 a courier of Caer Ysolde, on the first morning", i: true }], { alignment: AlignmentType.CENTER }));

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
  fs.writeFileSync(stagePath("KC_Sourcebook.docx"), buf);
  console.log("Written.");
});
