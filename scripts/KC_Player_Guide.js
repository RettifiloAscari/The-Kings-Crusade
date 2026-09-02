// KC_Player_Guide.js -- The Player Guide.
//
// Canon lives here. corpus/ and documents/ are generated from this file and are
// never edited by hand. See CLAUDE.md for the sign-off rules: anything in the
// "Not yet decided" table must not appear here until it has been approved.
//
// ESCAPE CONVENTION: all prose lives as \uXXXX escapes, never as literal
// typographic characters. tools/build.sh fails the build if that slips.
//
// SPOILER-SAFE BY DESIGN, AND AUTHORED INDEPENDENTLY -- per CLAUDE.md, this
// document is never produced by deleting paragraphs from the sourcebook.
// Every section below is written fresh, in a player-facing register, and
// checked against the sourcebook's own DM Only notes for what must stay out:
//   - Vale's actual motive (the closed vaults, the reading privilege). The
//     sourcebook is explicit that a table which works this out for itself
//     has had a better evening than one that was told. Players may know his
//     name and that he opened the wards himself; nothing more.
//   - Whether Vale is still human. Not raised as a question at all here.
//   - That Xavier will earn "the Wyvernheart" mid-campaign. Never named,
//     never hinted, in any read-aloud or player-facing text before it
//     happens at the table.
//   - Any Diverging Paths content, any NPC's eventual fate, any stat block,
//     the Branch Ledger, and the royal house's internal disagreement about
//     what Elduvaine should become -- all DM Only.
//   - The two-ending structure (hold Elduvaine or turn back) is never
//     described as a designed campaign fork here; it does not exist for the
//     players until it is a real choice at the table.
// Two-column layout (the document set's default); no change needed to
// tools/build.sh -- SINGLE_COL_MATCH only ever pointed at the Reference Guide.

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
  children: [new TextRun({ text: "The Player\u2019s Guide", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Before You Begin
c.push(H1("Before You Begin"));

c.push(P("This is what your character would know before ever meeting the rest of the party: the shape of the war, the realm you came from, and the reasons a person might have for answering a king\u2019s summons to march somewhere they have never been. It is not everything. Elduvaine has been closed to outsiders for three years, and most of what is said about it now is secondhand, half-remembered, or simply wrong \u2014 treat every claim below that isn\u2019t plainly stated fact as exactly that: a claim, not a promise of what you will find."));

c.push(P("Build your character against the world as described here. Let your DM fill in the rest at the table, as your character actually learns it."));

// ---------------------------------------------------------------- The Call
c.push(H1("The Call"));

c.push(P("Three years ago, the roads out of Elduvaine stopped keeping their old habits, and then stopped carrying anything at all. What little came out after that arrived on foot, thin, and reluctant to talk about what it had seen. The story that eventually settled was this: Elduvaine did not fall to an invasion. It fell to one of its own \u2014 a man named Maedoc Vale, who had spent nineteen years as Keeper of the great Archive at Caer Ysolde, and who one night opened every ward he had sworn to maintain and let an army walk in through the gap. Nobody who tells this story claims to know why he did it. He has not explained himself to anyone still able to ask."));

c.push(P("The royal house is taken, or scattered, or in hiding \u2014 accounts differ, and the truth is likely all three at once. The land itself, which in Elduvaine has never been ordinary, is reportedly failing in patches, the way a body fails when something is being taken from it faster than it can be replaced."));

c.push(P("King Xavier III of Harrowmark has called a crusade, and \u2014 unusually for a king \u2014 is leading it himself. Two other powers march under the same call: the Kingdom of Oksitan and the Grand Duchy of Auberitz, each for reasons of their own that they have been perfectly willing to state and perfectly unwilling to fully explain. A third, the Kingdom of Norvatch, has taken no side and made no promises, and deals with whoever is useful to it."));

c.push(P("You are not part of the army. You are what a king sends ahead of one: four to six people who can move faster than a column of thousands, go where it cannot, and be denied more plausibly if caught somewhere they should not be. Xavier chose this party himself, and remembers names."));

c.push(BOX("The march to Elduvaine is long, and your DM is not rushing you through it. Expect the road itself \u2014 the realms you cross, the people you meet on the way, the ordinary difficulty of moving anywhere with an army behind you \u2014 to be as much a part of this campaign as anything you eventually find at the other end of it."));

// ---------------------------------------------------------------- Where You're From
c.push(H1("Where You\u2019re From: Harrowmark"));

c.push(P("Harrowmark is cold, stony, and has never once been called marvellous by anyone who has actually been there. Grey stone under a greyer sky, hard winters, hill forts, and a people who take a great deal of convincing before they will call anything remarkable."));

c.push(P("Elduvish scholars have a habit of calling Harrowmark \u201cunmagical,\u201d which is a convenient phrase and not an accurate one. Harrowmark has wizards. It has priests whose prayers are answered, hedge-witches along the coast, war-mages in the king\u2019s own pay, and a cold, well-regarded college that has taught evocation for two centuries. Magic in Harrowmark works exactly as it works everywhere else in the world outside Elduvaine: it is studied, cast, carried by a person, and spent. A character from Harrowmark can be any spellcaster in the game, and should never be told they can\u2019t."));

c.push(P("What Harrowmark genuinely lacks is a land that participates. Its roads are exactly as long as they look. Its rivers hold nothing that is said to them. Its stone is dark once the sun goes down, and always has been. This is simply a fact about the place, not a flaw in it \u2014 and it is worth remembering the first time something in Elduvaine behaves otherwise."));

c.push(P("What Harrowmark has instead of wonder is wyverns. They nest in the high crags and always have, and dealing with them is not an adventure so much as a season of the working year, handled with long pikes, longer ropes, and losses the rest of the world finds alarming and Harrowmark finds ordinary. This produces a particular kind of person: not fearless, exactly, but hard to impress. Something enormous with wings overhead is a problem your grandmother already solved once. Whatever else your character carries out of Harrowmark, they carry that."));

// ---------------------------------------------------------------- The Coalition
c.push(H1("Who Marches With You"));

c.push(P("A character need not come from Harrowmark \u2014 Xavier\u2019s call reached every realm marching under it, and a member of the party may just as easily hail from Oksitan, Auberitz, or even Norvatch, arriving at the muster by their own road rather than Harrowmark\u2019s. The table below is what any of them would already know about the others."));

c.push(table(
  ["Realm", "Role in the War"],
  [24, 76],
  [
    ["Harrowmark", "Xavier III\u2019s own kingdom, and the crusade\u2019s spine. Cold, unmagical-in-the-land, and marching under its own king in person."],
    ["Kingdom of Oksitan", "An allied power under the call, marching for reasons of its own that its own court has not fully explained to anyone outside it."],
    ["Grand Duchy of Auberitz", "The second allied power, likewise under the call and likewise not marching purely out of sympathy for Elduvaine."],
    ["Kingdom of Norvatch", "Has taken no side. Trades with, and has some understanding with, whoever currently holds Elduvaine \u2014 and will deal with the coalition just as readily, for a price. Keeps whatever bargains it makes."]
  ]
));

c.push(P("None of the three visiting powers is marching purely for Elduvaine\u2019s sake, and none of them pretends otherwise to anyone who asks directly. What each of them actually wants out of a won war is not something your character would know yet \u2014 only that they all, plainly, want something."));

// ---------------------------------------------------------------- The Promise
c.push(H1("The Promise"));

c.push(P("A crusade this size is not raised on sympathy alone, and Xavier did not ask anyone to march for nothing. The terms offered to those who answer the call are simple to state and enormous in practice: a share of Elduvaine\u2019s own magic \u2014 land inside the Living Realm itself, where the old habits are said to work \u2014 and access to the Ysolde Archive, reputed to be the largest single collection of magical knowledge anywhere in the world."));

c.push(P("Both promises depend on one thing neither Xavier nor anyone else can currently guarantee: that Elduvaine is actually taken, and actually held, by the time the war ends. Your character may be marching for the promise itself, for reasons that have nothing to do with it, or for some mixture the player hasn\u2019t fully worked out yet \u2014 all three are equally valid places to start."));

// ---------------------------------------------------------------- Elduvaine, As It's Told
c.push(H1("Elduvaine, As It\u2019s Told"));

c.push(P("Almost nothing below comes from a firsthand source. Elduvaine has been closed for three years, and what travels out of a closed country is rumor, old memory, and the occasional survivor\u2019s account, filtered through however many tellings it took to reach a Harrowmark tavern. Treat it as the stories your character grew up hearing, not as a briefing."));

c.push(P("The stories agree on one strange thing: in Elduvaine, magic is said to live in the land itself rather than being worked by a person. Nobody outside Elduvaine has ever fully explained what that means, and the accounts that follow are the closest anyone has come."));

c.push(table(
  ["What\u2019s Said", "As the Stories Tell It"],
  [26, 74],
  [
    ["The Willing Road", "A road that shortens for a traveller who means well by whoever waits at the end of it. No one has ever agreed on what, exactly, the road is measuring."],
    ["The Listening Water", "Rivers and standing pools said to keep speech spoken at their edge, and give it back later in the speaker\u2019s own voice."],
    ["The Kept Season", "A wood that holds the season it was planted in, permanently, regardless of the calendar around it."],
    ["The Standing Light", "Worked stone that holds daylight poured into it, so that Elduvish cities are said to need no lamps at night."]
  ]
));

c.push(P("More recent accounts \u2014 from the handful of people who have actually come out since the wards opened \u2014 say some of this is failing. Roads that no longer shorten. Rivers that answer with voices nobody recognizes, or don\u2019t answer at all. Whether that\u2019s true, and what it would mean if it is, your character will have to find out for themselves."));

c.push(H2("Rumors from the Road"));

c.push(P("What follows is exactly what it sounds like \u2014 things travelers, refugees, and secondhand sources say. Some of it is certainly true. Some of it is certainly not. Nobody your character has ever met can tell them reliably which is which, and that includes their own party."));

c.push(BULLET([{ t: "\u201cThere\u2019s a stretch of the old west road where nothing works right anymore \u2014 not good, not bad, just nothing. A mile of ordinary dirt, in a country where dirt was never ordinary before.\u201d" }]));
c.push(BULLET([{ t: "\u201cA river between two market towns forgot eleven years of what was said to it. The towns started writing things down instead. They\u2019re bad at it.\u201d" }]));
c.push(BULLET([{ t: "\u201cThere\u2019s a wood that used to hold spring forever. Somebody who got out says it\u2019s winter there now, and getting colder every season, which shouldn\u2019t be possible for a wood like that at all.\u201d" }]));
c.push(BULLET([{ t: "\u201cThe man who took the kingdom was one of their own \u2014 kept the Archive for nineteen years, and everyone trusted him. Nobody agrees on why he turned. Nobody who\u2019d actually know is talking.\u201d" }]));
c.push(BULLET([{ t: "\u201cCaer Ysolde used to glow at night, they say \u2014 stone full of stored daylight, so the city never needed a lamp. Somebody who saw it recently says it\u2019s properly dark now, for the first time anyone can remember.\u201d" }]));

// ---------------------------------------------------------------- Building Your Character
c.push(H1("Building Your Character"));

c.push(P("This campaign runs on the 2014 edition of the fifth-edition rules (SRD 5.1). Characters begin at 5th level \u2014 Extra Attack and 3rd-level spells are already available from the first session, because Xavier did not choose raw recruits, he chose champions. Advancement is by milestone rather than by tallying experience: the party levels up when the story reaches a point that has earned it, not when a spreadsheet says so. The table expects four to six players, and every class, race, and background in the game is available \u2014 including full spellcasters from Harrowmark, whatever its reputation among the Elduvish."));

c.push(H2("Why Were You Chosen?"));

c.push(P("Xavier picked this party in person, which means he had a reason for each of you, even if that reason was never fully explained. Some possibilities, freely mixed or discarded:"));

c.push(BULLET([{ t: "You did something in Harrowmark\u2019s service \u2014 killed a wyvern nobody else would go near, held a hill fort, solved a problem quietly enough that only the king heard about it." }]));
c.push(BULLET([{ t: "You came recommended by someone Xavier trusts, and have never quite worked out whether that recommendation was earned or simply convenient." }]));
c.push(BULLET([{ t: "You are from one of the allied realms and were attached to this party for reasons your own court explained to you only partially." }]));
c.push(BULLET([{ t: "You have a stake in Elduvaine itself \u2014 a half-remembered family story, an old debt, a promise made to someone who never came back from there \u2014 that has nothing to do with the Promise and everything to do with why you said yes." }]));
c.push(BULLET([{ t: "You simply wanted to see a place people say is strange in ways nowhere else is, before whatever is happening to it finishes happening." }]));

c.push(P("You do not need to have this fully decided before the first session. Some of the best answers arrive during play, once your character has actually seen the road."));

// ---------------------------------------------------------------- Songs of the Muster
c.push(H1("Songs of the Muster"));

c.push(P("Every column marches to something. Harrowmark\u2019s troops, and by now most of the coalition\u2019s, have taken up a single verse said to date back to the muster at Duncarrow itself \u2014 sung, murmured, or simply thought at the end of a hard day, and by soldiers who could not tell you who wrote it first."));

c.push(VERSE([
  "By thought, and by word, and by deed,",
  "the king\u2019s own chosen kept their creed.",
  "Far from home, where the quiet land lay,",
  "they held the line, and would not stray."
]));

c.push(P("Your DM may close a session with it. If they do, it is worth listening to properly \u2014 it is the same four lines every time, and that is exactly the point of it."));

// ---------------------------------------------------------------- What You Carry
c.push(H1("What You Carry"));

c.push(P("You are leaving a cold, ordinary, thoroughly explicable kingdom to march toward one that, by every account, is neither cold nor ordinary nor explicable \u2014 and that is currently being taken apart by a man nobody can adequately explain either. The road between the two is long enough that you will be a different person by the time you reach the end of it. That is not a warning. It is the whole point of going."));

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
  fs.writeFileSync(stagePath("KC_Player_Guide.docx"), buf);
  console.log("Written.");
});
