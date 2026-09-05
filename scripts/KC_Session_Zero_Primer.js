// KC_Session_Zero_Primer.js -- Session Zero: Foundations. The table-setting primer.
//
// Canon lives here. corpus/ and documents/ are generated from this file and are
// never edited by hand. See CLAUDE.md for the sign-off rules: anything in the
// "Not yet decided" table must not appear here until it has been approved.
//
// ESCAPE CONVENTION: all prose lives as \uXXXX escapes, never as literal
// typographic characters. tools/build.sh fails the build if that slips.
//
// This is the one document in the set that is neither a module nor a handout.
// It is DM-facing and must NOT be added to PLAYER_FACING in tools/pipeline.conf:
// it names, in one place, every fact the campaign guards.
//
// It invents no canon. Every claim in it is drawn from an existing document --
// "Why Were You Chosen?" and "The Promise" in the Player Guide, the peoples,
// languages and Refrain in the sourcebook, the Occupation Survivor background in
// Character Options, and the summons, audience and loot of Module One. Two things
// are deliberately withheld rather than settled: no milestone schedule is
// prescribed (every module computes its encounter arithmetic at 5th level, and
// Module Nine says in as many words that the table will have outgrown it), and no
// starting magic item is granted (the loot of Module One hands the party its
// first, and that only lands once).

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
  children: [new TextRun({ text: "Session Zero: Foundations", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "The King\u2019s Crusade", i: true }], { alignment: AlignmentType.CENTER, spacing: { after: 60 } }));
c.push(PS([{ t: "A table-setting primer for the Dungeon Master \u2014 four to six players, characters starting at 5th level", i: true }],
  { alignment: AlignmentType.CENTER }));

// ---------------------------------------------------------------- Purpose
c.push(H1("Purpose of This Session"));

c.push(P("Session Zero is played out of character, and it produces five things, none of which is a character sheet. Whether the party already knows itself, and what it did together if it does. A reason, for each character, that the king knew their name \u2014 which hands the DM a personal thread to pull for eleven modules. A country for each character, decided separately from their race, because in this setting those are two questions and answering them separately produces a better character. What each of them wants out of the Promise, written down where the whole table can see it. And a shared understanding of tone, agreed before anybody is attached to anything. Everything below serves one of those five."));

c.push(P("This campaign does not open with the party assembled. It opens with a rider on a frozen road who says each name once, plainly, and does not wait to be thanked \u2014 and Module One runs whether or not those names already know each other. Which of the two it is, is a decision this session makes, and not one the DM should make alone."));

c.push(PS([DM("DM Only: "), { t: "this is the only book in the set that is neither a module nor a handout, and nothing in it is read aloud. It gathers, in one place, everything the campaign guards, which is exactly why it is not the document you leave on the table during a break. The two books that do go to the players are named under What the Players Should Know, below." }]));

// ------------------------------------------------------- Character creation
c.push(H1("Character Creation Rules"));

c.push(BUL("Level:", "5th. Extra Attack and 3rd-level spells are online from the first session, because Xavier did not choose raw levies \u2014 he chose champions, and the first thing he asks them to do is walk at a CR 6 animal of the kind Greywatch has been losing people to for nine centuries."));
c.push(BUL("Ability scores:", "27-point buy or the standard array (15, 14, 13, 12, 10, 8), DM\u2019s preference. Avoid rolled stats here. This campaign asks a party to negotiate with allies who want their own things about as often as it asks them to fight, and a party that dumped its face spends four modules being talked past by its own side."));
c.push(BUL("Hit points:", "maximum at 1st level, average (rounded up) at every level after."));
c.push(BUL("Rules edition:", "2014, SRD 5.1 \u2014 race rather than species, no weapon masteries, the 2014 exhaustion track, monster maths from the 2014 DMG. That was settled before the first stat block in this campaign existed, and it is not a thing your table has to relitigate at Session Zero. Say it once and move on."));
c.push(BUL("Race and realm:", "two questions, asked separately. See the next section; it is worth twenty minutes."));
c.push(BUL("Backgrounds:", "any, and the campaign\u2019s own book of them is player-facing and can be handed out today. At 5th level a background\u2019s feature matters less than usual and its bonds matter a great deal more \u2014 every background in that book is a claim about where the character stood on the night the wards opened, and the campaign will ask."));
c.push(BUL("Equipment:", "standard class and background starting equipment, plus the muster kit below, and no magic item at creation. That last is deliberate and worth holding to: the party\u2019s first magic weapon is handed to them at the end of Module One, out of Duncarrow\u2019s armory, by a king who has visibly decided not to make a speech about it. It only lands once, and it lands better if it is the first."));
c.push(BUL("Advancement:", "milestone, and the campaign deliberately prescribes no schedule. Every module\u2019s encounter arithmetic is computed against 5th level and Module Nine says outright that by then the table will have outgrown it \u2014 so whatever pace you choose, recompute the adjusted XP against your party\u2019s actual level from the first level-up onward. If you want a shape rather than a rule: the road survived (end of Module Two), the Ward freed (end of Module Four), the city broken (Module Seven), the field army broken (Module Nine). That is a suggestion. The modules assume nothing."));

c.push(B("The muster kit.", "Cold-weather gear fit for the last grey weeks before a Harrowmark winter, their own name on the muster roll at Duncarrow, entered by a clerk who did not ask them to spell it, twenty-five gold of the muster advance not yet spent, and a copy of the Call as it was published \u2014 in Ninefold Cant, which most of the people carrying one cannot read. Somebody at the muster will translate it for them, badly, for a coin."));

// ------------------------------------------------------------ Where you are from
c.push(H1("Where You Are From, and What You Are"));

c.push(P("Put this to the table in two parts and do not let it collapse into one. First: what do you want to play? Then, separately: which of these realms raised you? Nobody in this campaign is their species, and the setting is built so that the second question is the interesting one. An orc from Harrowmark\u2019s high country and an orc in the Sixth Free Legion\u2019s pay have almost nothing in common; a dwarf who hunts wyverns with a pike and a dwarf who drafts contracts in Norvatch have less in common with each other than either has with the neighbours they grew up beside, and both of them will tell you so at length. The Player Guide\u2019s table of the marching realms is the handout for this block."));

c.push(P("Three cases are worth a DM\u2019s attention before they arrive:"));

c.push(BUL("A character from an allied realm.", "Entirely supported. Xavier\u2019s call reached every realm marching under it, and a character out of Oksitan, Auberitz or Norvatch arrives at the muster by their own road rather than Harrowmark\u2019s. What their own court told them about why they are here is a thread worth keeping \u2014 because their court has not fully explained itself to anybody, including them."));
c.push(BUL("An Elduvish character.", "Also supported, and the Occupation Survivor background exists for exactly this. What such a character needs is a reason to be standing at Duncarrow, and the honest one is the one the sourcebook already gives: what came out of Elduvaine after the wards opened came out on foot, and did not want to talk about it. A player who takes this is volunteering to be the only person in the party who has seen the thing everyone else is marching toward, and the DM should ask them privately how much of it they intend to say out loud."));
c.push(BUL("Elduvish, the language.", "Being unable to speak it marks a traveller as foreign far more reliably than a face does, and that friction is a feature of the middle third of this campaign. One speaker in the party is a role \u2014 the person everybody looks at in a market. None is the intended experience. Everybody is a flattening, and if the table drifts that way, say so now rather than in Module Three."));

// ------------------------------------------------------------ Why the king
c.push(H1("Why the King Knows Your Name"));

c.push(P("Xavier chose this party himself and remembers names, and every player should leave Session Zero able to say why he knew theirs. Choose one below, or roll a d6. Each of them hands the DM a thread \u2014 a place, a patron, a debt, an old promise, somebody still waiting. Write them down; they are the cheapest personal hooks this campaign will ever get, and they cost the players nothing to give."));

c.push(ORD("Service, and somebody noticed.", "You did something in Harrowmark\u2019s service: killed a wyvern nobody else would go near, held a hill fort, solved a problem quietly enough that only the king heard about it. Thread: the people you did it for, and what they expect of you now."));
c.push(ORD("Recommended.", "Somebody Xavier trusts put your name forward, and you have never quite worked out whether that recommendation was earned or simply convenient. Thread: the recommender, and the day they ask you for something."));
c.push(ORD("Attached.", "You are from one of the allied realms, and your own court explained its reasons for sending you only partially. Thread: what your court actually wants, and when you find out."));
c.push(ORD("A stake in Elduvaine.", "A half-remembered family story, an old debt, a promise made to somebody who never came back from there. Thread: whatever is waiting at the other end of the road with your name on it."));
c.push(ORD("You simply wanted to see it.", "A place people say is strange in ways nowhere else is, before whatever is happening to it finishes happening. Thread: what you do when you get there and it is not what the stories said."));
c.push(ORD("A name on a list.", "Your name came up, and nobody has told you who put it there. Thread: whoever did, and what they wanted \u2014 which is a gift, because it can be wired to any faction in the campaign later."));

c.push(PS([DM("DM Only: "), { t: "whichever of the six a player takes, Xavier has one real answer and Module One states it \u2014 he asked his officers for the people other people trusted under pressure, rather than the people other people were impressed by. Hold it. It is worth a great deal more as an answer to a question a player asks at the audience than as a speech he gives unprompted, and a table that never asks should never be told." }]));

// ------------------------------------------------------------ Forging
c.push(H1("Forging the Party"));

c.push(P("Ask the gating question first, out loud, and let the table answer it together: do these people already know one another, or do they meet on the road? Module One is written to run either way \u2014 as an introduction, with the party crossing paths on the same road answering the same rider, or as a group that walks into Duncarrow already knowing who carries what. Neither is the better answer. What matters is that the table chooses, rather than discovering in the first ten minutes of play that half of them assumed the other."));

c.push(P("Then put the rest of these to the table and let the players build on each other\u2019s answers rather than taking turns:"));

c.push(BUL("If you already know each other:", "what did you do together that a king\u2019s officer heard about? It does not have to have gone well. A thing that went badly and was survived is a better shared history than a victory, and it gives the DM something to bring up at a bad moment."));
c.push(BUL("If you do not:", "you have two or three days\u2019 ride to Duncarrow. What did you trade on the way \u2014 a name, a suspicion, a flask, the little each of you knows about why a king would want you by name? Play that stretch briefly at the top of Module One rather than narrating it away."));
c.push(BUL("What did you leave,", "and who is looking after it? Harrowmark is going to be a long way behind you for eight months. Somebody is feeding your animals or minding your trade or waiting for a letter, and the campaign is better for knowing their name."));
c.push(BUL("Which of you would the others trust with a rope,", "and which with a decision? They need not be the same person and it is more interesting when they are not. Harrowmark\u2019s own answer to a crisis is the rope, the pike, and the child, in that order, and a party that can say where each of them stands in that list has done most of the work of knowing itself."));

c.push(H2("What You Want Out of It"));

c.push(P("Close Session Zero with this, and give it the full fifteen minutes. Xavier did not ask anyone to march for nothing: the Call promises those who answer a share of Elduvaine\u2019s own magic \u2014 land inside the Living Realm, where the habits work \u2014 and the run of the Ysolde Archive. Every character has some position on that, and the position may perfectly well be that they do not want either of the two things on offer."));

c.push(P("One sentence each, said aloud and written where the whole table can see it for the rest of the campaign. Not a backstory and not a secret: a want. Land in a magical country. A year in the largest library in the world. Neither, because you came for a person. Neither, because you were told to come. Let them argue about it \u2014 a table that spends ten minutes arguing about what it is owed has understood this crusade better than one that agrees quickly."));

c.push(PS([DM("DM Only: "), { t: "this is the single most useful thing Session Zero produces, and it costs one line per player. The Promise is the motive engine of the entire coalition, every ally the party meets has an answer to the same question, and the campaign will come back to what the party wrote here. Do not say when, do not say how, and do not hint that it is going to be tested. Write it down and let it sit." }]));

// ------------------------------------------------------------ Two-hour plan
c.push(H1("Running Session Zero: A Two-Hour Plan"));

c.push(P("Two hours, six blocks, in this order. The order matters in one place only: the tone conversation happens before anybody has a character they are attached to, because it is a different and much worse conversation afterward."));

c.push(table(
  ["Block", "Time", "What happens"],
  [23, 15, 62],
  [
    ["The pitch", "15 min", "Five sentences of setting, three of tone. Do not read the sourcebook aloud. A friendly magical kingdom, taken by one man who was already inside it; a king who is going himself; a road long enough to change the people walking it. Then stop talking."],
    ["Tone and lines", "20 min", "The conversation below, in full, before characters exist."],
    ["Realm and character", "20 min", "The two questions, asked separately. Watch for the whole party quietly becoming Harrowmark, which is fine, and for it quietly becoming five different realms, which is also fine but means nobody has a home to miss."],
    ["Why the king knows your name", "25 min", "Roll or choose, then round the table so everyone hears the others\u2019. This block runs long and is the one most worth letting."],
    ["Forging the party", "25 min", "The gating question first, then the rest. Answers build on each other."],
    ["What you want out of it", "15 min", "One sentence each. Written down where everyone can see it."]
  ]
));

// ------------------------------------------------------------ Tone
c.push(H1("Tone, Lines, and the Conversation Before the Campaign"));

c.push(P("This campaign is set in an occupied country and it takes the occupation seriously. There are permits, a grain levy published in advance, and clerks who took the work because they have children and a job that still exists \u2014 and neighbours who did not refuse it either. There is captivity, without mistreatment and without release. There is a long illness that does not get better. There are sieges of cities with people still living in them, death in battle and death by drowning, an ally who would burn a sacred wood in order to sanctify it, and a creature made to fight on a side it would never have chosen. And underneath all of it there is a country being deliberately drained of the thing that made it worth living in. All of it is written to be taken seriously rather than used for shock, and none of it works if somebody at the table is enduring it politely."));

c.push(B("What to say.", "Name the heavy material out loud, before characters exist. That paragraph takes a minute to read and it is the whole of the obligation. Then ask two questions and take the answers without discussion: is there anything here you would rather this campaign did not go near, and is there anything here you actively want it to."));
c.push(PS([DM("DM Only: "), { t: "the content paragraph that opens this section is written to be read aloud, and it deliberately names no plot. Four of the campaign\u2019s guarded facts sit one sentence away from it \u2014 why Vale opened the wards, what is happening to the sovereign, which king does not finish the march, and what is over Vindana in Module Seven \u2014 and a content warning is not the place to spend any of them. Do not improve it with specifics." }]));

c.push(B("The second question matters as much as the first.", "A player who says they want the occupation material to be real \u2014 the permits, the collaborators, the neighbour who took the work \u2014 is telling the DM where to aim, and the campaign is markedly better for knowing. The Occupation Survivor background exists because somebody will want it."));

c.push(B("How to stop something mid-session.", "Agree one word, out loud, that any player including the DM can say to move the camera. It needs no explanation at the time and no discussion afterward unless the person wants one. Agree it now, so that using it later is a procedure rather than an interruption."));

c.push(B("What this campaign does not do.", "State it plainly so nobody has to spend a session wondering. There is no sexual violence anywhere in this corpus, depicted, described or implied. There is no slavery and no bound labour. There is no torture on the page. Harm to children is referred to and never shown \u2014 the one place a child is at stake is a rescue, offered as a complication in Module One\u2019s fight, and the fight is a hungry animal rather than a cruelty. If your table wants any of those lines drawn differently, that is the table\u2019s to decide, and Session Zero is when."));

c.push(B("One thing worth saying about the villain.", "Nothing in this campaign is an argument about whether Maedoc Vale should be stopped, and a table braced for a story that will eventually reveal him to be right can relax. The complexity in this campaign is on the party\u2019s own side of the line \u2014 in what the crusade costs, in allies who want their own things and say so, and in what the Elduvish actually want once somebody thinks to ask them."));

// ------------------------------------------------------------ Know / not know
c.push(H1("What the Players Should Know, and What They Should Not"));

c.push(P("Hand out the Player Guide and the book of Character Options before or at Session Zero. Both are written to be read by players, both are checked on every build for spoiler leaks, and neither contains one. Everything a character would plausibly have picked up in a lifetime of Harrowmark rumour is in them."));

c.push(B("Safe to know.", "The Call and its terms. The realms marching, and that none of them is marching purely out of sympathy. Harrowmark, and that its land is mundane while its people are not. The four habits of Elduvaine as stories rather than as facts. The Nine Works, the Concord, that Elduvaine never built a church, and that the Order of the Tenth Work says openly what it intends. The House of Ysolde by name and public reputation. That the roads out stopped keeping their habits three years ago and then stopped carrying anything at all."));

c.push(B("Not to know, and worth guarding.", "Why Vale opened the wards \u2014 the refusal, in writing, five years before, by the rule he was himself charged with enforcing. That there are deepest vaults at all, and what he is reading toward. That the sovereign of Elduvaine is bound into the habits, which is why she is failing with them. That Xavier will earn a name he does not yet have, and where. That the second king does not survive the march. That the Refrain\u2019s last line changes, once, at the very end. And the shape of the ending itself, which is a real decision and not a trick, and which is ruined completely by a player who arrives knowing there is a decision coming."));

c.push(PS([DM("DM Only: "), { t: "the commonest way a campaign like this gets spoiled is a DM who is pleased with the mechanism and cannot resist signposting it. The draining, the two clocks, the ending \u2014 seed all of it constantly, confirm none of it, and let the first player who says it out loud at the table have that moment entirely to themselves. The one item on the guarded list that leaks by accident rather than by pride is the sovereign\u2019s binding: it reached the player-facing options book once, through an attunement line, which is a reminder that this material leaks through the small print rather than through the prose. Every build now scans both player-facing books for it." }]));

// ------------------------------------------------------------ Refrain
c.push(H1("A Note on the Refrain"));

c.push(P("Every module in this campaign ends on the same four lines, printed identically, read after the loot, and never remarked on. The verse is printed in the Player Guide, so the table may well have read it before they play \u2014 that is fine and does it no harm."));

c.push(P("Do not read it at the end of Session Zero. It is not a module, nothing has been survived yet, and the first time a table hears those four lines out loud should be at the end of a night on which somebody nearly did not come back from a ledge in the dark. Reading it today spends the device on a session that was played out of character. End Session Zero on the rider instead.", { keepNext: true }));

// ------------------------------------------------------------ d12
c.push(H1("Twelve Things to Carry Out of Harrowmark"));

c.push(P("One object each, that is not equipment. It has no statistics, weighs nothing worth tracking, and is the single cheapest way to make a character a person \u2014 and later, when the road has gone on for months, the cheapest way to make a loss land without a paragraph of explanation. Offer this table only if the block stalls; a player who has their own answer always has a better one."));

c.push(table(
  ["d12", "What you carry"],
  [10, 90],
  [
    ["1", "A muster token from a war your father marched in and came back from."],
    ["2", "A hearth-stone the size of an egg, warmed at your own fire the morning you left, long cold."],
    ["3", "Somebody\u2019s letter, unopened, that you have decided to read at the coast."],
    ["4", "A wyvern scute, drilled and strung \u2014 a hold\u2019s way of saying you were there."],
    ["5", "Your own name written out by a child who has just learned how."],
    ["6", "A debt token you have not redeemed, from a person you would rather not see again."],
    ["7", "The iron key to a door eight months behind you."],
    ["8", "A pouch of the seed you were going to sow this spring."],
    ["9", "A pipe you do not smoke, kept because of who gave it to you."],
    ["10", "A copy of the Call in Ninefold Cant that you cannot read and will not throw away."],
    ["11", "A grey stone off your own ground, ordinary, and the only ordinary thing you will own in six months."],
    ["12", "Nothing. You left with what you stood in, and everyone at this table is about to hear the story."]
  ]
));

// ------------------------------------------------------------ First five minutes
c.push(H1("The First Five Minutes of Session One"));

c.push(P("End Session Zero by describing where Session One opens, so that nobody arrives cold. It is the last cold week before a Harrowmark winter. The party is somewhere ordinary \u2014 a roadside inn, a garrison bunkhouse, a farmhold at the edge of the crags \u2014 and a rider is coming up the frozen ruts faster than anybody rides without a reason, with a lathered horse and one word: Duncarrow. He says each name once, the way a man reads a list he has already checked twice, and he does not wait to be thanked."));

c.push(P("Nobody in Harrowmark is summoned by name for a small reason. Say that, and stop there."));

c.push(P("~", { alignment: AlignmentType.CENTER, spacing: { before: 240, after: 120 } }));

c.push(PS([{ t: "\u201cI know your names,\u201d he says, before anyone has given one, and it is not a boast \u2014 he simply does.", i: true }],
  { alignment: AlignmentType.CENTER }));

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
  fs.writeFileSync(stagePath("KC_Session_Zero_Primer.docx"), buf);
  console.log("Written.");
});
