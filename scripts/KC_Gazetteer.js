// KC_Gazetteer.js -- the geography of the crusade, region by region.
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
const SB = (d) => { const out = []; out.push(new Paragraph({ spacing: { before: 240, after: 40 }, children: [new TextRun({ text: d.name, bold: true, size: 26, color: "5B1F1F" })] })); out.push(PS([{ t: d.meta, i: true }], { spacing: { after: 120 } })); out.push(B("Armor Class:", d.ac)); out.push(B("Hit Points:", d.hp)); out.push(B("Speed:", d.speed)); out.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [ new TableRow({ cantSplit: true, tableHeader: true, children: ["STR","DEX","CON","INT","WIS","CHA"].map(h => abCell(h, true)) }), new TableRow({ cantSplit: true, children: [d.str,d.dex,d.con,d.int,d.wis,d.cha].map(v => abCell(v + " (" + mod(v) + ")")) }) ] })); out.push(P("", { spacing: { after: 60 } })); if (d.saves) out.push(B("Saving Throws:", d.saves)); if (d.skills) out.push(B("Skills:", d.skills)); if (d.senses) out.push(B("Senses:", d.senses)); if (d.langs) out.push(B("Languages:", d.langs)); out.push(B("Challenge:", d.cr)); (d.traits||[]).forEach(t => out.push(PS([{ t: t.n + ". ", b: true, i: true }, { t: t.t }]))); if (d.actions && d.actions.length) { out.push(PS([{ t: "ACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.actions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } if (d.reactions && d.reactions.length) { out.push(PS([{ t: "REACTIONS", b: true }], { spacing: { before: 80, after: 80 } })); d.reactions.forEach(a => out.push(PS([{ t: a.n + ". ", b: true, i: true }, { t: a.t }]))); } return out; };

// ---------- content ----------
const c = [];

// A settlement entry: heading plus a compact italic stat line.
const SITE = (name, stat) => [H2(name), PS([{ t: stat, i: true }], { spacing: { after: 140 } })];

c.push(new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [new TextRun({ text: "The King\u2019s Crusade", bold: true, size: 40, color: "5B1F1F" })]
}));
c.push(PS([{ t: "A Gazetteer of the Long Road, from Duncarrow to the Braid", i: true }],
  { alignment: AlignmentType.CENTER }));

c.push(H1("Using This Gazetteer"));

c.push(P("Every entry below carries a stat line \u2014 how many people, who holds it, and what three years of occupation have done to it where that applies \u2014 followed by what the place is actually like and at least one reason for a party to go there. Populations for Elduvaine are given as they were before the fall, with the present figure after, because the difference is the campaign."));

c.push(P("Distances matter in this campaign more than in most, because the march is half of it. The table gives the legs the coalition actually travels. Times assume an army; a party of six moving alone and pushing hard does most of these in two thirds of the time, and the Willing Road does whatever the Willing Road does."));

c.push(table(
  ["Leg", "Time", "Notes"],
  [34, 16, 50],
  [
    ["Duncarrow to Ellendrake", "6 days", "Metalled road the whole way, the only one Harrowmark has. Xavier paid for it and mentions this."],
    ["Ellendrake to Caerwyn, by sea", "3 weeks", "With fair weather, which the coalition did not get. Storms scatter fleets onto Calanthe."],
    ["Duncarrow to Ashgate Ford", "12 days", "Good country, then hill country, then Baron Vell."],
    ["Ashgate Ford to the Cold Stair", "5 days", "Climbing the whole way. The last day is the pass itself."],
    ["The Cold Stair to Caerwyn", "3 weeks", "Down the far side and along the coast. Longer than the sea road and considerably less wet."],
    ["Caerwyn to Vindana", "9 days", "Four, on the days the Willing Road agrees to it. Nobody can arrange those days."],
    ["Vindana to Caer Ysolde", "11 days", "Upriver along the Braid. The last three days are inside Vale\u2019s reach and everybody knows it."],
    ["Caerwyn to Caer Ysolde, direct", "16 days", "Through the Orchard Marches. Beautiful, and in year three, wrong."]
  ]
));

// ================================================================ HARROWMARK
c.push(H1("Harrowmark"));

c.push(P("Hill country, thin soil, and stone that will take a wall better than it will take a plough. Harrowmark occupies a wedge of high ground between two ranges and has spent nine hundred years being a place nobody else particularly wanted, which is the entire secret of its survival and most of its personality. The land does nothing. The people do everything, unhurriedly, and are quietly certain this makes them the only serious kingdom on the continent."));

c.push(P("It is wyvern country. It has always been wyvern country. Every hold above the eight-hundred-foot line keeps a watch, every watch keeps a pike-rack, and every year some number of people do not come back down. Harrowmark has arranged its architecture, its livestock, its calendar and a good deal of its humour around this, and considers the arrangement normal."));

c.push(P("A visitor\u2019s first impression is grey \u2014 grey stone, grey weather, grey wool \u2014 and the second impression, arriving about four days later, is that everything in it is built extremely well. Nothing here is decorative. Everything here is still standing."));

c.push(...SITE("Duncarrow", "Population 14,000 \u00B7 Seat of Xavier III \u00B7 The muster ground of the crusade"));

c.push(P("The capital sits on a spur above the Draymere with the Marchhold at its head, and it is less a city than a fortress that acquired a town by accident and never quite decided what to do about it. Three concentric walls, all of them still garrisoned, all of them raised at different centuries and none of them demolished, so the streets run in rings and a stranger walking to the keep will pass the same bakery twice and begin to suspect the city of doing it on purpose."));

c.push(P("The Long Stair climbs from the river gate to the Marchhold in four hundred and ten steps, and by ancient custom a king of Harrowmark walks it rather than rides. Xavier has walked it three times a week for eleven years and can tell you what each landing smells like. Below it, the mustering fields have held eight thousand people since the Call and now smell of exactly that."));

c.push(P("The Ninefold House of Duncarrow is the largest Concord chapter in the kingdom and the place the Call was read from, in Ninefold Cant, on a wet morning to a crowd that mostly could not understand a word of it and understood the whole thing perfectly."));

c.push(BUL("Hook.", "Module One happens here. Beyond that: the Marchhold\u2019s lower vaults hold two centuries of Harrowmark\u2019s correspondence with Elduvaine, unread, and somebody in the Archive at Caer Ysolde wrote back."));

c.push(...SITE("Greywatch", "Population 900 \u00B7 The high wyvern-watch \u00B7 Huntmaster Brenna Vane"));

c.push(P("Eleven hundred feet up, four days from Duncarrow, and the reason the eastern valleys have livestock. Greywatch exists to kill wyverns, has existed to kill wyverns for nine centuries, and is very good at it in the way that only an institution with an unbroken casualty list can be. Dwarves take the pikes, because a pike-wall needs people who will not move. Orcs take the ropes, because the ropes go over the edge. Nobody at Greywatch regards either of those facts as being about dwarves or orcs; they are about pikes and ropes."));

c.push(P("The hold itself is one long hall dug back into the cliff with a stone yard in front of it, and the yard is where the hunt assembles, the dead are laid out, and the drinking happens, in that order and often on the same night. There is a wall inside the hall with names on it. It is not a memorial. It is a roster, and the dead are simply not crossed off."));

c.push(BUL("Hook.", "Brenna Vane will take anyone up who asks and is the campaign\u2019s best source of Harrowmark texture, camp levity, and blunt opinion. The wyvern-riders who go to Vindana with the army are hers, and she has views about lending them."));

c.push(...SITE("Ellendrake", "Population 6,000 \u00B7 Harrowmark\u2019s only true port \u00B7 The fleet sails from here"));

c.push(P("A working harbour town on a coast that does not want one, kept open by a mole three centuries old and a dredging levy everyone complains about and everyone pays. Ellendrake has never been beautiful and has stopped attempting it. What it has is deep water, a tide that behaves, and more rope, tar, salt fish and profane opinion per acre than anywhere else in the kingdom."));

c.push(P("Since the Call it has been the most crowded place in Harrowmark. Ninety-one hulls of wildly varying quality, requisitioned from four owners who are still arguing about compensation, and an Auberitz harbourmaster who arrived uninvited, took one look, and began reorganising the entire anchorage with the calm of a woman who has done this before."));

c.push(BUL("Hook.", "Module 2A sails from here. Before that: something in the requisitioned tonnage is not seaworthy, the owner knows, and the manifest says otherwise."));

c.push(...SITE("Stannock", "Population 2,400 \u00B7 Dwarf crag hold \u00B7 The armouries"));

c.push(P("Cut back into the western crags and reached by a road that switches back on itself eleven times, Stannock makes mail, pike-heads, and the particular long-hafted wyvern spear that no other kingdom has any use for. It has been making them for six hundred years, the pattern has changed twice, and both changes are still discussed."));

c.push(P("The crusade emptied the place. Stannock worked through two winters to arm the muster and then sent four hundred of its own with the result, which is a quarter of everyone it had. The forges are banked, the halls are quiet, and the people left are mostly over ninety or under twenty and all of them are waiting."));

c.push(BUL("Hook.", "The armourers kept the pattern-book. Anything the party needs made to Harrowmark standard can be made here \u2014 slowly, correctly, and with a great deal of comment about how it will be misused."));

c.push(...SITE("Corrieholt", "Population 1,800 \u00B7 Orc high country \u00B7 Horses, and the hill-fighting tradition"));

c.push(P("The high pasture north of Greywatch, held by orc families who have grazed it for as long as Harrowmark has had records and rather longer than Harrowmark has had a king. Corrieholt breeds the only horses in the kingdom worth the name \u2014 small, ugly, appallingly durable animals that will go up a slope no sensible mount would look at \u2014 and it supplies the light horse that the coalition has been quietly grateful for at every river crossing since the landing."));

c.push(P("It also supplies a fighting tradition: fast, dispersed, and built entirely around not being where the enemy expected. Auberitz officers find it undisciplined. Auberitz officers have also stopped saying so out loud since the second week of the march."));

c.push(BUL("Hook.", "A Corrieholt rider in the coalition has family in the Sixth Free Legion. This is true, it is known, and nobody in the column has made anything of it, which is exactly the campaign\u2019s position on the matter."));

c.push(...SITE("Fenmarrow", "Population 3,100 across a dozen villages \u00B7 Thin farm country \u00B7 Where the levy comes from"));

c.push(P("The low ground south of Duncarrow, and the only part of Harrowmark that grows enough of anything to matter. It is not good land. It is worked land, in narrow strips, by families who have improved it inch by inch for three centuries and will tell you the history of a particular hedge if you give them any opening at all."));

c.push(P("Most of the crusade\u2019s infantry is from here. Not the knights and not the professionals \u2014 the eight thousand. They answered a summons read in a language they did not speak, promising a place in a country they could not find on a map, and they are marching to it, and the DM should remember that every time the coalition is discussed as a political object."));

c.push(BUL("Hook.", "Somebody\u2019s mother in Fenmarrow gave the party a letter for her son in the column. He is in the lost column on the other road. This costs nothing to set up in Module One and pays for the rest of the campaign."));

c.push(...SITE("Kettleburn", "Population 400 \u00B7 A village with a scar \u00B7 Two days off the Duncarrow road"));

c.push(P("Eleven years ago a wyvern nested above Kettleburn and stayed for a season before Greywatch could bring it down, and the village has been the standing Harrowmark illustration of what that costs ever since. Half the houses were rebuilt. The rebuilt half is better than the original. The people are neither traumatised nor especially interested in discussing it, and a visitor who tries for pathos will be given a very dry account of roof repairs."));

c.push(BUL("Hook.", "The best low-stakes scene in Harrowmark for establishing what these people are like before the party leaves. They are not brave. They are unimpressed, and it is not the same thing."));

c.push(H2("Encounters in Harrowmark"));

c.push(P("Roll d12 once per day of travel, or when the pace wants breaking. Harrowmark is a settled kingdom in the middle of a muster: most of what a party meets is people."));

c.push(table(
  ["d12", "Encounter"],
  [12, 88],
  [
    ["1", "A wyvern (SRD, CR 6), hunting, and not remotely interested in a fair fight. It will take a pack animal and leave if allowed to."],
    ["2\u20133", "A levy company of forty walking to Duncarrow, badly, singing. Their sergeant is nineteen and terrified and hiding it well."],
    ["4", "Four bandits (SRD, CR 1/8) who used to be a levy company and thought better of it. They will surrender to anyone competent."],
    ["5", "A Concord priest of Voran walking the other way, who will feed anybody and asks nothing except news, which she then writes down."],
    ["6", "A Greywatch rope-team, off duty, three days into an argument about a knot. They will absolutely make it the party\u2019s problem."],
    ["7", "An Auberitz supply column, stuck. A wheel, a ford, a permit, or all three."],
    ["8", "Weather. Harrowmark weather, which means a full day lost and everyone soaked to the bone and unsurprised."],
    ["9", "A Norvatch factor riding to Duncarrow, ahead of the market, ledger already open."],
    ["10", "A pack of six wolves (SRD, CR 1/4) working the road because the flocks moved with the muster."],
    ["11", "A hill giant (SRD, CR 5) down out of the crags in a bad year, more hungry than hostile, and negotiable if anyone thinks to try."],
    ["12", "Riders of the Tenth Work, courteous, well-supplied, and extremely interested in where the party is going and why."]
  ]
));
// ================================================================ THE SEA ROAD
c.push(H1("The Sea Road"));

c.push(P("Three weeks of open water from Ellendrake to the Elduvish coast, in ninety-one hulls of which perhaps sixty were built for it. The sea road is faster, cheaper, and carries the siege train, which is the argument that won the council. It is also the road on which a storm can undo in one night what four months of organising achieved, and everybody who voted for it knew that too."));

c.push(...SITE("Calanthe", "Population 2,200 \u00B7 An independent island \u00B7 Warden Ivor Thane, self-styled"));

c.push(P("A hook of rock and gorse eleven miles long, lying inconveniently across the shipping lane, answering to no crown and never having been worth the cost of taking. Calanthe has one harbour, four hundred years of wrecking, and a population that has grown entirely comfortable with the distinction between salvage and theft being a matter of where you are standing."));

c.push(P("Ivor Thane holds it from a hall above the harbour and calls himself Warden, which is a title Calanthe invented for him and Calanthe can uninvent. He is not a pirate \u2014 he is careful about this \u2014 he is a man who charges for rescue, keeps what washes up, and has never once been proved to have moved a light. The distinction has kept him alive for nineteen years and will not survive contact with a coalition that needs its supplies back."));

c.push(BUL("Hook.", "Module 2A. Beyond it: Thane\u2019s cellar holds four hundred years of salvaged cargo manifests, and three of them are Elduvish, and one of them is from a ship that sailed after the wards opened."));

c.push(...SITE("The Ossary Shoals", "A wrecking coast \u00B7 Nine miles of it \u00B7 No population, by design"));

c.push(P("The shoals run north from Calanthe under six feet of water at the low tide and none at all at the high, and they are the reason Calanthe eats. Local pilots can take a hull through at a walking pace. Local pilots charge accordingly, and the price is a matter of what the cargo is, which they will establish first."));

c.push(P("At the low spring tide the wrecks stand out of the water \u2014 forty or fifty of them across four centuries, masts and ribs and one intact stern-castle sitting upright in the sand with weed in its windows. Calanthe children play in it. Nobody on the island finds this remarkable and every visitor does."));

c.push(H2("Encounters at Sea"));

c.push(table(
  ["d10", "Encounter"],
  [12, 88],
  [
    ["1\u20132", "Weather, worsening. A day of it, and a real chance of separation from the fleet."],
    ["3", "A hull in trouble \u2014 sprung seams, and forty people on deck doing arithmetic about the boats."],
    ["4", "Calanthe pilots, alongside, cheerful, and offering a service the party did not ask for at a price they will not like."],
    ["5", "A Norvatch merchantman on her lawful business, who will trade news and refuses passengers."],
    ["6", "Four merrow (SRD, CR 2) working the fleet\u2019s wake for anything that goes overboard, including people."],
    ["7", "A becalming. Three days, no wind, eight thousand soldiers on ninety-one boats, and tempers."],
    ["8", "A giant octopus (SRD, CR 1) on a small boat, which is far worse than it sounds if the small boat is yours."],
    ["9", "Wreckage from the other road, which is impossible, and is not."],
    ["10", "A young kraken\u2019s wake \u2014 no encounter, just a mile of flat water and every sailor aboard gone very quiet."]
  ]
));

// ============================================================ THE MOUNTAIN ROAD
c.push(H1("The Mountain Road"));

c.push(P("Longer, drier, and it cannot carry the siege engines, which is why it carries everything else. The mountain road climbs from Harrowmark\u2019s eastern hills through the Cold Stair and comes down the far side into three weeks of coastal going. Its virtue is that no storm can scatter it. Its cost is Baron Vell, the pass itself, and the fact that everything on it walks."));

c.push(...SITE("Ashgate Ford", "Population 700 \u00B7 Baron Vell\u2019s toll-keep \u00B7 The last easy crossing"));

c.push(P("The river is wide, shallow, fast and cold, and there has been a toll on it since somebody first noticed that the alternative is eleven days upstream. The keep is squat, competently sited, and older than the barony it now supports; the village below it exists to feed the keep and to sell rope, boots and hot food to people who have just discovered how cold the water is."));

c.push(P("Vell\u2019s toll is legal. That is the difficulty with Vell. He holds the ford by a charter four generations old, the charter says what he may charge, and what he has done is notice that the charter never contemplated eight thousand people and a war."));

c.push(BUL("Hook.", "Module 2B. Also: the charter is in the keep, it is genuine, and there is a clause in it that Vell has never read because it is on the back."));

c.push(...SITE("Kir Halloway", "Population 1,100 \u00B7 A mountain trade town \u00B7 Nobody\u2019s, officially"));

c.push(P("Two days above Ashgate, wedged into a valley that gets four hours of direct sun in winter, and existing because two roads meet there and both of them are bad. Kir Halloway is dwarves, humans, and a startling number of tieflings whose families came up from Norvatch three generations back and stayed for the trade. It has one inn with eleven beds, four warehouses, and a market that operates on the third day of every week regardless of weather, war, or whether anybody has come."));

c.push(P("It is also the last place on the mountain road where a party can buy anything at all. The next four hundred miles are the pass, and then Elduvaine, and Elduvaine has permits."));

c.push(BUL("Hook.", "A Norvatch house factor keeps a permanent office here and has done for sixty years. She knows what has been going up the road and, more usefully, what has been coming down it out of Elduvaine."));

c.push(...SITE("The Cold Stair", "The pass \u00B7 5,100 feet \u00B7 Open perhaps five months a year"));

c.push(P("A stair in name and nearly in fact: eleven miles of switchbacks cut into the eastern face, wide enough for two carts abreast where somebody widened it and wide enough for one where nobody did. The wind comes across it rather than along it. There is no water for the middle four miles, no shelter for the last three, and a set of stone refuge huts, most of them roofless, that a party will be extremely glad of at least once."));

c.push(P("At the summit there is a cairn. Travellers add a stone. It has been there long enough to be twenty feet across and it is the only thing anyone has ever built on the Cold Stair that the mountain has not taken back."));

c.push(BUL("Hook.", "Somebody has been robbing the refuge huts of their roofing timber, which is a small crime and, at five thousand feet in the wrong month, a lethal one."));

c.push(...SITE("The Old Workings", "Abandoned mines above Ashgate \u00B7 Kobold-held \u00B7 Older than the barony"));

c.push(P("Silver, four hundred years ago, and not enough of it. The workings run further into the hill than any surviving map shows and have been held for two generations by a kobold warren that arrived when the last human left and has been improving the place ever since \u2014 which is to say, trapping it comprehensively, signposting every trap in Draconic, and being extremely reasonable with anyone who reads Draconic."));

c.push(BUL("Hook.", "They want three things: to be left alone, a written guarantee of it, and somebody to do something about what is in the flooded lower gallery. They will pay for the third in silver they have no use for."));

c.push(...SITE("Barrowfell", "An upland of graves \u00B7 No living population \u00B7 Two days off the road"));

c.push(P("Four hundred barrows across nine miles of moor, raised by people nobody can now name, on a schedule nobody can reconstruct. Harrowmark\u2019s scholars think they predate the kingdom by a wide margin. Harrowmark\u2019s shepherds think they should be left alone, and have been proved right on eleven recorded occasions."));

c.push(PS([DM("DM Only: "), { t: "Barrowfell is the campaign\u2019s one piece of deliberate deep background and it has nothing whatever to do with Vale, Elduvaine, or the war. That is the point. A setting where every ruin is plot is a setting with no history in it. Put something genuinely old and genuinely unrelated down here, let the party spend an optional session on it, and never connect it to anything." }]));

c.push(H2("Encounters on the Mountain Road"));

c.push(table(
  ["d12", "Encounter"],
  [12, 88],
  [
    ["1", "Weather at altitude, which is a hazard rather than an encounter: exhaustion, lost days, and a real question about the refuge huts."],
    ["2\u20133", "A goat-track. It is shorter. It is also not what it looks like, and the party will be committed before they know."],
    ["4", "Six kobolds (SRD, CR 1/8) on a scouting loop, who would much rather talk and have prepared a speech."],
    ["5", "An ogre (SRD, CR 2) who has claimed a refuge hut and considers it his, which under the local rules of the road it now sort of is."],
    ["6", "A stranded Auberitz waggon, its crew, and a mathematical certainty that it cannot be got over the pass."],
    ["7", "Two griffons (SRD, CR 2) nesting above the switchbacks, hunting the pack animals."],
    ["8", "Pilgrims of the Tenth Work walking to Elduvaine unarmed, having refused an escort, three days from dying of it."],
    ["9", "A rockfall across the road. Four hours to clear, or a day to go around, and somebody is watching to see which."],
    ["10", "A troll (SRD, CR 5) under a bridge, doing the traditional thing, and genuinely open to the traditional solutions."],
    ["11", "Cold. Simple, unglamorous cold \u2014 a night that costs the party a level of exhaustion apiece unless somebody plans."],
    ["12", "The lost column\u2019s first courier, coming the wrong way, with news he does not want to be carrying."]
  ]
));
// ================================================================= THE VAUNT
c.push(H1("Elduvaine: The Vaunt"));

c.push(P("The coastal province, and the first Elduvaine anyone in this campaign sees. The Vaunt is low, green, wet in the good way, and threaded with tidal creeks that the halfling parishes have been navigating in flat-bottomed boats since before there were roads to prefer. It holds the kingdom\u2019s great port, its landing beaches, and about a fifth of its people."));

c.push(P("It is also where the occupation is heaviest, because it is where the goods leave from. A party arriving in the Vaunt is arriving in the most administered part of Elduvaine: permits, checkpoints, published levies, and a Norvatch counting-house three streets back from every quay."));

c.push(...SITE("Vindana", "Population 28,000, now perhaps 19,000 \u00B7 The great port \u00B7 Held by Marshal Ossian Drell"));

c.push(P("Elduvaine\u2019s window on the world, and the second city of the kingdom by every measure except the one that matters, which is that Caer Ysolde has the Archive. Vindana is built on a rise above a deep-water harbour, walled twice, and the walls are the reason the campaign spends two whole modules here. The inner wall is Standing Light stone: it held daylight, and for six centuries the city glowed faintly from its own defences on any clear night. It does not now."));

c.push(P("The harbour front is warehouses, rope-walks, and the smell of a working port. Behind it the city climbs in terraces to the Ward Gate and the garrison quarter, and behind that is the undercity \u2014 a drain and cistern system older than the walls above it, which the garrison stopped patrolling two years ago because something else moved in."));

c.push(P("Under the occupation Vindana has become the machine that empties Elduvaine. Everything that leaves the kingdom leaves through here, weighed, docketed, and carried in Norvatch bottoms under a standing contract. The people who load it are Elduvish and are paid."));

c.push(BUL("Hook.", "Modules 6, 7 and 8. Before and around them: the undercity, the counting-house, and Ottoline Vahn filing objections from what she insists on calling her chambers."));

c.push(...SITE("Caerwyn", "Population 1,600 \u00B7 The landing town \u00B7 Occupation grip: thin"));

c.push(P("A small coastal town of grey-gold stone and slate, eleven miles of orchard behind it, and a hedge at the top of the town that has sprites in it and always has. Caerwyn is far enough from both Vindana and the capital that the occupation posted a checkpoint, a clerk and eight guards and considered the matter closed, which is why the coalition chose it."));

c.push(P("It is the first wonder the party gets and it should be run as one. The light-stone in the doorsteps still holds a little of the afternoon. The baker\u2019s ovens have not gone out in two hundred years and the whole street smells of it at four in the morning. This is a place worth saving, and the module says so by showing rather than arguing."));

c.push(BUL("Hook.", "Module 3. Wyn Alder clerks the checkpoint and is the campaign\u2019s first honest look at what an ordinary complicit person is actually like."));

c.push(...SITE("The Dead Mile", "A mile of road \u00B7 Nothing lives on it \u00B7 Between Caerwyn and the interior"));

c.push(P("A stretch of the coast road where the draining was done early, thoroughly, and without any attempt at subtlety, and where the land has simply stopped. Not blasted \u2014 that would be easier to look at. The hedges are there. The grass is there. None of it is doing anything: no insects, no birds, no smell, and a silence that a party notices about four hundred yards in and cannot stop noticing afterwards."));

c.push(P("Wagons go through it at a trot. Nobody camps in it. There is no monster in the Dead Mile and there is nothing to fight, which is exactly why it is the most frightening place in the first half of the campaign."));

c.push(...SITE("Morgarth", "Population 2,900, now 800 \u00B7 A fishing town, emptied \u00B7 Two days south of Vindana"));

c.push(P("The Vaunt\u2019s second port, and the place the occupation made an example of in the first year \u2014 not by massacre, which would have been off-pattern, but by levy. Morgarth\u2019s boats were requisitioned, its catch was docketed, and its people were charged for the privilege of the arrangement until two thirds of them walked inland. It was entirely lawful. Every step of it is in the register."));

c.push(P("What remains is eight hundred people in a town built for three thousand, working the boats that were left, and a harbourmaster who has kept every single piece of paper."));

c.push(BUL("Hook.", "Those papers are the second-best account of the occupation\u2019s arithmetic in Elduvaine, after Norvatch\u2019s. Unlike Norvatch\u2019s, they are free, and the harbourmaster has been waiting three years for somebody to ask."));

c.push(H2("Encounters in the Vaunt"));

c.push(table(
  ["d12", "Encounter"],
  [12, 88],
  [
    ["1\u20132", "A checkpoint. Permits, a bored clerk, four guards, and the entire scene turning on whether anyone in the party can be boring convincingly."],
    ["3", "A Legion patrol of six, professional, uninterested, and willing to be talked past."],
    ["4", "A carter with a docketed load and a very careful account of exactly how legal it all is."],
    ["5", "A halfling family on a flat-bottomed boat who will move the party through the creeks for nothing and talk the entire way."],
    ["6", "Four sprites (SRD, CR 1/4) in a hedge, invisible, keeping up a running commentary."],
    ["7", "A stretch of road that is longer than it was. No encounter. Just the distance, quietly wrong."],
    ["8", "Three Norvatch carters, on schedule, who will not stop and will answer questions shouted alongside."],
    ["9", "Two ghouls (SRD, CR 1) in a drained orchard \u2014 the draining kills things, and not everything that dies in the Vaunt stays down."],
    ["10", "A Tenth Work chapter house going up on ground that had an Observance on it last year."],
    ["11", "An Unbound Clerk, hiding, half-starved, carrying four pages he will die rather than hand over."],
    ["12", "A drained dryad \u2014 see the Bestiary \u2014 no longer able to leave a wood that no longer exists."]
  ]
));

// ================================================================== THE BRAID
c.push(H1("Elduvaine: The Braid"));

c.push(P("Three rivers come down out of three different countries and meet in one valley, and everything Elduvaine is came out of that fact. The Braid is the heartland: the capital, the Archive, the oldest parishes, the seat of the house, and the densest concentration of resident magic anywhere in the world."));

c.push(P("It is also where Vale is, which means the last third of the campaign happens here and the party will spend it being extremely careful."));

c.push(...SITE("Caer Ysolde", "Population 41,000, now perhaps 30,000 \u00B7 The capital \u00B7 Held by Maedoc Vale in person"));

c.push(P("The city stands on and between the three rivers, on nine islands and both banks, connected by more bridges than anyone has successfully counted \u2014 the standing figure is sixty-one and it has been sixty-one for two hundred years regardless of how many have been built since. It is the most beautiful city in the world and it is not close, and every visiting ambassador for six centuries has said so in writing, usually while complaining about the damp."));

c.push(P("Caer Ysolde is built of light-stone almost entirely. On a clear night before the fall the whole city glowed the colour of late afternoon, from the ground up, bright enough to read by on the bridges and dim enough to see stars through. That is what it was famous for. That is what a Harrowmark levyman has heard about since he was a child."));

c.push(P("It is dark now. Not lightless \u2014 the occupation burns lamps like anybody else \u2014 but the stone is out, and it went out over about fourteen months, quarter by quarter, in a way the people who live there could watch happening."));

c.push(BUL("Hook.", "Module 11, and the whole campaign points here. Maelis Ysolde is held in her own apartments in the Ysolde Keep on the middle island."));

c.push(...SITE("The Ysolde Archive", "The largest collection of magical knowledge in the world \u00B7 Vale\u2019s, now"));

c.push(P("Nine floors above ground and an unstated number below, occupying most of the middle island, and governed for four hundred and eighty years by the single decision that makes this campaign possible: access by rule, not by power. A farmer with a legitimate question could read what the rule allowed a farmer to read. A monarch could not read past what the rule allowed a monarch. Neither could the Keeper."));

c.push(P("The reading rooms are still open. That is the detail that unsettles people most. Vale has kept the Archive functioning \u2014 clerks at the desks, requests filed, materials fetched \u2014 because he is not a vandal and because a working Archive is more use to him than a looted one. Elduvish scholars still come. Some of them still work there. They are paid."));

c.push(PS([DM("DM Only: "), { t: "what is in the deepest vaults has not been decided and must not be invented. The sourcebook states Vale\u2019s motive \u2014 he was refused, by rule, and he is spending a kingdom to finish reading \u2014 and it deliberately does not state his destination. Describe the doors. Describe the rule. Do not describe what is behind them. If the party gets that far, the campaign is more interesting for the DM answering it at their own table than for this book answering it here." }]));

c.push(PS([DM("DM Only: "), { t: "one dragon older than the current wards sleeps somewhere near or under the Archive and answers to neither Vale nor the party. This is a single deep-vault exception, it is not a general rule about Elduvish dragons, and it does not answer the question above. It sleeps until the campaign needs it to stop." }]));

c.push(...SITE("Sennoch Hall", "A great house of the Braid \u00B7 Two days downriver of the capital \u00B7 A gilded prison"));

c.push(P("A country seat of the house, walled, moated in the ornamental sense, and entirely unsuited to being a place of detention \u2014 which is exactly why the occupation chose it. Ninian Ysolde has been held here for three years in considerable comfort, with a library, a garden, a staff who are not permitted to leave, and a garrison that is embarrassed about the whole arrangement."));

c.push(BUL("Hook.", "Module 4. The garrison commander here has been correct, courteous and humane for three years and would very much like somebody to notice that before this ends badly for him."));

c.push(...SITE("Lisswater", "Population 2,100 \u00B7 A river parish \u00B7 Halfling country, and the kingdom\u2019s manners"));

c.push(P("Where the middle river runs slow for nine miles, there are eleven villages that consider themselves one place, and Lisswater is the name of all of them and none of them. It is halfling country: locks, weirs, eel traps, orchards on the levee, and the single best food in Elduvaine by a margin that the rest of the kingdom concedes without argument."));

c.push(P("The Listening Water is at its strongest along this stretch, and Lisswater has spent nine hundred years developing an etiquette around it so thorough that outsiders find it incomprehensible. You do not speak at the water\u2019s edge about anything you would not repeat. You do not stand at the edge when angry. Children are taught the rule before they are taught to swim, and adults who break it are not scolded but quietly and permanently regarded as unserious."));

c.push(BUL("Hook.", "The whole of the Four Voices technique works best here. So does the miller who has been miscounting the levy for two years and would like to stop being the only person doing it."));

c.push(...SITE("The Threefold Stair", "Where the rivers meet \u00B7 A mile below the capital \u00B7 An engineering work and a shrine"));

c.push(P("The three rivers do not simply join. They arrive at different heights, and nine centuries ago somebody built the Stair to marry them: three broad flights of worked channel, dropping the two higher rivers to the level of the third across about half a mile, in a controlled and continuous roar that can be heard from the city walls on a still night."));

c.push(P("It is the largest single piece of Standing Light work in Elduvaine, and at night the falling water used to carry the glow down with it. Every Observance in the Braid ends at the Stair. It is not holy. It is simply the place everybody goes."));

c.push(H2("Encounters in the Braid"));

c.push(table(
  ["d12", "Encounter"],
  [12, 88],
  [
    ["1", "A Legion officer of rank, escorted, on the road with a clear purpose and an interest in who else is."],
    ["2\u20133", "A permit inspection at a bridgehead \u2014 and there are sixty-one bridges, so this is the Braid\u2019s weather."],
    ["4", "An Observance in progress at a wayside. Twenty people, no priest, and a party that interrupts will be forgiven and remembered."],
    ["5", "Resistance: a signal chalked on a lock gate, meaning something the party could learn to read."],
    ["6", "A collaborator clerk who wants, desperately and dangerously, to give somebody a document."],
    ["7", "Four Legion scouts (Ashgate Scout block) shadowing the party for a full day before deciding anything."],
    ["8", "A stretch of Willing Road that shortens dramatically and inexplicably, in the party\u2019s favour, for no reason anyone will ever establish."],
    ["9", "A will-o\u2019-wisp (SRD, CR 2) over the flooded lower parishes, which the locals have a name for and do not fear correctly."],
    ["10", "An Unbound Clerk\u2019s dead drop, still stocked, its owner three weeks gone."],
    ["11", "The Tenth Work, consecrating ground, entirely lawfully, while the parish stands and watches."],
    ["12", "One of Vale\u2019s own \u2014 a battle-mage on the road, alone, unhurried, and not looking for a fight but entirely equal to one."]
  ]
));

// ====================================================== THE ORCHARD MARCHES
c.push(H1("Elduvaine: The Orchard Marches"));

c.push(P("The upland between the Braid and the northern hills, and the institution that made the elves of Elduvaine what they are. A Kept Season wood holds the season it was planted in, forever, and the Marches are four hundred years of people planting deliberately: a spring wood here, a high-summer wood there, an autumn stand kept for the timber it gives in that state, laid out across ninety miles like an argument nobody has finished."));

c.push(P("Walking the Marches in the old days meant crossing from full blossom into deep autumn in the space of a hedge, four times before lunch, and the roads were laid out to make the transitions land well. It was the single most-visited thing in Elduvaine and the kingdom made a great deal of money from people who wanted to see it."));

c.push(...SITE("Bryn Aeling", "The oldest orchard \u00B7 Planted four days into spring, seven hundred years ago \u00B7 Now nine weeks into a winter"));

c.push(P("The first Kept Season planting and the greatest of them, forty acres of birch and apple that has been four days into spring since before Harrowmark had a king. Generations of Elduvish came here to be married under it, because a wood four days into spring is four days into spring on the morning of your wedding regardless of what the calendar says, and that was the whole point."));

c.push(P("It is nine weeks into a winter it was never sown in, and has been for most of a year. The trees are dying \u2014 not dead, dying, slowly and in the wrong order \u2014 and an ecology has moved in behind the change: winter wolves along the edges, and a troll in the orchard-keeper\u2019s cottage, and the dryad who has been the spirit of this wood for three centuries and is now the spirit of a winter one and cannot leave."));

c.push(BUL("Hook.", "Module 5\u2019s Held Winter. The single most legible image of the draining in the campaign, and the one scene where a party can simply sit down and talk to the grief."));

c.push(...SITE("Nantcorrow", "Population 1,900 \u00B7 An orchard town \u00B7 The Season-keepers\u2019 seat"));

c.push(P("Where the Marches are administered from, insofar as they are administered at all: a town of long low halls and drying-lofts, given over almost entirely to fruit, timber, and the extremely serious business of deciding where the next wood goes and in which week it should be sown. A planting decision at Nantcorrow is a four-hundred-year decision and is argued like one."));

c.push(P("The Season-keepers are still here. They have not been arrested, because arresting them would stop the orchards producing, and the occupation would like the orchards to keep producing. They have therefore spent three years administering an institution on behalf of a man who is killing it, which is a position several of them can no longer live with."));

c.push(BUL("Hook.", "The Keepers know exactly which woods have turned and in what order. That map is the best measurement of the draining that exists outside Norvatch\u2019s ledgers, and it is not for sale, and they will give it to the right people."));

// ===================================================== THE STANDING MARCHES
c.push(H1("Elduvaine: The Standing Marches"));

c.push(P("Hill country west of the Braid, and the source of the other thing Elduvaine is famous for. Worked stone from these quarries holds light poured into it \u2014 an afternoon\u2019s sun, a lamp, a spell \u2014 and gives it back for hours or days depending on the cut. The whole of Caer Ysolde is built of it, and so is every doorstep in Caerwyn, and so is the inner wall of Vindana."));

c.push(...SITE("Cairn Ithel", "Population 3,400 \u00B7 The light-stone quarries \u00B7 Gnome country"));

c.push(P("Six working quarries and a town wedged between them, gnome-run for six hundred years, with a mason\u2019s guild whose apprenticeship is eleven years and whose masters can look at a face of rock and tell you how long a block cut from it will hold an evening. The cut is everything. The same stone cut two ways holds light for a day or for a week, and the difference is a craft secret that has never been written down."));

c.push(P("The occupation has the quarries working double. What it is producing is not building stone."));

c.push(PS([DM("DM Only: "), { t: "the party can work out what the quarries are producing for, and it is one of the campaign\u2019s better mid-game discoveries: light-stone cut to hold, then drained, is the most efficient way anyone has found to move Elduvaine\u2019s resident magic in a cart. Cairn Ithel is not being worked. It is being packaged." }]));

c.push(BUL("Hook.", "A master mason at Cairn Ithel has been deliberately cutting badly for two years \u2014 stone that holds for an hour instead of a week \u2014 and has cost the occupation a great deal, and knows exactly how long she has."));

c.push(...SITE("Ysgaron", "Population 1,200 \u00B7 A mason town \u00B7 Where the Standing Light was learned"));

c.push(P("Smaller, older, and higher than Cairn Ithel, and the place the craft was worked out, six centuries ago, by people whose names are on a wall in the guild hall. Ysgaron does the fine work: the cut faces, the inlays, the lamps that are not lamps. Most of what makes Caer Ysolde beautiful was made here by about nine hundred people over six hundred years."));

c.push(BUL("Hook.", "The guild hall\u2019s wall of names is also, unintentionally, a four-hundred-year technical record. Anyone who can read it can work out how the cut is done. Vale has not been here yet."));

c.push(H1("The Willing Road"));

c.push(P("It runs from the coast at Caerwyn to the capital and on north to the Marches, marked at intervals by waystones nobody can date, and it shortens for travellers who mean well. That is all four hundred years of scholarship has established, and it is not for want of effort \u2014 the Archive holds an entire shelf of failed attempts, most of them by people who were certain they had it."));

c.push(P("The waystones are unremarkable: waist-high, weathered, unlettered, and set at intervals that do not correspond to any unit anybody uses. There are three hundred and eleven of them. Nobody put them there. This is not a legend or an evasion; the Archive has looked, exhaustively, and there is no record of the road being built, no record of anyone claiming to have built it, and no record of a time before it."));

c.push(PS([DM("DM Only: "), { t: "see Travelling the Willing Road in the sourcebook. The essential rule bears repeating because it is the one thing in this setting that a well-meaning DM will ruin by being helpful: do not decide what it measures. Not even privately. The road is the only genuine mystery in a campaign whose villain, war and stakes are all fully explained, and it is load-bearing for exactly that reason." }]));

// ================================================== THE COALITION REALMS
c.push(H1("The Coalition Realms, in Brief"));

c.push(P("None of these are on the party\u2019s road, and all three are on the party\u2019s political map. What follows is what a well-briefed officer of the crusade would be able to tell you over dinner."));

c.push(H2("The Kingdom of Oksitan"));

c.push(P("River and horse country, four hundred miles of it, and every great house in it holds a ford. Humans in the main, with an old landed nobility that is substantially dragonborn \u2014 the houses of Vashkar, Ourrez, Sarrelan and Kaldiss between them hold every significant crossing on the Vaskren and have done since before the records begin. Oksitan is rich, well-mounted, and organised entirely around the control of movement."));

c.push(P("Its capital is Aurignan, on the Vaskren, which is less a city than a very large bridge with a city attached to both ends of it. Its king is Raimon V, who is seventy-one and has gone in person. Its army takes the road the party does not."));

c.push(H2("The Grand Duchy of Auberitz"));

c.push(P("The duchy that builds things. Humans, gnomes and halflings, mercantile and engineering rather than martial, wealthy in a quiet way that Oksitan finds vulgar and Harrowmark does not notice. Auberitz has no significant cavalry tradition, no chivalric literature to speak of, and the best siege train in the world \u2014 gnome-designed, halfling-quartermastered, and moved by people who have thought harder about roads than about glory."));

c.push(P("Its capital shares the duchy\u2019s name. Its arsenal is at Krenholt, which is where the engines that will open Vindana were built, over nine months, to a specification that assumed the walls would be worse than they are."));

c.push(PS([DM("DM Only: "), { t: "what Auberitz actually wants, beyond the Promise, has not been decided and is not invented here. Its role, its people, its engineers and its politics-of-logistics are all canon and all usable. Run its officers by role until the question is settled: they are competent, unimpressed by everyone else\u2019s war, and privately convinced this is a supply problem being mishandled by people who enjoy shouting." }]));

c.push(H2("The Kingdom of Norvatch"));

c.push(P("Guild-law country: dwarves and tieflings, humans throughout, and an entire realm whose standing in the world rests on the proposition that a Norvatch contract is honoured to the letter no matter what the letter turns out to have meant. It did not take the Call. It deals with either side. It answers to nobody, and it will keep a bargain it has made."));

c.push(P("Its great counting-city is Torvhal, where the Writ House sits \u2014 a court that has never in two centuries ruled on whether an agreement was fair, only on what it said. Norvatch does not want Elduvaine. Norvatch wants to remain the market for what leaves it, guaranteed in writing by whoever holds the place when the war ends, and it has held exactly that arrangement with the occupation for three years."));

c.push(BUL("Hook.", "House Kell\u2019s ledgers are the campaign\u2019s clock. Doria Kell will sell them. Her price is in Module 10 and it is worse than money."));

c.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300, after: 100 }, children: [new TextRun({ text: "~", size: 24 })] }));
c.push(PS([{ t: "\u201CThere are sixty-one bridges in Caer Ysolde. There have been sixty-one bridges in Caer Ysolde for two hundred years. We have built nine in that time and lost none, and it is still sixty-one, and we would all be grateful if you did not raise it again.\u201D", i: true }], { alignment: AlignmentType.CENTER }));
c.push(PS([{ t: "\u2014 the Bridgewarden of Caer Ysolde, to a visiting Auberitz surveyor", i: true }], { alignment: AlignmentType.CENTER }));
const doc = new Document({
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 260, hanging: 260 } } } }] }] },
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
  fs.writeFileSync(stagePath("KC_Gazetteer.docx"), buf);
  console.log("Written.");
});
