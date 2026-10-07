// Run with: node tests/rules.test.js
const R = require("../rules.js");
let pass = 0, fail = 0;
function eq(actual, expected, label) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (ok) pass++; else { fail++; console.log("FAIL:", label, "\n   got     ", JSON.stringify(actual), "\n   expected", JSON.stringify(expected)); }
}
const rules = (text, ctx, opts) => R.checkText(text, ctx, opts).map(i => i.rule);

// Table 23.2 examples from the DS Manual
[["3","10","10"],["71","72","72"],["96","117","117"],["100","104","104"],["1100","1113","1113"],
 ["101","109","9"],["808","833","33"],["1103","1104","4"],["321","328","28"],["498","532","532"],
 ["1087","1089","89"],["1496","1500","500"],["11564","11615","615"],["1984","2010","2010"]]
  .forEach(([s, e, want]) => eq(R.abbreviateRange(s, e), want, `range ${s}-${e}`));
eq(R.formatRange("642", "643"), "642–43", "formatRange en dash");
eq(R.formatRange("642", "643", { dash: "hyphen" }), "642-43", "formatRange hyphen");

// Scripture parsing and formatting (Citing the Bible in Turabian)
const p = s => R.parseReference(s);
eq(R.formatScripture(p("Jeremiah 31:4"), "paren", "NIV"), "(Jer. 31:4 NIV)", "Jer paren");
eq(R.formatScripture(p("jer 31:4"), "paren", "NIV", { abbrev: "short" }), "(Jer 31:4 NIV)", "Jer short");
eq(R.formatScripture(p("Prov 3:5a"), "paren", "NIV"), "(Prov. 3:5a NIV)", "Prov 3:5a");
eq(R.formatScripture(p("John 11:35"), "paren", "NIV"), "(John 11:35 NIV)", "John");
eq(R.formatScripture(p("Jn 11:35"), "running"), "John 11:35", "Jn running");
eq(R.formatScripture(p("1 thess 4:11, 5:2-5, 5:14"), "note"), "1 Thess. 4:11, 5:2–5, 5:14", "1 Thess list");
eq(R.formatScripture(p("I John 1:9"), "paren", null, { abbrev: "short" }), "(1 Jn 1:9)", "I John");
eq(R.formatScripture(p("1John 1:9"), "paren"), "(1 John 1:9)", "1John");
eq(R.formatScripture(p("Gal 5:22-23"), "note"), "Gal. 5:22–23", "Gal range");
eq(R.formatScripture(p("Psalm 23:1"), "paren", "ESV"), "(Ps. 23:1 ESV)", "Psalm");
eq(R.formatScripture(p("Psalms 23-24"), "note"), "Pss. 23–24", "Pss plural");
eq(R.formatScripture(p("Song of Songs 2:1"), "paren"), "(Song of Sol. 2:1)", "Song");
eq(R.formatScripture(p("Matthew 5"), "running"), "Matthew 5", "chapter only");
eq(R.formatScripture(p("Phil 4:13"), "paren"), "(Phil. 4:13)", "Phil");
eq(R.formatScripture(p("Philemon 1:6"), "paren"), "(Philem. 1:6)", "Philemon");
eq(R.parseReference("Rom 8:28 NRSV").version, "NRSV", "version parsed");
eq(!!R.parseReference("Hezekiah 3:1").error, true, "unknown book");
eq(R.BOOKS.length, 66, "66 books");

// Version rule
eq(R.versionDecision({}, "niv"), { show: true, addPrimaryFootnote: true, primary: "NIV" }, "first use");
eq(R.versionDecision({ primary: "NIV" }, "NIV"), { show: false, addPrimaryFootnote: false, primary: "NIV" }, "same");
eq(R.versionDecision({ primary: "NIV" }, "KJV"), { show: true, addPrimaryFootnote: false, primary: "NIV" }, "different");

// Title page layout: title line 8, name line 24
const lines = R.titlePageLines({ title: "Grace in Romans", subtitle: "A Study", name: "Jane Doe", course: "NT 501 New Testament", professor: "Dr. Smith", date: "October 5, 2026" });
eq(lines[7].text, "Grace in Romans:", "title line 8 with colon");
eq(lines[8].text, "A Study", "subtitle line 9");
eq(lines[23].text, "Jane Doe", "name line 24");
eq(lines[25].text, "NT 501 New Testament", "course line 26");
eq(lines[29].text, "October 5, 2026", "date line 30");
eq(R.titlePageLines({ title: "Only Title", name: "X" })[7].text, "Only Title", "no colon without subtitle");
eq(R.formatDate(new Date(2026, 9, 5)), "October 5, 2026", "date");

// Text checks
eq(rules("We read the bible daily.", "body"), ["cap-bible"], "lowercase bible");
eq(rules("This is a Biblical theme.", "body"), ["cap-lower"], "Biblical mid-sentence");
eq(rules("Biblical themes abound.", "body"), [], "Biblical at sentence start ok");
eq(rules("We trust the Holy Scriptures.", "body"), [], "Holy Scriptures ok");
eq(rules("In the gospel of Mark we see.", "body"), ["cap-gospel"], "gospel of Mark");
eq(rules("Jeremiah chapter 31 shows this.", "body"), ["bible-chapter-word"], "chapter word");
eq(rules("This is chapter 3 of my paper.", "body"), [], "no false positive 'is chapter'");
eq(rules("Love is patient (1 Cor. 13:4 NIV).", "body"), [], "good parenthetical");
eq(rules("As seen (verses 8-10) here.", "body"), ["bible-vv"], "verses in parens");
eq(rules("Read Gen. 1 and Ps 23 today.", "body"), ["bible-spell-out", "bible-spell-out"], "spell out chapters");
eq(rules("Psalm 23 and Matthew 5 and John 3 teach.", "body"), [], "full names ok");
eq(rules("The gospel spread.", "body"), [], "lowercase gospel ok");

// Footnotes
eq(rules("Klein, “Exegetical Rigor,” 33.", "footnote"), [], "good short note");
eq(rules("ibid., 43.", "footnote"), ["ibid-case"], "ibid lowercase");
eq(rules("Ibid, 43.", "footnote"), ["ibid-period"], "ibid no period");
eq(rules("Segal, A New Psalm, 42", "footnote"), ["note-period"], "missing period");
eq(rules("Hegel, Science of Logic, 642–643.", "footnote"), ["range"], "unabbreviated range");
eq(rules("Hegel, Science of Logic, 642–43.", "footnote"), [], "abbreviated range ok");
eq(rules("Mathewson, “Reading Heb 6:4-6,” Westminster Theological Journal 61 (1999): 209.", "footnote"), [], "verse range in title not flagged");
eq(rules("Garber, “Over,” 735, https://doi.org/10.1086/686960.", "footnote"), [], "doi not flagged");
eq(rules("Smith, Big Book, 1,234.", "footnote"), ["page-comma"], "comma in page number");
eq(rules("Klein, Ephesians, 23-36.", "footnote"), ["range"], "hyphen when en dash chosen");
eq(rules("Klein, Ephesians, 23-36.", "footnote", { dash: "hyphen" }), [], "hyphen ok when chosen");

// Bibliography
eq(R.checkBibliographyEntry("The Holy Bible: New International Version. Grand Rapids: Zondervan, 2011.").map(i => i.rule), ["bib-bible"], "bible in bib");
eq(R.checkBibliographyEntry("Author. “Note.” In The Orthodox Study Bible, edited by X. Nashville: Nelson, 2008.").map(i => i.rule), [], "study bible ok");
eq(R.checkBibliographyEntry("Merriam-Webster. s.v. “grace.”").map(i => i.rule), ["bib-common-dict"], "common dictionary");
eq(R.checkBibliographyEntry("Gomez, Sam. Email message to author, August 1, 2017.").map(i => i.rule), ["bib-personal"], "personal comm");
eq(R.checkBibliographyOrder(["Segal, Benjamin.", "Klein, William."]).length, 1, "order issue");
eq(R.sortBibliography(["Segal, B.", "“History.” Columbia", "Klein, W.", "The Making"]).map(s => s[0]), ["“", "K", "T", "S"], "sort ignores quotes/The");

// Parenthetical consistency
const found = R.findParentheticalRefs("Love (1 Cor. 13:4 NIV) and grace (Jn 3:16) and (Rom 8:28).");
eq(found.map(f => f.kind), ["traditional", "short", "short"], "detect styles");
eq(R.convertParenthetical(found[1], "traditional"), "(John 3:16)", "convert Jn");
eq(R.convertParenthetical(found[0], "short"), "(1 Cor 13:4 NIV)", "convert to short");

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
