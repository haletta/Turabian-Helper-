/*
 * Turabian Helper — rules engine
 * Pure functions with no Word dependency, so they can be unit-tested in Node.
 * Rules come from the Denver Seminary Writing Center guides:
 *   Turabian Formatting Checklist, Citing Common Sources in Turabian (2026),
 *   Guidance for Research Papers, Tips on Turabian (DS Manual), Title Page template.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.TurabianRules = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  /* ------------------------------------------------------------------ */
  /* Bible books (Citing the Bible in Turabian, 24.6.1 and 24.6.3)       */
  /* ------------------------------------------------------------------ */
  // [full name, traditional abbreviation, shorter abbreviation, extra aliases...]
  var BOOKS = [
    ["Genesis", "Gen.", "Gn", "Gen", "Ge"],
    ["Exodus", "Exod.", "Ex", "Exod", "Exo"],
    ["Leviticus", "Lev.", "Lv", "Lev"],
    ["Numbers", "Num.", "Nm", "Num"],
    ["Deuteronomy", "Deut.", "Dt", "Deut"],
    ["Joshua", "Josh.", "Jo", "Josh"],
    ["Judges", "Judg.", "Jgs", "Judg"],
    ["Ruth", "Ruth", "Ru"],
    ["1 Samuel", "1 Sam.", "1 Sm", "1 Sam"],
    ["2 Samuel", "2 Sam.", "2 Sm", "2 Sam"],
    ["1 Kings", "1 Kings", "1 Kgs", "1 Kgs."],
    ["2 Kings", "2 Kings", "2 Kgs", "2 Kgs."],
    ["1 Chronicles", "1 Chron.", "1 Chr", "1 Chron"],
    ["2 Chronicles", "2 Chron.", "2 Chr", "2 Chron"],
    ["Ezra", "Ezra", "Ezr"],
    ["Nehemiah", "Neh.", "Neh"],
    ["Esther", "Esther", "Est", "Esth", "Esth."],
    ["Job", "Job", "Jb"],
    ["Psalms", "Ps.", "Ps", "Psalm", "Pss", "Pss.", "Psa"],
    ["Proverbs", "Prov.", "Prv", "Prov"],
    ["Ecclesiastes", "Eccles.", "Eccl", "Eccles", "Eccl."],
    ["Song of Solomon", "Song of Sol.", "Sg", "Song of Songs", "Song", "Song of Sol"],
    ["Isaiah", "Isa.", "Is", "Isa"],
    ["Jeremiah", "Jer.", "Jer"],
    ["Lamentations", "Lam.", "Lam"],
    ["Ezekiel", "Ezek.", "Ez", "Ezek"],
    ["Daniel", "Dan.", "Dn", "Dan"],
    ["Hosea", "Hosea", "Hos"],
    ["Joel", "Joel", "Jl"],
    ["Amos", "Amos", "Am"],
    ["Obadiah", "Obad.", "Ob", "Obad"],
    ["Jonah", "Jon.", "Jon"],
    ["Micah", "Mic.", "Mi", "Mic"],
    ["Nahum", "Nah.", "Na", "Nah"],
    ["Habakkuk", "Hab.", "Hb", "Hab"],
    ["Zephaniah", "Zeph.", "Zep", "Zeph"],
    ["Haggai", "Hag.", "Hg", "Hag"],
    ["Zechariah", "Zech.", "Zec", "Zech"],
    ["Malachi", "Mal.", "Mal"],
    ["Matthew", "Matt.", "Mt", "Matt"],
    ["Mark", "Mark", "Mk"],
    ["Luke", "Luke", "Lk"],
    ["John", "John", "Jn"],
    ["Acts", "Acts", "Acts", "Acts of the Apostles"],
    ["Romans", "Rom.", "Rom", "Ro"],
    ["1 Corinthians", "1 Cor.", "1 Cor"],
    ["2 Corinthians", "2 Cor.", "2 Cor"],
    ["Galatians", "Gal.", "Gal"],
    ["Ephesians", "Eph.", "Eph"],
    ["Philippians", "Phil.", "Phil"],
    ["Colossians", "Col.", "Col"],
    ["1 Thessalonians", "1 Thess.", "1 Thes", "1 Thess", "1 Thes."],
    ["2 Thessalonians", "2 Thess.", "2 Thes", "2 Thess", "2 Thes."],
    ["1 Timothy", "1 Tim.", "1 Tm", "1 Tim"],
    ["2 Timothy", "2 Tim.", "2 Tm", "2 Tim"],
    ["Titus", "Titus", "Ti", "Tit"],
    ["Philemon", "Philem.", "Phlm", "Philem"],
    ["Hebrews", "Heb.", "Heb"],
    ["James", "James", "Jas"],
    ["1 Peter", "1 Pet.", "1 Pt", "1 Pet"],
    ["2 Peter", "2 Pet.", "2 Pt", "2 Pet"],
    ["1 John", "1 John", "1 Jn"],
    ["2 John", "2 John", "2 Jn"],
    ["3 John", "3 John", "3 Jn"],
    ["Jude", "Jude", "Jude"],
    ["Revelation", "Rev.", "Rv", "Rev", "Revelations"]
  ].map(function (row) {
    return { full: row[0], traditional: row[1], short: row[2], names: row };
  });

  function norm(s) {
    return s
      .toLowerCase()
      .replace(/\./g, "")
      .replace(/^(i{1,3})\s+/, function (m, r) { return r.length + " "; }) // I John -> 1 john
      .replace(/^(1st|first)\s+/, "1 ")
      .replace(/^(2nd|second)\s+/, "2 ")
      .replace(/^(3rd|third)\s+/, "3 ")
      .replace(/^([123])(?=[a-z])/, "$1 ") // 1John -> 1 john
      .replace(/\s+/g, " ")
      .trim();
  }

  var ALIAS = {};
  BOOKS.forEach(function (b) {
    b.names.forEach(function (n) { ALIAS[norm(n)] = b; });
  });

  function findBook(name) {
    return ALIAS[norm(name)] || null;
  }

  // Regex source matching any book name/alias (longest first so "1 John" beats "John").
  var bookAlternation = (function () {
    var all = [];
    BOOKS.forEach(function (b) {
      b.names.forEach(function (n) {
        var esc = n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/ /g, "\\s?");
        if (all.indexOf(esc) < 0) all.push(esc);
      });
    });
    all.sort(function (a, b) { return b.length - a.length; });
    return all.join("|");
  })();

  /* ------------------------------------------------------------------ */
  /* Dashes and page ranges (Table 23.2)                                 */
  /* ------------------------------------------------------------------ */
  var EN = "\u2013";

  function dashChar(opts) {
    return opts && opts.dash === "hyphen" ? "-" : EN;
  }

  // Abbreviate an inclusive number range per Turabian Table 23.2.
  function abbreviateRange(start, end) {
    var s = String(start), e = String(end);
    var a = parseInt(s, 10), b = parseInt(e, 10);
    if (isNaN(a) || isNaN(b) || b <= a) return e;
    if (a < 100 || a % 100 === 0) return e; // use all digits
    if (s.length !== e.length) return e;
    var k = 0;
    while (k < s.length && s[k] === e[k]) k++;
    if (a % 100 < 10) return e.slice(k); // 101-9, 808-33, 1103-4: changed part only
    return e.slice(Math.min(k, e.length - 2)); // 321-28, 498-532, 1496-500
  }

  function formatRange(start, end, opts) {
    return start + dashChar(opts) + abbreviateRange(start, end);
  }

  /* ------------------------------------------------------------------ */
  /* Scripture references                                                */
  /* ------------------------------------------------------------------ */
  // Parse "Jer 31:4", "1 Thess. 4:11, 5:2-5", "Romans 8", "Ps 23:1-6 NIV".
  function parseReference(input) {
    if (!input) return null;
    var text = String(input).trim().replace(/[\u2013\u2014]/g, "-");
    var m = text.match(/^((?:[123]|i{1,3}|1st|2nd|3rd|first|second|third)?\s*[A-Za-z][A-Za-z .]*?)\s*(\d.*)?$/i);
    if (!m) return null;
    var bookPart = m[1].trim();
    var rest = (m[2] || "").trim();
    var version = null;
    var vm = rest.match(/\s+([A-Z]{2,6})$/);
    if (vm) { version = vm[1]; rest = rest.slice(0, vm.index).trim(); }
    var book = findBook(bookPart);
    if (!book) return { error: 'I don\u2019t recognize the book "' + bookPart + '".' };
    if (rest && !/^[\d:,\-\s abv;]+$/i.test(rest)) {
      return { error: 'I can\u2019t read the chapter/verse part "' + rest + '".' };
    }
    return { book: book, passage: rest, version: version };
  }

  function formatPassage(passage, opts) {
    var d = dashChar(opts);
    return passage
      .replace(/\s*-\s*/g, d)
      .replace(/\s*,\s*/g, ", ")
      .replace(/\s*;\s*/g, "; ")
      .replace(/\s*:\s*/g, ":");
  }

  // style: "traditional" | "short" | "full"
  function bookName(book, style) {
    if (style === "short") return book.short;
    if (style === "full") return book.full;
    return book.traditional;
  }

  // Psalms: plural abbreviation when several psalms are cited (Pss. 23-24).
  function bookNameFor(book, passage, style) {
    if (book.full === "Psalms" && style !== "full" && passage && !/:/.test(passage) && /[-\u2013,;]/.test(passage)) {
      return style === "short" ? "Pss" : "Pss.";
    }
    return bookName(book, style);
  }

  /**
   * Build a scripture citation.
   * mode: "paren"   -> (Jer. 31:4 NIV)        quoting scripture
   *       "running" -> Jeremiah 31:4          naming a passage in your sentence
   *       "note"    -> Jer. 31:4              inside a footnote
   */
  function formatScripture(ref, mode, version, opts) {
    opts = opts || {};
    var style = mode === "running" ? "full" : opts.abbrev || "traditional";
    var passage = ref.passage ? formatPassage(ref.passage, opts) : "";
    var name = bookNameFor(ref.book, passage, style);
    var core = passage ? name + " " + passage : name;
    if (mode === "paren") return "(" + core + (version ? " " + version : "") + ")";
    if (mode === "note") return core + (version ? " " + version : "");
    return core;
  }

  // Decide whether to show the version, following the Writing Center rule:
  // first quotation names the version (and gets a footnote); afterwards only when it differs.
  function versionDecision(state, version) {
    var v = (version || "").toUpperCase().trim();
    if (!v) return { show: false, addPrimaryFootnote: false, primary: state.primary || null };
    if (!state.primary) return { show: true, addPrimaryFootnote: true, primary: v };
    return { show: v !== state.primary, addPrimaryFootnote: false, primary: state.primary };
  }

  function primaryVersionNote(version) {
    return "Unless otherwise noted, all scripture references are from the " + version + ".";
  }

  /* ------------------------------------------------------------------ */
  /* Title page                                                          */
  /* ------------------------------------------------------------------ */
  var MONTHS = ["January", "February", "March", "April", "May", "June", "July",
    "August", "September", "October", "November", "December"];

  function formatDate(d) {
    return MONTHS[d.getMonth()] + " " + d.getDate() + ", " + d.getFullYear();
  }

  // Lines of the title page, single-spaced, matching the Writing Center template:
  // title on line 8, colon after the title when a subtitle follows,
  // student name on line 24, then course, professor, date double-spaced.
  function titlePageLines(f) {
    var title = (f.title || "").trim();
    var subtitle = (f.subtitle || "").trim();
    if (subtitle && !/:$/.test(title)) title = title + ":";
    var lines = [];
    for (var i = 1; i < 8; i++) lines.push({ text: "" });
    lines.push({ text: title, bold: !!f.bold });
    if (subtitle) lines.push({ text: subtitle, bold: !!f.bold });
    while (lines.length < 23) lines.push({ text: "" });
    var tail = [f.name, f.course, f.professor, f.date].map(function (s) { return (s || "").trim(); });
    tail.forEach(function (t, idx) {
      if (idx > 0) lines.push({ text: "" });
      lines.push({ text: t });
    });
    return lines;
  }

  /* ------------------------------------------------------------------ */
  /* Text checks                                                         */
  /* ------------------------------------------------------------------ */
  function isSentenceStart(text, index) {
    var before = text.slice(0, index);
    return /^\s*["\u201C\u2018(]*$/.test(before) || /[.!?:]["\u201D\u2019)]*\s+["\u201C\u2018(]*$/.test(before);
  }

  function issue(rule, category, message, match, replacement, severity) {
    return {
      rule: rule, category: category, message: message, match: match,
      replacement: replacement === undefined ? null : replacement,
      severity: severity || "warn"
    };
  }

  /**
   * Check one paragraph of text.
   * context: "body" | "heading" | "footnote" | "bibliography"
   */
  function checkText(text, context, opts) {
    opts = opts || {};
    var out = [];
    if (!text || !text.trim()) return out;
    var m, re;
    var isProse = context === "body";

    // Capitalization (Writing Center "What to Capitalize")
    if (context !== "bibliography") {
      re = /\bbible\b/g;
      while ((m = re.exec(text))) {
        out.push(issue("cap-bible", "Capitalization", "Capitalize \u201CBible.\u201D", m[0], "Bible"));
      }
      re = /\bword of God\b/g;
      while ((m = re.exec(text))) {
        out.push(issue("cap-word", "Capitalization", "Capitalize \u201CWord of God.\u201D", m[0], "Word of God"));
      }
      re = /\bgospel of (Matthew|Mark|Luke|John)\b/g;
      while ((m = re.exec(text))) {
        out.push(issue("cap-gospel", "Capitalization", "Capitalize a specific Gospel (\u201CGospel of " + m[1] + "\u201D).", m[0], "Gospel of " + m[1]));
      }
    }
    if (context === "body" || context === "footnote") {
      re = /\b(Biblical|Scriptural|Scriptures?)\b/g;
      while ((m = re.exec(text))) {
        if (isSentenceStart(text, m.index)) continue;
        if (/Holy\s+$/.test(text.slice(0, m.index))) continue; // the Holy Scriptures
        if (/\bThe\s+$/.test(text.slice(0, m.index)) && m[1] === "Scriptures") continue;
        out.push(issue("cap-lower", "Capitalization",
          "Do not capitalize \u201C" + m[1].toLowerCase() + "\u201D (only the Holy Scriptures).",
          m[0], m[1].toLowerCase()));
      }
    }

    // Scripture usage
    if (context === "body" || context === "footnote") {
      re = new RegExp("\\b(" + bookAlternation + ")\\s+([Cc]hapters?|[Vv]erses?)\\s+\\d", "g");
      while ((m = re.exec(text))) {
        if (!findBook(m[1])) continue;
        out.push(issue("bible-chapter-word", "Scripture",
          "Don\u2019t write \u201C" + m[2] + "\u201D after a book name; use the form \u201CMatthew 5\u201D or \u201CGal. 5:22\u201323.\u201D",
          m[0].replace(/\s*\d$/, ""), null));
      }
      re = /\(([^()]*?)\bverses?\s+(\d)/g;
      while ((m = re.exec(text))) {
        out.push(issue("bible-vv", "Scripture",
          "Inside parentheses or footnotes, abbreviate verse/verses as v./vv.", m[0].slice(m[1].length + 1).replace(/\s*\d$/, ""), null));
      }
    }
    if (isProse) {
      // Abbreviated book for a whole chapter in running text (outside parentheses): spell it out.
      re = new RegExp("(^|[^(\\w])(" + bookAlternation + ")\\.?\\s+(\\d+)(?![\\d:])", "g");
      while ((m = re.exec(text))) {
        var b = findBook(m[2]);
        if (!b) continue;
        var written = m[0].slice(m[1].length).replace(/\s+\d+$/, "");
        if (written === b.full || norm(written) === norm(b.full)) continue;
        if (/^(Psalm|Song of Songs|Acts of the Apostles)$/.test(written)) continue; // accepted full forms
        if (insideParens(text, m.index + m[1].length)) continue;
        if (/^(Am|Is|Job|Acts|Jude|Joel|Mark|Luke|John|Ruth|Ezra|Titus|James|Amos|Hosea|Esther|Jon|Song|Ro|Ti|Jo|Na|Mi|Ob|Ez|Ex)$/.test(written)) continue; // common words / identical names
        out.push(issue("bible-spell-out", "Scripture",
          "In running text, spell out the book when citing a whole chapter (\u201C" + b.full + " " + m[3] + "\u201D).",
          written + " " + m[3], b.full + " " + m[3], "info"));
      }
      // Long quotation that should be a block quotation (five or more lines).
      re = /[\u201C"]([^\u201D"]{350,})[\u201D"]/g;
      while ((m = re.exec(text))) {
        var words = m[1].split(/\s+/).length;
        if (words >= 60) {
          out.push(issue("block-quote", "Quotations",
            "This quotation (~" + words + " words) is probably five or more lines: make it a block quotation (single-spaced, indented 0.5\u2033, no quotation marks).",
            m[0].slice(0, 40), null, "info"));
        }
      }
    }

    // Notes and bibliography
    if (context === "footnote" || context === "bibliography") {
      checkRanges(text, out, opts);
      re = /(?:^|[\s(])(\d{1,3},\d{3})(?=[\s.,;)\u2013-])/g;
      while ((m = re.exec(text))) {
        if (/\(\s*$/.test(text.slice(0, m.index + 1)) && /\)/.test(text.slice(m.index))) { /* fine to flag anyway */ }
        out.push(issue("page-comma", "Numbers", "Never use commas in page numbers.", m[1], m[1].replace(",", "")));
      }
    }
    if (context === "footnote") {
      var t = text.trim();
      if (/^ibid\b/.test(t)) out.push(issue("ibid-case", "Footnotes", "Capitalize \u201CIbid.\u201D at the start of a note.", "ibid", "Ibid"));
      if (/^Ibid(?!\.)/.test(t)) out.push(issue("ibid-period", "Footnotes", "\u201CIbid.\u201D needs a period.", "Ibid", "Ibid."));
      if (!/[.?!]["\u201D\u2019)]*$/.test(t) && !/[.?!]$/.test(t)) {
        out.push(issue("note-period", "Footnotes", "End each footnote with a period.", t.slice(-20), null));
      }
    }
    return out;
  }

  function insideParens(text, idx) {
    var open = text.lastIndexOf("(", idx);
    var close = text.lastIndexOf(")", idx);
    return open > close;
  }

  function checkRanges(text, out, opts) {
    var re = /(^|[^\d\/:.\w])(\d{2,5})(\s*[-\u2013]\s*)(\d{2,5})(?![\d\/])/g;
    var m;
    var want = dashChar(opts);
    while ((m = re.exec(text))) {
      var pre = text.slice(Math.max(0, m.index - 40), m.index + m[1].length);
      if (/(https?:\/\/|doi|www\.)\S*$/i.test(pre)) continue; // inside a URL/DOI
      if (/:\s*$/.test(pre) && /\b\d+\s*:\s*$/.test(pre)) continue; // verse ranges like 5:22-23 handled elsewhere
      var s = m[2], e = m[4];
      var a = parseInt(s, 10), b = parseInt(e, 10);
      if (b <= a) continue; // already abbreviated (e.g., 321-28) or not a range
      var ideal = s + want + abbreviateRange(s, e);
      var actual = s + m[3] + e;
      if (actual !== ideal) {
        var why = m[3].trim() !== want && abbreviateRange(s, e) === e
          ? "Use " + (want === EN ? "an en dash (\u2013)" : "a hyphen") + " in number ranges."
          : "Abbreviate this range per Turabian Table 23.2.";
        out.push(issue("range", "Numbers", why + " Suggested: " + ideal, actual, ideal, "info"));
      }
    }
  }

  /* ------------------------------------------------------------------ */
  /* Bibliography checks                                                 */
  /* ------------------------------------------------------------------ */
  function bibSortKey(entry) {
    return entry
      .replace(/^[\s"\u201C\u2018'\u2014\u2013-]+/, "")
      .replace(/^(the|a|an)\s+/i, "")
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
  }

  function checkBibliographyEntry(text) {
    var out = [];
    var t = text || "";
    if (/\b(Holy Bible|New International Version|English Standard Version|New Revised Standard Version|King James Version|New American Standard Bible|New Living Translation)\b/i.test(t) && !/\bStudy Bible\b|\bnotes?\b/i.test(t)) {
      out.push(issue("bib-bible", "Bibliography", "Don\u2019t list the Bible in the bibliography (17.8.2) unless you cite study notes or other supplemental parts.", t.slice(0, 40), null));
    }
    if (/\b(Merriam-Webster|Encyclop(a)?edia Britannica|Oxford English Dictionary|Dictionary\.com|Wikipedia)\b/i.test(t)) {
      out.push(issue("bib-common-dict", "Bibliography", "Common dictionaries/encyclopedias are cited in footnotes only, not the bibliography.", t.slice(0, 40), null));
    }
    if (/\b(email message|e-mail message|text message|interview by|Facebook message|direct message|personal communication)\b/i.test(t)) {
      out.push(issue("bib-personal", "Bibliography", "Interviews and personal communications are usually cited in notes only (17.6).", t.slice(0, 40), null, "info"));
    }
    if (/\(blog\)/i.test(t)) {
      out.push(issue("bib-blog", "Bibliography", "Citing a blog in the bibliography is usually unnecessary.", t.slice(0, 40), null, "info"));
    }
    return out;
  }

  function checkBibliographyOrder(entries) {
    var out = [];
    for (var i = 1; i < entries.length; i++) {
      var a = bibSortKey(entries[i - 1]), b = bibSortKey(entries[i]);
      if (a.localeCompare(b) > 0) {
        out.push({ index: i, message: "Out of alphabetical order: \u201C" + entries[i].slice(0, 30) + "\u2026\u201D should come before \u201C" + entries[i - 1].slice(0, 30) + "\u2026\u201D" });
      }
    }
    return out;
  }

  function sortBibliography(entries) {
    return entries.slice().sort(function (x, y) { return bibSortKey(x).localeCompare(bibSortKey(y)); });
  }

  /* ------------------------------------------------------------------ */
  /* Scripture consistency (traditional vs shorter abbreviations)        */
  /* ------------------------------------------------------------------ */
  // Find parenthetical scripture citations like "(Jn 3:16 NIV)" in text.
  function findParentheticalRefs(text) {
    var re = new RegExp("\\(((" + bookAlternation + ")\\.?\\s*\\d+[\\d:,;\\-\\u2013\\s]*(?:\\s+[A-Z]{2,6})?)\\)", "g");
    var res = [], m;
    while ((m = re.exec(text))) {
      var book = findBook(m[2]);
      if (!book) continue;
      var written = m[1].match(/^(.*?[A-Za-z]\.?)\s*\d/)[1].trim();
      var kind = written === book.traditional ? "traditional"
        : written === book.short ? "short"
        : norm(written) === norm(book.full) ? "full" : "other";
      if (book.traditional === book.short && kind !== "full") kind = "either";
      res.push({ text: m[0], inner: m[1], book: book, written: written, kind: kind });
    }
    return res;
  }

  function convertParenthetical(found, style, opts) {
    var ref = parseReference(found.inner);
    if (!ref || ref.error) return null;
    return formatScripture(ref, "paren", ref.version, { abbrev: style, dash: opts && opts.dash });
  }

  return {
    BOOKS: BOOKS,
    findBook: findBook,
    abbreviateRange: abbreviateRange,
    formatRange: formatRange,
    parseReference: parseReference,
    formatScripture: formatScripture,
    versionDecision: versionDecision,
    primaryVersionNote: primaryVersionNote,
    formatDate: formatDate,
    titlePageLines: titlePageLines,
    checkText: checkText,
    checkBibliographyEntry: checkBibliographyEntry,
    checkBibliographyOrder: checkBibliographyOrder,
    sortBibliography: sortBibliography,
    bibSortKey: bibSortKey,
    findParentheticalRefs: findParentheticalRefs,
    convertParenthetical: convertParenthetical
  };
});
