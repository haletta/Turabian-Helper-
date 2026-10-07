/* Turabian Helper — Word task pane logic (Office.js) */
/* global Office, Word, TurabianRules */
(function () {
  "use strict";
  var R = TurabianRules;

  var FONT = "Times New Roman";
  var TP_TAG = "turabian-title-page";
  var STYLE_BLOCK = "Turabian Block Quote";
  var STYLE_BIB = "Turabian Bibliography";
  var PT = { half: 36, single: 12, double: 24 };

  var api = { v13: false, v15: false };
  var prefs = loadPrefs();

  /* ---------------- small helpers ---------------- */
  function $(id) { return document.getElementById(id); }
  function store(key, val) { try { if (val === undefined) return localStorage.getItem(key); localStorage.setItem(key, val); } catch (e) { return null; } }
  function loadPrefs() {
    var p = { dash: "en", abbrev: "traditional" };
    try { var s = JSON.parse(localStorage.getItem("th.prefs") || "{}"); if (s.dash) p.dash = s.dash; if (s.abbrev) p.abbrev = s.abbrev; } catch (e) { /* ignore */ }
    return p;
  }
  function savePrefs() { store("th.prefs", JSON.stringify(prefs)); }
  function near(v, target, tol) { return typeof v === "number" && Math.abs(v - target) <= (tol || 1); }

  var toastTimer;
  function toast(msg, isError) {
    var t = $("toast");
    t.textContent = msg;
    t.className = "toast" + (isError ? " error" : "");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.className = "toast hidden"; }, isError ? 7000 : 3500);
  }
  function friendlyError(e) {
    var m = (e && (e.message || e.code)) || String(e);
    if (/InvalidSelection|ItemNotFound/i.test(m)) return "Put your cursor in the document first, then try again.";
    if (/AccessDenied|ReadOnly/i.test(m)) return "Word won’t let me edit right now. Is the document read-only, or are you in the middle of editing a header or comment?";
    return "Something went wrong: " + m;
  }
  function busy(btn, on) { if (btn) btn.disabled = on; }
  async function run(btn, fn) {
    busy(btn, true);
    try { await fn(); } catch (e) { console.error(e); toast(friendlyError(e), true); } finally { busy(btn, false); }
  }

  /* ---------------- formatting primitives ---------------- */
  function setBodyText(p) {
    p.font.name = FONT; p.font.size = 12; p.font.color = "#000000";
    p.alignment = "Left";
    p.leftIndent = 0; p.firstLineIndent = PT.half;
    p.lineSpacing = PT.double; p.spaceAfter = 0; p.spaceBefore = 0;
  }
  function setHeading(p, level) {
    p.styleBuiltIn = "Heading" + level;
    p.font.name = FONT; p.font.size = 12; p.font.color = "#000000";
    p.font.bold = level === 1 || level === 3;
    p.font.italic = false;
    p.alignment = level <= 2 ? "Centered" : "Left";
    p.leftIndent = 0; p.firstLineIndent = 0;
    p.lineSpacing = PT.double; p.spaceBefore = PT.single; p.spaceAfter = 0;
  }
  function setBlockQuote(p, first, last) {
    p.font.name = FONT; p.font.size = 12;
    p.alignment = "Left";
    p.leftIndent = PT.half; p.firstLineIndent = 0;
    p.lineSpacing = PT.single;
    p.spaceBefore = first ? PT.single : 0;
    p.spaceAfter = last ? PT.single : 0;
  }
  function setBibEntry(p) {
    p.font.name = FONT; p.font.size = 12; p.font.color = "#000000";
    p.alignment = "Left";
    p.leftIndent = PT.half; p.firstLineIndent = -PT.half;
    p.lineSpacing = PT.single; p.spaceBefore = 0; p.spaceAfter = PT.single;
  }
  function setFootnote(p) {
    p.font.name = FONT; p.font.size = 10; p.font.color = "#000000";
    p.alignment = "Left";
    p.leftIndent = 0; p.firstLineIndent = PT.half;
    p.lineSpacing = PT.single; p.spaceBefore = 0; p.spaceAfter = PT.single;
  }

  // Header paragraph: page number at top right, hidden on page 1 (the title page).
  // { IF { PAGE } > 1 "{ PAGE }" "" } avoids needing "different first page".
  function pageNumberOoxml() {
    var rpr = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/></w:rPr>';
    function r(inner) { return "<w:r>" + rpr + inner + "</w:r>"; }
    function fld(type, dirty) { return r('<w:fldChar w:fldCharType="' + type + '"' + (dirty ? ' w:dirty="true"' : "") + "/>"); }
    function instr(t) { return r('<w:instrText xml:space="preserve">' + t + "</w:instrText>"); }
    function txt(t) { return r("<w:t>" + t + "</w:t>"); }
    var page = fld("begin") + instr(" PAGE ") + fld("separate") + txt("2") + fld("end");
    var para =
      '<w:p><w:pPr><w:jc w:val="right"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0"/></w:pPr>' +
      fld("begin", true) + instr(" IF ") + page + instr(' &gt; 1 "') + page + instr('" "" ') +
      fld("separate") + txt("2") + fld("end") + "</w:p>";
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' +
      "</pkg:xmlData></pkg:part>" +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + para + "</w:body></w:document>" +
      "</pkg:xmlData></pkg:part></pkg:package>";
  }

  /* ---------------- document map (what is what) ---------------- */
  // Loads body paragraphs and classifies each one so formatting and checks
  // skip the title page, treat headings, block quotes and bibliography correctly.
  async function mapDocument(ctx) {
    var paras = ctx.document.body.paragraphs;
    paras.load("items/text,items/alignment,items/lineSpacing,items/firstLineIndent,items/leftIndent,items/styleBuiltIn,items/tableNestingLevel,items/font/name,items/font/size,items/font/bold");
    var ccs = ctx.document.contentControls.getByTag(TP_TAG);
    ccs.load("items");
    await ctx.sync();

    var titleCount = 0;
    if (ccs.items.length) {
      var tp = ccs.items[0].paragraphs;
      tp.load("items");
      await ctx.sync();
      titleCount = tp.items.length;
    } else {
      // Heuristic: leading blank/centered lines with at least 3 centered lines of text = a title page.
      var centered = 0, i = 0;
      for (; i < paras.items.length && i < 40; i++) {
        var q = paras.items[i];
        if (!q.text.trim()) continue;
        if (q.alignment === "Centered" && q.text.length < 160 && !/^Heading/.test(q.styleBuiltIn)) { centered++; continue; }
        break;
      }
      // A title page has at least title, name, course, professor and date (4+ centered lines);
      // fewer than that is more likely a centered heading, so don't skip it.
      if (centered >= 4 && i < 40) titleCount = i;
    }

    var bibIndex = -1;
    paras.items.forEach(function (p, idx) {
      if (idx >= titleCount && /^\s*bibliography\s*$/i.test(p.text)) bibIndex = idx;
    });

    var info = paras.items.map(function (p, idx) {
      var kind;
      var text = p.text || "";
      var hm = /^Heading([1-9])$/.exec(p.styleBuiltIn || "");
      if (idx < titleCount) kind = "title";
      else if (p.tableNestingLevel > 0) kind = "table";
      else if (!text.trim()) kind = "empty";
      else if (idx === bibIndex) kind = "bibHeading";
      else if (bibIndex >= 0 && idx > bibIndex) kind = "bib";
      else if (hm) kind = "heading";
      else if (p.alignment === "Centered" && text.length < 140 && !/[.!?]["”]?$/.test(text.trim())) kind = "heading"; // unstyled heading
      else if (p.leftIndent >= 30 && p.leftIndent <= 50 && p.firstLineIndent <= 1) kind = "block";
      else if (p.leftIndent > 0) kind = "indented"; // lists etc. — leave alone
      else kind = "body";
      return { p: p, idx: idx, kind: kind, text: text, level: hm ? parseInt(hm[1], 10) : 0 };
    });
    return { paras: paras, info: info, titleCount: titleCount, bibIndex: bibIndex };
  }

  async function loadFootnotes(ctx) {
    if (!api.v15) return null;
    var notes = ctx.document.body.footnotes;
    notes.load("items");
    await ctx.sync();
    notes.items.forEach(function (n) {
      n.body.paragraphs.load("items/text,items/lineSpacing,items/firstLineIndent,items/spaceAfter,items/font/name,items/font/size");
    });
    await ctx.sync();
    return notes;
  }

  /* ================= TITLE PAGE ================= */
  function initTitlePage() {
    ["Name", "Course", "Professor"].forEach(function (k) {
      var v = store("th.tp" + k);
      if (v) $("tp" + k).value = v;
    });
    $("tpDate").value = R.formatDate(new Date());
    $("btnTitlePage").onclick = function () { run(this, insertTitlePage); };
  }

  async function insertTitlePage() {
    var f = {
      title: $("tpTitle").value, subtitle: $("tpSubtitle").value, name: $("tpName").value,
      course: $("tpCourse").value, professor: $("tpProfessor").value, date: $("tpDate").value,
      bold: $("tpBold").checked
    };
    if (!f.title.trim()) { toast("Please enter a title.", true); $("tpTitle").focus(); return; }
    store("th.tpName", f.name); store("th.tpCourse", f.course); store("th.tpProfessor", f.professor);
    var lines = R.titlePageLines(f);

    await Word.run(async function (ctx) {
      var existing = ctx.document.contentControls.getByTag(TP_TAG);
      existing.load("items");
      await ctx.sync();
      existing.items.forEach(function (cc) { cc.delete(false); });
      await ctx.sync();

      var body = ctx.document.body;
      var created = [];
      var prev = null;
      lines.forEach(function (ln) {
        var p = prev ? prev.insertParagraph(ln.text, "After") : body.insertParagraph(ln.text, "Start");
        p.styleBuiltIn = "Normal";
        p.font.name = FONT; p.font.size = 12; p.font.color = "#000000";
        p.font.bold = !!ln.bold; p.font.italic = false;
        p.alignment = "Centered";
        p.leftIndent = 0; p.firstLineIndent = 0;
        p.lineSpacing = PT.single; p.spaceAfter = 0; p.spaceBefore = 0;
        created.push(p);
        prev = p;
      });
      var last = created[created.length - 1];
      last.getRange("Content").insertBreak("Page", "After");
      var whole = created[0].getRange("Start").expandTo(last.getRange("End"));
      var cc = whole.insertContentControl();
      cc.tag = TP_TAG;
      cc.title = "Title page (Turabian Helper)";
      cc.appearance = "Hidden";
      await ctx.sync();
      created[0].getRange("Start").select();
      await ctx.sync();
    });
    toast("Title page inserted. Page numbering starts with 2 on the next page once you run Format → Set up.");
  }

  /* ================= FORMAT ================= */
  function initFormat() {
    $("btnSetup").onclick = function () { run(this, setupPaper); };
    document.querySelectorAll("[data-heading]").forEach(function (b) {
      b.onclick = function () { var lvl = parseInt(b.getAttribute("data-heading"), 10); run(b, function () { return applyToSelection(function (ps) { ps.forEach(function (p) { setHeading(p, lvl); }); }, "Heading level " + lvl + " applied."); }); };
    });
    $("btnBlockQuote").onclick = function () {
      run(this, function () {
        return applyToSelection(function (ps) {
          ps.forEach(function (p, i) { if (api.v15) p.style = STYLE_BLOCK; setBlockQuote(p, i === 0, i === ps.length - 1); });
        }, "Block quotation applied. Remove any quotation marks around it.", true);
      });
    };
    $("btnNormal").onclick = function () { run(this, function () { return applyToSelection(function (ps) { ps.forEach(function (p) { p.styleBuiltIn = "Normal"; setBodyText(p); }); }, "Normal text applied."); }); };
    $("btnBibHeading").onclick = function () { run(this, addBibliographyPage); };
    $("btnBibEntries").onclick = function () {
      run(this, function () {
        return applyToSelection(function (ps) {
          ps.forEach(function (p) { if (api.v15) p.style = STYLE_BIB; setBibEntry(p); });
        }, "Bibliography entries formatted.", true);
      });
    };
    $("optDash").value = prefs.dash;
    $("optAbbrev").value = prefs.abbrev;
    $("optDash").onchange = function () { prefs.dash = this.value; savePrefs(); updateScripturePreview(); };
    $("optAbbrev").onchange = function () { prefs.abbrev = this.value; savePrefs(); updateScripturePreview(); };
  }

  async function applyToSelection(fn, doneMsg, needStyles) {
    await Word.run(async function (ctx) {
      if (needStyles && api.v15) await ensureStyles(ctx, []);
      var ps = ctx.document.getSelection().paragraphs;
      ps.load("items");
      await ctx.sync();
      if (!ps.items.length) throw new Error("InvalidSelection");
      fn(ps.items);
      await ctx.sync();
    });
    toast(doneMsg);
  }

  async function ensureStyles(ctx, log) {
    var doc = ctx.document;
    var styles = doc.getStyles();
    var names = ["Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Footnote Text", STYLE_BLOCK, STYLE_BIB];
    var got = {};
    names.forEach(function (n) { got[n] = styles.getByNameOrNullObject(n); got[n].load("nameLocal"); });
    await ctx.sync();
    [STYLE_BLOCK, STYLE_BIB].forEach(function (n) { if (got[n].isNullObject) got[n] = doc.addStyle(n, "Paragraph"); });
    await ctx.sync();

    function base(s, size) { s.font.name = FONT; s.font.size = size || 12; s.font.color = "#000000"; }
    var n = got["Normal"];
    if (!n.isNullObject) {
      base(n);
      n.paragraphFormat.lineSpacing = PT.double; n.paragraphFormat.firstLineIndent = PT.half;
      n.paragraphFormat.spaceAfter = 0; n.paragraphFormat.spaceBefore = 0; n.paragraphFormat.alignment = "Left";
      n.paragraphFormat.widowControl = true;
    }
    [1, 2, 3, 4].forEach(function (lvl) {
      var h = got["Heading " + lvl];
      if (h.isNullObject) return;
      base(h);
      h.font.bold = lvl === 1 || lvl === 3; h.font.italic = false;
      var pf = h.paragraphFormat;
      pf.alignment = lvl <= 2 ? "Centered" : "Left";
      pf.firstLineIndent = 0; pf.leftIndent = 0;
      pf.lineSpacing = PT.double; pf.spaceBefore = PT.single; pf.spaceAfter = 0; pf.keepWithNext = true;
    });
    var fn = got["Footnote Text"];
    if (!fn.isNullObject) {
      base(fn, 10);
      fn.paragraphFormat.lineSpacing = PT.single; fn.paragraphFormat.firstLineIndent = PT.half;
      fn.paragraphFormat.spaceAfter = PT.single; fn.paragraphFormat.spaceBefore = 0; fn.paragraphFormat.leftIndent = 0;
    }
    var bq = got[STYLE_BLOCK];
    base(bq);
    bq.paragraphFormat.leftIndent = PT.half; bq.paragraphFormat.firstLineIndent = 0;
    bq.paragraphFormat.lineSpacing = PT.single; bq.paragraphFormat.spaceBefore = PT.single; bq.paragraphFormat.spaceAfter = PT.single;
    var bb = got[STYLE_BIB];
    base(bb);
    bb.paragraphFormat.leftIndent = PT.half; bb.paragraphFormat.firstLineIndent = -PT.half;
    bb.paragraphFormat.lineSpacing = PT.single; bb.paragraphFormat.spaceBefore = 0; bb.paragraphFormat.spaceAfter = PT.single;
    await ctx.sync();
    log.push(["ok", "Styles set: Normal, Headings 1–4, Footnote Text, Block Quote, Bibliography"]);
  }

  async function setupPaper() {
    var log = [];
    await Word.run(async function (ctx) {
      if (api.v15) {
        try { await ensureStyles(ctx, log); } catch (e) { log.push(["warn", "Couldn’t update styles (" + (e.message || e) + "); applied direct formatting instead."]); }
      }
      var map = await mapDocument(ctx);
      var counts = { body: 0, heading: 0, block: 0, bib: 0 };
      map.info.forEach(function (it) {
        var p = it.p;
        if (it.kind === "title" || it.kind === "table") return;
        p.font.name = FONT;
        if (p.alignment === "Justified") p.alignment = "Left";
        if (it.kind === "body") { setBodyText(p); counts.body++; }
        else if (it.kind === "heading") { setHeading(p, it.level || (p.font.bold ? 1 : 2)); counts.heading++; }
        else if (it.kind === "block") { p.font.size = 12; p.lineSpacing = PT.single; p.firstLineIndent = 0; counts.block++; }
        else if (it.kind === "bibHeading") {
          p.styleBuiltIn = "Normal"; p.font.size = 12; p.font.bold = false; p.alignment = "Centered";
          p.firstLineIndent = 0; p.leftIndent = 0; p.lineSpacing = PT.single; p.spaceBefore = 0; p.spaceAfter = PT.double;
        }
        else if (it.kind === "bib") { setBibEntry(p); counts.bib++; }
        else if (it.kind !== "empty") { p.font.size = 12; }
      });
      await ctx.sync();
      log.push(["ok", counts.body + " body paragraphs: Times New Roman 12, double-spaced, 0.5″ indent, left-aligned"]);
      if (counts.heading) log.push(["ok", counts.heading + " headings formatted"]);
      if (counts.block) log.push(["ok", counts.block + " block quotation paragraphs single-spaced"]);
      if (counts.bib) log.push(["ok", counts.bib + " bibliography entries: hanging indent, single-spaced, blank line between"]);
      if (map.titleCount) log.push(["ok", "Title page left as is"]);

      var notes = await loadFootnotes(ctx);
      if (notes) {
        notes.items.forEach(function (n) { n.body.paragraphs.items.forEach(setFootnote); });
        await ctx.sync();
        log.push(["ok", notes.items.length + " footnotes: 10 pt, single-spaced, indented, blank line between"]);
      } else {
        log.push(["warn", "This version of Word can’t reach footnotes from add-ins. Format them by changing the Footnote Text style."]);
      }

      var secs = ctx.document.sections;
      secs.load("items");
      await ctx.sync();
      secs.items.forEach(function (s) { s.getHeader("Primary").insertOoxml(pageNumberOoxml(), "Replace"); });
      await ctx.sync();
      log.push(["ok", "Page numbers: top right, hidden on page 1, so the first text page shows 2"]);
    });

    // Margins need a newer API; try separately so a failure doesn't undo the rest.
    try {
      await Word.run(async function (ctx) {
        var secs = ctx.document.sections;
        secs.load("items");
        await ctx.sync();
        secs.items.forEach(function (s) {
          var ps = s.pageSetup;
          ps.topMargin = 72; ps.bottomMargin = 72; ps.leftMargin = 72; ps.rightMargin = 72;
        });
        await ctx.sync();
      });
      log.push(["ok", "Margins: 1″ on all sides"]);
    } catch (e) {
      log.push(["warn", "Set margins yourself: Layout → Margins → Normal (1″). Your Word doesn’t let add-ins change margins."]);
    }

    var ul = $("setupLog");
    ul.innerHTML = "";
    log.forEach(function (l) { var li = document.createElement("li"); li.className = l[0]; li.textContent = l[1]; ul.appendChild(li); });
    toast("Paper set up in Turabian format.");
  }

  async function addBibliographyPage() {
    await Word.run(async function (ctx) {
      var body = ctx.document.body;
      var last = body.paragraphs.getLast();
      last.load("text");
      await ctx.sync();
      var heading;
      if (last.text.trim()) {
        body.insertBreak("Page", "End");
        heading = body.insertParagraph("Bibliography", "End");
      } else {
        last.insertBreak("Page", "Before");
        last.insertText("Bibliography", "Replace");
        heading = last;
      }
      heading.styleBuiltIn = "Normal";
      heading.font.name = FONT; heading.font.size = 12; heading.font.bold = false; heading.font.italic = false; heading.font.color = "#000000";
      heading.alignment = "Centered"; heading.firstLineIndent = 0; heading.leftIndent = 0;
      heading.lineSpacing = PT.single; heading.spaceBefore = 0; heading.spaceAfter = PT.double; // two blank lines below
      var first = heading.insertParagraph("", "After");
      setBibEntry(first);
      first.font.bold = false;
      first.select("Start");
      await ctx.sync();
    });
    toast("Bibliography page added. Paste or insert your entries, then use “Format selected entries.”");
  }

  /* ================= SCRIPTURE ================= */
  var docState = { primary: null };

  function initScripture() {
    docState.primary = (Office.context && Office.context.document && Office.context.document.settings && Office.context.document.settings.get("th.primaryVersion")) || null;
    showPrimary();
    var v = store("th.lastVersion");
    $("scVersion").value = v || docState.primary || "NIV";
    ["scRef", "scVersion"].forEach(function (id) { $(id).addEventListener("input", updateScripturePreview); });
    document.querySelectorAll("input[name=scMode]").forEach(function (r) { r.addEventListener("change", updateScripturePreview); });
    $("scRef").addEventListener("keydown", function (e) { if (e.key === "Enter") $("btnScripture").click(); });
    $("btnScripture").onclick = function () { run(this, insertScripture); };
    $("btnResetPrimary").onclick = function () { setPrimary(null); toast("Primary version cleared. Your next quotation will set it."); };
    $("btnScriptureConsistent").onclick = function () { run(this, makeScriptureConsistent); };
    updateScripturePreview();
  }

  function scMode() { return document.querySelector("input[name=scMode]:checked").value; }

  function showPrimary() { $("scPrimary").textContent = docState.primary || "not set yet"; }

  function setPrimary(v) {
    docState.primary = v;
    showPrimary();
    try {
      var s = Office.context.document.settings;
      if (v) s.set("th.primaryVersion", v); else s.remove("th.primaryVersion");
      s.saveAsync(function () {});
    } catch (e) { /* outside Word */ }
  }

  function buildScripture() {
    var ref = R.parseReference($("scRef").value);
    if (!ref) return { error: "Type a reference, e.g. Jeremiah 31:4" };
    if (ref.error) return ref;
    var mode = scMode();
    var version = ($("scVersion").value || ref.version || "").toUpperCase().trim();
    var decision = mode === "paren" ? R.versionDecision(docState, version) : { show: false, addPrimaryFootnote: false };
    var text = R.formatScripture(ref, mode, decision.show ? version : null, { abbrev: prefs.abbrev, dash: prefs.dash });
    var note = "";
    if (mode === "paren" && decision.addPrimaryFootnote) note = "First quotation: the version is shown and a footnote will be added: “" + R.primaryVersionNote(version) + "”";
    else if (mode === "paren" && version && !decision.show) note = "Version left out because " + version + " is your primary version.";
    else if (mode === "paren" && decision.show) note = "Version shown because it differs from your primary version (" + docState.primary + ").";
    else if (mode === "running") note = "In running text, book names are spelled out.";
    return { text: text, version: version, decision: decision, mode: mode, note: note };
  }

  function updateScripturePreview() {
    var b = buildScripture();
    $("scPeriodWrap").style.display = scMode() === "paren" ? "" : "none";
    if (b.error) { $("scPreview").textContent = $("scRef").value.trim() ? b.error : "—"; $("scNote").textContent = ""; return; }
    var period = b.mode === "paren" && $("scPeriod").checked ? "." : "";
    $("scPreview").textContent = b.text + period;
    $("scNote").textContent = b.note;
  }
  document.addEventListener("change", function (e) { if (e.target && e.target.id === "scPeriod") updateScripturePreview(); });

  async function insertScripture() {
    var b = buildScripture();
    if (b.error) { toast(b.error, true); return; }
    var period = b.mode === "paren" && $("scPeriod").checked;
    var needNote = b.decision.addPrimaryFootnote;
    var noteAdded = false;
    await Word.run(async function (ctx) {
      var sel = ctx.document.getSelection();
      var lead = "";
      if (b.mode === "paren") {
        // Add a space before "(" unless the cursor already follows one.
        var upToCursor = sel.paragraphs.getFirst().getRange("Start").expandTo(sel.getRange("Start"));
        upToCursor.load("text");
        await ctx.sync();
        var t = upToCursor.text || "";
        if (t.length && !/\s$/.test(t)) lead = " ";
      }
      var inserted = sel.insertText(lead + b.text + (period ? "." : ""), "Replace");
      inserted.font.italic = false;
      if (needNote && api.v15) {
        inserted.insertFootnote(R.primaryVersionNote(b.version));
        noteAdded = true;
      }
      inserted.getRange("End").select();
      await ctx.sync();
    });
    if (b.mode === "paren" && b.decision.primary && !docState.primary) setPrimary(b.decision.primary);
    store("th.lastVersion", b.version);
    $("scRef").value = "";
    updateScripturePreview();
    if (needNote && !noteAdded) toast("Inserted. Add a footnote yourself: “" + R.primaryVersionNote(b.version) + "” (your Word can’t insert footnotes from add-ins).");
    else toast(needNote ? "Inserted with the primary-version footnote." : "Inserted.");
  }

  async function makeScriptureConsistent() {
    var changed = 0;
    await Word.run(async function (ctx) {
      var map = await mapDocument(ctx);
      var jobs = [];
      map.info.forEach(function (it) {
        if (it.kind === "title" || it.kind === "empty") return;
        R.findParentheticalRefs(it.text).forEach(function (f) {
          if (f.kind === prefs.abbrev || f.kind === "either") return;
          var replacement = R.convertParenthetical(f, prefs.abbrev, { dash: prefs.dash });
          if (!replacement || replacement === f.text) return;
          var hits = it.p.search(f.text, { matchCase: true });
          hits.load("items");
          jobs.push({ hits: hits, replacement: replacement });
        });
      });
      await ctx.sync();
      jobs.forEach(function (j) { j.hits.items.forEach(function (h) { h.insertText(j.replacement, "Replace"); changed++; }); });
      await ctx.sync();
    });
    toast(changed ? "Updated " + changed + " scripture citation" + (changed === 1 ? "" : "s") + "." : "All parenthetical scripture citations already match.");
  }

  /* ================= CHECK ================= */
  var lastIssues = [];

  function initCheck() {
    $("btnCheck").onclick = function () { run(this, runCheck); };
    $("btnFixAllFormat").onclick = function () { run(this, fixAllFormatting); };
  }

  var FORMAT_FIXES = { font: 1, size: 1, align: 1, spacing: 1, indent: 1, hanging: 1, bibspacing: 1, fnformat: 1, blockspacing: 1, headingalign: 1 };

  async function runCheck() {
    var issues = [];
    function add(o) { issues.push(o); }

    await Word.run(async function (ctx) {
      var map = await mapDocument(ctx);
      var notes = await loadFootnotes(ctx);
      var headerXml = ctx.document.sections.getFirst().getHeader("Primary").getOoxml();
      var tpCC = ctx.document.contentControls.getByTag(TP_TAG);
      tpCC.load("items");
      await ctx.sync();

      // Document-level
      if (!map.titleCount && !tpCC.items.length) {
        add({ category: "Structure", severity: "info", message: "No title page found. Use the Title page tab to add one.", loc: null });
      }
      if (!/instrText[^>]*>\s*PAGE\b|w:instr="\s*PAGE\b/.test(headerXml.value || "")) {
        add({ category: "Structure", message: "No page numbers. Turabian numbers pages top right, starting with 2 on the first page of text.", loc: null, fix: "pagenum" });
      }
      var noteCount = notes ? notes.items.length : 0;
      if (map.bibIndex < 0 && noteCount > 0) {
        add({ category: "Bibliography", message: "You have footnotes but no page titled “Bibliography.” Most papers need one (Format → Add Bibliography page).", loc: null });
      }

      // Paragraph checks
      var headingSeq = [];
      var refsKinds = { traditional: 0, short: 0 };
      var bibEntries = [];
      map.info.forEach(function (it) {
        var p = it.p, loc = { type: "p", index: it.idx };
        var snippet = it.text.slice(0, 60);
        if (it.kind === "title" || it.kind === "empty" || it.kind === "table") return;

        if (p.font.name && p.font.name !== FONT) add({ category: "Formatting", message: "Font is " + p.font.name + ". Use Times New Roman.", loc: loc, where: snippet, fix: "font" });
        if (p.font.size && p.font.size !== 12) add({ category: "Formatting", message: "Font size is " + p.font.size + " pt. Use 12 pt.", loc: loc, where: snippet, fix: "size" });
        if (p.alignment === "Justified") add({ category: "Formatting", message: "Text is justified. Align left so the right margin is jagged.", loc: loc, where: snippet, fix: "align" });

        if (it.kind === "body") {
          if (!near(p.lineSpacing, PT.double, 1.5)) add({ category: "Formatting", message: "Not double-spaced.", loc: loc, where: snippet, fix: "spacing" });
          if (!near(p.firstLineIndent, PT.half, 2)) add({ category: "Formatting", message: "First line should be indented 0.5″.", loc: loc, where: snippet, fix: "indent" });
          R.findParentheticalRefs(it.text).forEach(function (f) { if (refsKinds[f.kind] !== undefined) refsKinds[f.kind]++; });
        }
        if (it.kind === "block") {
          if (!near(p.lineSpacing, PT.single, 1.5)) add({ category: "Quotations", message: "Block quotations are single-spaced.", loc: loc, where: snippet, fix: "blockspacing" });
          if (/^\s*[“"]/.test(it.text) && /[”"]\s*\S{0,30}$/.test(it.text)) add({ category: "Quotations", message: "Don’t put quotation marks around a block quotation.", loc: loc, where: snippet });
        }
        if (it.kind === "heading") {
          var lvl = it.level || 1;
          headingSeq.push({ level: it.level, idx: it.idx, text: snippet });
          if (it.level && (it.level <= 2) !== (p.alignment === "Centered")) {
            add({ category: "Headings", message: "Level " + lvl + " headings are " + (lvl <= 2 ? "centered." : "flush left."), loc: loc, where: snippet, fix: "headingalign", level: lvl });
          }
        }
        if (it.kind === "bibHeading" && p.alignment !== "Centered") {
          add({ category: "Bibliography", message: "Center the word “Bibliography.”", loc: loc, where: snippet, fix: "headingalign", level: 2 });
        }
        if (it.kind === "bib") {
          bibEntries.push({ text: it.text, idx: it.idx });
          if (!(p.firstLineIndent < -20 && near(p.leftIndent, PT.half, 4))) add({ category: "Bibliography", message: "Use a 0.5″ hanging indent (runover lines indented).", loc: loc, where: snippet, fix: "hanging" });
          if (!near(p.lineSpacing, PT.single, 1.5)) add({ category: "Bibliography", message: "Single-space within entries, with a blank line between.", loc: loc, where: snippet, fix: "bibspacing" });
          R.checkBibliographyEntry(it.text).forEach(function (x) { add({ category: x.category, severity: x.severity, message: x.message, loc: loc, where: snippet }); });
        }

        var ctxName = it.kind === "heading" || it.kind === "bibHeading" ? "heading" : it.kind === "bib" ? "bibliography" : "body";
        R.checkText(it.text, ctxName, { dash: prefs.dash }).forEach(function (x) {
          add({ category: x.category, severity: x.severity, message: x.message, loc: loc, where: snippet, match: x.match, replacement: x.replacement });
        });
      });

      // Single subheading rule
      for (var i = 0; i < headingSeq.length; i++) {
        var h = headingSeq[i];
        if (!h.level) continue;
        var kids = 0, firstKid = null;
        for (var j = i + 1; j < headingSeq.length && headingSeq[j].level > h.level; j++) {
          if (headingSeq[j].level === h.level + 1) { kids++; firstKid = firstKid || headingSeq[j]; }
        }
        if (kids === 1) add({ category: "Headings", severity: "info", message: "Only one subheading under “" + h.text + "”. A subheading can’t stand alone: add another or remove it.", loc: { type: "p", index: firstKid.idx }, where: firstKid.text });
      }

      // Bibliography order
      R.checkBibliographyOrder(bibEntries.map(function (b) { return b.text; })).forEach(function (o) {
        add({ category: "Bibliography", message: o.message, loc: { type: "p", index: bibEntries[o.index].idx }, where: bibEntries[o.index].text.slice(0, 60) });
      });

      // Scripture consistency
      if (refsKinds.traditional && refsKinds.short) {
        add({ category: "Scripture", message: "You mix traditional (Jer.) and shorter (Jer) abbreviations. Pick one and be consistent.", loc: null, fix: "scripture" });
      }

      // Footnotes
      if (notes) {
        var hasPrimaryNote = false;
        var ibidChecks = [];
        notes.items.forEach(function (n, ni) {
          n.body.paragraphs.items.forEach(function (fp, pi) {
            var t = fp.text || "";
            if (/Unless otherwise noted/i.test(t)) hasPrimaryNote = true;
            var loc = { type: "fn", index: ni, para: pi };
            var snippet = (ni + 1) + ". " + t.slice(0, 55);
            if ((fp.font.size && fp.font.size !== 10) || (fp.font.name && fp.font.name !== FONT) || !near(fp.lineSpacing, PT.single, 1.5) || !near(fp.firstLineIndent, PT.half, 2)) {
              add({ category: "Footnotes", message: "Footnotes: Times New Roman 10 pt, single-spaced, first line indented 0.5″, blank line between.", loc: loc, where: snippet, fix: "fnformat" });
            }
            R.checkText(t, "footnote", { dash: prefs.dash }).forEach(function (x) {
              add({ category: x.category, severity: x.severity, message: x.message, loc: loc, where: snippet, match: x.match, replacement: x.replacement });
            });
            if (/^\s*Ibid/.test(t)) {
              var hits = fp.search("Ibid", { matchCase: true });
              hits.load("items/font/italic");
              ibidChecks.push({ hits: hits, loc: loc, snippet: snippet });
            }
          });
        });
        if (ibidChecks.length) {
          await ctx.sync();
          ibidChecks.forEach(function (c) {
            if (c.hits.items.length && c.hits.items[0].font.italic) add({ category: "Footnotes", message: "“Ibid.” is not italicized.", loc: c.loc, where: c.snippet, match: "Ibid", fix: "unitalic" });
          });
        }
        var anyVersion = map.info.some(function (it) { return /\([^()]*\d+:\d+[^()]*\b(NIV|ESV|NRSV|NASB|KJV|NLT|CSB|NKJV|RSV)\)/.test(it.text); });
        if (anyVersion && !hasPrimaryNote) {
          add({ category: "Scripture", severity: "info", message: "Footnote your first scripture quotation: “Unless otherwise noted, all scripture references are from the …”.", loc: null });
        }
      }
    });

    lastIssues = issues;
    renderIssues(issues);
  }

  var ORDER = ["Structure", "Formatting", "Headings", "Quotations", "Footnotes", "Bibliography", "Scripture", "Capitalization", "Numbers"];

  function renderIssues(issues) {
    var sum = $("checkSummary"), out = $("checkResults");
    out.innerHTML = "";
    sum.classList.remove("hidden");
    if (!issues.length) {
      sum.className = "summary clean";
      sum.textContent = "Looks good. No problems found.";
      $("btnFixAllFormat").classList.add("hidden");
      return;
    }
    var warn = issues.filter(function (i) { return i.severity !== "info"; }).length;
    sum.className = "summary issues";
    sum.textContent = issues.length + " item" + (issues.length === 1 ? "" : "s") + " to review" + (warn < issues.length ? " (" + (issues.length - warn) + " suggestions)" : "");
    var fmtCount = issues.filter(function (i) { return FORMAT_FIXES[i.fix]; }).length;
    $("btnFixAllFormat").classList.toggle("hidden", fmtCount < 2);
    $("btnFixAllFormat").textContent = "Fix all " + fmtCount + " formatting issues";

    var groups = {};
    issues.forEach(function (it, i) { it.id = i; (groups[it.category] = groups[it.category] || []).push(it); });
    Object.keys(groups).sort(function (a, b) { return ORDER.indexOf(a) - ORDER.indexOf(b); }).forEach(function (g) {
      var wrap = document.createElement("div");
      wrap.className = "group";
      var h = document.createElement("h3");
      h.textContent = g + " (" + groups[g].length + ")";
      wrap.appendChild(h);
      groups[g].forEach(function (it) { wrap.appendChild(issueCard(it)); });
      out.appendChild(wrap);
    });
  }

  function issueCard(it) {
    var d = document.createElement("div");
    d.className = "issue" + (it.severity === "info" ? " info" : "");
    d.id = "issue-" + it.id;
    var m = document.createElement("p"); m.className = "msg"; m.textContent = it.message; d.appendChild(m);
    if (it.where) { var w = document.createElement("p"); w.className = "where"; w.textContent = "“" + it.where + "…”"; d.appendChild(w); }
    var a = document.createElement("div"); a.className = "actions";
    if (it.loc) { var s = document.createElement("button"); s.textContent = "Show"; s.onclick = function () { run(s, function () { return showIssue(it); }); }; a.appendChild(s); }
    var canFix = it.fix || (it.replacement && it.match);
    if (canFix) {
      var f = document.createElement("button");
      f.textContent = it.replacement ? "Change to “" + it.replacement + "”" : "Fix";
      f.onclick = function () { run(f, async function () { await fixIssue(it); d.classList.add("fixed"); f.disabled = true; f.textContent = "Fixed"; }); };
      a.appendChild(f);
    }
    if (a.childNodes.length) d.appendChild(a);
    return d;
  }

  function locate(ctx, loc) {
    if (loc.type === "p") {
      var ps = ctx.document.body.paragraphs;
      ps.load("items");
      return { ps: ps, get: function () { return ps.items[loc.index]; } };
    }
    var notes = ctx.document.body.footnotes;
    notes.load("items");
    return {
      ps: notes,
      get: function () { return null; },
      async getFn() {
        await ctx.sync();
        var fpars = notes.items[loc.index].body.paragraphs;
        fpars.load("items");
        await ctx.sync();
        return fpars.items[loc.para];
      }
    };
  }

  async function getPara(ctx, loc) {
    var l = locate(ctx, loc);
    if (loc.type === "p") { await ctx.sync(); return l.get(); }
    return l.getFn();
  }

  async function showIssue(it) {
    await Word.run(async function (ctx) {
      var p = await getPara(ctx, it.loc);
      if (!p) throw new Error("ItemNotFound");
      if (it.match) {
        var hits = p.search(it.match.slice(0, 250), { matchCase: true });
        hits.load("items");
        await ctx.sync();
        if (hits.items.length) { hits.items[0].select(); await ctx.sync(); return; }
      }
      p.select();
      await ctx.sync();
    });
  }

  async function fixIssue(it, ctxIn) {
    var work = async function (ctx) {
      if (it.fix === "pagenum") {
        var secs = ctx.document.sections; secs.load("items"); await ctx.sync();
        secs.items.forEach(function (s) { s.getHeader("Primary").insertOoxml(pageNumberOoxml(), "Replace"); });
        await ctx.sync(); return;
      }
      if (it.fix === "scripture") { await makeScriptureConsistent(); return; }
      var p = await getPara(ctx, it.loc);
      if (!p) throw new Error("ItemNotFound");
      switch (it.fix) {
        case "font": p.font.name = FONT; break;
        case "size": p.font.size = 12; break;
        case "align": p.alignment = "Left"; break;
        case "spacing": p.lineSpacing = PT.double; p.spaceAfter = 0; p.spaceBefore = 0; break;
        case "indent": p.firstLineIndent = PT.half; break;
        case "hanging": case "bibspacing": setBibEntry(p); break;
        case "blockspacing": p.lineSpacing = PT.single; break;
        case "fnformat": setFootnote(p); break;
        case "headingalign": p.alignment = it.level <= 2 ? "Centered" : "Left"; break;
        case "unitalic": {
          var hits = p.search("Ibid", { matchCase: true }); hits.load("items"); await ctx.sync();
          hits.items.forEach(function (h) { h.font.italic = false; });
          break;
        }
        default:
          if (it.replacement && it.match) {
            var r = p.search(it.match.slice(0, 250), { matchCase: true }); r.load("items"); await ctx.sync();
            if (!r.items.length) throw new Error("Couldn’t find that text any more; it may already be fixed.");
            r.items[0].insertText(it.replacement, "Replace");
          }
      }
      await ctx.sync();
    };
    if (ctxIn) return work(ctxIn);
    await Word.run(work);
  }

  async function fixAllFormatting() {
    var targets = lastIssues.filter(function (i) { return FORMAT_FIXES[i.fix]; });
    await Word.run(async function (ctx) {
      for (var i = 0; i < targets.length; i++) await fixIssue(targets[i], ctx);
    });
    targets.forEach(function (t) { var el = $("issue-" + t.id); if (el) el.classList.add("fixed"); });
    toast("Fixed " + targets.length + " formatting issues. Run the check again to confirm.");
  }

  /* ================= startup ================= */
  function initTabs() {
    document.querySelectorAll(".tab").forEach(function (t) {
      t.onclick = function () {
        document.querySelectorAll(".tab").forEach(function (x) { x.classList.toggle("active", x === t); });
        document.querySelectorAll(".panel").forEach(function (p) { p.classList.toggle("active", p.id === "panel-" + t.getAttribute("data-tab")); });
      };
    });
  }

  function start(inWord) {
    initTabs(); initTitlePage(); initFormat(); initScripture(); initCheck();
    if (!inWord) $("notInWord").classList.remove("hidden");
  }

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function (info) {
      var inWord = info && info.host === Office.HostType.Word;
      if (inWord) {
        api.v13 = Office.context.requirements.isSetSupported("WordApi", "1.3");
        api.v15 = Office.context.requirements.isSetSupported("WordApi", "1.5");
      }
      start(inWord);
    });
  } else {
    document.addEventListener("DOMContentLoaded", function () { start(false); });
  }
})();
