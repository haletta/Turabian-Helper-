// Minimal Office.js / Word mock for exercising the task pane in a browser test.
(function () {
  var log = (window.__wordLog = []);
  function rangeMock(name) {
    var target = { text: "", font: {}, items: [] };
    return new Proxy(target, {
      get: function (t, k) {
        if (k in t) return t[k];
        if (k === "then") return undefined;
        return function () { log.push(name + "." + String(k)); return rangeMock(name + "." + String(k)); };
      },
      set: function (t, k, v) { t[k] = v; return true; }
    });
  }
  function coll(items) {
    return { items: items, load: function () { return this; }, getFirst: function () { return items[0]; }, getLast: function () { return items[items.length - 1]; } };
  }
  function para(text, props) {
    var p = Object.assign({
      text: text, alignment: "Left", lineSpacing: 24, firstLineIndent: 36, leftIndent: 0,
      styleBuiltIn: "Normal", tableNestingLevel: 0, spaceAfter: 0,
      font: { name: "Times New Roman", size: 12, bold: false, italic: false }
    }, props || {});
    p.load = function () { return p; };
    p.search = function (t) { var hits = text.indexOf(t) >= 0 ? [Object.assign(rangeMock("hit"), { font: { italic: /Ibid/.test(t) && p.__ibidItalic } })] : []; return coll(hits); };
    p.select = function () { log.push("select:" + text.slice(0, 20)); };
    p.insertParagraph = function (t) { log.push("insertParagraph:" + t); return para(t); };
    p.getRange = function () { return rangeMock("range"); };
    p.insertBreak = function () { log.push("insertBreak"); };
    p.insertText = function (t) { log.push("insertText:" + t); return rangeMock("ins"); };
    return p;
  }
  var paras = [
    para("Grace in Romans:", { alignment: "Centered", firstLineIndent: 0, lineSpacing: 12 }),
    para("Jane Doe", { alignment: "Centered", firstLineIndent: 0, lineSpacing: 12 }),
    para("NT 501", { alignment: "Centered", firstLineIndent: 0, lineSpacing: 12 }),
    para("Dr. Smith", { alignment: "Centered", firstLineIndent: 0, lineSpacing: 12 }),
    para("October 5, 2026", { alignment: "Centered", firstLineIndent: 0, lineSpacing: 12 }),
    para("Introduction", { alignment: "Centered", styleBuiltIn: "Heading1", firstLineIndent: 0, font: { name: "Times New Roman", size: 12, bold: true } }),
    para("We read the bible and this Biblical theme in Jeremiah chapter 31 shows grace (Jn 3:16) and love (1 Cor. 13:4 NIV).", { font: { name: "Calibri", size: 11 }, alignment: "Justified", lineSpacing: 12, firstLineIndent: 0 }),
    para("Only Child Heading", { alignment: "Left", styleBuiltIn: "Heading2", firstLineIndent: 0 }),
    para("The quoted passage is long.", { leftIndent: 36, firstLineIndent: 0, lineSpacing: 24 }),
    para("Bibliography", { alignment: "Centered", firstLineIndent: 0 }),
    para("Segal, Benjamin J. A New Psalm. Jerusalem: Gefen, 2013.", { firstLineIndent: 0, leftIndent: 0, lineSpacing: 24 }),
    para("Klein, William W. “Exegetical Rigor.” 23-36. Grand Rapids: Eerdmans, 2003.", { firstLineIndent: -36, leftIndent: 36, lineSpacing: 12 }),
    para("The Holy Bible: New International Version. Grand Rapids: Zondervan, 2011.", { firstLineIndent: -36, leftIndent: 36, lineSpacing: 12 })
  ];
  var fnParas = [
    para("Benjamin J. Segal, A New Psalm (Jerusalem: Gefen, 2013), 25.", { font: { name: "Times New Roman", size: 10 }, lineSpacing: 12 }),
    para("ibid., 642–643", { font: { name: "Times New Roman", size: 12 }, lineSpacing: 24 }),
    Object.assign(para("Ibid., 43.", { font: { name: "Times New Roman", size: 10 }, lineSpacing: 12 }), { __ibidItalic: true })
  ];
  var notes = coll(fnParas.map(function (fp) { return { body: { paragraphs: coll([fp]) } }; }));
  var styleMock = function () { return { isNullObject: false, font: {}, paragraphFormat: {}, load: function () { return this; } }; };
  var header = { text: "", load: function () { return this; }, getOoxml: function () { return { value: this.xml || "" }; }, insertOoxml: function (x) { header.xml = x; log.push("header.insertOoxml:" + (x.indexOf(" IF ") > 0)); } };
  var section = { getHeader: function () { return header; }, pageSetup: {} };
  var doc = {
    body: {
      paragraphs: coll(paras), footnotes: notes,
      insertParagraph: function (t) { log.push("body.insertParagraph:" + t); return para(t); },
      insertBreak: function () { log.push("body.insertBreak"); }
    },
    contentControls: { getByTag: function () { return coll([]); } },
    getSelection: function () {
      var r = rangeMock("sel");
      r.paragraphs = coll([paras[6]]);
      r.paragraphs.getFirst = function () { return paras[6]; };
      r.insertText = function (t) { log.push("sel.insertText:" + t); return rangeMock("ins"); };
      return r;
    },
    sections: Object.assign(coll([section]), { getFirst: function () { return section; } }),
    getStyles: function () { return { getByNameOrNullObject: styleMock }; },
    addStyle: function (n) { log.push("addStyle:" + n); return styleMock(); }
  };
  var settings = {};
  window.Word = { run: async function (fn) { return fn({ document: doc, sync: async function () {} }); } };
  window.Office = {
    HostType: { Word: "Word" },
    onReady: function (cb) { setTimeout(function () { cb({ host: "Word" }); }, 0); },
    context: {
      requirements: { isSetSupported: function () { return true; } },
      document: { settings: { get: function (k) { return settings[k]; }, set: function (k, v) { settings[k] = v; }, remove: function (k) { delete settings[k]; }, saveAsync: function (cb) { cb && cb(); } } }
    }
  };
})();
