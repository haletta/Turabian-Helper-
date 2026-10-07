# Turabian Helper — project context

A Microsoft Word add-in (Office.js task pane) that automates Turabian 9th ed. notes-bibliography formatting for Denver Seminary papers. Built from the Denver Seminary Writing Center guides: Turabian Formatting Checklist, Citing Common Sources in Turabian (2026), Guidance for Research Papers, Tips on Turabian (DS Manual), and the Turabian Title Page template.

## Who uses it and how
- Owner: a seminary student on a **Mac** (Word for Mac desktop; Word on the web also works).
- Zotero handles sources, footnote text and bibliography text. This add-in covers everything around them. Don't rebuild Zotero features.
- Hosted on GitHub Pages: repo `haletta/Turabian-Helper-` (note the trailing hyphen) → `https://haletta.github.io/Turabian-Helper-/`.
- Installed on Mac by copying `manifest.xml` into `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`, then restarting Word.
- After changing any file: commit and push to `main`; GitHub Pages redeploys in 1–2 min. Only re-copy the manifest into `wef` if `manifest.xml` itself changed.

## Files
| File | Role |
|---|---|
| `manifest.xml` | Add-in manifest (XML, not the unified JSON manifest). All URLs point to haletta.github.io/Turabian-Helper-. Requires WordApi 1.3. |
| `taskpane.html` / `.css` / `.js` | The panel: tabs Title page, Format, Scripture, Check. |
| `rules.js` | Pure rules with no Word dependency (UMD, works in Node and browser): 66 Bible books with traditional/shorter abbreviations, Table 23.2 range abbreviation, scripture parsing/formatting, title page line layout, text/footnote/bibliography checks. |
| `commands.html`, `help.html`, `assets/` | Required by the manifest (function file, help URL, icons 16/32/64/80/128). |
| `tests/rules.test.js` | `node tests/rules.test.js` — 76 unit tests, all must pass. |
| `tests/ui_test.py` + `tests/office-mock.js` | Playwright run of the panel against a mocked Word API: `python3 tests/ui_test.py <screenshot-dir>`. |
| `test-files/Turabian Helper Test Paper.docx` | Practice paper with planted mistakes, used for manual tests (README Step 4). |

## Rules implemented (source of truth: the Writing Center guides)
- Body: Times New Roman 12, double-spaced, 0.5″ first-line indent, left aligned (ragged right), 1″ margins.
- Page numbers: top right, none on title page, first text page shows 2. Implemented as header field `{ IF { PAGE } > 1 "{ PAGE }" "" }` so no "different first page" setting is needed.
- Title page: single-spaced centered lines; title on line 8 (colon added when a subtitle follows); student name on line 24; course, professor, date double-spaced. Wrapped in a hidden content control tagged `turabian-title-page` so re-inserting replaces it.
- Headings: L1 centered bold, L2 centered regular, L3 flush-left bold, L4 flush-left regular sentence-case; extra space above (12 pt) so it reads as a triple space. Warn when a section has only one subheading.
- Block quotation (5+ lines): single-spaced, 0.5″ left indent, no quotation marks.
- Footnotes: 10 pt, single-spaced, 0.5″ first-line indent, blank line between, end with a period; "Ibid." capitalized, with period, not italic.
- Bibliography: new page, centered "Bibliography" with two blank lines below; entries single-spaced, blank line between, 0.5″ hanging indent, alphabetical (ignore leading quotes and The/A/An). Flag the Bible (unless study notes), common dictionaries, personal communications/interviews, blogs.
- Scripture: parenthetical `(Jer. 31:4 NIV)` when quoting; full book name in running text for whole chapters; first quotation names the version and adds the footnote "Unless otherwise noted, all scripture references are from the NIV."; later citations show the version only when it differs. Use "v./vv." inside parentheses and notes. Never "chapter"/"verse" after a book name.
- Capitalization: Bible, Holy Scriptures, Word of God, Gospel of Mark; lowercase biblical, scripture(s), scriptural, gospel (content).
- Numbers: Table 23.2 range abbreviation (321–28, 101–9, 1496–500); no commas in page numbers.

## Decisions made with the owner
- Where the guides disagree, follow the newest (Citing Common Sources, 2026): **en dash** in ranges by default (user can switch to hyphen in Format → Preferences); **traditional** Bible abbreviations by default (switchable to shorter).
- Spell-out-book check only fires for whole-chapter references in running text, because the guides contradict each other about verse references.
- Scope is Turabian only for now. SBL and APA (used by counseling programs) are possible future additions.
- Preferences and the remembered name/course/professor live in the pane's localStorage; the primary Bible version is stored per document in Office document settings.

## Known limits / things to verify in real Word
- Never tested inside real Word yet, only against a mock. Watch: page number field on page 2, footnote features (need WordApi 1.5), margins (pageSetup isn't available in every Word version; the pane falls back to telling the user).
- Older Word builds can't reach footnotes from add-ins; the pane detects this and says so.
- Hand-typed citations aren't parsed into structured sources (by design; Zotero does that).

## Working conventions
- Keep `rules.js` free of Office.js so it stays unit-testable; put Word calls in `taskpane.js`.
- Add a unit test in `tests/rules.test.js` for every rule change.
- User-facing text is plain language for a non-developer; no jargon in the panel.