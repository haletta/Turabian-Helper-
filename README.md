# Turabian Helper — setup and test guide

A Word add-in that handles the repetitive parts of a Turabian paper, built from the Denver Seminary Writing Center guides (Formatting Checklist, Citing Common Sources, Guidance for Research Papers, Tips on Turabian, Title Page template).

It opens as a panel on the right side of Word with four tabs:

| Tab | What it does |
|---|---|
| **Title page** | Fills in the Writing Center title page from a short form: title on line 8, colon before the subtitle, name on line 24, course/professor/date double-spaced, no page number. Remembers your name, course and professor. |
| **Format** | One click sets up the whole paper: Times New Roman 12, double spacing, 0.5″ indents, left aligned, 10 pt single-spaced footnotes, page numbers top right starting with 2, 1″ margins. Buttons for heading levels 1–4, block quotations, a Bibliography page and bibliography entries. |
| **Scripture** | Type `jn 3:16` and get `(John 3:16 NIV).` Uses traditional or shorter abbreviations consistently, adds the “Unless otherwise noted…” footnote on your first quotation, and only repeats the version when it changes. Can fix every parenthetical citation in the paper at once. |
| **Check** | Reviews the paper against the checklist: fonts, spacing, indents, headings, block quotations, footnotes, Ibid., page ranges (Table 23.2), capitalization (Bible, biblical, scripture, Gospel of Mark), Bible abbreviations, and bibliography problems (Bible or common dictionaries listed, alphabetical order, hanging indents). Each item has **Show** and, where safe, **Fix**. |

Use **Zotero** for your sources, footnotes and bibliography text; use Turabian Helper for everything around them.

**Requires** Microsoft 365 (Windows, Mac or Word on the web) or Word 2021/2024. Footnote features need an up-to-date Microsoft 365.

---

## Step 1 — Put the add-in online (GitHub Pages, free, ~10 minutes)

Word loads add-ins from a web address, so the files need a free home on the internet. Nothing personal is stored there; it's just the add-in's code.

1. Go to **github.com** and create a free account. Your **username** matters: it becomes part of the address. (Example below: `jdoe`.)
2. Click **+** (top right) → **New repository**.
   - Repository name: `turabian-helper` (exactly this)
   - Choose **Public**
   - Click **Create repository**.
3. On the next page click **uploading an existing file**.
4. Unzip `turabian-helper.zip` on your computer. Open the `turabian-helper` folder, select **everything inside it** (including the `assets` folder) and drag it into the GitHub page. Click **Commit changes**.
5. Go to the repository's **Settings** → **Pages** (left menu).
   - Source: **Deploy from a branch**
   - Branch: **main**, folder **/ (root)** → **Save**.
6. Wait 1–2 minutes, then open this in your browser (with your username):
   `https://jdoe.github.io/turabian-helper/taskpane.html`
   You should see the Turabian Helper panel with a yellow note saying it only works inside Word. That means hosting works.

## Step 2 — Point the manifest at your address

The **manifest.xml** file tells Word where the add-in lives.

1. Open `manifest.xml` (from the unzipped folder on your computer) in **Notepad** (Windows) or **TextEdit** (Mac; use Format → Make Plain Text).
2. **Find & Replace** `YOUR-GITHUB-USERNAME` with your GitHub username in **lowercase** (e.g. `jdoe`). Use **Replace All**.
3. Save.

## Step 3 — Install it in Word (“sideloading”)

Pick the one you use.

### Word on the web (easiest; works on any computer)
1. Go to **office.com**, open any Word document.
2. **Home** tab → **Add-ins** → **More Add-ins** (or **Advanced**) → **My Add-ins** → **Upload My Add-in**.
3. Choose your edited `manifest.xml` → **Upload**.
4. **Turabian Helper** appears on the **Home** tab, at the right end.

### Word for Windows
1. Make a folder, e.g. `C:\TurabianAddin`, and copy your edited `manifest.xml` into it.
2. Right-click the folder → **Properties** → **Sharing** → **Share…** → **Share** → **Done**. Copy the **network path** shown (looks like `\\YOUR-PC\TurabianAddin`).
3. In Word: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.
4. Paste the network path into **Catalog Url** → **Add catalog** → tick **Show in Menu** → **OK** → **OK**.
5. Close and reopen Word.
6. **Home → Add-ins → More Add-ins** (or **Insert → My Add-ins**) → **SHARED FOLDER** tab → **Turabian Helper** → **Add**.

### Word for Mac
1. In Finder: **Go → Go to Folder…** and paste
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
   (If `wef` doesn't exist, go to `…/Documents` and create a folder named `wef`.)
2. Copy your edited `manifest.xml` into that folder.
3. Quit and reopen Word, open a document.
4. **Home → Add-ins** (or **Insert → My Add-ins** dropdown) → **Turabian Helper**.

---

## Step 4 — Test it (about 15 minutes)

Use the practice paper `test-files/Turabian Helper Test Paper.docx`. It has mistakes planted on purpose. (On Word on the web, upload it to OneDrive first.)

Open the paper and click **Turabian Helper** on the Home tab.

### Test 1 — Check finds the problems
1. **Check** tab → **Check my paper**.
2. You should see items under these groups (exact counts may vary slightly):
   - **Structure**: no title page, no page numbers
   - **Formatting**: Calibri font, 11 pt, justified text, not double-spaced, missing 0.5″ indents
   - **Headings**: level 1 should be centered
   - **Quotations**: Moo's long quotation should be a block quotation
   - **Footnotes**: wrong size/spacing, `ibid.` not capitalized, missing period, `Ibid` needs a period, italic Ibid
   - **Bibliography**: no hanging indent, Bible listed, Merriam-Webster listed, out of alphabetical order
   - **Scripture**: mixed `Rom.` and `Rom`, `Ps 23` should be spelled out, “Jeremiah chapter 31”, “verses” inside parentheses
   - **Capitalization**: `bible`, `Biblical`, `gospel of John`
   - **Numbers**: `642–643` → `642–43`, `1,234` → `1234`, `23-36` → `23–36`
3. Click **Show** on a few. Word should jump to and highlight the text.
4. Click **Change to “Bible”**. The word changes in the paper and the card fades.

✅ Pass if the groups above appear and Show/Change work.

### Test 2 — One-click setup
1. **Format** tab → **Set up this paper in Turabian**.
2. A list of ✓ lines appears. The paper should now be Times New Roman 12, double-spaced, indented, left-aligned; footnotes 10 pt with a blank line between; bibliography entries with hanging indents.
3. **Check** tab → **Check my paper** again. The Formatting items should be gone; content items (Bible in bibliography, Ibid, etc.) remain for you to decide.

✅ Pass if formatting issues disappear on the second check.

### Test 3 — Title page and page numbers
1. **Title page** tab. Fill in title, subtitle, your name, course, professor. Date is pre-filled.
2. **Insert title page**.
3. Compare page 1 with the Writing Center “Turabian Title Page” template: title around a third of the way down with a colon before the subtitle, your name lower down, then course, professor, date with blank lines between.
4. Scroll to page 2: the number **2** should be at the top right. Page 1 has no number.
5. Change the title and click **Insert title page** again: it replaces the old one (no duplicate).
6. Close the add-in and reopen it: your name, course and professor are still filled in.

✅ Pass if layout matches and numbering starts with 2 on page 2.

### Test 4 — Scripture
1. In the paper, click right after the closing quotation mark of a quote (or at the end of any sentence, before the period).
2. **Scripture** tab → Reference `prov 3:5a`, Version `NIV`, **After a quotation**. Preview shows `(Prov. 3:5a NIV).` and a note about the footnote.
3. **Insert at cursor**. The citation appears with a period and a new footnote: *Unless otherwise noted, all scripture references are from the NIV.* “Primary version” now shows **NIV**.
4. Insert `jn 11:35` with NIV → `(John 11:35).` (no version, it's your primary).
5. Insert `jn 11:35` with KJV → `(John 11:35 KJV).`
6. Choose **Named in my sentence**, type `1 thess 4:11` → preview `1 Thessalonians 4:11`.
7. Click **Make all scripture citations consistent** → `(Rom 5:8)` becomes `(Rom. 5:8)` and `(Jn 3:16)` becomes `(John 3:16)`.
8. Under **Format → Preferences** switch to **Shorter** and run it again: now `(Rom 5:8)`, `(Jn 3:16)`.

✅ Pass if previews and inserted text match.

### Test 5 — Headings, block quotation, bibliography
1. Click in “The Meaning of Justification” → **Format → Level 2** → centered, regular type.
2. Cut Moo's quotation into its own paragraph, remove the quotation marks, click in it → **Block quotation** → single-spaced, indented 0.5″.
3. Select the four bibliography entries → **Format selected entries**.
4. Move the cursor to the end of a fresh document → **Add Bibliography page** → a new page with a centered “Bibliography” and space below.

✅ Pass if each looks like the Turabian Sample Paper.

### Test 6 — With Zotero (optional)
1. Insert two or three citations with the Zotero Word plugin (style: *Chicago Manual of Style (full note)*), then add the bibliography.
2. Run **Format → Set up this paper** and **Check my paper**. Zotero's footnotes and bibliography should be formatted, and Check should flag only content questions.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Button doesn't appear in Word | Check that every `YOUR-GITHUB-USERNAME` is replaced, lowercase. Reopen Word. On Windows, make sure **Show in Menu** is ticked. |
| Panel is blank or says “can't load” | Open `https://yourname.github.io/turabian-helper/taskpane.html` in a browser. If that fails, GitHub Pages isn't on yet (Step 1.5) or files are in a subfolder. `taskpane.html` must be at the top level of the repository. |
| “This version of Word can't reach footnotes…” | Update Microsoft 365. Older Word can still do everything else. |
| Margins weren't changed | Some Word versions don't let add-ins change margins: **Layout → Margins → Normal**. |
| Page number shows on page 1 | Click in the header and press **F9**, or close and reopen the document. |
| Made a change you don't want | **Ctrl+Z** / **⌘Z** undoes add-in changes like any other edit. |

## Updating the add-in later
Edit or replace the files in your GitHub repository (**Add file → Upload files**). Word picks up the new version the next time the panel opens. You only re-sideload if `manifest.xml` changes.

## Files
| File | Purpose |
|---|---|
| `manifest.xml` | Tells Word where the add-in lives (edit your username in) |
| `taskpane.html`, `taskpane.css`, `taskpane.js` | The panel and what it does in Word |
| `rules.js` | All the Turabian rules (Bible abbreviations, Table 23.2, checks) |
| `commands.html`, `help.html`, `assets/` | Required by Word: ribbon helper, help page, icons |
| `tests/` | Automated tests (`node tests/rules.test.js`) — not needed in Word |
| `test-files/` | Practice paper with planted mistakes |

## Rule choices you can change
Your guides disagree in a few places. Defaults follow the newest guide (*Citing Common Sources*, 2026):
- **Number ranges** use an en dash (642–43). Switch to hyphens in **Format → Preferences**.
- **Bible abbreviations** default to traditional (Jer., 1 Cor.). Switch to shorter in **Preferences**.
- In running text the checker only asks you to spell out a book when you cite a **whole chapter** (“Psalm 23,” not “Ps 23”), since the slides disagree about verse references.
