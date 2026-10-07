# Drives the task pane in headless Chromium against a mock Word document.
# Run: python3 tests/ui_test.py <output-dir-for-screenshots>
import pathlib, sys
from playwright.sync_api import sync_playwright

root = pathlib.Path(__file__).resolve().parent.parent
out = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else ".")
errors = []

with sync_playwright() as p:
    b = p.chromium.launch()
    pg = b.new_page(viewport={"width": 340, "height": 900})
    pg.on("pageerror", lambda e: errors.append(str(e)))
    pg.on("console", lambda m: m.type == "error" and errors.append(m.text))
    pg.route("**/office.js", lambda r: r.fulfill(body="", content_type="text/javascript"))
    pg.add_init_script(path=str(root / "tests" / "office-mock.js"))
    pg.goto((root / "taskpane.html").as_uri())
    pg.wait_for_timeout(300)

    # Title page
    pg.fill("#tpTitle", "Grace and Works in Romans")
    pg.fill("#tpSubtitle", "A Study")
    pg.fill("#tpName", "Jane Doe")
    pg.click("#btnTitlePage")
    pg.wait_for_timeout(200)
    pg.screenshot(path=str(out / "ui-title.png"), full_page=True)

    # Format
    pg.click("[data-tab=format]")
    pg.click("#btnSetup")
    pg.wait_for_timeout(300)
    print("SETUP LOG:", pg.inner_text("#setupLog"))
    pg.click("[data-heading='2']")
    pg.click("#btnBlockQuote")
    pg.click("#btnBibHeading")
    pg.wait_for_timeout(200)
    pg.screenshot(path=str(out / "ui-format.png"), full_page=True)

    # Scripture
    pg.click("[data-tab=scripture]")
    pg.fill("#scRef", "jer 31:4")
    pg.fill("#scVersion", "NIV")
    print("PREVIEW 1:", pg.inner_text("#scPreview"), "|", pg.inner_text("#scNote"))
    pg.click("#btnScripture")
    pg.wait_for_timeout(200)
    print("PRIMARY:", pg.inner_text("#scPrimary"))
    pg.fill("#scRef", "John 11:35")
    print("PREVIEW 2:", pg.inner_text("#scPreview"))
    pg.fill("#scVersion", "KJV")
    print("PREVIEW 3:", pg.inner_text("#scPreview"))
    pg.check("input[value=running]")
    print("PREVIEW 4:", pg.inner_text("#scPreview"))
    pg.click("#btnScriptureConsistent")
    pg.wait_for_timeout(200)
    pg.screenshot(path=str(out / "ui-scripture.png"), full_page=True)

    # Check
    pg.click("[data-tab=check]")
    pg.click("#btnCheck")
    pg.wait_for_timeout(400)
    print("SUMMARY:", pg.inner_text("#checkSummary"))
    print("RESULTS:\n" + pg.inner_text("#checkResults"))
    pg.click(".issue .actions button")  # Show first
    fixes = pg.locator(".issue .actions button:has-text('Change to')")
    if fixes.count():
        fixes.first.click()
    if pg.is_visible("#btnFixAllFormat"): pg.click("#btnFixAllFormat")
    pg.wait_for_timeout(300)
    pg.screenshot(path=str(out / "ui-check.png"), full_page=True)

    print("WORD CALLS (sample):", [c for c in pg.evaluate("window.__wordLog") if not c.startswith("range")][:25])
    b.close()

print("ERRORS:", errors or "none")
sys.exit(1 if errors else 0)
