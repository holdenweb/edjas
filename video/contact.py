"""Capture frames of the introducer and tile them into a contact sheet.

    uv run --group video python video/contact.py              # 0-34s, one-second intervals
    uv run --group video python video/contact.py 21 22.5 23   # just those instants
    uv run --group video python video/contact.py 20:26:0.5    # a range, start:stop:step

Frames land in video/frames/ and the sheet in video/contact.png.

A contact sheet is the only honest way to judge pacing: a sequence that reads well when
scrubbed by hand can turn out to hold a single unchanging picture for three seconds, and
that is invisible until the frames are laid out side by side.
"""

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_PAGE = "introducer.html"
END = 34


def chromium(playwright):
    """Playwright's own Chromium, or failing that whatever is in its download cache.

    The cache is often already populated by some other tool even where `playwright
    install` has never been run in this checkout, and reusing it saves a 150MB download.
    """
    try:
        return playwright.chromium.launch()
    except Exception:
        cached = sorted(Path.home().glob("Library/Caches/ms-playwright/chromium-*"
                                         "/chrome-mac/Chromium.app/Contents/MacOS/Chromium"))
        if not cached:
            raise
        return playwright.chromium.launch(executable_path=str(cached[-1]))


def frame(browser, width=1400, height=900, pin=True, page_name=DEFAULT_PAGE):
    """The introducer, loaded from disk with its build passes done.

    fit(1) matters for capture: the page sizes itself to the viewport, so without it the
    frames would come out at whatever scale the browser window happened to imply.  Pass
    pin=False to leave the page at the scale it chose, which is what the scale-invariance
    check needs -- pinning it there would make that check unable to fail.
    """
    path = HERE / page_name
    if not path.exists():
        raise SystemExit(f"{path} does not exist -- build it first")
    page = browser.new_page(viewport={"width": width, "height": height})
    page.goto(path.as_uri())
    page.wait_for_function("typeof seek === 'function' && FLY.length && TAGS.length")
    if pin:
        page.evaluate("fit(1)")
    return page


def times(args, end=END):
    """Instants to capture: bare numbers, start:stop:step ranges, or the whole timeline."""
    if not args:
        return [float(t) for t in range(int(end) + 1)]
    out = []
    for arg in args:
        if ":" in arg:
            start, stop, step = (float(x) for x in arg.split(":"))
            out += [round(start + i * step, 3)
                    for i in range(int(round((stop - start) / step)) + 1)]
        else:
            out.append(float(arg))
    return out


def page_end(page_name=DEFAULT_PAGE):
    """How long the animation runs, read out of the page's own timeline."""
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        browser = chromium(playwright)
        page = frame(browser, page_name=page_name)
        end = page.evaluate("T.end")
        browser.close()
    return end


def capture(wanted, page_name=DEFAULT_PAGE):
    from playwright.sync_api import sync_playwright

    out = HERE / "frames"
    out.mkdir(exist_ok=True)
    for stale in out.glob("t*.png"):
        stale.unlink()
    with sync_playwright() as playwright:
        browser = chromium(playwright)
        page = frame(browser, page_name=page_name)
        stage = page.locator("#stage")
        for t in wanted:
            page.evaluate(f"seek({t})")
            stage.screenshot(path=str(out / f"t{t:07.2f}.png"))
        browser.close()
    return sorted(out.glob("t*.png"))


def tile(files, columns=5, width=400):
    from PIL import Image, ImageDraw

    height, label = round(width * 9 / 16), 20
    rows = (len(files) + columns - 1) // columns
    sheet = Image.new("RGB", (columns * width, rows * (height + label)), "#2a2533")
    draw = ImageDraw.Draw(sheet)
    for i, f in enumerate(files):
        x, y = (i % columns) * width, (i // columns) * (height + label)
        sheet.paste(Image.open(f).convert("RGB").resize((width - 4, height - 4)),
                    (x + 2, y + label + 2))
        draw.text((x + 8, y + 5), f.stem.lstrip("t").lstrip("0") + "s", fill="#e8e2f5")
    return sheet


if __name__ == "__main__":
    argv = sys.argv[1:]
    # --page picks which animation to shoot; everything else is an instant or a range
    page_name = DEFAULT_PAGE
    if "--page" in argv:
        i = argv.index("--page")
        page_name = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]
    # each animation states its own length; ask the page rather than guessing from its name
    wanted = times(argv, end=page_end(page_name))
    files = capture(wanted, page_name)
    sheet = tile(files, columns=5 if len(files) > 8 else 2,
                 width=400 if len(files) > 8 else 620)
    path = HERE / "contact.png"
    sheet.save(path)
    print(f"wrote {path}  ({len(files)} frames, {sheet.width}x{sheet.height})")
