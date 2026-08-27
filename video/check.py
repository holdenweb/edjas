"""Two properties the introducer must hold, checked by driving it rather than reading it.

    uv run --group video python video/check.py

  - **seek() is total.**  Stepping the whole timeline raises nothing, and no element ever
    leaves the 1280x720 frame.  Every element's state is computed from t alone, so this is
    a real exhaustive check rather than a sample: there is no accumulated state for a
    later frame to depend on.

  - **The frame is scale-invariant.**  seek() works in frame coordinates and divides out
    the display scale, so the geometry must be identical at every viewport size.  The
    animation measures the live layout -- thread endpoints, flyer targets, range outlines
    -- and one forgotten division by SCALE would make all of it drift with the window.

Exits non-zero if either fails.
"""

import sys

from contact import chromium, frame

# The flyers are the only things that move far enough to leave the frame, and they carry
# their own transforms, so their boxes are what a bounds check has to look at.
BOUNDS = """() => {
  const s = document.getElementById('stage').getBoundingClientRect(), bad = [];
  for (let t = 0; t <= 35.0001; t += 0.05) {
    seek(+t.toFixed(2));
    for (const f of FLY) {
      if (+f.el.style.opacity < 0.05) continue;
      const r = f.el.getBoundingClientRect();
      if (r.left - s.left < -4 || r.top - s.top < -4 ||
          r.right - s.left > s.width + 4 || r.bottom - s.top > s.height + 4)
        bad.push([+t.toFixed(2), f.key,
                  Math.round(r.left - s.left), Math.round(r.top - s.top)]);
    }
  }
  return bad;
}"""

GEOMETRY = """() => {
  const s = document.getElementById('stage').getBoundingClientRect();
  const at = r => [Math.round((r.left - s.left) / SCALE * 10) / 10,
                   Math.round((r.top - s.top) / SCALE * 10) / 10];
  seek(22.0);                                  /* mid-morph: everything is in motion */
  return {
    scale: +SCALE.toFixed(4),
    shown: [Math.round(s.width), Math.round(s.height)],
    where: FLY.map(f => [f.key, ...at(f.el.getBoundingClientRect())])
             .concat([...document.querySelectorAll('#labels .chip')]
                     .map(c => [c.textContent, ...at(c.getBoundingClientRect())])),
  };
}"""

VIEWPORTS = [(1400, 900), (900, 620), (2200, 1300), (700, 500)]


def main():
    from playwright.sync_api import sync_playwright

    failures = []
    with sync_playwright() as playwright:
        browser = chromium(playwright)

        errors = []
        page = frame(browser)
        page.on("pageerror", lambda e: errors.append(str(e)))
        offstage = page.evaluate(BOUNDS)
        page.close()
        print(f"seek() over 0-35s at 0.05s: {len(errors)} page errors, "
              f"{len(offstage)} elements out of frame")
        if errors:
            failures.append(f"page errors: {errors[:3]}")
        if offstage:
            failures.append(f"out of frame: {offstage[:5]}")

        reference = None
        for width, height in VIEWPORTS:
            page = frame(browser, width, height, pin=False)
            seen = page.evaluate(GEOMETRY)
            page.close()
            if reference is None:
                reference = seen["where"]
            same = seen["where"] == reference
            print(f"{width}x{height}: scale {seen['scale']:<7} shown {seen['shown']}  "
                  f"geometry identical: {same}")
            if not same:
                failures.append(f"geometry drifted at {width}x{height}")
        browser.close()

    for failure in failures:
        print(f"FAIL  {failure}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
