# The EDJAS introducer

A ~34-second animation of what EDJAS does, generated from the real package: a crowded
government workbook, a small specification, three references resolving, and the JSON they
produce. It exists to be recorded as a promotional video, with a voiceover added
afterwards.

Nothing on screen is a mock-up. The opening wall of data is the Scottish local-government
finance workbook that ships in `edjas-examples`; the grid is read with openpyxl; the
specification is the `quarter.toml` in this directory; and the JSON is what `read_spec()`
actually returns. Re-running the generator after a change to EDJAS produces a truthful
picture of the changed EDJAS, which is the whole point of building it this way.

None of this ships in the wheel or the sdist. It is a checkout-only tool, like `docs/`.

```
uv sync --group video
uv run python video/make_video.py
```

That writes `video/introducer.html` — self-contained, no server, no network. Open it in
any browser and use the scrubber; the frame scales to whatever window you give it.

## How it is put together

`make_video.py` is in two halves. The Python reads your files and produces a JSON payload;
the template string is the page, and its `seek(t)` computes **every** element's position,
scale, opacity and colour from `t` alone. No accumulated state, no timers driving the
visuals. That single constraint is what makes the scrubber, the contact sheets and any
future frame capture agree with each other, and it is why `check.py` can step the entire
timeline and treat the result as exhaustive rather than as a sample.

Everything positional is **measured from the live layout** rather than typed in: thread
endpoints, the boxes the keys and values fly to in the JSON, the outlines around the named
ranges. Change the font or move the panels and the animation still converges, because the
targets are read back out of the rendered page.

The three references in `quarter.toml` are deliberately of different kinds — a bare cell
address, a named range extracted as a list, a named range extracted as an object — and the
sheet is annotated to say so. The `Sales` and `Hours` badges are the workbook's own
property, so they appear with the workbook; `B2` is not in the workbook at all, so it
appears with the line of the specification that types it. Both labels are derived from the
spec at build time, not hardcoded.

## Checking it

```
uv run python video/check.py
uv run python video/contact.py             # 0-34s at one-second intervals
uv run python video/contact.py 20:26:0.5   # a range, start:stop:step
```

`check.py` asserts that `seek()` is total over the whole timeline and that the frame
geometry is identical at every viewport size. `contact.py` tiles captured frames into
`contact.png`, which is the only reliable way to see pacing: a sequence that scrubs nicely
can still hold one unchanging picture for three seconds, and that is invisible until the
frames are side by side.

Both drive a headless Chromium, reusing whatever is already in Playwright's download cache
before fetching one of its own.

## Still to do

  - Frame capture to a video file: step `seek(t)` at 30fps and pipe the PNGs through
    ffmpeg. The animation is already deterministic, so this is mechanical.
  - The voiceover, which has to be recorded.
  - Pacing. `contact.py` shows roughly eight seconds across the 34 with no visible
    change — the opener holds 0-3s, there is a gap around 9.5-12s between the rows
    lifting and the threads starting, and the fan-out sits still from about 31.5s.
