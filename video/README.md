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

## Editing the timeline

The timeline is Python, in `timeline()` in `make_video.py`. Every instant and duration is
named, and each is fixed **either absolutely or against one already fixed**:

```python
t.morphAt  = t.threadTo + 3.5                # the hold, after the threads arrive
t.stripAt  = t.morphAt + 0.4                 # the sheet goes as the flyers take over
t.flightAt = t.morphAt + t.growFor + 0.6     # only once the growing has finished
```

Two rules make that safe, and both are enforced:

  - **using a point before it is fixed is an error**, so the function reads top to bottom
    and nothing can refer forwards;
  - **fixing a point twice is an error**, so a name means one thing.

Between them they rule out cycles by construction, which is why resolving the timeline is
nothing more than running the function.

### Retiming

Change an anchor and everything hanging off it follows. Moving one number — `t.threadAt`
from `t.liftTo + 2.5` to `t.liftTo + 0.5` — takes two seconds of dead air out of the middle
and carries eight later beats and three captions with it, while `ctaAt` and `end` stay
where they are because they are fixed absolutely rather than relatively.

That is the whole point of the arrangement: **what is relative moves, what is absolute
pins**. If you want the closing card to land at a particular second regardless of what
happens earlier, fix it absolutely, as `t.ctaAt` is. If you want it to follow the fan-out,
fix it against `t.fanAt`.

To see what your edit actually resolved to:

```
uv run python video/make_video.py --timeline
```

### What is checked for you

  - **the relationships inside a beat**, because they are written as arithmetic rather than
    as two numbers that have to agree: the flight cannot start before the growing finishes,
    the JSON cannot shrink before it has assembled, the lit cells go out exactly as the
    sheet finishes fading;
  - **the order of the beats**, which arithmetic cannot guarantee — a negative offset or an
    over-enthusiastic anchor is refused by name, `threadAt (1.5s) does not follow liftAt
    (8.4s)`;
  - **that the closing card ends before the timeline does**;
  - **captions**, which are anchored to their beats in `captions()` and so cannot be left
    behind by a retimed beat.

What is *not* checked is whether the result is any good. Nothing knows that three seconds
of an unchanging picture is too long.

### The knobs

Anchors are the entries fixed to a plain number — `cutAt`, `ctaAt`, `end`. Everything else
hangs off them. Suffixes are consistent: `At` is when something starts, `For` how long it
takes, `To` when it ends, `Per` the stagger between the three keys or the five destinations.

One entry is not a time at all: `threadLit` (0.85) is the fraction of a thread's draw at
which its cells begin to light, so it follows the thread however you retime it.

### In the browser

`window.T` is the resolved table, so a single value can still be nudged from the console to
get a feel for it:

```javascript
T.threadAt = 10; seek(11)
```

But the timeline is resolved in Python, so **the console only holds numbers, not
relationships** — nothing else moves with it, and the couplings will be inconsistent until
you reload. Use it to eyeball one value; make the actual change in `timeline()`, where the
relationships live.

## Checking it

```
uv run --group video python video/check.py
uv run --group video python video/contact.py            # 0-34s, one-second intervals
uv run --group video python video/contact.py 20:26:0.5  # a range, start:stop:step
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
