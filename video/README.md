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

Every time in the animation lives in one object, `T`, at the top of the `<script>` in
`make_video.py`. Seconds throughout. Its entries come in two kinds:

  - **the ones you set** — the beats and their durations;
  - **the ones `derive()` computes** — `showAt`, `stripAt`, `stripTo`, `litTo`, `padTo`,
    `cardAt`, `cardTo`, `landsAt`. Never write these by hand. They exist so that the
    relationships which *must* hold cannot be broken by changing one number and forgetting
    the one that had to agree with it.

**So the rule is: change a beat, then call `derive()`.** In the file that happens on its
own at load. In the browser console it does not, so say it yourself — otherwise the
derived values keep their old figures and the couplings come quietly apart.

Suffixes are consistent: `At` is when something starts, `For` how long it takes, `To` when
it ends, `Per` the stagger between the three keys (or the five destinations), and `Lag` an
offset from the beat it follows.

### Try it in the browser first

Open `video/introducer.html`, open the console, and move something:

```javascript
T.threadAt = 10; derive(); seek(11)
```

Reload to get back — nothing is saved. This is much the fastest loop: scrub, tweak, scrub
again, and only edit the file once you know the number you want.

### The knobs

| Beat | Set these |
|---|---|
| The opener, then the cut | `cutAt` 3.0, `cutFor` 1.4 |
| The named ranges appear | `namesAt` 4.6, `namesFor` 1.0 |
| The specification slides in | `specAt` 5.0, `specFor` 2.2 |
| The rows lift, each with its reference | `liftAt` 8.4, `liftPer` 0.35, `liftFor` 0.4, `chipFor` 0.8 |
| Threads draw, then cells light | `threadAt` 12.0, `threadPer` 0.4, `threadTo` 17.5, `litOutFor` 0.35 |
| Grow, and strip the rest away | `morphAt` 21.0, `growFor` 1.6, `padFor` 0.9, `chromeGone` 22.4, `stripLag` 0.4, `stripFor` 1.6 |
| The flight, and the handover | `flightAt` 23.2, `flightTo` 25.4, `flightPer` 0.18, `padOutFor` 1.0, `cardLag` 0.1, `cardFor` 0.9, `handLead` 0.1, `slotLead` 0.05, `handKey` 0.55, `handBlock` 0.40 |
| Where the JSON goes next | `shrinkAt` 26.6, `shrinkTo` 29.0, `fanAt` 27.0, `fanFor` 1.0, `destAt` 27.2, `destFor` 2.0, `destPer` 0.22 |
| The close | `ctaAt` 32.9, `ctaTo` 34.1, `end` 35 |

One entry is not a time: `threadLit` (0.85) is the fraction of a thread's draw at which its
cells begin to light, so it follows the thread however you retime it.

### Four things that do not follow automatically

`derive()` handles the couplings *within* the morph. Four constraints run between beats,
and those are yours to keep:

  1. **Captions do not move with their beat.** They sit in `CAPTIONS` just below `T`, each
     with its own absolute window. Retime the threads and the caption over them stays put.
  2. **`flightAt` must not come before `morphAt + growFor`** — the keys and values finish
     growing before they fly. Currently 23.2 against 22.6.
  3. **`shrinkAt` must not come before `T.landsAt`** — the JSON has to be assembled before
     it shrinks to make room for the destinations. Currently 26.6 against 26.16.
  4. **`end` must not come before `ctaTo`**, or the last of the video is unreachable.

Beyond those, ordinary reading order: the names before the specification, the specification
before the rows lift, the rows before the threads that leave them.

There is also no "shift everything after here" operation. Each beat's `At` is absolute, so
taking two seconds out of the middle means moving every `At` that follows. If retiming turns
into a lot of that, the fix is to derive the later beats from the earlier ones, exactly as
`derive()` already does inside the morph.

### Then make it permanent

```
uv run python video/make_video.py
uv run --group video python video/check.py
uv run --group video python video/contact.py 20:26:0.5
```

Regenerate, check, and look at the frames either side of what you moved.

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
