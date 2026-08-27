"""The narrated introducer: a 30-second animation cut to a written voiceover.

An experiment alongside make_video.py rather than a replacement for it.  The two differ in
more than length.  This one opens on the workbook itself -- the wall-of-data shot that
motivated the first version is cut, to keep the whole thing inside thirty seconds -- and it
assembles the JSON rather than cross-fading into it: the cells stack, the labels align to
the top of each block, and only then does the punctuation arrive around them.

The script is in SCRIPT below, verbatim, with the second each line is spoken.  Everything
else follows from it: the timeline's anchors are those seconds and the subtitles are those
words, so a change to the script is a change to the film and the two cannot drift apart.

    uv run python video/narrated.py             # -> video/narrated.html
    uv run python video/narrated.py --timeline  # ...and print what the beats resolved to
"""

import json
import sys
import tomllib
from pathlib import Path

from openpyxl import load_workbook

from edjas import read_spec
from make_video import COLOURS, Timeline, reference, resolve

HERE = Path(__file__).resolve().parent

# The thirty-second script, verbatim: the beat it cues, the second it is spoken, the
# animation it calls for, and the words.  Retiming the video means retiming the script,
# which is the right way round when the words come first and the pictures follow.
SCRIPT = [
    ("sheet", 0.0, "spreadsheet fades in",
     "Here's an example spreadsheet,"),
    ("names", 2.0, "labels appear",
     "with a couple of named ranges."),
    ("spec", 4.0, "specifications fade in",
     "EDJAS lets you describe and name the data you want. For example"),
    ("title", 10.0, "title row highlights and begins snaking towards the spreadsheet",
     "a single cell"),
    ("sales", 11.0, "sales row highlights and begins snaking towards the spreadsheet",
     "a columnar table"),
    ("hours", 12.0, "hours row highlights and begins snaking towards the spreadsheet",
     "or a set of keys and values"),
    ("hold", 14.0, "all three animated threads reach their destination ranges, "
     "which highlight",
     "Your spreadsheet is never modified."),
    ("strip", 17.0, "background to specification and spreadsheet begins to fade, "
     "data and keys are surrounded by an outline",
     "The names and values"),
    ("stack", 19.0, "the data items are stacked vertically, and the labels are vertically "
     "aligned with the top of the relevant data item",
     "are packaged into a JSON data set."),
    ("flow", 22.0, "the outlines fade away, the text shrinks, and the remainder of the "
     "JSON appears",
     "Almost everything can consume JSON,"),
    ("box", 26.0, "the JSON document shrinks to a box labelled JSON, and the API, report, "
     "database, dashboard and web site icons appear over it and zoom out into position",
     "turning your spreadsheets into valuable data sources."),
]
CUE = {beat: at for beat, at, _, _ in SCRIPT}


# API, report, database, dashboard, web site -- in the order the script names them, with the
# offset from the centre each travels to.  Stroke-drawn on currentColor so they inherit.
ICONS = [
    ("An API", -430, -170,
     '<path d="M9 6 4 12l5 6"/><path d="M15 6l5 6-5 6"/>'),
    ("A report", 330, -195,
     '<rect x="5" y="3" width="14" height="18" rx="2"/><path d="M8.5 8h7M8.5 12h7M8.5 16h4"/>'),
    ("A database", -470, 75,
     '<ellipse cx="12" cy="6" rx="7" ry="3"/><path d="M5 6v12c0 1.7 3.1 3 7 3s7-1.3 7-3V6"/>'
     '<path d="M5 12c0 1.7 3.1 3 7 3s7-1.3 7-3"/>'),
    ("A dashboard", 380, 55,
     '<path d="M4 17a8 8 0 1 1 16 0"/><path d="M12 17l3.5-4.5"/><path d="M4 17h16"/>'),
    ("A web site", -40, 235,
     '<circle cx="12" cy="12" r="8"/><path d="M4 12h16"/>'
     '<path d="M12 4a12 12 0 0 1 0 16 12 12 0 0 1 0-16"/>'),
]


def timeline():
    """The beats, taken straight off the voiceover's timings.

    Each anchor is the second the corresponding line is spoken, so retiming the video means
    retiming the script -- which is the right way round when the words come first.
    """
    t = Timeline()
    t.end = 32                                   # 30s of script, and a moment to breathe

    t.sheetAt = CUE["sheet"] + 0.2               # "Here's an example spreadsheet,"
    t.sheetFor = 1.4

    t.namesAt = CUE["names"]                     # "with a couple of named ranges."
    t.namesFor = 1.2

    # "EDJAS lets you describe and name the data you want.  For example" runs for six
    # seconds, so the panel arrives and then names itself a line at a time rather than
    # landing whole and leaving the screen still while the sentence finishes.
    t.specAt = CUE["spec"]
    t.specFor = 1.0
    t.specRowAt = t.specAt + 0.6
    t.specRowPer = 1.7
    t.specRowFor = 0.9
    t.specRowsTo = t.specRowAt + 2 * t.specRowPer + t.specRowFor

    # "a single cell / a columnar table / or a set of keys and values", a row at a time
    t.threadAt = CUE["title"]
    t.threadPer = CUE["sales"] - CUE["title"]
    t.threadTo = CUE["hold"]                     # and a shared arrival
    t.liftFor = 0.5                              # a specification row lighting up
    t.litFor = 0.6                               # a range highlighting as its thread lands

    t.stripAt = CUE["strip"]                     # "The names and values"
    t.stripFor = 1.4
    t.showAt = t.stripAt - 0.02                  # the flyers take over just before
    t.outlineAt = t.stripAt + 0.3
    t.outlineFor = 0.9

    t.stackAt = CUE["stack"]                     # "are packaged into a JSON data set."
    t.stackFor = 2.6
    t.stackTo = t.stackAt + t.stackFor

    t.flowAt = CUE["flow"]                       # "Almost everything can consume JSON,"
    t.flowFor = 2.8
    t.flowTo = t.flowAt + t.flowFor
    t.outlineOutFor = 0.8
    t.synAt = t.flowAt + 0.9                     # the remainder of the JSON arrives
    t.synFor = 1.8
    t.cardAt = t.flowAt + 1.0                    # the card is solid before anything lands
    t.cardTo = t.cardAt + 1.2
    t.handFor = 0.7                              # flyers out as the real tokens come in

    t.boxAt = CUE["box"]                         # "turning your spreadsheets into..."
    t.boxFor = 1.6
    t.iconAt = t.boxAt + 0.5
    t.iconPer = 0.5
    t.iconFor = 1.3
    t.iconsTo = t.iconAt + (len(ICONS) - 1) * t.iconPer + t.iconFor

    if CUE["hours"] - CUE["sales"] != t.threadPer:
        raise ValueError("the three rows are not evenly spaced in the script; the animation "
                         "staggers them by one interval and would drift from the words")
    if t.specRowsTo > t.threadAt:
        raise ValueError(f"the specification is still arriving at {t.specRowsTo}s when the "
                         f"first row is due to light at {t.threadAt}s")

    beats = ["sheetAt", "namesAt", "specAt", "threadAt", "stripAt", "stackAt",
             "flowAt", "boxAt"]
    for earlier, later in zip(beats, beats[1:]):
        if getattr(t, later) <= getattr(t, earlier):
            raise ValueError(f"{later} ({getattr(t, later)}s) does not follow "
                             f"{earlier} ({getattr(t, earlier)}s)")
    if t.iconsTo > t.end:
        raise ValueError(f"the last icon settles at {t.iconsTo}s, past the end at {t.end}s")
    return t


def subtitles(t):
    """The voiceover, on screen, so the animation can be timed against the words.

    Each line runs to the start of the next, which is how the script is written; the last
    holds to the end.  They are the reason the video exists in this shape, so they are
    shown by default and switched off from the transport when it is time to film.
    """
    lines = [(at, line) for _, at, _, line in SCRIPT]
    ends = [start for start, _ in lines[1:]] + [t.end]
    return [[text, start, end] for (start, text), end in zip(lines, ends)]


def blocks(book, ranges, rows):
    """What each reference selects, as a grid of (address, value) -- the map everything else
    is built from.

    The three shapes come from the specification's own expression forms, so the animation
    cannot describe a shape the spec does not ask for: a bare reference is one cell, `[...]`
    with `records` is a heading row over records, and `{...}` is two columns of pairs.
    """
    ws = load_workbook(book, data_only=True).active
    out = {}
    for key, a1 in ranges.items():
        top_left, bottom_right = a1.split(":")     # single-letter columns; this sheet is A-I
        c1, r1 = ord(top_left[0]), int(top_left[1:])
        c2, r2 = ord(bottom_right[0]), int(bottom_right[1:])
        grid_ = [[(f"{chr(c)}{r}", ws.cell(row=r, column=c - ord("A") + 1).value)
                  for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
        expr = rows[key]
        kind = ("records" if "records" in expr else
                "object" if expr.strip().startswith("{") else "scalar")
        out[key] = {"kind": kind, "grid": grid_}
    return out


def _cell(value):
    return "" if value is None else str(value)


def sheet_html(path, max_row=9, max_col=9):
    """The workbook's cells, each value wrapped in a span of its own.

    The morph measures text, not the boxes text happens to sit in: a cell is 26px tall
    whatever the 14px inside it is doing, and scaling by the box would land the words at the
    wrong size.  The span gives all three waypoints -- sheet, stack, JSON -- the same kind of
    thing to measure.
    """
    ws = load_workbook(path, data_only=True).active
    letters = [chr(ord("A") + i) for i in range(max_col)]
    rows = ["<tr><th></th>" + "".join(f"<th>{c}</th>" for c in letters) + "</tr>"]
    for r in range(1, max_row + 1):
        cells = [f"<th>{r}</th>"]
        for c in range(1, max_col + 1):
            addr = f"{letters[c - 1]}{r}"
            value = _cell(ws.cell(row=r, column=c).value)
            span = f'<span class="cv" id="cv-{addr}">{value}</span>' if value else ""
            cells.append(f'<td id="c{addr}">{span}</td>')
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return f'<table class="sheet">{"".join(rows)}</table>' 


def json_html(parts):
    """The JSON, token by token, with every value carrying the cell it came from.

    A `tok` is a value that exists in the workbook and so can fly out of it; a `syn` is
    everything else -- braces, quotes, commas, indentation, and the record keys that repeat
    on the second and third rows.  Those are the "remainder of the JSON" the script has
    appearing last, which is exactly what they are.
    """
    def syn(text):
        return f'<span class="syn">{text}</span>'

    def tok(addr, text):
        return f'<span class="tok" id="j-{addr}">{text}</span>'

    def key(name):
        return f'<span class="tok jk" id="jk-{name}">{name}</span>'

    out = [syn("{")]
    keys = list(parts)
    for n, name in enumerate(keys):
        tail = "," if n < len(keys) - 1 else ""
        block = parts[name]
        rows = block["grid"]
        if block["kind"] == "scalar":
            addr, value = rows[0][0]
            out.append(syn('  "') + key(name) + syn('": "') + tok(addr, _cell(value))
                       + syn(f'"{tail}'))
        elif block["kind"] == "records":
            heads = rows[0]
            out.append(syn('  "') + key(name) + syn('": ['))
            for i, row in enumerate(rows[1:]):
                bits = []
                for (haddr, head), (addr, value) in zip(heads, row):
                    label = tok(haddr, head) if i == 0 else f'<span class="syn">{head}</span>'
                    quoted = isinstance(value, str)
                    bits.append(syn('"') + label + syn('": ' + ('"' if quoted else ''))
                                + tok(addr, _cell(value)) + syn('"' if quoted else ''))
                comma = "," if i < len(rows) - 2 else ""
                out.append(syn("    { ") + syn(", ").join(bits) + syn(" }" + comma))
            out.append(syn(f"  ]{tail}"))
        else:
            out.append(syn('  "') + key(name) + syn('": {'))
            for i, ((kaddr, kval), (vaddr, vval)) in enumerate(rows):
                comma = "," if i < len(rows) - 1 else ""
                out.append(syn('    "') + tok(kaddr, _cell(kval)) + syn('": "')
                           + tok(vaddr, _cell(vval)) + syn(f'"{comma}'))
            out.append(syn(f"  }}{tail}"))
    out.append(syn("}"))
    return "".join(f'<div class="jl">{line}</div>' for line in out)


def stack_html(parts):
    """The halfway house: the blocks stacked vertically, each label level with its own top.

    This is measured, not drawn -- it is laid out off-screen so the flyers can read real
    positions out of it, exactly as they read their finish out of the JSON.  Writing the
    intermediate arrangement as ordinary markup is much less trouble than computing it, and
    it stays right when the data changes shape.
    """
    out = []
    for name, block in parts.items():
        rows = "".join(
            "<tr>" + "".join(
                f'<td><span id="s-{addr}">{_cell(value)}</span></td>' for addr, value in row
            ) + "</tr>" for row in block["grid"])
        out.append(f'<div class="sblock" style="--c:{COLOURS[name]}">'
                   f'<span class="slab" id="sk-{name}">{name}</span>'
                   f"<table>{rows}</table></div>")
    return f'<div id="stackInner">{"".join(out)}</div>'


def build(show_timeline=False):
    book, spec = HERE / "quarter.xlsx", HERE / "quarter.toml"
    t = timeline()
    data = read_spec(str(book), str(spec))
    rows = tomllib.loads(spec.read_text())["extract"]
    wb = load_workbook(book)
    labels = {k: reference(v) for k, v in rows.items()}
    ranges = {k: resolve(labels[k]["text"], wb) for k in rows}
    parts = blocks(book, ranges, rows)

    payload = {
        "grid": sheet_html(book),
        "specRows": "".join(
            f'<div class="srow" id="s{k}" style="--c:{COLOURS[k]}">'
            f'<span class="k">{k}</span> = <span class="v">"{v}"</span></div>'
            for k, v in rows.items()),
        "json": json_html(parts),
        "stack": stack_html(parts),
        "cells": {k: [a for row in v["grid"] for a, _ in row] for k, v in parts.items()},
        "ranges": ranges,
        "labels": labels,
        "colours": COLOURS,
        "icons": [[n, dx, dy, d] for n, dx, dy, d in ICONS],
        "T": t.values(),
        "subtitles": subtitles(t),
    }
    html = TEMPLATE.replace("__PAYLOAD__", json.dumps(payload))
    (HERE / "narrated.html").write_text(html, encoding="utf-8")
    print(f"wrote {HERE / 'narrated.html'}")
    if show_timeline:
        print(t.listing())
    return data


TEMPLATE = r"""<!doctype html>
<meta charset="utf-8"><title>EDJAS narrated introducer</title>
<style>
  :root { --ink:#1c1a22; --rule:#d8d2e4; --paper:#e8e2f5; }
  * { box-sizing:border-box; }
  html,body { height:100%; overflow:hidden; }
  body { margin:0; background:#3a3348; font-family:system-ui,-apple-system,"Segoe UI",sans-serif;
         display:flex; flex-direction:column; align-items:center; justify-content:center; gap:10px; }
  #sizer { position:relative; flex:0 0 auto; }
  #fit { position:absolute; left:0; top:0; width:1280px; height:720px; transform-origin:0 0; }
  #stage { position:relative; width:1280px; height:720px; background:#fff; overflow:hidden;
           color:var(--ink); box-shadow:0 12px 60px rgba(0,0,0,.5); }
  .layer { position:absolute; inset:0; }

  table.sheet { border-collapse:collapse; font-size:14px; position:absolute; left:500px; top:150px; }
  .sheet th { background:#eee9f5; color:#655c78; font-weight:600; font-size:12px;
              border:1px solid var(--rule); width:34px; height:26px; }
  .sheet td { border:1px solid var(--rule); padding:3px 7px; min-width:58px; height:26px;
              white-space:nowrap; }

  #labels .rbox { position:absolute; border-radius:3px; outline-offset:1px; }
  #labels .rbox.name { outline:2px dashed var(--c); }
  #labels .rbox.address { outline:2px solid var(--c); }
  #labels .chip { position:absolute; height:19px; display:flex; align-items:center;
                  padding:0 7px; border-radius:5px; font-size:12px; font-weight:700;
                  white-space:nowrap; }
  #labels .chip.name { background:var(--c); color:#fff; }
  #labels .chip.address { background:#fff; color:var(--c); border:1.5px solid var(--c);
                          font-family:ui-monospace,"SF Mono",Menlo,monospace; }

  /* The panel's vertical position is set from the measured layout, not here: each row has to
     sit below the bottom of every range to its left, or its thread reaches the far side of
     the sheet only by cutting straight through one of the others. */
  #spec { position:absolute; left:32px; top:236px; width:400px; background:#faf8fd;
          border:1px solid var(--rule); border-radius:10px; padding:16px 18px;
          font-family:ui-monospace,"SF Mono",Menlo,monospace; font-size:15px; }
  #spec .hdr { color:#8a7fa6; margin-bottom:8px; }
  .srow { padding:8px; border-radius:6px; white-space:nowrap; }
  .srow .k { color:var(--c); font-weight:700; }
  .srow .v { color:#3d3550; }

  #threads { position:absolute; inset:0; width:1280px; height:720px; pointer-events:none; }
  #threads path { fill:none; stroke-width:3; stroke-linecap:round;
                  stroke-dasharray:var(--len); stroke-dashoffset:var(--len); }

  /* the staging layout: measured, never seen */
  /* Laid out where the stack is meant to appear and simply not painted, so the middle of
     every flyer's journey is a real position on the stage rather than a computed guess. */
  #stack { position:absolute; left:50%; top:50%; transform:translate(-50%,-50%);
           opacity:0; pointer-events:none; font-size:20px; }
  #stack .sblock { display:flex; align-items:flex-start; gap:22px; margin-bottom:20px; }
  #stack .slab { font-family:ui-monospace,Menlo,monospace; font-weight:700; color:var(--c);
                 min-width:78px; line-height:30px; }
  #stack table { border-collapse:collapse; }
  #stack td { padding:0 16px 0 0; height:30px; white-space:nowrap; vertical-align:middle; }

  #json { position:absolute; left:50%; top:50%; transform:translate(-50%,-50%);
          font-family:ui-monospace,Menlo,monospace; font-size:17px; line-height:1.55;
          background:#1e1a2b; color:var(--paper); padding:20px 28px; border-radius:12px;
          white-space:pre; opacity:0; box-shadow:0 20px 60px rgba(0,0,0,.4); }
  #json .syn, #json .tok { opacity:0; }
  #json .jk { font-weight:700; }

  #box { position:absolute; left:50%; top:50%; transform:translate(-50%,-50%);
         background:#1e1a2b; color:#fff; border-radius:12px; padding:16px 30px; opacity:0;
         font-family:ui-monospace,Menlo,monospace; font-size:26px; font-weight:700;
         box-shadow:0 16px 44px rgba(0,0,0,.34); }

  #morph { pointer-events:none; opacity:0; }
  .fly { position:absolute; transform-origin:0 0; white-space:nowrap; display:flex;
         align-items:center; }
  .fly .ring { position:absolute; inset:-4px -7px; border-radius:6px; opacity:0;
               outline:2px solid var(--c); }
  .fly.key { font-family:ui-monospace,"SF Mono",Menlo,monospace; font-weight:700;
             color:var(--c); }

  #icons .dest { position:absolute; left:50%; top:50%; display:flex; flex-direction:column;
                 align-items:center; gap:7px; opacity:0; color:var(--ink); }
  #icons svg { width:46px; height:46px; fill:none; stroke:currentColor; stroke-width:1.7;
               stroke-linecap:round; stroke-linejoin:round; }
  #icons .name { font-size:16px; font-weight:600; }

  #sub { position:absolute; left:60px; right:60px; bottom:38px; text-align:center;
         font-size:27px; font-weight:600; color:var(--ink); opacity:0; line-height:1.3; }
  #scrub { display:flex; gap:12px; align-items:center; color:#fff; flex:0 0 auto;
           font-family:system-ui; font-size:13px; }
  #scrub input[type=range] { flex:1; }
  #scrub label { display:flex; gap:5px; align-items:center; cursor:pointer; }
</style>

<div id="sizer"><div id="fit"><div id="stage">
  <div class="layer" id="sheetLayer"></div>
  <div class="layer" id="labels"></div>
  <div class="layer"><div id="spec"><div class="hdr">[extract]</div><div id="specRows"></div></div></div>
  <svg class="layer" id="threads"></svg>
  <div class="layer"><div id="json"></div></div>
  <div class="layer"><div id="box">JSON</div></div>
  <div class="layer" id="morph"></div>
  <div class="layer" id="icons"></div>
  <div id="stack"></div>
  <div id="sub"></div>
</div></div></div>
<div id="scrub">
  <button id="play">play</button>
  <input id="t" type="range" min="0" max="32" step="0.02" value="0">
  <span id="tv">0.00s</span>
  <label><input id="subs" type="checkbox" checked> subtitles</label>
</div>

<script>
const D = __PAYLOAD__;
const T = D.T, SUBS = D.subtitles;
const stage = document.getElementById('stage');
const KEYS = Object.keys(D.ranges);
document.getElementById('sheetLayer').innerHTML = D.grid;
document.getElementById('specRows').innerHTML = D.specRows;
document.getElementById('json').innerHTML = D.json;
document.getElementById('stack').innerHTML = D.stack;
document.getElementById('icons').innerHTML = D.icons.map(([n,,,d])=>
  `<div class="dest"><svg viewBox="0 0 24 24">${d}</svg><span class="name">${n}</span></div>`
).join('');
window.T = T;

const clamp=(x,a=0,b=1)=>Math.max(a,Math.min(b,x));
const ease =x=>x<.5?4*x*x*x:1-Math.pow(-2*x+2,3)/2;
const ph=(t,a,b)=>ease(clamp((t-a)/(b-a)));
const win=(t,a,b,f=.35)=>Math.min(ph(t,a,a+f), 1-ph(t,b-f,b));
const rgb=h=>[1,3,5].map(i=>parseInt(h.slice(i,i+2),16));
const mix=(a,b,k)=>{const A=rgb(a),B=rgb(b);
  return '#'+A.map((v,i)=>Math.round(v+(B[i]-v)*k).toString(16).padStart(2,'0')).join('');};

let SCALE=1;
function fit(force){
  const pad=20, bar=46;
  SCALE = force || Math.min((innerWidth-pad*2)/1280, (innerHeight-pad*2-bar)/720);
  document.getElementById('fit').style.transform=`scale(${SCALE})`;
  const sz=document.getElementById('sizer');
  sz.style.width=(1280*SCALE)+'px'; sz.style.height=(720*SCALE)+'px';
  document.getElementById('scrub').style.width=(1280*SCALE)+'px';
}
window.fit=fit; addEventListener('resize',()=>fit());
const box=el=>{const s=stage.getBoundingClientRect(), r=el.getBoundingClientRect();
  return {x:(r.left-s.left)/SCALE, y:(r.top-s.top)/SCALE, w:r.width/SCALE, h:r.height/SCALE,
          cx:(r.left-s.left+r.width/2)/SCALE, cy:(r.top-s.top+r.height/2)/SCALE};};

/* Drop the panel until every row clears the ranges to its left.  A row level with the sales
   block can only reach the hours columns by crossing it, which reads as a mistake however
   neatly it is drawn -- so the layout is corrected from the measurements rather than tuned
   by hand in the stylesheet. */
function placeSpec(){
  const spec=document.getElementById('spec');
  let shift=0;
  KEYS.forEach((k,i)=>{
    const row=box(document.getElementById('s'+k));
    KEYS.slice(0,i).forEach(j=>{
      const b=box(document.getElementById('c'+D.ranges[j].split(':')[1]));
      shift=Math.max(shift, b.y+b.h+14-row.cy);
    });
  });
  if(shift>0) spec.style.top=(parseFloat(getComputedStyle(spec).top)+shift)+'px';
}

let TAGS=[];
function buildLabels(){
  const L=document.getElementById('labels'); L.innerHTML=''; TAGS=[];
  for(const key of KEYS){
    const [tl,br]=D.ranges[key].split(':');
    const a=box(document.getElementById('c'+tl)), b=box(document.getElementById('c'+br));
    const w=b.x+b.w-a.x, h=b.y+b.h-a.y, k=D.labels[key].kind;
    const g=document.createElement('div');
    g.style.cssText=`--c:${D.colours[key]};position:absolute;inset:0`;
    g.innerHTML=
      `<div class="rbox ${k}" style="left:${a.x}px;top:${a.y}px;width:${w}px;height:${h}px"></div>`+
      `<div class="chip ${k}" style="left:${a.x}px;top:${a.y-22}px">${D.labels[key].text}</div>`;
    L.appendChild(g); TAGS.push(g);
  }
}

const svg=document.getElementById('threads');
function buildThreads(){
  svg.innerHTML='';
  KEYS.forEach(key=>{
    const row=box(document.getElementById('s'+key));
    const [tl,br]=D.ranges[key].split(':');
    const a=box(document.getElementById('c'+tl)), b=box(document.getElementById('c'+br));
    /* Aim below anything already passed.  Entering at the middle of the range is the
       natural choice, but the curve sags on its way across, and a thread that clips the
       corner of the block above reads as a mistake however neatly it is drawn.  Dropping
       the aim point keeps it under the obstacle and still lands squarely on the range. */
    const floor=Math.max(...KEYS.map(o=>{
      if(o===key) return 0;
      const q=box(document.getElementById('c'+D.ranges[o].split(':')[1]));
      const p0=box(document.getElementById('c'+D.ranges[o].split(':')[0]));
      return (p0.x < a.x) ? q.y+q.h+18 : 0;            /* only what lies to the left */
    }), 0);
    const x1=row.x+row.w, y1=row.cy, x2=a.x;
    const y2=Math.min(Math.max((a.y+b.y+b.h)/2, floor), b.y+b.h-10);
    const p=document.createElementNS('http://www.w3.org/2000/svg','path');
    p.setAttribute('d',`M${x1},${y1} C${x1+80},${y1} ${x2-80},${y2} ${x2},${y2}`);
    p.setAttribute('stroke',D.colours[key]); p.id='p'+key;
    svg.appendChild(p);
    p.style.setProperty('--len', p.getTotalLength());
  });
}

/* Three measured positions per flyer: where it is on the sheet, where it stands in the
   stacked arrangement, and where its own text sits in the finished JSON.  None is computed
   -- the staging layout is real markup off-screen, so the middle of the journey is as
   truthful as either end. */
let FLY=[];
function buildFlyers(){
  const morph=document.getElementById('morph');
  morph.innerHTML=''; FLY=[];
  const j=document.getElementById('json');
  const wasO=j.style.opacity; j.style.opacity=1;
  const toks=[...j.querySelectorAll('.tok')];
  toks.forEach(el=>el.style.opacity=1);

  const add=(src,mid,dst,text,cls,colour)=>{
    if(!src||!mid||!dst) return;
    const a=box(src), b=box(mid), c=box(dst);
    const em=el=>parseFloat(getComputedStyle(el).fontSize);
    const f=document.createElement('div');
    f.className='fly '+cls;
    f.style.cssText=`--c:${colour};left:${a.x}px;top:${a.y}px;font-size:${em(src)}px`;
    f.innerHTML=`<span class="ring"></span><span class="txt">${text}</span>`;
    morph.appendChild(f);
    FLY.push({el:f, ring:f.querySelector('.ring'), txt:f.querySelector('.txt'),
              x0:a.x, y0:a.y, x1:b.x, y1:b.y, x2:c.x, y2:c.y,
              s1:em(mid)/em(src), s2:em(dst)/em(src),
              from:cls==='key'?colour:'#1c1a22',
              to:cls==='key'?mix(colour,'#ffffff',0.3):'#e8e2f5'});
  };

  KEYS.forEach(key=>{
    add(document.querySelector('#s'+key+' .k'), document.getElementById('sk-'+key),
        document.getElementById('jk-'+key), key, 'key', D.colours[key]);
    D.cells[key].forEach(addr=>{
      const cv=document.getElementById('cv-'+addr);      /* the value, not the cell holding
                                                            it: a td is padded, and starting
                                                            from its corner leaves the text
                                                            visibly doubled at the handover */
      if(cv) add(cv, document.getElementById('s-'+addr),
                 document.getElementById('j-'+addr), cv.textContent, 'val', D.colours[key]);
    });
  });
  toks.forEach(el=>el.style.opacity='');
  j.style.opacity=wasO;
}

function seek(t){
  document.getElementById('t').value=t;
  document.getElementById('tv').textContent=t.toFixed(2)+'s';

  /* the workbook, then the names that belong to it, then the specification */
  const strip=ph(t,T.stripAt,T.stripAt+T.stripFor);
  const sheetIn=ph(t,T.sheetAt,T.sheetAt+T.sheetFor)*(1-strip);
  document.getElementById('sheetLayer').style.opacity=sheetIn;
  document.getElementById('labels').style.opacity=
    ph(t,T.namesAt,T.namesAt+T.namesFor)*(1-strip);

  const specIn=ph(t,T.specAt,T.specAt+T.specFor);
  const spec=document.getElementById('spec');
  spec.style.opacity=specIn*(1-strip);
  spec.style.transform=`translateX(${(specIn-1)*40}px)`;

  /* Each line of the specification arrives on its own, which gives the six seconds of
     "describe and name the data you want" something to do; then a row lights and its
     thread snakes across to the range it names. */
  KEYS.forEach((k,i)=>{
    const arrive=T.specRowAt+i*T.specRowPer;
    const written=ph(t,arrive,arrive+T.specRowFor);
    const at=T.threadAt+i*T.threadPer;
    const up=ph(t,at,at+T.liftFor)*(1-strip);
    const row=document.getElementById('s'+k);
    row.style.opacity=written;
    row.style.transform=`translateX(${(written-1)*16}px)`;
    row.style.background = up ? `rgba(255,255,255,${up})` : '';
    row.style.boxShadow  = up ? `0 ${8*up}px ${26*up}px rgba(0,0,0,${0.22*up})` : '';
    row.style.outline    = up ? `${2*up}px solid ${D.colours[k]}` : '';

    const p=document.getElementById('p'+k); if(!p) return;
    const drawn=ph(t, at, T.threadTo);
    p.style.strokeDashoffset=p.getTotalLength()*(1-drawn);
    p.style.opacity=1-ph(t,T.stripAt,T.stripAt+T.stripFor*0.7);

    const lit=ph(t,T.threadTo,T.threadTo+T.litFor)*(1-strip);
    D.cells[k].forEach(addr=>{
      const cell=document.getElementById('c'+addr); if(!cell) return;
      cell.style.boxShadow = lit ? `inset 0 0 0 ${3*lit}px ${D.colours[k]}` : '';
    });
  });

  /* the names and values leave the sheet, stack up, and become the JSON */
  const outline=ph(t,T.outlineAt,T.outlineAt+T.outlineFor)
                *(1-ph(t,T.flowAt,T.flowAt+T.outlineOutFor));
  const stack=ph(t,T.stackAt,T.stackTo);
  const flow=ph(t,T.flowAt,T.flowTo);
  const shrink=ph(t,T.boxAt,T.boxAt+T.boxFor);
  const hand=ph(t,T.flowTo-T.handFor,T.flowTo);      /* flyers out, real tokens in */
  document.getElementById('morph').style.opacity = t>=T.showAt ? 1-hand : 0;
  FLY.forEach(f=>{
    const x=f.x0+(f.x1-f.x0)*stack+(f.x2-f.x1)*flow;
    const y=f.y0+(f.y1-f.y0)*stack+(f.y2-f.y1)*flow;
    const s=1+(f.s1-1)*stack+(f.s2-f.s1)*flow;
    f.el.style.transform=`translate(${x-f.x0}px,${y-f.y0}px) scale(${s})`;
    f.ring.style.opacity=outline;
    f.txt.style.color=mix(f.from,f.to,flow);
  });

  const j=document.getElementById('json');
  j.style.opacity=ph(t,T.cardAt,T.cardTo)*(1-ph(t,T.boxAt+T.boxFor*0.35,T.boxAt+T.boxFor));
  j.style.transform=`translate(-50%,-50%) scale(${1-0.55*shrink})`;
  const syn=ph(t,T.synAt,T.synAt+T.synFor);
  [...j.querySelectorAll('.syn')].forEach(el=>el.style.opacity=syn);
  [...j.querySelectorAll('.tok')].forEach(el=>el.style.opacity=hand);

  /* ...which becomes a box, and feeds everything else */
  document.getElementById('box').style.opacity=ph(t,T.boxAt+T.boxFor*0.5,T.boxAt+T.boxFor);
  [...document.querySelectorAll('#icons .dest')].forEach((d,i)=>{
    const at=T.iconAt+i*T.iconPer, k=ph(t,at,at+T.iconFor);
    const [,dx,dy]=D.icons[i];
    d.style.transform=`translate(-50%,-50%) translate(${dx*k}px,${dy*k}px) scale(${0.25+0.75*k})`;
    d.style.opacity=Math.min(1,k*2.2);
  });

  const line=SUBS.find(s=>t<s[2]) || SUBS[SUBS.length-1];
  const sub=document.getElementById('sub');
  sub.textContent=line[0];
  sub.style.opacity=document.getElementById('subs').checked ? win(t,line[1],line[2],.3) : 0;
}
window.seek=seek;

let playing=false, t0=0, base=0;
function frame(ts){ if(!playing) return;
  const t=base+(ts-t0)/1000; if(t>=T.end){ seek(T.end); playing=false; return; }
  seek(t); requestAnimationFrame(frame); }
document.getElementById('play').onclick=()=>{
  playing=!playing; base=+document.getElementById('t').value;
  if(playing) requestAnimationFrame(ts=>{t0=ts; frame(ts);}); };
document.getElementById('t').oninput=e=>{ playing=false; seek(+e.target.value); };
document.getElementById('subs').onchange=()=>seek(+document.getElementById('t').value);
document.getElementById('t').max=T.end;

addEventListener('load',()=>{
  fit(); placeSpec(); buildLabels(); buildThreads(); buildFlyers(); seek(0);
});
</script>
"""

if __name__ == "__main__":
    build(show_timeline="--timeline" in sys.argv)
