"""Spike: emit Markdown for Zensical from the same sources the HTML build uses.

The point of the experiment is that the interesting half of ``build_site.py`` is not
about HTML at all. Composing canonical files, extracting ``site.xlsx`` with EDJAS,
reading the function registry and the command-line help from the live objects -- all of
that produces *content*, and content does not care what renders it. This module reuses
those functions unchanged and writes Markdown instead of pages.

Zensical is Material for MkDocs' successor, by the same authors: a Rust core with a thin
Python front end, configured by ``zensical.toml`` rather than ``mkdocs.yml``. Everything
below the configuration writer is the same code the MkDocs spike used, which is rather
the point -- the renderer is the replaceable part.

    uv run python docs/build_md.py                        # -> docs/md/, docs/zensical.toml
    uv run zensical build -f docs/zensical.toml --strict  # -> docs/_zensical/
    uv run zensical serve -f docs/zensical.toml           # preview on localhost:8000
"""

import re
import shutil
from pathlib import Path

import build_site as bs   # the existing build: content functions reused verbatim

DOCS = Path(__file__).resolve().parent
OUT = DOCS / "md"
ROOT = DOCS.parent

# The same repository URLs the HTML build rewrites, now pointing at Markdown pages.
LINKS = {
    f"{bs.REPO}/blob/main/edjas-examples/README.md": "why.md",
    f"{bs.REPO}/tree/main/edjas-examples": "gallery/index.md",
    f"{bs.REPO}/blob/main/edjas-examples/src/edjas_examples/ATTRIBUTION.md": "colophon.md",
    f"{bs.RAW}/images/example_workbook.png": "images/example_workbook.png",
}
for stem, slug, _ in bs.DIAGRAMS:
    LINKS[f"{bs.REPO}/blob/main/images/{stem}.svg"] = f"architecture.md#{slug}"


def relink(text, depth):
    """Point repository links at site pages, and 'site:' links at the right depth."""
    up = "../" * depth
    for source, target in LINKS.items():
        text = text.replace(source, up + target)
    text = re.sub(r"\bsite:([\w/.-]+)\.html", lambda m: f"{up}{m.group(1)}.md", text)
    return text


def demote(text, levels):
    """Shift ATX headings down, leaving fenced code alone."""
    if not levels:
        return text
    out, fenced = [], False
    for line in text.splitlines():
        if line.lstrip().startswith("```"):
            fenced = not fenced
        elif not fenced and line.startswith("#"):
            line = "#" * levels + line
        out.append(line)
    return "\n".join(out)


def part(spec, depth):
    """One manifest part, as Markdown."""
    path = (DOCS / spec["file"]).resolve()
    text = path.read_text(encoding="utf-8")
    if spec.get("as") == "code":
        lang = spec.get("lang", path.suffix.lstrip("."))
        return f"```{lang}\n{text.strip()}\n```\n"
    if spec.get("skip_first_heading"):
        text = bs._strip_first_heading(text)
    return relink(demote(text, spec.get("demote", 0)), depth) + "\n"


def write(name, text):
    target = OUT / name
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(text.rstrip() + "\n", encoding="utf-8")
    return name


def prose_pages():
    """The pages whose content is composed markdown, from pages.toml."""
    written = []
    for page in bs.load_manifest():
        if "example" in page or not page.get("parts"):
            continue
        name = page["url"].replace(".html", ".md")
        depth = name.count("/")
        body = "".join(part(p, depth) for p in page.get("parts", []))
        tail = "".join(part(p, depth) for p in page.get("tail_parts", []))
        # Only a level-one heading counts. content/index.md opens at `##` because the
        # bespoke build's home.html supplies the title; testing for a bare "#" left that
        # page with no H1 at all, and Zensical titled it from the filename ("Index").
        heading = "" if body.lstrip().startswith("# ") else f"# {page['title']}\n\n"
        written.append(write(name, heading + body + tail))
    return written


def reference_pages():
    """Function, API and CLI reference, generated from the live objects as before."""
    rows = "\n".join(
        f"| `{e['name']}` | `{e['input']}` | {e['result'] or e['doc']} |"
        for e in bs.function_reference()
    )
    lead = part({"file": "content/reference/functions-lead.md"}, 1)
    write("reference/functions.md",
          f"# Built-in functions\n\n{lead}\n"
          f"| Function | Typical input | Result |\n|---|---|---|\n{rows}\n")

    api = "\n\n".join(
        f"## `{e['name']}`\n\n```python\n{e['signature']}\n```\n\n{e['doc']}"
        for e in bs.api_reference()
    )
    write("reference/api.md",
          f"# Python API\n\n{part({'file': 'content/reference/api-lead.md'}, 1)}\n{api}\n")

    cli = "\n\n".join(
        f"## `{e['command']}`\n\n```\n{e['help']}\n```" for e in bs.cli_reference()
    )
    write("reference/cli.md",
          f"# Command line\n\n{part({'file': 'content/reference/cli-lead.md'}, 1)}\n{cli}\n")

    tables = bs.spec_tables()
    forms = "\n".join(f"| `{r['form']}` | {r['produces']} | `{r['example']}` |"
                      for r in tables["forms"])
    kinds = "\n".join(f"| {r['kind']} | `{r['written as']}` | {r['notes']} |"
                      for r in tables["ref_kinds"])
    write("reference/spec-language.md",
          "# The specification language\n\n"
          + part({"file": "content/reference/spec-language.md"}, 1)
          + "\n## The three forms\n\n| Written as | Produces | For example |\n|---|---|---|\n"
          + forms
          + "\n\n## The three kinds of reference\n\n| Kind | Written as | Notes |\n|---|---|---|\n"
          + kinds + "\n\n"
          + part({"file": "content/reference/spec-language-tail.md"}, 1))
    return ["reference/functions.md", "reference/api.md", "reference/cli.md",
            "reference/spec-language.md"]


def gallery_pages(reports):
    """One page per example, from the spec's own comment header, as before."""
    import edjas_examples
    cards = ["| Example | Shape | What it demonstrates | Publisher |", "|---|---|---|---|"]
    for row in bs.site_data()["gallery"]:
        cards.append(f"| [{row['name']}]({row['name']}.md) | {row['shape']} "
                     f"| {row['teaser']} | {row['publisher']} |")
    write("gallery/index.md",
          "# Gallery\n\n" + part({"file": "content/gallery/index.md"}, 1)
          + "\n" + "\n".join(cards) + "\n")

    written = ["gallery/index.md"]
    for name, cfg in edjas_examples.EXAMPLES.items():
        spec = edjas_examples.DATA / cfg["spec"]
        body = (
            f"# {name}\n\n{relink(bs.comment_header(spec), 1)}\n\n"
            f"## The specification\n\nThe whole of `{cfg['spec']}`, applied to the "
            f"unmodified `{cfg['xlsx']}`.\n\n```toml\n{bs.spec_body(spec)}\n```\n\n"
            f"## The report it produces\n\n"
            f'<iframe src="../{reports[name]}" title="EDJAS report: {name}" '
            f'loading="lazy" style="width:100%;height:34rem;border:1px solid #ccc"></iframe>\n\n'
            f"[Open it full width](../{reports[name]})\n\n"
            + demote(relink(bs.attribution_section(cfg["xlsx"]), 1), 0)
        )
        written.append(write(f"gallery/{name}.md", body))
    return written


def architecture_page():
    """The diagrams as images -- MkDocs serves them, so no id or CSS namespacing."""
    body = ["# How EDJAS works\n", part({"file": "content/architecture.md"}, 0)]
    for stem, slug, title in bs.DIAGRAMS:
        desc = bs.svg_description(bs.IMAGES / f"{stem}.svg")
        body.append(f"\n## {title} {{ #{slug} }}\n\n"
                    f"![{title}](images/{stem}.svg)\n\n*{desc}*\n")
    return [write("architecture.md", "\n".join(body))]


def assets(reports):
    images = OUT / "images"
    images.mkdir(parents=True, exist_ok=True)
    for stem, _, _ in bs.DIAGRAMS:
        shutil.copy(bs.IMAGES / f"{stem}.svg", images / f"{stem}.svg")
    png = bs.IMAGES / "example_workbook.png"
    if png.is_file():
        shutil.copy(png, images / png.name)


def nav_tree():
    """The navigation, derived from pages.toml rather than restated here.

    ``pages.toml`` already says which pages are nav entries (``nav``) and which hang
    below one (``nav_group``), and it generates the seven gallery entries from
    ``edjas_examples.EXAMPLES`` so a new example cannot be forgotten. Reading it back
    keeps that guarantee; the MkDocs spike listed the pages by hand and lost it.
    """
    pages = bs.load_manifest()
    children = {}
    for page in pages:
        if group := page.get("nav_group"):
            children.setdefault(group, []).append(page["url"])

    lines = []
    for page in pages:
        if not page.get("nav"):
            continue
        url, kids = page["url"], children.get(page["url"], [])
        label = toml_string(page["nav"])
        if not kids:
            lines.append(f"  {{ {label} = {toml_string(as_md(url))} }},")
            continue
        # navigation.indexes makes the section heading itself the first page, so the
        # section's own file leads the list rather than sitting beside it.
        lines.append(f"  {{ {label} = [")
        for target in [url, *kids]:
            lines.append(f"    {toml_string(as_md(target))},")
        lines.append("  ] },")
    return "\n".join(lines)


def as_md(url):
    return url.replace(".html", ".md")


def toml_string(value):
    return '"' + value.replace("\\", "\\\\").replace('"', '\\"') + '"'


def zensical_config():
    """Write ``docs/zensical.toml``.

    Two settings carry more weight than they look. ``use_directory_urls = false`` keeps
    ``gallery/retail.md`` at ``gallery/retail.html`` instead of ``gallery/retail/``, so
    the depth-relative links this module writes -- and, more to the point, the raw
    ``<iframe src="../reports/retail.html">`` that Markdown passes through untouched --
    resolve from a ``file://`` preview, from a project page under ``/edjas/``, and from a
    custom domain, exactly as the bespoke build's output does. And ``markdown_extensions``
    is given explicitly because supplying the table at all *replaces* Zensical's generous
    default set: the pages are composed from README and docstring text that was never
    written with SmartSymbols or MagicLink in mind, and silently rewriting a canonical
    file's punctuation is precisely what this site exists not to do.
    """
    (DOCS / "zensical.toml").write_text(f"""\
# Generated by docs/build_md.py -- edit that, not this.
[project]
site_name = "EDJAS"
site_description = "Extract data in JSON from any spreadsheet"
site_url = "https://holdenweb.github.io/edjas/"
site_author = "Steve Holden"
repo_url = "{bs.REPO}"
repo_name = "holdenweb/edjas"

# Both are relative to this file, and Zensical insists they sit under it.
docs_dir = "md"
site_dir = "_zensical"

# See the docstring above: relative links and the report iframes depend on this.
use_directory_urls = false

nav = [
{nav_tree()}
]

[project.validation]
# The bespoke build fails on an unmapped repository link; this is the equivalent, and
# `zensical build --strict` turns both of these from warnings into a non-zero exit.
invalid_links = true
invalid_link_anchors = true

[project.theme]
language = "en"
features = [
  "content.code.copy",
  "navigation.indexes",
  "navigation.sections",
  "navigation.top",
  "navigation.tracking",
  "search.highlight",
  "toc.follow",
]

[[project.theme.palette]]
media = "(prefers-color-scheme: light)"
scheme = "default"
primary = "deep purple"
toggle.icon = "lucide/sun"
toggle.name = "Switch to dark mode"

[[project.theme.palette]]
media = "(prefers-color-scheme: dark)"
scheme = "slate"
primary = "deep purple"
toggle.icon = "lucide/moon"
toggle.name = "Switch to light mode"

[project.markdown_extensions]
# toc and tables are always on and need no entry. attr_list carries the `{{ #slug }}`
# anchors the architecture page's diagram headings need; md_in_html lets the report
# iframes sit inside Markdown. The rest of Zensical's defaults are deliberately absent.
attr_list = {{}}
admonition = {{}}
def_list = {{}}
md_in_html = {{}}
sane_lists = {{}}
toc.permalink = true
pymdownx.highlight.anchor_linenums = true
pymdownx.superfences = {{}}
""", encoding="utf-8")


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    reports = bs.build_reports(OUT)          # the seven reports, as before
    assets(reports)
    pages = prose_pages() + reference_pages() + gallery_pages(reports) + architecture_page()
    zensical_config()
    # The later passes rewrite pages prose_pages() has already emitted -- the reference
    # pages need their generated tables, the gallery index its cards -- so the list holds
    # each of those twice, and the count has to be of distinct files.
    print(f"wrote {len(set(pages))} markdown pages and {len(reports)} reports into {OUT}")


if __name__ == "__main__":
    main()
