"""Tests for the documentation-site generator (``docs/build_site.py``).

The generator needs Jinja2, Markdown and the ``edjas-examples`` companion package, which
arrive together through the ``docs`` dependency group; the module skips when any is
absent, so a core-only environment still collects and runs the rest of the suite. The
site is built once for the whole session into a temporary directory, and every test
inspects that output.

``build_site`` is importable because ``pythonpath = ["docs"]`` is set in pyproject.toml.
"""

import posixpath
import re

import pytest

pytest.importorskip("jinja2")
pytest.importorskip("markdown")
edjas_examples = pytest.importorskip("edjas_examples")
build_site = pytest.importorskip("build_site")

from edjas.functions import DEFAULT_FUNCTIONS  # noqa: E402


@pytest.fixture(scope="session")
def site(tmp_path_factory):
    """Build the whole site once, strictly, and hand back its root."""
    out = tmp_path_factory.mktemp("site")
    build_site.build(out, strict=True)
    return out


def _pages(site):
    return sorted(site.rglob("*.html"))


# --- structure ---------------------------------------------------------------

def test_every_declared_page_is_written(site):
    """Every page in the manifest, including the generated gallery, reaches the output."""
    for page in build_site.load_manifest():
        written = site / page["url"]
        assert written.is_file(), f"{page['url']} was not written"
        assert written.stat().st_size > 500, f"{page['url']} is suspiciously small"


def test_no_template_syntax_leaks(site):
    """No page contains unrendered Jinja2 markup."""
    for page in _pages(site):
        text = page.read_text(encoding="utf-8")
        assert "{{" not in text and "{%" not in text, page.name


def test_every_internal_link_resolves(site):
    """Every relative href/src points at a file that exists, and every #fragment at an id.

    This one test catches broken navigation, a wrong ``url_for`` depth, root-relative
    links (which would 404 under a project Pages subpath), missing assets and stale
    anchors.
    """
    problems = []
    for page in _pages(site):
        rel = page.relative_to(site).as_posix()
        html = page.read_text(encoding="utf-8")
        for value in re.findall(r'(?:href|src)="([^"]+)"', html):
            if value.startswith(("http://", "https://", "mailto:", "data:", "#")):
                continue
            assert not value.startswith("/"), f"{rel}: root-relative link {value}"
            path, _, fragment = value.partition("#")
            target = site / posixpath.normpath(
                posixpath.join(posixpath.dirname(rel), path)
            )
            if not target.exists():
                problems.append(f"{rel}: missing {value}")
            elif fragment and f'id="{fragment}"' not in target.read_text(encoding="utf-8"):
                problems.append(f"{rel}: no anchor #{fragment} in {path}")
    assert not problems, problems


def test_no_unrewritten_repository_links(site):
    """Links into the repository tree are rewritten to their place on the site.

    The bare repository home is a genuine outbound link and stays; a ``/blob/`` or
    ``/tree/`` URL, or a raw asset URL, means a page escaped the rewriter.
    """
    for page in _pages(site):
        html = page.read_text(encoding="utf-8")
        for pattern in ("github.com/holdenweb/edjas/blob/",
                        "github.com/holdenweb/edjas/tree/",
                        "raw.githubusercontent.com"):
            assert pattern not in html, f"{page.name} still links to {pattern}"


def test_nav_marks_exactly_one_current_page(site):
    for page in _pages(site):
        if (site / "reports") in page.parents:
            continue  # the embedded reports carry no site navigation
        html = page.read_text(encoding="utf-8")
        assert html.count('aria-current="page"') == 1, page.name


# --- the inlined diagrams -----------------------------------------------------

def test_architecture_inlines_four_diagrams_without_collisions(site):
    """The four SVGs are inlined, and neither their ids nor their CSS can collide.

    Inline SVG is document-scoped: all four diagrams define ``<marker id="arrow">`` and
    share class names (``.title``, ``.desc``, ``.hdr``) with *differing* values, and one
    styles ``.line``, which would otherwise leak onto the ledger rows that base.html
    styles. So ids are namespaced and every SVG selector is scoped to its own root.
    """
    html = (site / "architecture.html").read_text(encoding="utf-8")

    assert len(re.findall(r"<svg\b", html)) == 4
    assert '<img src="' not in html or ".svg" not in html  # inlined, not linked

    ids = re.findall(r'\bid="([^"]+)"', html)
    duplicates = {name for name in ids if ids.count(name) > 1}
    assert not duplicates, f"duplicate ids: {duplicates}"

    # Skip the first <style>: that is base.html's own stylesheet, which is page-wide by
    # design. Every selector in the four inlined SVG stylesheets must be scoped.
    for style in re.findall(r"<style>(.*?)</style>", html, re.DOTALL)[1:]:
        for selector in re.findall(r"(?:^|\})\s*([^{}@]+?)\s*\{", style):
            for part in selector.split(","):
                assert part.strip().startswith("#"), f"unscoped SVG selector: {part!r}"

    for slug in ("c4-system-context", "c4-container-diagram",
                 "c4-component-diagram", "internal-structure"):
        assert f'id="{slug}"' in html, f"missing anchor target {slug}"


# --- the gallery and its reports ----------------------------------------------

def test_gallery_covers_every_example(site):
    """A page exists for every registered example -- an eighth cannot be forgotten."""
    built = {page.stem for page in (site / "gallery").glob("*.html")} - {"index"}
    assert built == set(edjas_examples.EXAMPLES)


def test_reports_published_and_self_contained(site):
    """Each report is published under a consistent name and remains droppable anywhere."""
    reports = sorted((site / "reports").glob("*.html"))
    assert {r.stem for r in reports} == set(edjas_examples.EXAMPLES)
    for report in reports:
        html = report.read_text(encoding="utf-8")
        assert html.startswith("<!doctype html>")
        relative = [
            value for value in re.findall(r'(?:href|src)="([^"]+)"', html)
            if not value.startswith(("http://", "https://", "mailto:", "#"))
        ]
        assert not relative, f"{report.name} depends on local assets: {relative}"


def test_example_pages_carry_their_own_prose(site):
    """Iframe content is invisible to search and to a screen reader's document flow, so
    each example page must carry the spec, its explanation and its attribution itself."""
    page = (site / "gallery" / "slgfs.html").read_text(encoding="utf-8")
    assert "Scottish Local Government Finance Statistics" in page   # attribution
    assert "[extract]" in page                                       # the spec source
    assert "revenue account" in page                                 # the spec's own header
    assert 'title="EDJAS report: slgfs"' in page                     # accessible iframe


# --- reference pages cannot drift from the code -------------------------------

def test_function_reference_lists_every_builtin(site):
    html = (site / "reference" / "functions.html").read_text(encoding="utf-8")
    for name in DEFAULT_FUNCTIONS:
        assert f'id="fn-{name}"' in html, f"{name} missing from the function reference"


def test_api_reference_shows_real_signatures(site):
    html = (site / "reference" / "api.html").read_text(encoding="utf-8")
    for name in ("read_spec", "render_report", "json_default"):
        assert f'id="api-{name}"' in html
    assert "headings=None" in html  # read from the live signature, not transcribed


def test_spec_language_documents_the_quote_rules(site):
    """The single/double quote distinction is documented only in ``_scan``'s docstring in
    the source; the site must state it too, or the rule stays invisible to users."""
    html = (site / "reference" / "spec-language.html").read_text(encoding="utf-8")
    assert "doubled" in html and "Bob" in html
    assert "case-insensitively" in html


def test_site_workbook_matches_the_live_registries():
    """The workbook the site extracts itself from must not drift from the code.

    Both tables name things that belong to code, so a stale hand-edit of the committed
    .xlsx would document functions or examples that do not exist. Same guard as
    ``test_demo_heading_map_keys_exist_in_data`` applies to the heading map.
    """
    data = build_site.site_data()
    assert {row["name"] for row in data["functions"]} == set(DEFAULT_FUNCTIONS)
    assert {row["name"] for row in data["gallery"]} == set(edjas_examples.EXAMPLES)


def test_workbook_driven_tables_reach_the_pages(site):
    """The content extracted from site.xlsx actually appears -- the dogfooding is real."""
    gallery = (site / "gallery" / "index.html").read_text(encoding="utf-8")
    for row in build_site.site_data()["gallery"]:
        assert row["teaser"] in gallery, f"{row['name']}'s teaser did not reach the page"
