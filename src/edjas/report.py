"""Render extracted spreadsheet data into an HTML report with Jinja2.

Jinja2 is an **optional** dependency: importing :mod:`edjas` (and this module) never
pulls it in. Install it with ``pip install edjas[demo]`` to call :func:`render_report`;
until then the function raises a clear :class:`ModuleNotFoundError` pointing at the
extra, so the core package stays free of the dependency.

The division of labour mirrors the rest of EDJAS: :func:`edjas.read_spec` turns a
workbook plus a TOML spec into a plain data dict, and this module only concerns itself
with presenting that dict. The extracted data is handed to the template as ``data``.
"""

from .spec import read_spec

__all__ = ["render_report"]


def _require_jinja2():
    """Import Jinja2 on demand, with an actionable message if the extra is missing."""
    try:
        import jinja2
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            "render_report needs Jinja2, an optional dependency of EDJAS. "
            "Install it with:  pip install edjas[demo]"
        ) from exc
    return jinja2


def render_report(spreadsheet, spec, template, templates_dir, functions=None, **context):
    """Extract ``spreadsheet`` per ``spec`` and render ``template`` to an HTML string.

    ``templates_dir`` is the directory Jinja2 loads templates from and ``template`` is
    the entry template's filename within it. The extracted data dict is passed to the
    template as ``data``; any extra keyword arguments are merged into the render
    context. ``functions`` is forwarded to :func:`edjas.read_spec` to add or override
    extraction functions. Autoescaping is on, so values from the spreadsheet are safe
    to interpolate into HTML.
    """
    jinja2 = _require_jinja2()
    data = read_spec(spreadsheet, spec, functions=functions)
    env = jinja2.Environment(
        loader=jinja2.FileSystemLoader(str(templates_dir)),
        autoescape=jinja2.select_autoescape(["html", "xml"]),
        trim_blocks=True,
        lstrip_blocks=True,
    )
    return env.get_template(template).render(data=data, **context)
