import argparse
import json
import sys

from importlib.metadata import PackageNotFoundError, version

from .spec import read_spec, open_workbook
from .report import render_report
from .names import describe, report
from .functions import DEFAULT_FUNCTIONS, json_default

try:
    __version__ = version("edjas")
except PackageNotFoundError:  # not installed (e.g. running from a source checkout)
    __version__ = "0.0.0+unknown"

__all__ = [
    "read_spec", "render_report", "open_workbook", "describe", "report",
    "DEFAULT_FUNCTIONS", "json_default", "__version__",
]


def build_parser():
    """The command-line parser, kept separate so the documentation can render its help."""
    parser = argparse.ArgumentParser(
        prog="edjas",
        description="Extract data in JSON from any spreadsheet, as directed by a spec.",
    )
    parser.add_argument("spreadsheet", help="path to the spreadsheet to read")
    parser.add_argument(
        "spec", nargs="?",
        help="path to the TOML extraction spec (not needed with --list-names)",
    )
    parser.add_argument(
        "--list-names", action="store_true",
        help="list the named ranges and Excel Tables the spreadsheet offers, "
             "instead of extracting from it",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true",
        help="report in more detail; with --list-names, list every unusable name "
             "individually rather than counting them",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.list_names:
        if args.spec is not None:
            parser.error("--list-names reads the spreadsheet alone; do not give a spec")
        try:
            workbook = open_workbook(args.spreadsheet)
        except (OSError, ValueError) as exc:
            parser.exit(1, f"{parser.prog}: error: {exc}\n")
        report(workbook, args.spreadsheet, verbose=args.verbose)
        return

    if args.spec is None:
        parser.error("a spec is required unless --list-names is given")
    try:
        data = read_spec(args.spreadsheet, args.spec)
    except (OSError, ValueError) as exc:  # missing file, bad TOML, bad spec/expression
        parser.exit(1, f"{parser.prog}: error: {exc}\n")
    json.dump(data, sys.stdout, default=json_default)
    sys.stdout.write("\n")
