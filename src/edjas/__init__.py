import argparse
import json
import sys

from importlib.metadata import PackageNotFoundError, version

from .spec import read_spec
from .functions import DEFAULT_FUNCTIONS, json_default

try:
    __version__ = version("edjas")
except PackageNotFoundError:  # not installed (e.g. running from a source checkout)
    __version__ = "0.0.0+unknown"

__all__ = ["read_spec", "DEFAULT_FUNCTIONS", "json_default", "__version__"]

def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="edjas",
        description="Extract data in JSON from any spreadsheet, as directed by a spec.",
    )
    parser.add_argument("spreadsheet", help="path to the spreadsheet to read")
    parser.add_argument("spec", help="path to the TOML extraction spec")
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    args = parser.parse_args(argv)
    try:
        data = read_spec(args.spreadsheet, args.spec)
    except (OSError, ValueError) as exc:  # missing file, bad TOML, bad spec/expression
        parser.exit(1, f"{parser.prog}: error: {exc}\n")
    json.dump(data, sys.stdout, default=json_default)
    sys.stdout.write("\n")