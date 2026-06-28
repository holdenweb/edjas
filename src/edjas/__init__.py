import argparse
import json
import sys

from importlib.metadata import PackageNotFoundError, version

from .read_params import read_file

try:
    __version__ = version("edjas")
except PackageNotFoundError:  # not installed (e.g. running from a source checkout)
    __version__ = "0.0.0+unknown"

__all__ = ["read_file", "__version__"]

def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="edjas",
        description="Extract data in JSON from any spreadsheet.",
    )
    parser.add_argument("file", help="path to the spreadsheet to read")
    parser.add_argument(
        "-r",
        "--range",
        default="Parameters",
        help="named range to use as the starting point (default: Parameters)",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    args = parser.parse_args(argv)
    json.dump(read_file(args.file, args.range), sys.stdout)
    sys.stdout.write("\n")