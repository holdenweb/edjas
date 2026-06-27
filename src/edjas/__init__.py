import json
import sys

from importlib.metadata import PackageNotFoundError, version

from .read_params import read_file

try:
    __version__ = version("edjas")
except PackageNotFoundError:  # not installed (e.g. running from a source checkout)
    __version__ = "0.0.0+unknown"

__all__ = ["read_file", "__version__"]

def main():
    json.dump(read_file(*sys.argv[1:]), sys.stdout)