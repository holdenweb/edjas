# ADR-0001: The workbook is never modified

**Status:** Accepted

**Date:** 2026-08-26

## Context

EDJAS exists because organisations run real work in spreadsheets and are not going
to stop. Every tool that has tried to improve that situation by asking people to
change their workbook — add a macro, insert a marker column, restructure a sheet so
it is machine-readable — has run into the same wall: the workbook belongs to someone
else, it is in use, and it is often the authoritative copy. A tool that needs the
file changed before it can help is a tool that needs permission, a change-control
conversation, and a person willing to take responsibility if the change breaks a
formula three sheets away.

There is a sharper version of the problem for published data. Government
statistical releases are the exact case EDJAS was built for, and they are immutable
by definition: the published file is the record. Anything that writes to it is not
merely inconvenient, it is wrong.

## Decision

EDJAS only ever reads a spreadsheet. It opens the file, takes the values it was
asked for, and closes it; there is no code path that writes to a workbook, and no
option to enable one. Everything EDJAS needs to know about an extraction comes from
somewhere other than the file being extracted. This is the property the whole design
is arranged around, not a feature of it.

## Consequences

Anything EDJAS needs to be told has to be told somewhere else, which is what forces
[ADR-0002](0002-extraction-is-described-beside-the-spreadsheet.md) and everything
downstream of it. A user cannot annotate a workbook to make it easier to extract;
they must describe the extraction separately, which is more work the first time and
less work every time after.

It also means EDJAS cannot repair what it reads. A workbook with values in the wrong
place stays that way, and the specification has to accommodate it. Several of the
uglier corners of the worked examples — padding cells, nil markers, footnotes inside
the data range — exist because fixing them at source is not an option available to us.

The promise is worth nothing if it is only intended, so it is tested: rendering a
report asserts the source file is byte-for-byte unchanged afterwards.

## Alternatives considered

**Write extraction instructions into the workbook.** This was the original design and
is recorded as rejected in [ADR-0002](0002-extraction-is-described-beside-the-spreadsheet.md).

**Offer an opt-in write mode** — normalise the sheet, add named ranges, tidy headings.
Rejected because the guarantee is only useful if it is absolute. "EDJAS never writes
to your file, unless you asked it to, in which case check which mode you are in" is
not a guarantee anyone can rely on, and the audience least able to check is the one
that most needs the assurance.
