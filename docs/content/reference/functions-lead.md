Any extraction expression may append a **pipeline** of functions, separated by `|` and
applied left to right after the value is extracted, so `[Grid | transpose | records]`
transposes the range and then builds objects from it.

> **A specification cannot run arbitrary code.** Functions are looked up in a fixed,
> built-in registry rather than evaluated. A specification can only invoke the functions
> listed here, plus any you supply yourself when calling
> [`read_spec`](site:reference/api.html); it can never execute code of its own.

A function may take arguments after its name, separated by spaces. An argument is a
**number** (`2`), a **double-quoted string** (`", "`), or a **bare word**, which is read
as another range reference. The extracted value is always passed first, so
`[Price | round 2]` means `round(Price, 2)`.

The list below is generated from the registry itself, so it cannot describe a function
that does not exist.
