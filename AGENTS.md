# Project scope

This project converts a number of Excel IP address calculation functions to Power Query, with the aim of building a reusable library for IP address manipulation.

# IP address calculation references

The local reference sources are:

- [JavaScript implementation](ip-calc.js)
- [Extracted VBA module](ipcalc_module.bas)

The original upstream references are:

- [JavaScript source](http://trk.free.fr/ipcalc/ip-calc.js)
- [Excel workbook containing the VBA source](http://trk.free.fr/ipcalc/ip-calc.xlsm)
- [IP calculation project page](http://trk.free.fr/ipcalc/)

Power Query references:

- [Power Query M formula language](https://learn.microsoft.com/en-us/powerquery-m/)
- [Power Query M language specification: consolidated grammar](https://learn.microsoft.com/en-us/powerquery-m/m-spec-consolidated-grammar)
- [Power Query M Primer](https://bengribaudo.com/power-query-m-primer)

## Standalone Power Query design guidance

- Before implementing or reviewing any Power Query M function, read [Power Query M IPAM build notes](POWERQUERY-M-IPAM-NOTES.md) and apply its relevant guidance. Treat it as project-specific implementation policy, not optional background material; document deliberate deviations in the implementation or focused validation.
- Prefer idiomatic, readable, maintainable Power Query M and established M best practices. Use the VBA and JavaScript implementations as behavioral references. Preserve their intended IPAM semantics and documented examples, while designing idiomatic Power Query M contracts and correcting implementation artifacts or inconsistencies where appropriate.
- M `number` arithmetic and equality use double precision by default. IPv4 integers, octets, masks, and prefix lengths are safe as numbers, but IPv6 must not be represented as one numeric value. Prefer text at public boundaries and a list of bytes or `binary` internally; use records for parsed address/subnet state.
- `null` is distinct from false and zero: `null = null` is true, while arithmetic and relational operations involving `null` generally yield `null`. Make validation and fallback behavior explicit.
- Lists are ordered and zero-based; records provide named fields. Use `Record.Field` for dynamic field names and fixed record lookup for static names.
- Preserve explicit table schemas for empty results; do not rely on input-derived column inference when downstream steps require those columns.
- Use `Table.RenameColumns` for actual table renames. Treat `Value.ReplaceType` as a type annotation only: table type details apply positionally, and implementation bugs can make ascribed names disagree with operations such as filtering.
- When several output fields depend on one parsed or source value, bind that value once in the relevant `let`/row scope and reuse it; do not invoke the parser or source independently for each field.
- Use `Value.Is`/`Type.Is` for type compatibility; reserve `=` between type values for deliberate identity or host-subtype-claim checks, since separately constructed equivalent type values may compare unequal.
- Treat table column type annotations as truthful output contracts, not conversion or validation. Use `Table.TransformColumnTypes` when values must be converted and checked; use `Value.ReplaceType` only when the existing values already satisfy the claim.
- Write public IPAM functions with a descriptive block comment, a readable `let` implementation, an adjacent `<FunctionName>Type` carrying `Documentation.*` metadata for the function and parameters, and a final `<FunctionName>Doc = Value.ReplaceType(...)` binding. Documentation metadata improves the host experience but does not replace explicit runtime validation.
- Choose join shape deliberately: `Table.Join` immediately flattens matches and can multiply rows, while `Table.NestedJoin` preserves one left row with a nested matches table until expansion is explicitly requested.
- Use `JoinKind.LeftSemi`/`LeftAnti` for existence/non-existence filters without row multiplication; treat a cross join as an explicit Cartesian-product operation with deliberate cardinality checks.
- M is dependency-ordered, immutable, and streaming-oriented. Tables and lists are not guaranteed snapshots; use `Table.Buffer` or `List.Buffer` only when a repeatable in-memory snapshot is required and the memory cost and loss of folding are justified. Buffer after filtering/reduction where possible, and remember buffering is shallow.

## Reference implementation policy

- Treat `ip-calc.js` and `ipcalc_module.bas` as reference implementations that
  describe the intended IPAM functionality.
- Prefer clear, idiomatic, maintainable Power Query M over mechanical
  translation.
- Preserve meaningful functional behavior and documented examples, but do not
  preserve implementation artifacts, Excel-specific APIs, mutation patterns,
  permissive parsing, or apparent bugs.
- When the references differ or are ambiguous, choose an explicit M-native
  contract based on the documented purpose, mathematical behavior, and useful
  validation rules.
- Record important interpretation decisions in tests or implementation notes.
- Public M functions should expose Power Query values and schemas, not VBA
  `Range`, `Variant`, `ByRef`, or worksheet-array conventions.

## Bead dependency workflow

- When given a Bead to convert a function from the VBA and/or JavaScript reference implementation, inspect its function dependencies before implementing it. If a dependency has not yet been converted, find the Bead for that dependency and ask the user whether they want the dependency implemented first before proceeding with the requested Bead.
