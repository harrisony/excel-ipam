# Repository Guidelines

## Purpose and Structure

This repository converts Excel IP address calculations into a reusable Power Query M library.

- `IpQ_ipv4.m`: Power Query IPv4 implementation.
- `ip-calc.js`: JavaScript behavioral reference.
- `ipcalc_module.bas`: extracted VBA behavioral reference.
- `POWERQUERY-M-IPAM-NOTES.md`: mandatory project-specific M guidance.
- `.beads/`: task and dependency tracking.

There is no Power Query runtime, and there never will be one. All validation must be source-level, reference-based, and deliberate.

## Research and Design Gate

Before implementing or reviewing an M function:

1. Read `POWERQUERY-M-IPAM-NOTES.md` and identify every applicable rule.
2. Use the Microsoft Learn MCP (`mcp__microsoftLearnDocs__microsoft_docs_search`
   followed by `mcp__microsoftLearnDocs__microsoft_docs_fetch` when detail is
   needed) to verify M language semantics, native functions, and type behavior.
   Do not use web search, browser search, or direct `learn.microsoft.com` web
   requests for this repository; the Microsoft Learn MCP is the authoritative
   documentation path.
3. Consult the JavaScript and VBA files for intended behavior, not as translation templates.
4. Inspect dependencies and their Beads. Ask before implementing an unfinished dependency.
5. Identify the function's semantic operation before writing code (for example,
   parsing, validation, projection, transformation, lookup, or aggregation).
   Implement that operation using M-native values and composition; do not
   mechanically reproduce the reference implementation's control flow,
   mutation, or intermediate representation.

“Idiomatic” and “best practice” require evidence. Search for native `Number.*`, `Text.*`, `List.*`, `Record.*`, and `Table.*` operations before writing manual logic. Explain rejected alternatives and document deliberate deviations.

## M Coding Standards

Use readable `let` expressions, explicit contracts and schemas, and function documentation metadata. Parse or derive each value once and reuse it. Make null and malformed-input behavior explicit. Represent IPv6 as text, byte lists, or binary—not one number. Use truthful type annotations and deliberate table joins. Buffer only when a repeatable snapshot is required.

## Measure Twice, Cut Once

Before completion, perform two separate reviews:

- **Behavioral review:** documented examples, boundaries, nulls, malformed input, empty results, and agreement with the references.
- **M review:** native-function usage, type correctness, schema accuracy, dependency ordering, duplicate evaluation, buffering, and accidental row multiplication.

Use `git diff --check` for basic source checks. If behavior cannot be established confidently from the notes, Microsoft Learn MCP, references, and source inspection, leave the uncertainty documented rather than declaring the work complete.

## Commits

Use concise prefixes consistent with project history, such as `feat:`, `fix:`, `docs:`, and `chore:`.
