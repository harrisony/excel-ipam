# Power Query M build notes for IPAM

This is the implementation-oriented extraction from the Power Query M Primer
articles reviewed for this project. It is guidance for the standalone `.pq`
library, not a replacement for the Microsoft language specification.

## Query and function shape

- A `.pq` file is a top-level M expression that must return a value or raise an
  error. The usual structure is `let` bindings followed by an `in` expression.
- A `let` binding is an expression, not a mutable assignment or procedural
  statement. Bindings are comma-separated; the `in` expression is the result.
- Functions are values. They can be assigned to names, passed to library
  functions, returned from functions, defined inline, and used as closures.
- `each expression` is shorthand for `(_) => expression`; `[Field]` inside an
  `each` expression is shorthand for `_[Field]`. Prefer named functions when
  the implicit argument would obscure the IP calculation.
- Recursive self-reference uses `@FunctionName`. Prefer list transforms or
  accumulators unless recursion makes the algorithm clearer.

## Execution model

- M evaluates by dependency, not by the physical order of bindings. Keep the
  source order readable, but never rely on it as control flow.
- Values are immutable. There is no reassignment or ordinary procedural loop
  state; use `List.Transform`, `List.Select`, `List.Accumulate`, or
  `List.Generate` for iteration.
- `let` expressions, lists, records, and tables are evaluated lazily. A value
  that is not needed for the result may never be evaluated.
- List and table values can have streaming semantics. A variable referring to
  one is not automatically a materialized, stable snapshot. Buffer only when
  repeatable access is required and the memory cost is justified.

## Names and fields

- Use regular identifiers for internal functions and steps.
- Use quoted identifiers such as `#"IP Address"` when preserving a field or
  step name containing spaces, keywords, or other special characters.
- Fixed field access uses `record[Field]`. A field name held in a variable must
  be accessed with `Record.Field(record, fieldName)` or
  `Record.FieldOrDefault`.
- List indexes are zero-based. Optional list/field access with `?` returns
  `null` instead of raising a missing-item or missing-field error.

## Values relevant to IP calculations

### Text

- Text literals use double quotes; an embedded double quote is written as `""`.
- Text comparison is case-sensitive unless a comparer is supplied.
- `&` does not implicitly convert numbers to text. Use `Text.From` or
  `Number.ToText` deliberately.
- Concatenating text with `null` produces `null`; handle missing values before
  formatting an address.
- `#(cr)`, `#(lf)`, and `#(tab)` are available text escapes.

### Numbers

- Numeric literals include decimal, exponential, and hexadecimal forms such as
  `0xFF`.
- Ordinary arithmetic and equality use double precision. M `number` is not an
  arbitrary-width integer representation.
- IPv4 values, octets, masks, prefix lengths, and address offsets fit safely in
  the exact range needed for ordinary double-based arithmetic. Do not collapse
  an IPv6 address into one M number; use text, a list of bytes/words, or
  `binary`.
- Decimal precision is an explicit choice through functions such as
  `Value.Add` with `Precision.Decimal`; it is not needed for ordinary IP
  bit/byte calculations.

### Logical and null

- Logical values are `true` and `false`. Numeric conversion treats zero as
  false and any other number as true, so validation should use explicit
  comparisons when that distinction matters.
- `null` represents absence or unknown data. `null = null` is `true`, while
  arithmetic and relational operations involving `null` generally produce
  `null`. `and` and `or` have three-valued behavior.
- Choose a null policy for every public function: reject it, propagate it, or
  replace it with a documented default. Do not let concatenation or comparison
  choose accidentally.

### Binary, lists, and records

- `binary` is an exact byte sequence and can represent IPv6 data without the
  precision problem of a numeric representation. Convert to a byte list when
  list-oriented processing is clearer.
- Lists are ordered, zero-based sequences and are a natural representation for
  IPv4 octets or IPv6 bytes. They may contain mixed values unless a stricter
  type is applied.
- Records group named values and are a natural representation for parsed state,
  for example `[Address, PrefixLength, Bytes]`.
- Record field order is not part of record equality, while list order is part of
  list equality. Merging records with `&` lets the right-hand record override
  duplicate fields.
- Record fields and `let` bindings are evaluated on demand. Once a record field
  or ordinary `let` value is evaluated, its value or error is fixed; streaming
  list/table sources are the important exception.

## Tables and performance

- Streaming lets consumers such as `Table.FirstN` request only the rows needed
  from upstream transformations.
- Operations such as sorting, grouping, joins, pivoting, and buffering may hold
  rows in memory. Apply filters and column reduction before these operations
  where semantics permit.
- Query folding can move supported transformations into an external source,
  but it is relevant only if the standalone library later reads from a
  foldable source. Do not design pure IP functions around folding assumptions.
- `Table.Buffer` and `List.Buffer` stabilize a value for one query execution by
  materializing it. They are deliberate resource trade-offs, not general fixes
  for evaluation surprises.

## Lazy table-construction nuggets

- A table is a streamed construct, not necessarily a fully materialized set of
  cell values. A consumer such as `Table.RowCount` may enumerate rows without
  evaluating their cells, so errors or expensive work in unused columns may not
  occur.
- Column values are evaluated when requested. Narrowing a table to selected
  columns can therefore avoid work needed only by removed columns. Keep expensive
  parsing or source access inside the expressions that actually need it.
- When multiple columns depend on one value, bind that value once in a local
  `let` expression for the row and reuse it. M saves the computed value for that
  identifier during its lifetime, preventing repeated parsing/source calls across
  dependent columns while retaining lazy evaluation until a dependent column is
  requested.
- This sharing is scoped to the identifier/evaluation that created it. A later
  independent streaming of the row can recompute the expression. It is not a
  general cache; use buffering only when a stable repeated snapshot is required.

## Buffering-specific nuggets

- `Table.Buffer` and `List.Buffer` eagerly enumerate their immediate source by
  default and retain a stable sequence of rows/items for the buffer's lifetime.
  This can prevent repeated source reads, but it also consumes memory and can
  prevent later query folding. Filter and reduce columns before buffering when
  the semantics allow it.
- Buffering is shallow. Scalar row/cell values are captured, but record fields
  remain lazy and nested tables, lists, and binary values are not recursively
  enumerated. Buffer the nested value separately if that is genuinely required.
- Errors attached to an item, row, or cell are saved and re-raised when that
  buffered value is accessed. An error encountered while enumerating the source
  itself is different: it causes `Table.Buffer`/`List.Buffer` to fail while
  populating, so buffering can change when that error is observed and caught.
- `Table.Buffer` supports `BufferMode.Delayed`, which defers population until
  the first row is requested. `List.Buffer` is eager; do not assume buffering a
  value is free merely because no consumer later reads its contents.

## Empty-table and paging nuggets

- Never let a table's output schema depend only on observed input rows.
  `Table.FromList` with an empty input and no explicit column list can produce a
  zero-column table, so later expansion or field selection fails with a missing
  column error. Supply the expected column names (and, where useful, an explicit
  table type) at the construction boundary.
- Treat empty results as a first-class contract case. A helper that returns no
  matching IPAM rows should still return the same named columns and types as a
  populated result; test empty, singleton, and populated inputs.
- Paged results have a second empty-case trap: collecting per-page tables and
  expanding them can preserve an empty page's holder row as an apparent row
  containing `null` in every output column. Filter empty page tables before
  combining them, or use a schema-preserving aggregation that handles empty
  pages explicitly.
- Filtering empty pages can require an extra read of each page table. That may
  mean extra source requests, so the choice is a correctness/performance tradeoff
  when the source is not already cached. Pure IPAM helpers should avoid this
  uncertainty by operating on materialized in-memory values where practical.

## Table-syntax nuggets

- `#table({"ColumnA", "ColumnB"}, {{valueA, valueB}, ...})` constructs a
  table from an ordered column-name list and an outer list of row-value lists.
  Keep the column list explicit when the result is part of a public contract.
- Positional table selection, `table{n}`, returns row `n` as a record and uses
  zero-based indexing. `table{n}?` returns `null` when that position is absent.
  Value-based selection, `table{[Key = value]}`, returns at most one record:
  `?` suppresses a no-match error but does not suppress an error when multiple
  rows match.
- Column access, `table[Column]`, returns that column as an ordered list. This
  lets list functions operate on a column without a separate table-specific
  distinct/sum operation.
- Table projection syntax uses hard-coded column names, such as
  `table[[Address], [PrefixLength]]`. An optional projection can create missing
  columns filled with `null`; use `Table.SelectColumns` when the column list is
  dynamic.
- Table equality compares the data rather than schema details, metadata, or
  column order: names, values, and row positions matter, but column types and
  ordering do not. Do not use table equality as a schema/type assertion.
- The `&` operator combines tables by column name, not position. Columns absent
  from one side are filled with `null`, output columns follow the first table's
  order followed by new columns from the second, and incompatible column type
  claims do not prevent combination. Validate or normalize values explicitly
when a uniform output type is required.

## Additional table-processing nuggets

- Table keys are optimization hints, not correctness guarantees. Add a key only
  when the column values have been verified to identify rows, and measure
  whether the specific join or lookup benefits. The operations that use key
  metadata and the performance benefit are implementation details.
- Host-level caching of native query results is optional and environment
  specific. It may explain why a source request runs fewer times, but it must
  never be relied on for data stability or correctness; buffer or restructure
  the M expression when a stable snapshot is required.
- Privacy levels primarily govern whether data from one source may be folded
  into a native request sent to another source. They are a data-protection
  boundary, not a general prohibition on combining sources. Do not disable the
  protection layer by default; if a firewall error occurs, separating each
  source into its own query and combining those query results is a useful
  partitioning pattern.
- Native requests may execute more than once for schema discovery, previews,
  firewall analysis, or connector behavior. Treat external requests as
  repeatable reads; never depend on a native query causing a write exactly once.

## Join nuggets

- `Table.Join` directly pairs matching rows and can multiply a left row when
  multiple right rows match. `Table.NestedJoin` instead preserves the left row
  and adds a nested table of matches; row multiplication occurs only if that
  nested table is later expanded. Choose based on whether the relationship or
  the flattened rows are the desired output.
- A nested join is useful when an IP address, subnet, or allocation should retain
  its matching records for later counts, summaries, or custom logic. Work with
  the nested table directly, or use `Table.AggregateTableColumn`, rather than
  expanding and then regrouping.
- `Table.Join` normally includes all columns from both sides, so duplicate names
  must be renamed before joining. `Table.ExpandTableColumn` can select only the
  right-side columns needed and rename them during expansion, which helps keep
  a stable output schema.
- `Table.AddJoinColumn` is a less configurable equivalent of `Table.NestedJoin`;
  prefer `Table.NestedJoin` when explicit join options are needed.
- Semi-joins are existence filters: `JoinKind.LeftSemi` keeps each left row
  having at least one right match, without returning right-side data or
  multiplying rows. Anti-joins (`JoinKind.LeftAnti`/`RightAnti`) keep rows with
  no match. Prefer the explicit join kind over simulating existence with a
  nested join and then discarding the nested column when the connector can
  optimize the declared intent.
- M has no dedicated cross-join kind. To form a Cartesian product, add the
  right table as a nested value on every left row and expand it; the output row
  count is the product of the two input row counts, so use this deliberately.
  Supplying `Value.Type(rightTable)` when adding the nested table can preserve
  useful type information for the subsequent expansion.

## Initial IPAM design implications

- Keep public inputs and outputs in the project’s human-readable IP notation.
- Use a validated internal representation for arithmetic: numbers for IPv4
  octets/offsets and lists or binary for IPv6 bytes.
- Return structured records from parsing helpers so later calculations do not
  repeatedly reparse the same text.
- Make invalid input, `null`, prefix-length bounds, and overflow/wrap behavior
  explicit in function contracts and tests.
- Prefer small pure helpers whose intermediate values are visible in `let`;
  this makes the standalone file testable without relying on the Query Editor’s
  Applied Steps UI.
- Public function convention: place a block comment immediately above the
  implementation; keep the implementation readable and `let`-based; define an
  adjacent `<FunctionName>Type` with `Documentation.*` metadata on the function
  and its parameters; then expose `<FunctionName>Doc =
  Value.ReplaceType(<FunctionName>, <FunctionName>Type)`. Use the documented
  examples as part of the public contract, but keep parsing, range checks, and
  cross-field validation in the implementation.
- Create private helpers when they make parsing, validation, normalization,
  projection, or repeated calculations clearer. Keep them local and give them
  explicit parameter and return types when that improves correctness or
  readability. Private helpers do not need `Documentation.*` metadata, a public
  `Doc` value, or `Value.ReplaceType` unless they are intentionally exposed as
  part of the library API. Do not add helpers solely to mirror reference
  control flow or introduce unnecessary indirection.

### Function documentation metadata contract

- Document a public function by defining types for its parameters, defining a
  function type, attaching metadata to that type, and applying it with
  `Value.ReplaceType`. Documentation metadata is host-facing; it does not
  replace runtime validation or change the function's implementation.
- Function-level metadata used by this project is `Documentation.Name` (the
  displayed function name), `Documentation.LongDescription` (the function
  description), and `Documentation.Examples` (a list of records whose optional
  text fields are `Description`, `Code`, and `Result`). Microsoft Learn does not
  define a function-level `Documentation.Description` field; do not emit it.
- Parameter-level metadata used by this project is
  `Documentation.FieldCaption`, `Documentation.FieldDescription`,
  `Documentation.SampleValues`, and `Documentation.AllowedValues`. The latter
  two are UI hints; `Documentation.AllowedValues` does not prevent a user from
  supplying another value in a query.
- The Microsoft Learn `handling-documentation` article is a reference for
  these fields, not a routine prerequisite for implementing an IPAM function.
  Use targeted Microsoft Learn search when this contract is insufficient; do
  not fetch the entire article merely to confirm these rules.

## Reference-to-M interpretation

- Use the reference implementations to discover the function inventory,
  examples, algorithms, and intended IPAM concepts.
- Re-design each function around M-native inputs, outputs, validation, and
  composition.
- A reference implementation is not authoritative when it relies on mutation,
  implicit conversion, Excel worksheet behavior, or inconsistent error
  handling.
- Prefer one consistent library-wide contract over reproducing inconsistent
  behavior between reference functions.
- For ambiguous behavior, document the chosen rule and add boundary tests.
- Validation error precedence is specified only where an implementation creates
  an explicit dependency between checks. When multiple independent arguments
  are invalid, M's dependency-driven evaluation does not define which error is
  observed first; callers and tests must not depend on a particular precedence.
  Test one invalid component at a time unless a public function explicitly
  documents an ordered validation contract.

## Control-flow nuggets

- M has one general branching form: `if test then valueA else valueB`. The test
  must produce a logical value, and an `else` is mandatory because every
  expression must return a value or raise an error.
- Only the selected branch of an `if` expression is evaluated. Put guarded
  parsing, conversion, or error-prone work in the appropriate branch instead
  of evaluating it before the condition.
- There is no `switch`, `for`, `foreach`, or `while` syntax. Use a record or
  table as a lookup map for many fixed cases, and use list functions for
  iteration. `List.Generate` can produce a sequence, `List.Transform` applies
  work, `List.Select` filters, and `List.Accumulate` carries state through a
  fold.
- Streaming can make bounded iteration stop early: a consumer such as
  `List.First` may request only enough values from a generated/transformed list
  to produce its result. Do not add eager materialization when early stopping
  is useful.
- The `??` null-coalescing operator is a concise alternative to
  `if value <> null then value else fallback`; verify host compatibility if
  targeting older Power Query environments.

## Error-handling nuggets

- Every branch still has to return a value or raise an error. Use `error
  Error.Record(reason, message, detail)` for deliberate validation failures;
  the structured reason/message/detail makes invalid IP input diagnosable.
- Errors are contained by the value that encounters them (a `let` binding,
  record field, table cell, or similar) and are re-raised when that value is
  accessed. Lazy evaluation means an error in an unused field or cell may never
  be evaluated.
- `try expression otherwise fallback` is appropriate when every error has the
  same recovery policy. `try expression catch (e) => ...` allows selective
  recovery by inspecting the error record and re-raising with `error e` when it
  should not be hidden. The catch function must be defined inline.
- Plain `try expression` returns a record containing either
  `[HasError = false, Value = ...]` or `[HasError = true, Error = ...]`, which
  is useful when a parser should return diagnostics instead of immediately
  aborting.
- Put `try` at the level where the failing expression is evaluated. Wrapping a
  whole lazy table does not catch errors stored inside table cells, and a
  transform function may not receive a cell value if eager argument evaluation
  errors before the function is invoked. For cell-level recovery, access the
  value inside a row-level `try`, or use `Table.ReplaceErrorValues` when a
  uniform replacement is sufficient.
- For this library, scalar parsing helpers should normally raise structured
  errors for invalid input. Adapters can use `try` when they intentionally need
  nullable results or error diagnostics; tests should cover both paths.
- Error `Detail` may be a record, not only text. For IPAM validation, preserve
  machine-readable context such as `[Input = input, Component = "octet",
  Position = 3, Expected = "0..255"]` in the detail record. Keep the human
  explanation in `Message`, and have tests inspect detail fields rather than
  parsing the displayed message.
- Structured errors can separate a human-readable template from its values:
  use `Message.Format` with zero-based placeholders such as `#{0}` and
  `Message.Parameters` for the corresponding list. A caught error retains the
  rendered `Message`, the template, and the parameters, so diagnostics do not
  need to parse interpolated text. Keep stable machine-readable validation
  context in `Detail`; structured messages complement it rather than replace it.
- Structured error messages are an optional newer feature. If the `.pq` file
  must run across older hosts, verify support or retain a plain `Message` path.

## Inspection and debugging nuggets

- M does not directly convert tables, lists, or records to text, but JSON is a
  useful diagnostic representation for nested values:
  `(input as any) as nullable text => if input = null then null else
  Text.FromBinary(Json.FromValue(input))`.
- This is appropriate for inspecting parsed IP records, byte lists, and test
  tables in one result. It is a debugging/rendering helper, not the canonical
  IP serialization format; keep public IP text formatting explicit and stable.
- `Json.FromValue` is not universal: errors, type values, and function values
  may be incompatible. If diagnostics need to handle those, recursively walk
  the structure and replace unsupported values with explicit text placeholders
  before JSON conversion.

## Type-contract nuggets

- M is dynamically typed: variables do not have declared types; runtime values
  do. Type annotations on function parameters and return values are runtime
  compatibility checks, not static variable declarations.
- Untyped parameters and return values default to `any`. Use explicit
  signatures for public IPAM functions, such as `(input as text) as record =>`.
- `nullable T` explicitly accepts either values compatible with `T` or `null`.
  Do not use `any` merely to accommodate nulls; it also admits unrelated value
  kinds. Decide whether each function rejects, propagates, or accepts null.
- `as` asserts compatibility and raises an error when it fails; it does not
  perform conversions. Use `Text.From`, `Number.From`, and similar functions
  when conversion is intended.
- `is`/`as` use hard-coded primitive type names. Use `Value.Is`/`Value.As` when
  the type value must be supplied dynamically. `as none` is a useful contract
  for a helper whose only valid outcome is raising an error.
- Broad `list`, `record`, and `table` annotations describe only the general
  kind of value. Add more specific custom types later if consumers need
  enforceable field, item, or column structure.

## Facet nuggets

- Facets are informational metadata, primarily for connectors, hosts, and
  tooling. They do not change M or mashup-engine behavior.
- `Int64.Type`, `Int32.Type`, `Currency.Type`, `Percentage.Type`, and similar
  names are type claims decorating the base type `number`; they do not create
  integer or fixed-point arithmetic and do not solve IPv6 precision.
- `Value.ReplaceType` changes the type information associated with a
  structurally compatible value; it does not convert the value or validate a
  type claim. Use explicit `.From` conversions or
  `Table.TransformColumnTypes` when conversion/range validation is intended.
- Table column types are largely courtesy information: the engine needs the
  table's column structure, while individual values determine runtime behavior.
  A `#table` type or `Value.ReplaceType` can claim `number`, `text`, or another
  type without changing or checking the values. Such claims still matter to
  hosts and destinations, so they must be truthful.
- `Table.TransformColumnTypes` is different: it applies the relevant `.From`
  conversion to each targeted value, sets the resulting column type, and
  records errors for values that cannot be converted. Use it at deliberate
  input/output boundaries when conversion or validation is required; use cheap
  ascription only after compatibility is already guaranteed.
- For table types, `Value.ReplaceType` applies column details positionally,
  not by matching column names. An upstream reorder can therefore attach the
  wrong names or type claims to the data unless the ascribed type is updated in
  the same order. Keep table construction and schema order explicit.
- Do not use `Value.ReplaceType` as the mechanism for renaming real table
  columns. The article reports a Power Query implementation bug where
  `Table.Schema` and direct field access can observe an ascribed new name while
  operations such as `Table.SelectRows` still use the old name. Use
  `Table.RenameColumns` for structural renames, then apply type annotations only
  after the schema is final; test any advanced ascription against the target
  host.
- Only emit facets/type claims that are true. They can influence how Excel or
  Power BI stores and displays output tables even though M itself may accept a
  value that does not satisfy the claim.

## Custom-type and conformance nuggets

- Custom list, record, table, and function types are useful for describing
  structure, documentation, and host/tool integration, but child types are not
  generally validated or enforced by the mashup engine. A list of text can be
  ascribed `type { number }` without its items being checked.
- Therefore, a custom type is not a substitute for IP validation. Explicitly
  inspect octet/byte values, field presence, prefix bounds, and cross-field
  invariants in parser/validator functions.
- Open record types (`...`) express a minimum required shape and optional
  fields express optional presence. Table row types are more constrained and do
  not support arbitrary extra columns in the same way.
- Type ascription requires compatible base/overall structure, but child field,
  item, column, and function assertion details may still be unchecked. For
  records and tables, ascription associates child details positionally; use
  `Record.RenameFields` or `Table.RenameColumns` for robust name-based renames.
- Ascribing a new function type changes reported information, not the function
  implementation’s original argument/return checks. Do not use it to retrofit
  behavior or validation onto an existing helper.
- Type equality is primarily identity-oriented: the same type value compares
  equal to itself, but separately created type expressions can compare false
  even when they describe the same type. Metadata and (in the flagship engine)
  facets are ignored for equality, which makes `=` unsuitable as a general
  structural type comparison.
- Distinguish identity from compatibility. Use `Type.Is`, `Value.Is`, `Value.As`,
  or `is`/`as` when asking whether a value conforms to a type. `Int64.Type`,
  `Currency.Type`, and `type number` can classify numeric values compatibly even
  though their type values are distinct for equality and may convey different
  host-facing subtype claims.
- For this project, start with explicit validators and ordinary typed function
  signatures. Add custom types only when their documentation, table schema, or
  host-integration value is demonstrated by the output contract.

## Metadata and scope nuggets

- Metadata is an ordinary record attached to a value, not to a variable. It is
  optional descriptive or control information; it has no intrinsic meaning to
  the M engine unless code, a library function, or the host chooses to read it.
- Most operators produce a new value with fresh/empty metadata. If metadata is
  important across a transform, reattach it deliberately. Required IPAM state
  belongs in explicit record fields, not hidden metadata.
- Metadata on a function’s type can provide Query Editor documentation such as
  a name, long description, examples, and parameter captions. This is useful
  later for a polished reusable library, but it is host-facing documentation,
  not runtime validation.
- Identifier resolution is local-first: a binding shadows same-named values in
  parent, section, or global scopes. Avoid ambiguous names in nested `let` and
  `each` expressions; rename an outer value before entering a nested scope when
  both values are needed.
- Sections organize named top-level expressions. `shared` members are exposed
  through the global environment; unshared members are intended for local use.
  This matters when the library grows beyond one standalone expression, but a
  normal `.pq` function/query should not assume it can edit section documents.
- `#shared` exposes the current global environment and `#sections` exposes
  sections in advanced environments. Standard-library names such as
  `Table.SelectRows` are names with dots, not class/member access.

## Dynamic evaluation and closures

- `Expression.Evaluate` evaluates M source text in a clean environment by
  default. Its optional record supplies the entire environment, so required
  library functions must be injected explicitly. Treat dynamic M text as code:
  never evaluate untrusted input with `#shared` unless the security impact is
  understood. The IPAM library should avoid dynamic evaluation.
- Functions are closures: they retain the environment in which they were
  defined, not the caller’s later environment. This makes configurable helper
  generators reliable and prevents a caller from silently rebinding their
  dependencies.
- Closures can implement object-like state by returning a record of functions
  that each return a new record. This is useful for connector/query-folding
  infrastructure, but is unnecessary complexity for ordinary pure IP helpers.

## References

- [Power Query M formula language](https://learn.microsoft.com/en-us/powerquery-m/)
- [Primer part 1: expressions and `let`](https://bengribaudo.com/blog/2017/11/17/4107/power-query-m-primer-part1-introduction-simple-expressions-let)
- [Primer part 2: defining functions](https://bengribaudo.com/blog/2017/11/28/4199/power-query-m-primer-part2-functions-defining)
- [Primer part 3: function values and `each`](https://bengribaudo.com/blog/2017/12/08/4270/power-query-m-primer-part3-functions-function-values-passing-returning-defining-inline-recursion)
- [Primer part 4: variables and identifiers](https://bengribaudo.com/blog/2018/01/19/4321/power-query-m-primer-part4-variables-identifiers)
- [Primer part 5: paradigm](https://bengribaudo.com/blog/2018/02/28/4391/power-query-m-primer-part5-paradigm)
- [Primer part 6: text](https://bengribaudo.com/blog/2018/06/26/4470/power-query-m-primer-part6-types-intro-text)
- [Primer part 7: numbers](https://bengribaudo.com/blog/2018/07/31/4497/power-query-m-primer-part7-types-numbers)
- [Primer part 9: logical, null, and binary](https://bengribaudo.com/blog/2018/09/13/4617/power-query-m-primer-part9-types-logical-null-binary)
- [Primer part 10: lists and records](https://bengribaudo.com/blog/2018/10/30/4644/power-query-m-primer-part10-types-list-record)
- [Primer part 11: table syntax](https://bengribaudo.com/blog/2019/09/19/4713/power-query-m-primer-part11-tables-syntax)
- [Primer part 12: table streaming, folding, and buffering](https://bengribaudo.com/blog/2019/12/10/4778/power-query-m-primer-part12-tables-table-think-i)
- [Primer part 13: table keys, caching, privacy, and repeated requests](https://bengribaudo.com/blog/2019/12/20/4805/power-query-m-primer-part13-tables-table-think-ii)
- [Primer part 14: control structure and iteration](https://bengribaudo.com/blog/2020/01/06/4844/power-query-m-primer-part14-control-structure)
- [Primer part 15: error handling](https://bengribaudo.com/blog/2020/01/15/4883/power-query-m-primer-part-15-error-handling)
- [Primer part 16: type-system basics](https://bengribaudo.com/blog/2020/02/05/4948/power-query-m-primer-part16-type-system-i-basics)
- [Primer part 17: type-system facets](https://bengribaudo.com/blog/2020/02/28/5009/power-query-m-primer-part17-type-system-ii-facets)
- [Primer part 18: custom types](https://bengribaudo.com/blog/2020/06/02/5259/power-query-m-primer-part18-type-system-iii-custom-types)
- [Primer part 19: ascription, conformance, and equality](https://bengribaudo.com/blog/2020/09/03/5408/power-query-m-primer-part19-type-system-iv-ascription-conformance-and-equalitys-strange-behaviors)
- [Primer part 20: metadata](https://bengribaudo.com/blog/2021/03/17/5523/power-query-m-primer-part20-metadata)
- [Primer part 21: identifier scope and sections](https://bengribaudo.com/blog/2021/07/12/5809/power-query-m-primer-part21-identifier-scope-sections)
- [Primer part 22: global environment and closures](https://bengribaudo.com/blog/2021/09/01/5989/power-query-m-primer-part22-identifier-scope-ii-controlling-the-global-environment-closures)
- [Enhancing an Error's Detail](https://bengribaudo.com/blog/2022/02/21/6561/enhancing-an-errors-detail)
- [Zero Rows Can Bite (part 1): the mysterious missing column](https://bengribaudo.com/blog/2022/05/03/6691/zero-rows-can-bite-part-1-the-mysterious-missing-column)
- [Zero Rows Can Bite (part 2): the mysterious all-null row](https://bengribaudo.com/blog/2022/06/01/6776/zero-rows-can-bite-part-2-the-mysterious-all-null-row)
- [Value.ReplaceType & Table Column Renames (Bug Warning!)](https://bengribaudo.com/blog/2023/02/15/7268/value-replacetype-table-column-renames-bug-warning)
- [Render Tables, Lists, Records -> Text](https://bengribaudo.com/blog/2023/10/12/7406/render-tables-lists-records-text)
- [Lazy, Streamed, Immutable: Try Building a Table](https://bengribaudo.com/blog/2023/03/03/7292/lazy-streamed-immutable-try-building-a-table)
- [Exploring Power Query Buffering: How Table.Buffer and List.Buffer Work](https://bengribaudo.com/blog/2024/10/03/7489/exploring-power-query-buffering-how-table-buffer-and-list-buffer-work)
- [New M Feature: Structured Error Messages](https://bengribaudo.com/blog/2022/05/24/6753/new-m-feature-structured-error-messages)
- [Type Equality](https://bengribaudo.com/blog/2025/12/24/7562/type-equality)
- [Column Types Don’t Matter, or Do They?](https://bengribaudo.com/blog/2026/01/12/7602/column-types-dont-matter-or-do-they)
- [Deep Dive Into Joins (Part 1): Join vs. Nested Join](https://bengribaudo.com/blog/2026/05/05/7714/deep-dive-into-joins-part-1-join-vs-nested-join)
- [Deep Dive Into Joins (Part 2): Not So Common, But Real-World Useful](https://bengribaudo.com/blog/2026/05/22/7767/deep-dive-into-joins-part-2-not-so-common-but-real-world-useful)
