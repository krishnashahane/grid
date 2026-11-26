# Formula compatibility

GRID follows the widely used spreadsheet conventions for formulas, values and
errors, so that existing knowledge and existing formulas carry over. This page
lists the conventions GRID implements and every **known, deliberate difference**.

## Conventions implemented

| Area | Behaviour |
| --- | --- |
| Operator precedence | `:` range → unary `-` → `%` → `^` → `* /` → `+ -` → `&` → comparisons. Unary minus binds tighter than `^` (`-2^2 = 4`), and `^` is left-associative (`2^3^2 = 64`). |
| Comparison order | numbers < text < booleans. Text comparison is case-insensitive. A blank compares as `0`, `""` or `FALSE` depending on the other operand. |
| Coercion | Text that looks like a number, percent, currency or date is converted in arithmetic (`"3"+1 = 4`). Other text gives `#VALUE!`. `TRUE`/`FALSE` count as 1 and 0. |
| Aggregates | `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT` and the rest ignore text and booleans inside ranges and arrays, but coerce values typed directly as arguments (`SUM("x")` is `#VALUE!`). |
| Rounding | `ROUND` rounds halves away from zero and corrects binary representation error (`ROUND(1.005,2) = 1.01`). |
| Floating point | Arithmetic results are normalised to 15 significant digits, so `0.1+0.2=0.3` is `TRUE`. |
| Dates | The 1900 date system: serial 1 = 1900-01-01, including the historical fictitious 1900-02-29 (serial 60) so serial numbers match the established convention. Times are fractions of a day. |
| Criteria | `*IF`/`*IFS` criteria accept `">10"`, `"<>x"`, `"=apple"`, wildcards `*` `?` with `~` escapes, booleans and blanks (`""`, `"<>"`). Criteria arguments broadcast over arrays. |
| Lookup | `VLOOKUP`/`HLOOKUP`/`MATCH` default to approximate (sorted) matching. Exact match supports wildcards. `XLOOKUP`/`XMATCH` default to exact matching and support next-smaller/next-larger, wildcard and binary-search modes. |
| Dynamic arrays | Array results spill. A blocked spill area gives `#SPILL!`. Empty results give `#CALC!`. Operators and scalar functions broadcast (`=A1:A3*10`, `=LEN(A1:A3)`). |
| References | Relative, absolute and mixed references, whole columns and rows, cross-sheet references, named ranges, `INDIRECT` (A1 and R1C1), `OFFSET`, `INDEX` returning rows and columns. |
| Structure | Inserting or deleting rows and columns rewrites references on every sheet. References to deleted cells become `#REF!`. Renaming a sheet rewrites references to it, and deleting a sheet turns them into `#REF!`. |
| Errors | `#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#SPILL!`, `#CALC!` propagate through formulas and can be caught with `IFERROR`, `IFNA`, `ISERROR` and similar. |

## Documented differences

1. **Circular references** produce the dedicated error `#CIRC!` on every cell in
   the cycle (and on cells that depend on it), and the status bar shows a warning.
   Iterative calculation is not offered.
2. **Volatile functions** (`RAND`, `RANDBETWEEN`, `RANDARRAY`, `NOW`, `TODAY`,
   `INDIRECT`, `OFFSET`) recalculate on every edit. Because `INDIRECT` and `OFFSET`
   resolve references at run time, they are always recalculated rather than tracked
   in the dependency graph.
3. **`INDEX` used as a range endpoint** (for example `A1:INDEX(...)`) is not
   supported. `INDEX` returns values, whole rows and whole columns.
4. **Implicit intersection** is not applied. A range used where a single value is
   expected spills (or broadcasts) as a dynamic array.
5. **Cut and paste** moves cells and retargets formulas that point at the moved
   block, provided they refer to cells or ranges entirely inside it. Partially
   overlapping ranges are left unchanged.
6. **Whole-row/column formatting** is applied to the used area plus a margin, rather
   than to all 5 million rows.
7. **Fill handle with a single number** copies the value. Dates and text with a
   trailing number are extended as a series, and two or more numbers extend as a
   linear trend. Hold Alt while releasing the handle to force copying.
8. **`LET`** supports named values and references. Recursive or lambda-style
   functions are not available.
9. **Statistical distributions**: `NORM.DIST` uses an approximation of the error
   function accurate to about 1.5e-7.
10. **Workbook file format**: GRID saves its own `.grid` JSON format and
    imports and exports CSV/TSV. Other binary workbook formats are not read or written.

## Function coverage

Run `functionList()` from `src/engine/functions` (or open **Insert → Function**)
for the complete, searchable list with signatures and descriptions. Every
category requested in the product specification is covered and tested in
`tests/engine.test.ts`.
