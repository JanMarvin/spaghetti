
<!-- README.md is generated from README.Rmd. Please edit that file -->

# spaghetti

Bidirectional translator between spreadsheet user-facing formulas (as
displayed to the end user) and the OOXML storage format (as found inside
`.xlsx` XML).

**Key features:**

- Automatic prefix injection (`_xlfn.`, `_xlws.`, `_xlpm.`) for
  modern (post-2010) functions
- Spill operator (`#`) and implicit intersection (`@`) translation,
  including for function calls (`@SUM(...)` → `_xlfn.SINGLE(SUM(...))`)
- LAMBDA and LET parameter scoping with `_xlpm.` prefixes — the
  parameter / body distinction is detected correctly so named ranges
  inside the body are not mistakenly prefixed
- Full lex of sheet-qualified, 3D, quoted-sheet, external-workbook, and
  structured table references — including spill (`#`) and intersection
  (`@`) on table refs
- Error literals (`#REF!`, `#N/A`, …) and array literals
  (`{1,2;3,4}`) round-trip cleanly across locale conversion
- Localised formula support (translate between languages) with
  multi-segment locale fallback (`de-DE` → `de`, `zh-Hans-CN` →
  `zh-Hans` → `zh`)
- Vectorised — works with single formulas or character vectors
- Formula linting with spelling suggestions for typos

## Why does this exist?

When you write a formula with `openxlsx2::wb_add_formula()`, the string
goes directly into the XML. Modern functions like `LAMBDA`, `XLOOKUP`,
`FILTER`, or `SEQUENCE` need namespace prefixes in storage that you
never type as a user:

| You type                    | XML must contain                        |
|-----------------------------|-----------------------------------------|
| `=SEQUENCE(10)`             | `=_xlfn.SEQUENCE(10)`                   |
| `=FILTER(A1:A10, B1:B10>5)` | `=_xlfn._xlws.FILTER(A1:A10, B1:B10>5)` |
| `=LAMBDA(x, x * 2)`         | `=_xlfn.LAMBDA(_xlpm.x, _xlpm.x * 2)`   |
| `=SUM(A1#)`                 | `=SUM(_xlfn.ANCHORARRAY(A1))`           |
| `=@SUM(A1:A10)`             | `=_xlfn.SINGLE(SUM(A1:A10))`            |

Without the prefixes the formula opens as `#NAME?`. `spaghetti` handles
the translation so you don't have to remember the rules.

## Installation

``` r
# install.packages("remotes")
remotes::install_github("JanMarvin/spaghetti")
```

## Quick start

``` r
library(spaghetti)

# Single formula
to_xml("=SEQUENCE(10)")
#> [1] "=_xlfn.SEQUENCE(10)"

# Vector of formulas — same function
to_xml(c(
  "=SEQUENCE(10)",
  "=FILTER(A1:A10, B1:B10 > 5)",
  "=SUM(A1#)"
))
#> [1] "=_xlfn.SEQUENCE(10)"
#> [2] "=_xlfn._xlws.FILTER(A1:A10, B1:B10 > 5)"
#> [3] "=SUM(_xlfn.ANCHORARRAY(A1))"

# Round-trip with localisation
from_xml("=_xlfn.SEQUENCE(10)", locale = "de")
#> [1] "=SEQUENZ(10)"
```

## Usage with openxlsx2

### Example 1 — Dynamic array formulas (SEQUENCE, UNIQUE, XLOOKUP)

The most common pain point: dynamic-array functions that return spilled
arrays. Write them in plain user-facing syntax; `to_xml()` adds the
required prefixes before handing off to openxlsx2.

``` r
library(spaghetti)
library(openxlsx2)

products <- data.frame(
  id    = c(3, 1, 2, 1, 3, 2),
  name  = c("Apple", "Banana", "Cherry", "Banana", "Apple", "Cherry"),
  sales = c(120, 85, 200, 95, 110, 175)
)

wb <- wb_workbook() |>
  wb_add_worksheet("Data") |>
  wb_add_data(x = products, dims = "A1")

# UNIQUE + SORT: deduplicated sorted product list spilling from E2
wb <- wb |>
  wb_add_data(dims = "E1", x = "Unique Products") |>
  wb_add_formula(
    dims = "E2",
    x = to_xml("=UNIQUE(SORT(B2:B7))"),
    cm = TRUE
  )

# XLOOKUP: look up total sales for each unique product using the spill ref
wb <- wb |>
  wb_add_data(dims = "F1", x = "Total Sales") |>
  wb_add_formula(
    dims = "F2",
    x = to_xml("=XLOOKUP(E2#, B2:B7, C2:C7, 0, 0)"),
    cm = TRUE
  )

if (interactive()) wb$open()
```

### Example 2 — LAMBDA and LET for reusable logic

`LAMBDA` parameters need an additional `_xlpm.` prefix inside the XML —
something that's nearly impossible to remember manually. spaghetti
detects which identifiers are bound parameters vs. body expressions (so
a named range inside the body is *not* prefixed).

``` r
library(spaghetti)
library(openxlsx2)

temps_f <- data.frame(city = c("Berlin", "Paris", "Oslo"),
                      temp_f = c(35.6, 42.8, 28.4))

wb <- wb_workbook() |>
  wb_add_worksheet("Weather") |>
  wb_add_data(x = temps_f, dims = "A1")

wb <- wb_add_data(wb, x = "Temp °C", dims = "C1")

formulas_lambda <- to_xml(
  sprintf("=LAMBDA(f, (f - 32) * 5/9)(B%s)", 2:4)
)

wb <- wb_add_formula(wb, dims = "C2", x = formulas_lambda, cm = TRUE)

# LET: named intermediate variables for a more complex calculation
let <- paste0(
  "=LET(",
  "tc, (B%s-32)*5/9, ",        # Celsius
  "rh, 0.6, ",                 # assume 60% relative humidity
  "tc * ATAN(0.151977 * (rh * 100 + 8.313659) ^ 0.5)",
  ")"
)
formulas_let <- to_xml(sprintf(let, 2:4))

wb <- wb_add_data(wb, x = "Wet Bulb °C", dims = "D1") |>
  wb_add_formula(dims = "D2", x = formulas_let, cm = TRUE)

if (interactive()) wb$open()
```

### Example 3 — Reading back OOXML formulas and translating for display

When you load an existing `.xlsx` with `wb_to_df(show_formula = TRUE)`,
the formulas come back in OOXML storage format with all prefixes intact.
`from_xml()` strips them back to readable form — useful for auditing,
diffing, or displaying formulas to users.

``` r
library(spaghetti)
library(openxlsx2)

wb <- wb_workbook() |>
  wb_add_worksheet("Report") |>
  wb_add_data(x = data.frame(x = 1:10, y = rnorm(10)), dims = "A1") |>
  wb_add_formula(dims = "C2", x = to_xml("=SEQUENCE(5)"), cm = TRUE) |>
  wb_add_formula(dims = "D2", x = to_xml("=XLOOKUP(C2#, A2:A11, B2:B11)"), cm = TRUE)

raw_formulas <- wb_to_df(wb, dims = "C2:D2", show_formula = TRUE, col_names = FALSE)

# The formulas come back with prefixes — not user-friendly
# C2: =_xlfn.SEQUENCE(5)
# D2: =_xlfn.XLOOKUP(_xlfn.ANCHORARRAY(C2), A2:A11, B2:B11)

readable <- as.data.frame(lapply(raw_formulas, function(col) {
  ifelse(is_ooxml(col), from_xml(col), col)
}))

# C2: =SEQUENCE(5)
# D2: =XLOOKUP(C2#, A2:A11, B2:B11)
print(readable)
#>              C                             D
#> 1 =SEQUENCE(5) =XLOOKUP(C2#, A2:A11, B2:B11)

# Translate to German for a localised formula audit report
readable_de <- as.data.frame(lapply(raw_formulas, function(col) {
  ifelse(is_ooxml(col), from_xml(col, locale = "de"), col)
}))
# D2: =XVERWEIS(C2#; A2:A11; B2:B11)
print(readable_de)
#>             C                              D
#> 1 =SEQUENZ(5) =XVERWEIS(C2#; A2:A11; B2:B11)
```

## Core functions

| Function                          | Direction       | Description                                                                                          |
|-----------------------------------|-----------------|------------------------------------------------------------------------------------------------------|
| `to_xml(formula, locale)`         | user → OOXML    | Add `_xlfn.`, `_xlws.`, `_xlpm.` prefixes. Single string or character vector.                        |
| `from_xml(formula, locale)`       | OOXML → user    | Strip all prefixes. Single string or character vector.                                               |
| `round_trip(formula, …)`          | both            | Convert and back-convert. Returns `list(xml = …, formula = …)`.                                      |
| `is_ooxml(formula)`               | —               | Detect whether a formula is already prefixed (any of `_xlfn.`, `_xlws.`, `_xlpm.`).                  |
| `function_prefix(fn)`             | —               | Check which tier a function name falls into (`legacy`, `xlfn`, `xlws`).                              |
| `supported_locales()`             | —               | List available locale codes.                                                                         |
| `function_description(fn)`        | —               | English description for a function name.                                                             |
| `function_table()`                | —               | Full data frame of `fn`, `description`, and per-locale translated names.                             |
| `check_formula(formula, locale)`  | —               | Lint for unknown function names with spelling suggestions.                                           |
| `xlex(formula, locale)`           | —               | Tokenise and pretty-print the parse tree (functions, refs, separators, error/array literals, etc.). |

## Localisation

`openxlsx2` only accepts English function names. If you have formulas
authored in a localised application (e.g. German `SUMMEWENN` instead of
`SUMIF`), pass the `locale` argument to `to_xml()` to translate first.

``` r
to_xml("=SUMMEWENNS(C2:C10; A2:A10; \"Berlin\")", locale = "de")
#> [1] "=SUMIFS(C2:C10, A2:A10, \"Berlin\")"

round_trip("=SUMMEWENNS(C2:C10; A2:A10; \"Berlin\")", locale = "de", out_locale = "es")
#> $xml
#> [1] "=SUMIFS(C2:C10, A2:A10, \"Berlin\")"
#>
#> $formula
#> [1] "=SUMAR.SI.CONJUNTO(C2:C10; A2:A10; \"Berlin\")"
```

Locale codes are matched case-insensitively and multi-segment codes
fall back: `de-DE` → `de`, `zh-Hans-CN` → `zh-Hans` → `zh`.

Supported locales depend on the bundled function table. Use
`supported_locales()` to list them.

## Formula linting

`check_formula()` validates function names and suggests corrections for
typos:

``` r
check_formula("=SUIM(A1:A10)")
#>         formula   fn     suggestion
#> 1 =SUIM(A1:A10) SUIM DSUM, SIN, SUM

check_formula(c(
  "=VLOKUP(A1, B:C, 2, 0)",
  "=FLITER(A1:A10, B1:B10 > 0)"
))
#>                       formula     fn               suggestion
#> 1      =VLOKUP(A1, B:C, 2, 0) VLOKUP HLOOKUP, LOOKUP, VLOOKUP
#> 2 =FLITER(A1:A10, B1:B10 > 0) FLITER     FILES, FISHER, FIXED

# Works with localised formulas too
check_formula("=SUMMEWENNS(C2:C10; A2:A10; \"Berlin\")", locale = "de")
#> No unknown functions found.
```

## Reference coverage

The lexer recognises the following reference shapes as single REF
tokens, so spill (`#`) and implicit intersection (`@`) operators apply
to the whole reference:

| Form                                | Example                            |
|-------------------------------------|------------------------------------|
| Cell                                | `A1`, `$A$1`                       |
| Cell range                          | `A1:B10`, `$A$1:$B$10`             |
| Whole column / row range            | `A:Z`, `1:10`                      |
| Sheet-qualified                     | `Sheet1!A1`                        |
| Quoted sheet (spaces, punctuation)  | `'My Sheet'!A1`                    |
| 3D range across sheets              | `Sheet1:Sheet5!A1`                 |
| External workbook                   | `[Book1]Sheet1!A1`, `[1]Sheet1!A1` |
| Quoted external                     | `'[Book1.xlsx]Sheet1'!A1`          |
| Structured table column             | `Table1[Col1]`                     |
| Structured table special items      | `Table1[#Headers]`, `Table1[@Col]` |
| Nested structured                   | `Table1[[#All],[Col1]:[Col2]]`     |

User-typed whitespace inside a range (`A1 : B10`) is normalised to the
compact form on `to_xml()`. The range-intersection operator (a literal
space between two refs) is preserved.
