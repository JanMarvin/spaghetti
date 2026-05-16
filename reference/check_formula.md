# Check a formula for unknown function names

Tokenises a formula and reports any function names that are not in the
function registry, with spelling suggestions for likely typos.

## Usage

``` r
check_formula(formula, locale = NULL)
```

## Arguments

- formula:

  Character scalar or vector of formulas.

- locale:

  Locale code or NULL. When set, localised names are translated before
  checking.

## Value

A data frame with columns `formula`, `fn`, `suggestion`. `fn` is the
unknown function name. `suggestion` is a comma-separated string of close
matches, or `NA` if no suggestion was found. Returns an empty data frame
(invisibly) if no issues are found.

## Details

Unlike
[`to_xml()`](https://janmarvin.github.io/spaghetti/reference/to_xml.md),
this does not translate — it is a pure linting pass.

## Examples

``` r
check_formula("=SUIM(A1:A10)")          # typo: SUIM -> SUM
#>         formula   fn     suggestion
#> 1 =SUIM(A1:A10) SUIM SUM, DSUM, SIN
check_formula("=VLOKUP(A1,B:B,2,0)")    # typo: VLOKUP -> VLOOKUP
#>               formula     fn               suggestion
#> 1 =VLOKUP(A1,B:B,2,0) VLOKUP VLOOKUP, HLOOKUP, LOOKUP
check_formula(c("=SUM(A1)", "=FLITER(A1:A10,B1:B10>0)"))
#>                    formula     fn            suggestion
#> 1 =FLITER(A1:A10,B1:B10>0) FLITER FILTER, FISHER, FIXED
```
