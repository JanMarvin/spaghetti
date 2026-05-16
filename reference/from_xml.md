# Convert an OOXML storage formula to user-facing format

Convert an OOXML storage formula to user-facing format

## Usage

``` r
from_xml(formula, locale = NULL)
```

## Arguments

- formula:

  Character scalar or vector. OOXML formula(s), with or without `=`.

- locale:

  Two-letter locale code or NULL. When set, function names are
  translated to the target locale and the locale separator is used in
  output.

## Value

Character scalar or vector: user-facing formula(s) starting with `=`.

## Examples

``` r
from_xml("=_xlfn.SEQUENCE(10)")
#> [1] "=SEQUENCE(10)"
from_xml("=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(100)")
#> [1] "=LAMBDA(temp, (5/9) * (temp-32))(100)"
from_xml("=_xlfn._xlws.FILTER(A1:A10,B1:B10>5)")
#> [1] "=FILTER(A1:A10,B1:B10>5)"
from_xml("=SUM(_xlfn.ANCHORARRAY(A1))")
#> [1] "=SUM(A1#)"
if (FALSE) from_xml("=_xlfn.SEQUENCE(10)", locale = "de") # \dontrun{}
from_xml(c("=_xlfn.SEQUENCE(5)", "=SUM(_xlfn.ANCHORARRAY(A1))"))
#> [1] "=SEQUENCE(5)" "=SUM(A1#)"   
```
