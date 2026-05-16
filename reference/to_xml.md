# Convert a user-facing formula to OOXML storage format

Convert a user-facing formula to OOXML storage format

## Usage

``` r
to_xml(formula, locale = NULL, warn_unknown = TRUE)
```

## Arguments

- formula:

  Character scalar or vector. Formula(s), with or without `=`.

- locale:

  Two-letter locale code (`"de"`, `"fr"`, …) or NULL. When set,
  localised function names are translated to English and the locale
  argument separator (`;` for many European locales) is accepted.

- warn_unknown:

  Logical; warn for unknown function names (default TRUE).

## Value

Character scalar or vector: OOXML formula(s) starting with `=`.

## Examples

``` r
to_xml("=SEQUENCE(10)")
#> [1] "=_xlfn.SEQUENCE(10)"
to_xml("=LAMBDA(temp, (5/9) * (temp-32))(100)")
#> [1] "=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(100)"
to_xml("=FILTER(A1:A10, B1:B10 > 5)")
#> [1] "=_xlfn._xlws.FILTER(A1:A10, B1:B10 > 5)"
to_xml("=SUM(A1#)")
#> [1] "=SUM(_xlfn.ANCHORARRAY(A1))"
to_xml("=LET(tc,(B2-32)*5/9,rh,0.6,tc*ATAN(0.151977*(rh*100+8.313659)^0.5))")
#> [1] "=_xlfn.LET(_xlpm.tc,(B2-32)*5/9,_xlpm.rh,0.6,_xlpm.tc*ATAN(0.151977*(_xlpm.rh*100+8.313659)^0.5))"
if (FALSE) to_xml("=SUMMEWENN(A1:A10;\"x\";B1:B10)", locale = "de") # \dontrun{}
to_xml(c("=SUM(A1:A10)", "=SEQUENCE(5)", "=FILTER(A1:A10, B1:B10 > 0)"))
#> [1] "=SUM(A1:A10)"                           
#> [2] "=_xlfn.SEQUENCE(5)"                     
#> [3] "=_xlfn._xlws.FILTER(A1:A10, B1:B10 > 0)"
```
