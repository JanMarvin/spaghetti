# Round-trip a formula through OOXML and back

Converts to OOXML then back to user-facing form. Useful for testing
idempotency.

## Usage

``` r
round_trip(formula, locale = NULL, out_locale = NULL)
```

## Arguments

- formula:

  Character scalar.

- locale:

  Locale for input formula (passed to
  [`to_xml()`](https://janmarvin.github.io/spaghetti/reference/to_xml.md)).

- out_locale:

  Locale for output (passed to
  [`from_xml()`](https://janmarvin.github.io/spaghetti/reference/from_xml.md)).

## Value

Named list with `xml` (OOXML form) and `formula` (round-tripped
user-facing form).

## Examples

``` r
round_trip("=LAMBDA(x, x * 2)(5)")
#> $xml
#> [1] "=_xlfn.LAMBDA(_xlpm.x, _xlpm.x * 2)(5)"
#> 
#> $formula
#> [1] "=LAMBDA(x, x * 2)(5)"
#> 
```
