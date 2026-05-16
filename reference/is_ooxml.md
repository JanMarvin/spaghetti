# Check whether a formula is already in OOXML storage format

A formula is considered "already OOXML" if it contains at least one
`_xlfn.`, `_xlws.`, or `_xlpm.` token.

## Usage

``` r
is_ooxml(formula)
```

## Arguments

- formula:

  Character scalar.

## Value

Logical.

## Examples

``` r
is_ooxml("=_xlfn.SEQUENCE(10)")   # TRUE
#> [1] TRUE
is_ooxml("=SEQUENCE(10)")          # FALSE
#> [1] FALSE
```
