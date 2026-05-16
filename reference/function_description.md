# Look up the English description for a function.

Reads from the cached Microsoft Terminology Collection table; available
only after
[`setup_terminology()`](https://janmarvin.github.io/spaghetti/reference/setup_terminology.md)
has been run. Returns NA for every input if the cache is missing.

## Usage

``` r
function_description(fn)
```

## Arguments

- fn:

  Character vector of function names (English, any case).

## Value

Character vector of descriptions (NA where not found / available).

## Examples

``` r
function_description(c("SUM", "XLOOKUP", "LAMBDA"))
#> [1] NA NA NA
```
