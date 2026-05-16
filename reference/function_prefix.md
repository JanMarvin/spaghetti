# Identify what prefix a function will receive

Useful for inspecting the registry without running a full translation.

## Usage

``` r
function_prefix(fn)
```

## Arguments

- fn:

  Character vector of function names.

## Value

Character vector: `"legacy"`, `"xlfn"`, or `"xlws"`.

## Examples

``` r
function_prefix(c("SUM", "SEQUENCE", "FILTER", "LAMBDA", "XLOOKUP"))
#> [1] "legacy" "xlfn"   "xlws"   "xlfn"   "xlfn"  
```
