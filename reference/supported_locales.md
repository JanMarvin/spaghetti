# List supported locale codes.

Returns the locale columns present in the cached function table, i.e.
the locales for which at least partial translation data was loaded.
Returns `character(0)` if
[`setup_terminology()`](https://janmarvin.github.io/spaghetti/reference/setup_terminology.md)
hasn't been run.

## Usage

``` r
supported_locales()
```

## Value

Character vector of locale codes (e.g. `c("de", "fr", "es", …)`).

## See also

[`setup_terminology()`](https://janmarvin.github.io/spaghetti/reference/setup_terminology.md),
[`has_terminology()`](https://janmarvin.github.io/spaghetti/reference/has_terminology.md)

## Examples

``` r
supported_locales()
#> character(0)
```
