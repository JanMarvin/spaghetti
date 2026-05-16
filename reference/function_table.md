# Return the full function translation table.

Columns: `fn` (English name), `description` (if available), and one
column per supported locale with the localised function name (NA if not
translated for that locale).

## Usage

``` r
function_table()
```

## Value

data.frame

## Details

If
[`setup_terminology()`](https://janmarvin.github.io/spaghetti/reference/setup_terminology.md)
hasn't been run, returns a single-column data frame with just `fn`.

## See also

[`setup_terminology()`](https://janmarvin.github.io/spaghetti/reference/setup_terminology.md),
[`supported_locales()`](https://janmarvin.github.io/spaghetti/reference/supported_locales.md)

## Examples

``` r
head(function_table())
#>         fn
#> 1      ABS
#> 2   ABSREF
#> 3  ACCRINT
#> 4 ACCRINTM
#> 5     ACOS
#> 6    ACOSH
```
