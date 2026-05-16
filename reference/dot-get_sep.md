# Return the formula argument separator for a given locale.

Locales using "," as their decimal separator use ";" in formulas. The
list lives in `.spaghetti_env$SEMICOLON_LOCALES` (see R/aaa.R).

## Usage

``` r
.get_sep(locale)
```

## Arguments

- locale:

  Locale code or NULL.

## Value

";" or ","
