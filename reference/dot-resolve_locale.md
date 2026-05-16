# Resolve a user-supplied locale code to a column name in FUNCTIONS.

Tries an exact case-insensitive match first; failing that, splits on
'-'/'*' and tries the leading segment (so `de-DE` falls back to `de`,
`zh-Hans-CN` falls back to `zh-Hans` then `zh`). Returns NA_character*
if nothing matches.

## Usage

``` r
.resolve_locale(locale)
```
