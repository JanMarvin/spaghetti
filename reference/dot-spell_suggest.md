# Find the closest known function name(s) to an unknown token.

Uses generalised edit distance (utils::adist) over the
worksheet-function registry (LEGACY_WORKSHEET ∪ XLFN ∪ XLWS). XLM macro
names are excluded from suggestions since user typos in worksheet
contexts are almost never typos of XLM names.

## Usage

``` r
.spell_suggest(fn, n = 3L, max_dist = 3L)
```

## Arguments

- fn:

  Unknown function name (uppercase).

- n:

  Maximum number of suggestions to return (default 3).

- max_dist:

  Maximum edit distance to consider a useful suggestion (default 3).

## Value

Character vector of suggestions, possibly empty.

## Details

Returns up to `n` suggestions, or character(0) if the closest match is
more than `max_dist` edits away.
