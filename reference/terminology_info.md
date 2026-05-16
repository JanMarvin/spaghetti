# Metadata about the currently loaded terminology cache.

Returns the provenance attributes that were attached to the cached RDS
at download time: source URL, observed SHA-256, download timestamp, and
the spaghetti version that produced the cache. Returns `NULL` if no
terminology is currently loaded.

## Usage

``` r
terminology_info()
```

## Value

A named list, or `NULL`.

## Examples

``` r
terminology_info()
#> NULL
```
