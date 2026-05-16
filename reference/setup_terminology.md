# Download and parse the Microsoft Terminology Collection.

Source()s the parser script bundled in `inst/extdata/parse_locales.R`,
then calls its `download_and_parse_tbx()` function to fetch the zip from
Microsoft's public download URL, validate the SHA-256 (if you supply
one), unzip, parse, and write the resulting translation table to a
per-user cache directory.

## Usage

``` r
setup_terminology(
  expected_sha256 = .MTC_EXPECTED_SHA256,
  force = FALSE,
  workers = max(1L, parallel::detectCores() - 1L),
  quiet = FALSE
)
```

## Arguments

- expected_sha256:

  SHA-256 hex digest expected for the downloaded zip. Defaults to the
  digest of the version of the zip known to this release of spaghetti. A
  mismatch produces a warning (not an error), since Microsoft may
  republish the file. Pass `NULL` to skip the check entirely.

- force:

  If TRUE, re-download even if a cache exists.

- workers:

  Number of parallel TBX-parsing workers. Default detects cores - 1,
  capped to 8.

- quiet:

  If TRUE, suppress progress messages.

## Value

Invisibly, the path to the cached RDS.

## Details

Subsequent R sessions load the cache automatically on package attach.

## Dependencies

The parser uses the `openxlsx2` package (for XML parsing) and `digest`
(for SHA-256 verification). Both are declared in `Suggests:` and
installed only if you call this function.

## Licensing

Microsoft has not published an explicit license for the contents of the
Terminology Collection. The data is downloaded directly from Microsoft;
this package does not redistribute it.

## See also

[`has_terminology()`](https://janmarvin.github.io/spaghetti/reference/has_terminology.md),
[`clear_terminology()`](https://janmarvin.github.io/spaghetti/reference/clear_terminology.md),
[`terminology_info()`](https://janmarvin.github.io/spaghetti/reference/terminology_info.md)

## Examples

``` r
if (FALSE) { # \dontrun{
setup_terminology()
# Skip verification entirely:
setup_terminology(expected_sha256 = NULL)
} # }
```
