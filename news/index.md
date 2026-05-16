# Changelog

## spaghetti 0.3.0

### Breaking changes

- Locale translation data is no longer bundled with the package.
  Previously, `inst/extdata/excel_functions.rds` shipped a pre-built
  function-name translation table extracted from the Microsoft
  Terminology Collection. Microsoft does not publish an explicit license
  for the Terminology Collection contents, so this package now ships
  only the parser. Users invoke
  [`setup_terminology()`](https://janmarvin.github.io/spaghetti/reference/setup_terminology.md)
  once per machine to download (~100 MB) and parse the data into a
  per-user cache directory (`tools::R_user_dir("spaghetti", "data")`).

- [`to_xml()`](https://janmarvin.github.io/spaghetti/reference/to_xml.md)
  /
  [`from_xml()`](https://janmarvin.github.io/spaghetti/reference/from_xml.md)
  /
  [`check_formula()`](https://janmarvin.github.io/spaghetti/reference/check_formula.md)
  calls with a non-NULL `locale` argument will now error if the
  terminology cache has not been built. The error message directs the
  user to
  [`setup_terminology()`](https://janmarvin.github.io/spaghetti/reference/setup_terminology.md).
  Calls with `locale = NULL` (the default) work as before with no setup
  required.

- R version requirement bumped to R \>= 4.0.0 (for
  [`tools::R_user_dir`](https://rdrr.io/r/tools/userdir.html)).

### New features

- `setup_terminology(expected_sha256, force, workers, quiet)`: downloads
  the Microsoft Terminology Collection zip from Microsoft’s public
  download URL, verifies it against a known-good SHA-256 digest
  (overridable; mismatches warn but do not abort, since Microsoft can
  republish the file at any time), unzips, parses, and writes the cached
  RDS. Pass `expected_sha256 = NULL` to skip verification. Thin R
  wrapper around the parser script in `inst/extdata/parse_locales.R`.

- [`terminology_info()`](https://janmarvin.github.io/spaghetti/reference/terminology_info.md):
  returns the provenance attributes attached to the cached RDS at
  download time — source URL, observed SHA-256, timestamp, spaghetti
  version, cache path, function count, locale count. Useful for
  verifying which version of the Microsoft data is currently loaded.

- [`has_terminology()`](https://janmarvin.github.io/spaghetti/reference/has_terminology.md):
  returns TRUE if a terminology cache is loaded.

- [`clear_terminology()`](https://janmarvin.github.io/spaghetti/reference/clear_terminology.md):
  removes the cached RDS.

- `inst/extdata/parse_locales.R` gains a `download_and_parse_tbx()`
  function that the R wrapper invokes. The script can also be sourced
  manually
  (`source(system.file("extdata", "parse_locales.R", package = "spaghetti"))`)
  for users who want to drive the parser themselves.

### Internal

- `openxlsx2` and `digest` are declared in `Suggests`;
  [`setup_terminology()`](https://janmarvin.github.io/spaghetti/reference/setup_terminology.md)
  checks for them at runtime.

- `.onLoad()` looks for a cached RDS in the user data directory; if
  found it loads it silently, otherwise leaves the locale tables empty.

## spaghetti 0.2.0

### Breaking changes

- [`round_trip()`](https://janmarvin.github.io/spaghetti/reference/round_trip.md)
  now returns a list with elements `xml` and `formula`. The previous
  element name was `excel` (and earlier `locale`, which was a bug).
  Update any code reading `rt$excel` or `rt$locale` to `rt$formula`.

- The `LEGACY` function registry has been split into `LEGACY_WORKSHEET`
  (modern worksheet functions) and `LEGACY_XLM` (Excel 4 macro
  language). External code accessing `spaghetti:::.spaghetti_env$LEGACY`
  must use one of the new names.
  [`function_prefix()`](https://janmarvin.github.io/spaghetti/reference/function_prefix.md)
  continues to return `"legacy"` for both.

- The R/Zzz.R file was renamed to R/zzz.R to match standard convention
  and fix incorrect documentation about ASCII source order.

- IDENT tokens (LAMBDA / LET parameters, named ranges) are no longer
  passed through the locale lookup in
  [`from_xml()`](https://janmarvin.github.io/spaghetti/reference/from_xml.md).
  The previous behaviour would mis-translate a LAMBDA parameter named
  e.g. `sum` to the local function name for SUM. If you relied on
  locale-translation of bare identifiers, this is a breaking change.

### New features

- Lexer now distinguishes a dedicated `REF` token type from `IDENT`.
  Cell references — including sheet-qualified, quoted-sheet, 3D,
  external-workbook, and structured table refs — are emitted as a single
  REF token:

  - `Sheet1!A1`, `'My Sheet'!A1`
  - `Sheet1:Sheet5!A1` (3D range across sheets)
  - `[Book1]Sheet1!A1`, `[1]Sheet1!A1`, `'[Book1.xlsx]Sheet1'!A1`
  - `Table1[Col]`, `Table1[#Headers]`, `Table1[@Col]`,
    `Table1[[#All],[Col]]`

- `@FUNC(...)` now wraps as `_xlfn.SINGLE(FUNC(...))` instead of
  dropping the function call.

- Excel error literals (`#REF!`, `#N/A`, `#DIV/0!`, `#VALUE!`, `#NAME?`,
  `#NUM!`, `#NULL!`, `#GETTING_DATA`, `#SPILL!`, `#CALC!`, `#BLOCKED!`,
  `#FIELD!`, `#PYTHON!`, etc.) lex as opaque values rather than being
  fragmented as `#` (anchor) + identifier.

- Array literals `{...}` are lexed as a single opaque token. Internal
  column (`,`) and row (`;`) separators are no longer normalised by the
  outer locale-separator pass, so `=SUM({1,2;3,4};A1)` round-trips
  correctly under German locale.

- User-typed whitespace inside ranges (`A1 : B10`) is normalised to the
  compact form (`A1:B10`) in
  [`to_xml()`](https://janmarvin.github.io/spaghetti/reference/to_xml.md)
  output. The range-intersection operator (a space between two refs,
  `A1:B10 C5:D15`) is preserved.

- Locale codes can now be multi-segment (`de-DE`, `pt-BR`, `zh-Hans`,
  `zh-Hans-CN`). Lookup tries the full string first then falls back
  segment-by-segment (`zh-Hans-CN → zh-Hans → zh`).

- [`xlex()`](https://janmarvin.github.io/spaghetti/reference/xlex.md)
  now labels three new token kinds explicitly: `error` for error
  literals, `array` for `{...}` literals, and `intersection` for a space
  between two refs.

### Bug fixes

- LAMBDA / LET parameter detection no longer prefixes named ranges that
  appear in the body. The rule is: at depth==1 of the innermost
  LAMBDA/LET, an IDENT is a bound name iff its next token is `,`.

- Number lexer no longer consumes trailing `+`/`-` as part of a number.
  `1+2` and similar arithmetic now tokenises correctly. Scientific
  notation `1e-3`, `1.5E+2` still works.

- [`is_ooxml()`](https://janmarvin.github.io/spaghetti/reference/is_ooxml.md)
  regex now detects `_xlws.` in addition to `_xlfn.` and `_xlpm.`.

- [`.spell_suggest()`](https://janmarvin.github.io/spaghetti/reference/dot-spell_suggest.md)
  orders results by edit distance instead of returning matches in
  registry order.

- `TEXTJOIN` is no longer duplicated across the LEGACY and XLFN tiers.

- [`function_table()`](https://janmarvin.github.io/spaghetti/reference/function_table.md)
  no longer leaks the internal `term_id` column.

- [`from_xml()`](https://janmarvin.github.io/spaghetti/reference/from_xml.md)
  recursively transforms the inner tokens of `_xlfn.ANCHORARRAY(...)`
  and `_xlfn.SINGLE(...)`, so nested prefixed function calls inside the
  wrapper get their prefixes stripped and their names localised
  correctly.

- `parse_locales.R` no longer auto-executes its bottom-of-file call when
  the script is [`source()`](https://rdrr.io/r/base/source.html)d. Wrap
  it in `if (sys.nframe() == 0L && ...)`.

### Performance

- O(n²) list-growth eliminated from
  [`.tokenise()`](https://janmarvin.github.io/spaghetti/reference/dot-tokenise.md),
  `.transform_to_xml()`, and `.transform_from_xml()` by preallocating
  output lists and using amortised-growth indexing.

- Lexer switched from per-character `paste0(s, c)` accumulation to
  substring extraction, reducing character-level allocation in long
  formulas.

- Locale lookups in
  [`.locale_to_english()`](https://janmarvin.github.io/spaghetti/reference/dot-locale_to_english.md)
  /
  [`.english_to_locale()`](https://janmarvin.github.io/spaghetti/reference/dot-english_to_locale.md)
  are now O(1) hashed environment lookups instead of linear column scans
  with per-call [`toupper()`](https://rdrr.io/r/base/chartr.html).

- [`.spell_suggest()`](https://janmarvin.github.io/spaghetti/reference/dot-spell_suggest.md)
  and
  [`check_formula()`](https://janmarvin.github.io/spaghetti/reference/check_formula.md)
  use a cached `ALL_KNOWN` vector built once at `.onLoad()`.

- [`.get_sep()`](https://janmarvin.github.io/spaghetti/reference/dot-get_sep.md)
  in vapply hot loops is now lifted once outside the loop body in
  [`to_xml()`](https://janmarvin.github.io/spaghetti/reference/to_xml.md),
  [`from_xml()`](https://janmarvin.github.io/spaghetti/reference/from_xml.md),
  and
  [`check_formula()`](https://janmarvin.github.io/spaghetti/reference/check_formula.md).

### Internal

- Shared semicolon-locales list moved from a hardcoded constant in
  `translate.R` into `.spaghetti_env$SEMICOLON_LOCALES`.

- Per-locale lookup environments (`LOC_TO_EN`, `EN_TO_LOC`) built once
  at `.onLoad()` from the function table.

- `Imports: utils` narrowed to `importFrom(utils, adist)`.

## spaghetti 0.1.0

Initial release.
