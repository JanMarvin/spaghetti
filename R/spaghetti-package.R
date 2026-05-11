#' spaghetti: Bidirectional Excel Formula to OOXML Translator
#'
#' @description
#' Translates Excel formulas between the user-facing format (as seen in the
#' formula bar) and the OOXML storage format (as found in `.xlsx` XML source).
#'
#' ## Key transformations
#'
#' | User-facing | OOXML storage | Rule |
#' |---|---|---|
#' | `SEQUENCE(10)` | `_xlfn.SEQUENCE(10)` | Future function prefix |
#' | `FILTER(A1:A10,...)` | `_xlfn._xlws.FILTER(A1:A10,...)` | Web-service namespace |
#' | `LAMBDA(x, x*2)` | `_xlfn.LAMBDA(_xlpm.x, _xlpm.x*2)` | Lambda param prefix |
#' | `A1#` | `_xlfn.ANCHORARRAY(A1)` | Spilled range operator |
#' | `@A1:A10` | `_xlfn.SINGLE(A1:A10)` | Implicit intersection |
#'
#' ## Main functions
#'
#' - [to_xml()]: Excel formula bar → OOXML storage
#' - [from_xml()]: OOXML storage → Excel formula bar
#' - [function_prefix()]: inspect prefix tier for any function name
#' - [supported_locales()]: list available locale codes
#' - [round_trip()]: convert and back-convert for testing
#'
#' ## Localisation
#'
#' Excel ships with translated function names in non-English locales.
#' Pass `locale = "de"` (German), `"fr"` (French), `"es"` (Spanish), etc.
#' to [to_xml()] to translate local names to English before prefixing,
#' or to [from_xml()] to output localised names.
#'
#' Supported locales: `r paste(supported_locales(), collapse = ", ")`.
#'
#' @references
#' - ECMA-376 Part 1 §18.17: Built-in function list
#' - XlsxWriter: <https://xlsxwriter.readthedocs.io/working_with_formulas.html>
#' - EPPlus wiki: <https://github.com/EPPlusSoftware/EPPlus/wiki/Function-prefixes>
#' - libxlsxwriter LAMBDA: <https://libxlsxwriter.github.io/lambda_8c-example.html>
#'
#' @importFrom utils adist
#' @keywords internal
"_PACKAGE"
