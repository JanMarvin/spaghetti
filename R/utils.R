# R/utils.R

#' Translate a vector of Excel formulas to OOXML storage format
#'
#' Vectorised wrapper around [to_xml()].
#'
#' @param formulas Character vector of Excel formulas.
#' @param locale   Locale code, or NULL. Applied to all formulas.
#' @param warn_unknown Logical.
#'
#' @return Character vector of OOXML formulas.
#' @export
#' @examples
#' to_xml_v(c("=SUM(A1:A10)", "=SEQUENCE(5)", "=FILTER(A1:A10, B1:B10 > 0)"))
to_xml_v <- function(formulas, locale = NULL, warn_unknown = TRUE) {
  vapply(formulas, to_xml, character(1),
         locale = locale, warn_unknown = warn_unknown,
         USE.NAMES = FALSE)
}

#' Translate a vector of OOXML formulas to user-facing Excel format
#'
#' Vectorised wrapper around [from_xml()].
#'
#' @param formulas Character vector of OOXML formulas.
#' @param locale   Target locale code, or NULL.
#'
#' @return Character vector of Excel formulas.
#' @export
#' @examples
#' from_xml_v(c("=_xlfn.SEQUENCE(5)", "=SUM(_xlfn.ANCHORARRAY(A1))"))
from_xml_v <- function(formulas, locale = NULL) {
  vapply(formulas, from_xml, character(1), locale = locale, USE.NAMES = FALSE)
}

#' Identify what prefix a function will receive
#'
#' Useful for inspecting the registry without running a full translation.
#'
#' @param fn Character vector of function names.
#' @return Character vector: `"legacy"`, `"xlfn"`, or `"xlws"`.
#' @export
#' @examples
#' function_prefix(c("SUM", "SEQUENCE", "FILTER", "LAMBDA", "XLOOKUP"))
function_prefix <- function(fn) {
  vapply(toupper(fn), .prefix_for, character(1), USE.NAMES = FALSE)
}

#' Check whether a formula is already in OOXML storage format
#'
#' A formula is considered "already OOXML" if it contains at least one
#' `_xlfn.` or `_xlpm.` token.
#'
#' @param formula Character scalar.
#' @return Logical.
#' @export
#' @examples
#' is_ooxml("=_xlfn.SEQUENCE(10)")   # TRUE
#' is_ooxml("=SEQUENCE(10)")          # FALSE
is_ooxml <- function(formula) {
  grepl("_xlfn\\.|_xlpm\\.", formula, fixed = FALSE)
}

#' Round-trip a formula through OOXML and back
#'
#' Converts to OOXML then back to Excel. Useful for testing idempotency.
#'
#' @param formula  Character scalar.
#' @param locale   Locale for input formula (passed to [to_xml()]).
#' @param out_locale Locale for output (passed to [from_xml()]).
#'
#' @return Named list with `xml` (OOXML) and `excel` (round-tripped) formulas.
#' @export
#' @examples
#' round_trip("=LAMBDA(x, x * 2)(5)")
round_trip <- function(formula, locale = NULL, out_locale = NULL) {
  xml   <- to_xml(formula, locale = locale)
  locale <- from_xml(xml,   locale = out_locale)
  list(xml = xml, locale = locale)
}
