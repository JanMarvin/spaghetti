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

#' Check a formula for unknown function names
#'
#' Tokenises a formula and reports any function names that are not in the
#' Excel function registry, with spelling suggestions for likely typos.
#'
#' Unlike `to_xml()`, this does not translate — it is a pure linting pass.
#'
#' @param formula  Character scalar or vector of Excel formulas.
#' @param locale   Locale code or NULL. When set, localised names are
#'                 translated before checking.
#'
#' @return A data frame with columns `formula`, `fn`, `suggestion`.
#'   `fn` is the unknown function name. `suggestion` is a comma-separated
#'   string of close matches, or `NA` if no suggestion was found.
#'   Returns an empty data frame (invisibly) if no issues are found.
#'
#' @export
#' @examples
#' check_formula("=SUIM(A1:A10)")          # typo: SUIM -> SUM
#' check_formula("=VLOKUP(A1,B:B,2,0)")    # typo: VLOKUP -> VLOOKUP
#' check_formula(c("=SUM(A1)", "=FLITER(A1:A10,B1:B10>0)"))
check_formula <- function(formula, locale = NULL) {
  stopifnot(is.character(formula))

  all_known <- c(
    .spaghetti_env$LEGACY,
    .spaghetti_env$XLFN,
    .spaghetti_env$XLWS
  )

  issues <- list()

  for (f in formula) {
    if (is.na(f) || !nzchar(f)) next
    local_sep <- .get_sep(locale)
    tokens    <- .tokenise(f, sep = local_sep)

    for (tok in tokens) {
      if (tok$type != TOKEN_TYPES$FUNC) next

      fn_raw <- .strip_prefix(tok$val)
      fn_en  <- toupper(.locale_to_english(fn_raw, locale))

      if (!fn_en %in% all_known) {
        suggestions <- .spell_suggest(fn_en)
        issues[[length(issues) + 1L]] <- data.frame(
          formula    = f,
          fn         = fn_en,
          suggestion = if (length(suggestions) > 0L)
            paste(suggestions, collapse = ", ")
          else NA_character_,
          stringsAsFactors = FALSE
        )
      }
    }
  }

  if (length(issues) == 0L) {
    message("No unknown functions found.")
    return(invisible(data.frame(
      formula = character(0), fn = character(0),
      suggestion = character(0), stringsAsFactors = FALSE
    )))
  }

  do.call(rbind, issues)
}
