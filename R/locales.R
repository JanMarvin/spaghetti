#' Translate a function name from a locale to English
#'
#' @param fn Function name string.
#' @param locale Two-letter locale code (e.g. "de", "fr"). NULL for auto-detect.
#' @return English function name, or fn unchanged if not found.
#' @keywords internal
.locale_to_english <- function(fn, locale = NULL) {
  fn_up <- toupper(fn)
  if (!is.null(locale) && locale %in% names(.spaghetti_env$LOCALES)) {
    m <- .spaghetti_env$LOCALES[[locale]]
    if (fn_up %in% names(m)) return(m[[fn_up]])
    return(fn)
  }
  # Auto-detect: search all locales
  for (m in .spaghetti_env$LOCALES) {
    if (fn_up %in% names(m)) return(m[[fn_up]])
  }
  fn
}

#' Translate a function name from English to a locale
#'
#' @param fn English function name.
#' @param locale Two-letter locale code.
#' @return Localised name, or fn if not found.
#' @keywords internal
.english_to_locale <- function(fn, locale) {
  if (is.null(locale) || !locale %in% names(.spaghetti_env$LOCALES_REV)) return(fn)
  fn_up <- toupper(fn)
  m <- .spaghetti_env$LOCALES_REV[[locale]]
  if (fn_up %in% names(m)) return(m[[fn_up]])
  fn
}

#' List supported locales
#'
#' @return Character vector of supported locale codes.
#' @export
#' @examples
#' supported_locales()
supported_locales <- function() {
  names(.spaghetti_env$LOCALES)
}
