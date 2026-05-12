# R/locales.R
#
# Locale lookup helpers.
#
# Source of truth: .spaghetti_env$FUNCTIONS — a data frame with columns
#   fn        : English function name (uppercase)
#   <locale>  : localised names per supported locale (NA where not
#               translated). Column names are the locale codes detected
#               in the TBX files; case-insensitive matched on lookup.
#
# .onLoad() populates per-locale hash environments for O(1) lookup:
#   .spaghetti_env$LOC_COLS   : map tolower(code) -> actual column name
#   .spaghetti_env$LOC_TO_EN  : list, per locale: env(LOCAL_UP = "EN", ...)
#   .spaghetti_env$EN_TO_LOC  : list, per locale: env(EN_UP = "Local", ...)
# Both directions use *uppercase* keys; lookups upper-case the input.
#
# If the user has not run setup_terminology(), .spaghetti_env$FUNCTIONS
# is NULL and any call with a non-NULL locale errors out with a clear
# message directing the user to setup_terminology().

#' Resolve a user-supplied locale code to a column name in FUNCTIONS.
#'
#' Tries an exact case-insensitive match first; failing that, splits on
#' '-'/'_' and tries the leading segment (so `de-DE` falls back to `de`,
#' `zh-Hans-CN` falls back to `zh-Hans` then `zh`). Returns NA_character_
#' if nothing matches.
#' @keywords internal
.resolve_locale <- function(locale) {
  if (is.null(locale)) return(NA_character_)
  cols <- .spaghetti_env$LOC_COLS
  if (is.null(cols) || length(cols) == 0L) return(NA_character_)

  key <- tolower(locale)
  if (key %in% names(cols)) return(cols[[key]])

  parts <- strsplit(key, "[-_]")[[1]]
  if (length(parts) > 1L) {
    for (i in seq.int(length(parts) - 1L, 1L)) {
      sub_key <- paste(parts[seq_len(i)], collapse = "-")
      if (sub_key %in% names(cols)) return(cols[[sub_key]])
    }
  }
  NA_character_
}

#' Translate a function name from a locale to English.
#'
#' @param fn_name Function name string (any case).
#' @param locale  Locale code or NULL (returns fn unchanged).
#' @return English function name, or fn unchanged if no mapping found.
#' @keywords internal
.locale_to_english <- function(fn_name, locale) {
  if (is.null(locale)) return(fn_name)
  if (!has_terminology()) .stop_no_terminology()
  if (is.na(fn_name) || !nzchar(fn_name)) return(fn_name)
  col <- .resolve_locale(locale)
  if (is.na(col)) return(fn_name)
  tbl <- .spaghetti_env$LOC_TO_EN[[col]]
  if (is.null(tbl)) return(fn_name)
  hit <- tbl[[toupper(fn_name)]]
  if (is.null(hit)) fn_name else hit
}

#' Translate a function name from English to a locale.
#'
#' @param fn_en  English function name (any case).
#' @param locale Locale code or NULL (returns fn unchanged).
#' @return Localised name, or fn unchanged if no mapping found.
#' @keywords internal
.english_to_locale <- function(fn_en, locale) {
  if (is.null(locale)) return(fn_en)
  if (!has_terminology()) .stop_no_terminology()
  if (is.na(fn_en) || !nzchar(fn_en)) return(fn_en)
  col <- .resolve_locale(locale)
  if (is.na(col)) return(fn_en)
  tbl <- .spaghetti_env$EN_TO_LOC[[col]]
  if (is.null(tbl)) return(fn_en)
  hit <- tbl[[toupper(fn_en)]]
  if (is.null(hit) || is.na(hit)) fn_en else hit
}

#' List supported locale codes.
#'
#' Returns the locale columns present in the cached function table, i.e.
#' the locales for which at least partial translation data was loaded.
#' Returns `character(0)` if `setup_terminology()` hasn't been run.
#'
#' @return Character vector of locale codes (e.g. `c("de", "fr", "es", …)`).
#' @seealso [setup_terminology()], [has_terminology()]
#' @export
#' @examples
#' supported_locales()
supported_locales <- function() {
  df <- .spaghetti_env$FUNCTIONS
  if (is.null(df)) return(character(0))
  setdiff(names(df), c("fn", "description", "term_id"))
}

#' Look up the English description for a function.
#'
#' Reads from the cached Microsoft Terminology Collection table; available
#' only after `setup_terminology()` has been run. Returns NA for every
#' input if the cache is missing.
#'
#' @param fn Character vector of function names (English, any case).
#' @return Character vector of descriptions (NA where not found / available).
#' @export
#' @examples
#' function_description(c("SUM", "XLOOKUP", "LAMBDA"))
function_description <- function(fn) {
  df <- .spaghetti_env$FUNCTIONS
  if (is.null(df) || !"description" %in% names(df))
    return(rep(NA_character_, length(fn)))
  idx <- match(toupper(fn), df$fn)
  df$description[idx]
}

#' Return the full function translation table.
#'
#' Columns: `fn` (English name), `description` (if available), and one
#' column per supported locale with the localised function name (NA if
#' not translated for that locale).
#'
#' If `setup_terminology()` hasn't been run, returns a single-column
#' data frame with just `fn`.
#'
#' @return data.frame
#' @seealso [setup_terminology()], [supported_locales()]
#' @export
#' @examples
#' head(function_table())
function_table <- function() {
  df <- .spaghetti_env$FUNCTIONS
  if (is.null(df)) {
    return(data.frame(
      fn = sort(unique(c(.spaghetti_env$LEGACY_WORKSHEET,
                         .spaghetti_env$LEGACY_XLM,
                         .spaghetti_env$XLFN,
                         .spaghetti_env$XLWS))),
      stringsAsFactors = FALSE
    ))
  }
  # Keep description; only hide term_id (internal bookkeeping).
  df[, setdiff(names(df), "term_id"), drop = FALSE]
}
