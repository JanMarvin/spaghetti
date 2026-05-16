# R/setup.R
#
# Cache management for the Microsoft Terminology Collection translation
# table. The actual download/parse logic lives in
# inst/extdata/parse_locales.R — the parser is intentionally kept outside
# of R/ so users who don't need locale translation never pay for it
# (openxlsx2 etc. only become required if you actually run setup).
#
# To run setup, this file source()s the inst script and invokes its
# `download_and_parse_tbx()` function.

#' Cache directory for the terminology RDS.
#' @keywords internal
.terminology_dir <- function() {
  tools::R_user_dir("spaghetti", which = "data")
}

#' Path to the cached terminology RDS.
#' @keywords internal
.terminology_path <- function() {
  file.path(.terminology_dir(), "excel_functions.rds")
}

#' Path to the bundled parser script.
#' @keywords internal
.parser_script <- function() {
  system.file("extdata", "parse_locales.R", package = "spaghetti")
}

#' Check whether locale terminology has been loaded.
#'
#' @return Logical scalar.
#' @seealso [setup_terminology()], [clear_terminology()]
#' @export
#' @examples
#' has_terminology()
has_terminology <- function() {
  isTRUE(.spaghetti_env$has_terminology)
}

#' Known-good SHA-256 of the Microsoft Terminology Collection zip.
#' Recorded once; mismatches warn but do not abort, since Microsoft may
#' republish the zip without notice.
#' @keywords internal
.MTC_EXPECTED_SHA256 <- "fed7d16955fd4063731712704cbd6584869a9009acf8a9c3c3de31e0d4ebdfe6"

#' Download and parse the Microsoft Terminology Collection.
#'
#' Source()s the parser script bundled in `inst/extdata/parse_locales.R`,
#' then calls its `download_and_parse_tbx()` function to fetch the zip
#' from Microsoft's public download URL, validate the SHA-256 (if you
#' supply one), unzip, parse, and write the resulting translation table
#' to a per-user cache directory.
#'
#' Subsequent R sessions load the cache automatically on package attach.
#'
#' @section Dependencies:
#'   The parser uses the `openxlsx2` package (for XML parsing) and
#'   `digest` (for SHA-256 verification). Both are declared in
#'   `Suggests:` and installed only if you call this function.
#'
#' @section Licensing:
#'   Microsoft has not published an explicit license for the contents of
#'   the Terminology Collection. The data is downloaded directly from
#'   Microsoft; this package does not redistribute it.
#'
#' @param expected_sha256 SHA-256 hex digest expected for the downloaded
#'   zip. Defaults to the digest of the version of the zip known to this
#'   release of spaghetti. A mismatch produces a warning (not an error),
#'   since Microsoft may republish the file. Pass `NULL` to skip the
#'   check entirely.
#' @param force      If TRUE, re-download even if a cache exists.
#' @param workers    Number of parallel TBX-parsing workers. Default
#'                   detects cores - 1, capped to 8.
#' @param quiet      If TRUE, suppress progress messages.
#'
#' @return Invisibly, the path to the cached RDS.
#' @seealso [has_terminology()], [clear_terminology()], [terminology_info()]
#' @export
#' @examples
#' \dontrun{
#' setup_terminology()
#' # Skip verification entirely:
#' setup_terminology(expected_sha256 = NULL)
#' }
setup_terminology <- function(expected_sha256 = .MTC_EXPECTED_SHA256,
                              force      = FALSE,
                              workers    = max(1L, parallel::detectCores() - 1L),
                              quiet      = FALSE) {
  for (pkg in c("openxlsx2", "digest")) {
    if (!requireNamespace(pkg, quietly = TRUE)) {
      stop("Package '", pkg, "' is required for setup_terminology(). ",
           "Install it with:\n  install.packages(\"", pkg, "\")",
           call. = FALSE)
    }
  }

  script <- .parser_script()
  if (!nzchar(script) || !file.exists(script)) {
    stop("Bundled parser script not found at inst/extdata/parse_locales.R. ",
         "Reinstall the package.", call. = FALSE)
  }

  rds_path  <- .terminology_path()
  cache_dir <- .terminology_dir()
  if (!force && file.exists(rds_path)) {
    if (!quiet) message("Terminology cache exists at:\n  ", rds_path,
                        "\nUse force = TRUE to re-download.")
    .load_terminology_cache()
    return(invisible(rds_path))
  }

  dir.create(cache_dir, showWarnings = FALSE, recursive = TRUE)

  # Source the parser into a private environment so its top-level helpers
  # don't pollute globalenv. `local()` returns the environment via
  # envir=local_env trick.
  parser_env <- new.env(parent = globalenv())
  source(script, local = parser_env, echo = FALSE, verbose = FALSE)

  if (!exists("download_and_parse_tbx", envir = parser_env, inherits = FALSE)) {
    stop("Bundled parser script does not define download_and_parse_tbx(). ",
         "The installed package may be from an older version; please ",
         "reinstall.", call. = FALSE)
  }

  parser_env$download_and_parse_tbx(
    dest_rds        = rds_path,
    expected_sha256 = expected_sha256,
    workers         = workers,
    quiet           = quiet
  )

  .load_terminology_cache()
  invisible(rds_path)
}

#' Remove the cached terminology RDS.
#'
#' @return Invisibly TRUE if a cache was removed, FALSE if none existed.
#' @export
clear_terminology <- function() {
  rds_path <- .terminology_path()
  if (!file.exists(rds_path)) {
    .reset_terminology_state()
    return(invisible(FALSE))
  }
  unlink(rds_path)
  .reset_terminology_state()
  message("Terminology cache cleared.")
  invisible(TRUE)
}

#' Metadata about the currently loaded terminology cache.
#'
#' Returns the provenance attributes that were attached to the cached
#' RDS at download time: source URL, observed SHA-256, download timestamp,
#' and the spaghetti version that produced the cache. Returns `NULL` if
#' no terminology is currently loaded.
#'
#' @return A named list, or `NULL`.
#' @export
#' @examples
#' terminology_info()
terminology_info <- function() {
  df <- .spaghetti_env$FUNCTIONS
  if (is.null(df)) return(NULL)
  list(
    source_url        = attr(df, "source_url"),
    source_sha256     = attr(df, "source_sha256"),
    downloaded_at     = attr(df, "downloaded_at"),
    spaghetti_version = attr(df, "spaghetti_version"),
    cache_path        = .terminology_path(),
    n_functions       = nrow(df),
    n_locales         = length(supported_locales())
  )
}

#' Load the cached RDS into .spaghetti_env. Called from .onLoad() and
#' from setup_terminology() after writing the cache. Silent on missing
#' cache.
#' @keywords internal
.load_terminology_cache <- function() {
  rds_path <- .terminology_path()
  if (!file.exists(rds_path)) {
    .reset_terminology_state()
    return(invisible(FALSE))
  }
  df <- tryCatch(readRDS(rds_path), error = function(e) NULL)
  if (is.null(df) || !is.data.frame(df) || !"fn" %in% names(df)) {
    .reset_terminology_state()
    return(invisible(FALSE))
  }
  .spaghetti_env$FUNCTIONS       <- df
  .spaghetti_env$has_terminology <- TRUE
  .build_locale_lookup_tables(df)
  invisible(TRUE)
}

#' Reset locale state to "no terminology".
#' @keywords internal
.reset_terminology_state <- function() {
  .spaghetti_env$FUNCTIONS       <- NULL
  .spaghetti_env$has_terminology <- FALSE
  .spaghetti_env$LOC_COLS        <- character(0)
  .spaghetti_env$LOC_TO_EN       <- list()
  .spaghetti_env$EN_TO_LOC       <- list()
}

#' Build per-locale lookup environments from a FUNCTIONS data frame.
#' @keywords internal
#' @importFrom stats setNames
.build_locale_lookup_tables <- function(df) {
  loc_cols <- setdiff(names(df), c("fn", "description", "term_id"))
  loc_to_en <- list()
  en_to_loc <- list()
  for (col in loc_cols) {
    vals    <- df[[col]]
    fn_vals <- df$fn
    not_na  <- !is.na(vals)

    le <- new.env(hash = TRUE, parent = emptyenv(),
                  size = max(8L, sum(not_na)))
    el <- new.env(hash = TRUE, parent = emptyenv(),
                  size = max(8L, sum(not_na)))

    keys_loc <- toupper(vals[not_na])
    keys_en  <- toupper(fn_vals[not_na])
    vals_en  <- fn_vals[not_na]
    vals_loc <- vals[not_na]
    for (i in seq_along(keys_loc)) {
      assign(keys_loc[i], vals_en[i],  envir = le)
      assign(keys_en[i],  vals_loc[i], envir = el)
    }
    loc_to_en[[col]] <- le
    en_to_loc[[col]] <- el
  }
  .spaghetti_env$LOC_COLS  <- setNames(loc_cols, tolower(loc_cols))
  .spaghetti_env$LOC_TO_EN <- loc_to_en
  .spaghetti_env$EN_TO_LOC <- en_to_loc
}

#' Stop with a clear message when locale work is requested but no cache.
#' @keywords internal
.stop_no_terminology <- function() {
  stop(
    "Locale translation requested, but the Microsoft Terminology ",
    "Collection has not been downloaded.\n",
    "Run `spaghetti::setup_terminology()` once to download and cache it.\n",
    "(One-time setup: ~100 MB download, 1-5 minute parse.)",
    call. = FALSE
  )
}
