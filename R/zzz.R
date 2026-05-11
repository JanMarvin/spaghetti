# R/zzz.R
# Package initialisation.
#
# R sources files in R/ in ASCII order (locale-independent under
# R CMD INSTALL). Top-level assignments in aaa.R and registry.R run at
# install/load time before .onLoad() is invoked, so the ordering of
# function definitions across files is not load-time-critical. .onLoad()
# itself fires after every R/ file has been sourced.

.onLoad <- function(libname, pkgname) {

  rds_path <- system.file("extdata", "excel_functions.rds", package = pkgname)

  EXCEL_FUNCTIONS <- c(.spaghetti_env$LEGACY_WORKSHEET,
                       .spaghetti_env$LEGACY_XLM,
                       .spaghetti_env$XLFN,
                       .spaghetti_env$XLWS)

  if (nzchar(rds_path) && file.exists(rds_path)) {
    .spaghetti_env$FUNCTIONS <- readRDS(rds_path)
  } else {
    .spaghetti_env$FUNCTIONS <- data.frame(
      fn          = EXCEL_FUNCTIONS,
      description = NA_character_,
      stringsAsFactors = FALSE
    )
  }

  # Cache the full known-function vector for fast unknown-name checks
  # and spell suggestions.
  .spaghetti_env$ALL_KNOWN <- EXCEL_FUNCTIONS

  # Build per-locale lookup tables for case-insensitive O(1) translation
  # in both directions. Each table is an environment so missing keys
  # return NULL (named character vectors error on missing keys).
  df         <- .spaghetti_env$FUNCTIONS
  loc_cols   <- setdiff(names(df), c("fn", "description", "term_id"))
  loc_to_en  <- list()
  en_to_loc  <- list()
  for (col in loc_cols) {
    vals     <- df[[col]]
    fn_vals  <- df$fn
    not_na   <- !is.na(vals)

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

  invisible(NULL)
}
