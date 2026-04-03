# R/zzz.R
# Package initialisation.
#
# R loads package files in alphabetical order, which means:
#   lexer.R -> locales.R -> registry.R -> spaghetti-package.R -> translate.R -> utils.R -> zzz.R
#
# locales.R references .spaghetti_env (created in registry.R), and
# locales.R also builds LOCALES_REV from LOCALES at parse time, both of
# which are unsafe to rely on without an explicit init sequence.
#
# .onLoad() runs after all R/ files are sourced, so by the time it fires
# the environment exists. We use it to (re-)build any derived structures
# that depend on cross-file state.

.onLoad <- function(libname, pkgname) {

  # 1. Ensure the package environment exists (defensive)
  if (!exists(".spaghetti_env", envir = parent.env(environment()))) {
    assign(".spaghetti_env", new.env(parent = emptyenv()),
           envir = parent.env(environment()))
  }

  # 2. Build reverse locale maps (English -> locale) from the forward maps.
  #    This is safe here because LOCALES is populated by locales.R which has
  #    already been sourced before .onLoad fires.
  .spaghetti_env$LOCALES_REV <- lapply(.spaghetti_env$LOCALES, function(m) {
    rev_m <- list()
    for (loc_name in names(m)) {
      en_name <- m[[loc_name]]
      # Keep first mapping if duplicates exist (e.g. ISNA / ISND both -> ISNA)
      if (is.null(rev_m[[en_name]])) rev_m[[en_name]] <- loc_name
    }
    rev_m
  })

  invisible(NULL)
}
