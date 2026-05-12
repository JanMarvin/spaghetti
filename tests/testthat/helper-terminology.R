# tests/testthat/helper-terminology.R
#
# Local-only mock for the Microsoft Terminology Collection.
#
# Locale-translation tests need a populated function table. Real
# `setup_terminology()` downloads ~100 MB from Microsoft, which is
# unsuitable for CI. This helper installs a small in-memory mock table
# at test session start IFF:
#
#   - the package does not already have terminology loaded
#     (so a developer with a real cache doesn't get their cache
#      shadowed by the test mock), and
#   - we're not on CI (CI runs without locale data so the
#     "no terminology" code paths get exercised).
#
# Tests that require the table call:
#   skip_if_not(has_terminology(), "needs terminology cache")
# so they pass on CI by skipping rather than failing.

.in_ci <- function() {
  isTRUE(as.logical(Sys.getenv("CI", "false")))
}

local({
  if (.in_ci())                          return(invisible(NULL))
  if (spaghetti::has_terminology())      return(invisible(NULL))

  fn  <- c("SUM", "IF", "SUMIF", "SUMIFS", "VLOOKUP", "XLOOKUP", "FILTER",
           "SEQUENCE", "LAMBDA", "LET", "ANCHORARRAY", "SINGLE", "SORT",
           "UNIQUE", "TEXTJOIN", "BYROW", "MAP")
  de  <- c("SUMME", "WENN", "SUMMEWENN", "SUMMEWENNS", "SVERWEIS",
           "XVERWEIS", "FILTER", "SEQUENZ", "LAMBDA", "LET",
           NA, NA, "SORTIEREN", "EINDEUTIG", "TEXTVERKETTEN", NA, NA)
  fr  <- c("SOMME", "SI", "SOMME.SI", "SOMME.SI.ENS", "RECHERCHEV",
           "RECHERCHEX", "FILTRE", "SEQUENCE", "LAMBDA", "LET",
           NA, NA, "TRIER", "UNIQUE", "JOINDRE.TEXTE", NA, NA)
  es  <- c("SUMA", "SI", "SUMAR.SI", "SUMAR.SI.CONJUNTO", "BUSCARV",
           "BUSCARX", "FILTRAR", "SECUENCIA", "LAMBDA", "LET",
           NA, NA, "ORDENAR", "UNICOS", "UNIRCADENAS", NA, NA)

  master <- data.frame(fn = fn, de = de, fr = fr, es = es,
                       stringsAsFactors = FALSE, check.names = FALSE)

  env <- spaghetti:::.spaghetti_env
  env$FUNCTIONS       <- master
  env$has_terminology <- TRUE
  spaghetti:::.build_locale_lookup_tables(master)
})
