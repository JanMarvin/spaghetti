# tests/testthat/helper-terminology.R
#
# Inject a minimal locale lookup table for the duration of the test session.
# Tests can call mock_terminology() in setup, restore_terminology() in
# teardown. Or for simplicity we just install it once at file load and
# restore at session end.
#
# This mirrors what setup_terminology() would build, but with a small
# hand-crafted set of German/French/Spanish translations chosen to cover
# the existing test cases.

local({
  fn  <- c("SUM", "IF", "SUMIF", "SUMIFS", "VLOOKUP", "XLOOKUP", "FILTER",
           "SEQUENCE", "LAMBDA", "LET", "ANCHORARRAY", "SINGLE", "SORT",
           "UNIQUE", "TEXTJOIN", "BYROW", "MAP")
  de  <- c("SUMME", "WENN", "SUMMEWENN", "SUMMEWENNS", "SVERWEIS",
           "XVERWEIS", "FILTER", "SEQUENZ", "LAMBDA", "LET",
           NA, NA, "SORTIEREN", "EINDEUTIG", "TEXTVERKETTEN",
           NA, NA)
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
