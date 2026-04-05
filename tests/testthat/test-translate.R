library(testthat)

# ── 1. Legacy functions: no prefix ──────────────────────────────────────────
test_that("legacy functions get no prefix", {
  expect_equal(to_xml("=SUM(A1:A10)"),       "=SUM(A1:A10)")
  expect_equal(to_xml("=VLOOKUP(A1,B:C,2)"), "=VLOOKUP(A1,B:C,2)")
  expect_equal(to_xml("=IF(A1>0,1,0)"),      "=IF(A1>0,1,0)")
  expect_equal(to_xml("=AVERAGE(A1:A5)"),    "=AVERAGE(A1:A5)")
  expect_equal(to_xml("=IFERROR(A1,0)"),     "=IFERROR(A1,0)")
})

# ── 2. Future functions: _xlfn. prefix ──────────────────────────────────────
test_that("future functions get _xlfn. prefix", {
  expect_equal(to_xml("=SEQUENCE(10)"),    "=_xlfn.SEQUENCE(10)")
  expect_equal(to_xml("=UNIQUE(A1:A10)"), "=_xlfn.UNIQUE(A1:A10)")
  expect_equal(to_xml("=XLOOKUP(A1,B:B,C:C)"), "=_xlfn.XLOOKUP(A1,B:B,C:C)")
  expect_equal(to_xml("=XMATCH(A1,B:B)"), "=_xlfn.XMATCH(A1,B:B)")
  expect_equal(to_xml("=RANDARRAY(5,2)"), "=_xlfn.RANDARRAY(5,2)")
  expect_equal(to_xml("=LET(X,1,X)"),    "=_xlfn.LET(_xlpm.X,1,_xlpm.X)")
})

# ── 3. Web-service functions: _xlfn._xlws. prefix ───────────────────────────
test_that("FILTER and SORT get _xlfn._xlws. prefix", {
  expect_equal(to_xml("=FILTER(A1:A10,B1:B10>5)"),
               "=_xlfn._xlws.FILTER(A1:A10,B1:B10>5)")
  expect_equal(to_xml("=SORT(A1:A10)"),
               "=_xlfn._xlws.SORT(A1:A10)")
})

# ── 4. LAMBDA with _xlpm. parameter prefixes ────────────────────────────────
test_that("LAMBDA parameters receive _xlpm. prefix", {
  result <- to_xml("=LAMBDA(temp, (5/9) * (temp-32))(100)")
  expect_match(result, "_xlfn.LAMBDA", fixed = TRUE)
  expect_match(result, "_xlpm.temp",  fixed = TRUE)
})

test_that("LAMBDA with multiple parameters", {
  result <- to_xml("=LAMBDA(x, y, x + y)(1, 2)")
  expect_match(result, "_xlpm.x", fixed = TRUE)
  expect_match(result, "_xlpm.y", fixed = TRUE)
})

# ── 5. Spilled range operator # → ANCHORARRAY ───────────────────────────────
test_that("spill operator # wraps in ANCHORARRAY", {
  expect_equal(to_xml("=SUM(A1#)"),
               "=SUM(_xlfn.ANCHORARRAY(A1))")
  expect_equal(to_xml("=COUNTA(B2#)"),
               "=COUNTA(_xlfn.ANCHORARRAY(B2))")
})

# ── 6. Implicit intersection @ → SINGLE ─────────────────────────────────────
test_that("@ operator becomes SINGLE()", {
  result <- to_xml("=@A1:A10")
  expect_match(result, "_xlfn.SINGLE(A1:A10)", fixed = TRUE)
})

# ── 7. Nested future functions ───────────────────────────────────────────────
test_that("nested future functions all get prefixed", {
  result <- to_xml("=UNIQUE(SORT(A1:A10))")
  expect_match(result, "_xlfn.UNIQUE",           fixed = TRUE)
  expect_match(result, "_xlfn._xlws.SORT",       fixed = TRUE)
})

# ── 8. String literals are not modified ─────────────────────────────────────
test_that("string literals pass through unchanged", {
  result <- to_xml('=IF(A1="SEQUENCE",1,0)')
  expect_false(grepl("_xlfn.SEQUENCE", result, fixed = TRUE))
  expect_match(result, '"SEQUENCE"', fixed = TRUE)
})

test_that("escaped quotes in strings are preserved", {
  result <- to_xml('=IF(A1="He said ""hello""","yes","no")')
  expect_match(result, '""hello""', fixed = TRUE)
})

# ── 9. from_xml() reversal ───────────────────────────────────────────────────
test_that("from_xml strips _xlfn. prefix", {
  expect_equal(from_xml("=_xlfn.SEQUENCE(10)"), "=SEQUENCE(10)")
  expect_equal(from_xml("=_xlfn.UNIQUE(A1:A10)"), "=UNIQUE(A1:A10)")
})

test_that("from_xml strips _xlfn._xlws. prefix", {
  expect_equal(from_xml("=_xlfn._xlws.FILTER(A1:A10,B1:B10>5)"),
               "=FILTER(A1:A10,B1:B10>5)")
  expect_equal(from_xml("=_xlfn._xlws.SORT(A1:A10)"), "=SORT(A1:A10)")
})

test_that("from_xml strips _xlpm. prefix from LAMBDA params", {
  result <- from_xml("=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(100)")
  expect_false(grepl("_xlpm.", result, fixed = TRUE))
  expect_false(grepl("_xlfn.", result, fixed = TRUE))
  expect_match(result, "LAMBDA", fixed = TRUE)
})

test_that("from_xml converts ANCHORARRAY back to #", {
  expect_equal(from_xml("=SUM(_xlfn.ANCHORARRAY(A1))"), "=SUM(A1#)")
})

test_that("from_xml converts SINGLE back to @", {
  expect_equal(from_xml("=_xlfn.SINGLE(A1:A10)"), "=@A1:A10")
})

# ── 10. Idempotency ──────────────────────────────────────────────────────────
test_that("to_xml is idempotent (already-prefixed formula unchanged)", {
  xml1 <- to_xml("=SEQUENCE(10)")
  xml2 <- to_xml(xml1)
  expect_equal(xml1, xml2)
})

test_that("from_xml is idempotent (already stripped formula unchanged)", {
  excel1 <- from_xml("=_xlfn.SEQUENCE(10)")
  excel2 <- from_xml(excel1)
  expect_equal(excel1, excel2)
})

# ── 11. Round-trip ───────────────────────────────────────────────────────────
test_that("round_trip() returns consistent xml and excel", {
  rt <- round_trip("=FILTER(A1:A10, B1:B10 > 5)")
  expect_match(rt$xml,   "_xlfn._xlws.FILTER", fixed = TRUE)
  expect_match(rt$locale, "FILTER",              fixed = TRUE)
  expect_false(grepl("_xlfn", rt$locale, fixed = TRUE))
})

# ── 12. Localisation: German ─────────────────────────────────────────────────
test_that("German function names are translated to English in to_xml", {
  result <- to_xml("=SUMME(A1:A10)", locale = "de")
  expect_match(result, "SUM", fixed = TRUE)
  expect_false(grepl("SUMME", result, fixed = TRUE))
})

test_that("from_xml can output German function names", {
  result <- from_xml("=SUM(A1:A10)", locale = "de")
  expect_match(result, "SUMME", fixed = TRUE)
})

test_that("German SVERWEIS becomes VLOOKUP in OOXML", {
  result <- to_xml("=SVERWEIS(A1,B:C,2,0)", locale = "de")
  expect_match(result, "VLOOKUP", fixed = TRUE)
})

# ── 13. Localisation: French ─────────────────────────────────────────────────
test_that("French SOMME is translated correctly", {
  result <- to_xml("=SOMME(A1:A10)", locale = "fr")
  expect_match(result, "SUM", fixed = TRUE)
})

test_that("French RECHERCHEV is translated to VLOOKUP", {
  result <- to_xml("=RECHERCHEV(A1,B:C,2,0)", locale = "fr")
  expect_match(result, "VLOOKUP", fixed = TRUE)
})

# ── 14. Vectorised wrappers ──────────────────────────────────────────────────
test_that("to_xml_v handles a vector of formulas", {
  input  <- c("=SUM(A1:A5)", "=SEQUENCE(10)", "=FILTER(A:A,B:B>0)")
  result <- to_xml_v(input)
  expect_length(result, 3)
  expect_equal(result[1], "=SUM(A1:A5)")
  expect_match(result[2], "_xlfn.SEQUENCE",    fixed = TRUE)
  expect_match(result[3], "_xlfn._xlws.FILTER", fixed = TRUE)
})

test_that("from_xml_v handles a vector of OOXML formulas", {
  input  <- c("=_xlfn.SEQUENCE(10)", "=_xlfn._xlws.FILTER(A:A,B:B>0)")
  result <- from_xml_v(input)
  expect_equal(result[1], "=SEQUENCE(10)")
  expect_equal(result[2], "=FILTER(A:A,B:B>0)")
})

# ── 15. function_prefix() utility ────────────────────────────────────────────
test_that("function_prefix returns correct tiers", {
  expect_equal(function_prefix("SUM"),      "legacy")
  expect_equal(function_prefix("SEQUENCE"), "xlfn")
  expect_equal(function_prefix("FILTER"),   "xlws")
  expect_equal(function_prefix("LAMBDA"),   "xlfn")
})

# ── 16. is_ooxml() utility ───────────────────────────────────────────────────
test_that("is_ooxml correctly identifies prefixed formulas", {
  expect_true( is_ooxml("=_xlfn.SEQUENCE(10)"))
  expect_false(is_ooxml("=SEQUENCE(10)"))
  expect_true( is_ooxml("=_xlfn.LAMBDA(_xlpm.x, x)"))
  expect_false(is_ooxml("=SUM(A1:A10)"))
})

# ── 17. supported_locales() ──────────────────────────────────────────────────
test_that("supported_locales returns expected codes", {
  locs <- supported_locales()
  expect_true("de" %in% locs)
  expect_true("fr" %in% locs)
  expect_true("es" %in% locs)
})

# ── 18. Formula without leading '=' ──────────────────────────────────────────
test_that("formulas without '=' are handled gracefully", {
  expect_equal(to_xml("SUM(A1:A10)"), "=SUM(A1:A10)")
  expect_match(to_xml("SEQUENCE(10)"), "_xlfn.SEQUENCE", fixed = TRUE)
})

# ── 19. MAP / REDUCE / BYROW with LAMBDA args ────────────────────────────────
test_that("MAP with LAMBDA gets correct prefixes", {
  result <- to_xml("=MAP(A1:A5, LAMBDA(x, x*2))")
  expect_match(result, "_xlfn.MAP",    fixed = TRUE)
  expect_match(result, "_xlfn.LAMBDA", fixed = TRUE)
})

test_that("BYROW with LAMBDA gets correct prefixes", {
  result <- to_xml("=BYROW(A1:C3, LAMBDA(row, SUM(row)))")
  expect_match(result, "_xlfn.BYROW",  fixed = TRUE)
  expect_match(result, "_xlfn.LAMBDA", fixed = TRUE)
})

# ── 20. Complex real-world formula ───────────────────────────────────────────
test_that("complex nested formula is handled correctly", {
  f <- "=LET(data, FILTER(A1:D10, C1:C10>100), UNIQUE(SORT(data)))"
  result <- to_xml(f)
  expect_match(result, "_xlfn.LET",           fixed = TRUE)
  expect_match(result, "_xlfn._xlws.FILTER",  fixed = TRUE)
  expect_match(result, "_xlfn.UNIQUE",         fixed = TRUE)
  expect_match(result, "_xlfn._xlws.SORT",     fixed = TRUE)
})
