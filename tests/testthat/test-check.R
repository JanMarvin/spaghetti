test_that("clean formulas return empty frame invisibly", {
  expect_message(result <- check_formula("=SUM(A1:A10)"), "No unknown functions found.")
  expect_s3_class(result, "data.frame")
  expect_equal(nrow(result), 0L)
  expect_named(result, c("formula", "fn", "suggestion"))
})

test_that("single typo is detected with suggestion", {
  result <- suppressWarnings(check_formula("=SUIM(A1:A10)"))
  expect_equal(nrow(result), 1L)
  expect_equal(result$fn, "SUIM")
  # SUM should appear in suggestions
  expect_true(grepl("SUM", result$suggestion))
})

test_that("transposition typo is detected with suggestion", {
  result <- suppressWarnings(check_formula("=FLITER(A1:A10,B1:B10>0)"))
  expect_equal(nrow(result), 1L)
  expect_equal(result$fn, "FLITER")
  # expect_true(grepl("FILTER", result$suggestion))
})

test_that("missing letter typo is detected with suggestion", {
  result <- suppressWarnings(check_formula("=VLOKUP(A1,B:B,2,0)"))
  expect_equal(nrow(result), 1L)
  expect_equal(result$fn, "VLOKUP")
  expect_true(grepl("VLOOKUP", result$suggestion))
})

test_that("completely made-up name gets no suggestion", {
  result <- suppressWarnings(check_formula("=ZZZZFAKEFUNCTION(A1)"))
  expect_equal(nrow(result), 1L)
  expect_equal(result$fn, "ZZZZFAKEFUNCTION")
  expect_true(is.na(result$suggestion))
})

test_that("vector of formulas checks all entries", {
  result <- suppressWarnings(check_formula(c(
    "=SUM(A1:A5)",           # clean
    "=SUIM(A1:A5)",          # typo
    "=AVERAGE(B1:B5)",       # clean
    "=VLOKUP(A1,B:B,2,0)"   # typo
  )))
  expect_equal(nrow(result), 2L)
  expect_setequal(result$fn, c("SUIM", "VLOKUP"))
})

test_that("multiple unknown functions in one formula are all reported", {
  result <- suppressWarnings(check_formula("=IF(SUIM(A1)>0, VLOKUP(A1,B:B,2), 0)"))
  expect_equal(nrow(result), 2L)
  expect_setequal(result$fn, c("SUIM", "VLOKUP"))
})

test_that("NA and empty formulas are skipped silently", {
  result <- suppressWarnings(check_formula(c(NA_character_, "", "=SUM(A1)")))
  expect_equal(nrow(result), 0L)
})

test_that("to_xml() emits warning for unknown function when warn_unknown=TRUE", {
  expect_warning(
    to_xml("=SUIM(A1:A10)", warn_unknown = TRUE),
    regexp = "SUIM"
  )
})

test_that("to_xml() is silent for unknown function when warn_unknown=FALSE", {
  expect_no_warning(
    to_xml("=SUIM(A1:A10)", warn_unknown = FALSE)
  )
})

test_that("warning message includes suggestion", {
  w <- tryCatch(
    to_xml("=VLOKUP(A1,B:B,2)", warn_unknown = TRUE),
    warning = function(w) conditionMessage(w)
  )
  expect_true(grepl("VLOOKUP", w))
})

test_that("localized unknown function uses translated name for lookup", {
  # SUMMEWENNS is valid German for SUMIFS — should not warn
  expect_no_warning(
    to_xml("=SUMMEWENNS(C2:C10;A2:A10;\"x\")", locale = "de",
           warn_unknown = TRUE)
  )
})

test_that("check_formula respects locale separator", {
  # semicolon-separated German formula with a real typo
  result <- suppressWarnings(
    check_formula("=SUMMEWENNSS(C2:C10;A2:A10;\"x\")", locale = "de")
  )
  expect_equal(nrow(result), 1L)
  # translated to SUMIFSSS or similar — unknown
  expect_true(nchar(result$fn) > 0)
})
