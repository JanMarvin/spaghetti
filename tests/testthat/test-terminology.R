# tests/testthat/test-terminology.R

test_that("has_terminology reflects cached state", {
  expect_true(has_terminology())   # helper installed a mock cache
})

test_that("locale call errors clearly when no terminology is loaded", {
  env <- spaghetti:::.spaghetti_env
  saved_funcs    <- env$FUNCTIONS
  saved_loccols  <- env$LOC_COLS
  saved_loctoen  <- env$LOC_TO_EN
  saved_entoloc  <- env$EN_TO_LOC
  saved_has      <- env$has_terminology

  on.exit({
    env$FUNCTIONS       <- saved_funcs
    env$LOC_COLS        <- saved_loccols
    env$LOC_TO_EN       <- saved_loctoen
    env$EN_TO_LOC       <- saved_entoloc
    env$has_terminology <- saved_has
  })

  spaghetti:::.reset_terminology_state()
  expect_false(has_terminology())

  # No locale: fine
  expect_equal(to_xml("=SUM(A1:A10)"), "=SUM(A1:A10)")

  # Locale requested: error with actionable message
  expect_error(
    to_xml("=SUMME(A1:A10)", locale = "de"),
    "setup_terminology",
    fixed = TRUE
  )
  expect_error(
    from_xml("=SUM(A1:A10)", locale = "de"),
    "setup_terminology",
    fixed = TRUE
  )
})

test_that("function_table falls back to bare fn list when no terminology", {
  env <- spaghetti:::.spaghetti_env
  saved <- env$FUNCTIONS
  on.exit(env$FUNCTIONS <- saved)

  env$FUNCTIONS <- NULL
  tbl <- function_table()
  expect_true("fn" %in% names(tbl))
  expect_true("SUM"     %in% tbl$fn)
  expect_true("XLOOKUP" %in% tbl$fn)
})

test_that("function_table hides term_id but keeps description when present", {
  env <- spaghetti:::.spaghetti_env
  saved <- env$FUNCTIONS
  on.exit(env$FUNCTIONS <- saved)

  # Override the helper's table to include a description column
  env$FUNCTIONS <- data.frame(
    fn          = c("SUM", "XLOOKUP"),
    description = c("Sums args", "Modern lookup"),
    term_id     = c("t1", "t2"),
    de          = c("SUMME", NA_character_),
    stringsAsFactors = FALSE
  )

  tbl <- function_table()
  expect_false("term_id"     %in% names(tbl))
  expect_true("description"  %in% names(tbl))
  expect_true("fn"           %in% names(tbl))
  expect_true("de"           %in% names(tbl))
})

test_that("function_description reads from the FUNCTIONS table when present", {
  env <- spaghetti:::.spaghetti_env
  saved <- env$FUNCTIONS
  on.exit(env$FUNCTIONS <- saved)

  env$FUNCTIONS <- data.frame(
    fn          = c("SUM", "XLOOKUP"),
    description = c("Sums args", "Modern lookup"),
    stringsAsFactors = FALSE
  )

  expect_equal(function_description("SUM"),     "Sums args")
  expect_equal(function_description("XLOOKUP"), "Modern lookup")
  expect_true(is.na(function_description("BOGUS")))
})

test_that("function_description returns NA when no description column", {
  expect_true(all(is.na(function_description(c("SUM", "XLOOKUP")))))
})
