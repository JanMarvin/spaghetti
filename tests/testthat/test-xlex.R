test_that("xlex returns a data frame invisibly", {
  df <- xlex("=SUM(A1:A10)", print = FALSE)
  expect_s3_class(df, "data.frame")
  expect_named(df, c("depth", "val", "label"))
  expect_true(nrow(df) > 0L)
})

test_that("xlex identifies function token", {
  df <- xlex("=SUM(A1:A10)", print = FALSE)
  expect_true(any(df$label == "function" & df$val == "SUM"))
})

test_that("xlex identifies text token", {
  df <- xlex('=IF(A1="hello",1,0)', print = FALSE)
  expect_true(any(df$label == "text"))
})

test_that("xlex identifies separator token", {
  df <- xlex("=SUM(A1,A2)", print = FALSE)
  expect_true(any(df$label == "separator"))
})

test_that("xlex identifies fun_open and fun_close tokens", {
  df <- xlex("=SUM(A1:A10)", print = FALSE)
  expect_true(any(df$label == "fun_open"))
  expect_true(any(df$label == "fun_close"))
})

test_that("xlex nests arguments inside function (depth)", {
  df <- xlex("=SUM(A1:A10)", print = FALSE)
  fn_depth  <- df$depth[df$label == "function"][1L]
  arg_depth <- df$depth[df$val == "A1:A10"][1L]
  # Arguments should be deeper than the function itself
  expect_true(arg_depth > fn_depth)
})

test_that("xlex handles nested functions", {
  df <- xlex("=IF(A1>0,SUM(B1:B5),0)", print = FALSE)
  fns <- df$val[df$label == "function"]
  expect_true("IF" %in% fns)
  expect_true("SUM" %in% fns)
  # SUM should be deeper than IF
  if_depth  <- df$depth[df$label == "function" & df$val == "IF"][1L]
  sum_depth <- df$depth[df$label == "function" & df$val == "SUM"][1L]
  expect_true(sum_depth > if_depth)
})

test_that("xlex handles locale separator", {
  df <- xlex("=SUMME(A1:A5;B1:B5)", locale = "de", print = FALSE)
  expect_true(any(df$label == "separator" & df$val == ";"))
})

test_that("xlex handles OOXML-prefixed formula from to_xml()", {
  ooxml <- to_xml("=FILTER(A1:A10,B1:B10>0)", warn_unknown = FALSE)
  df    <- xlex(ooxml, print = FALSE)
  expect_true(any(df$label == "function"))
  fns <- df$val[df$label == "function"]
  expect_true(any(grepl("FILTER", fns)))
})

test_that("xlex handles empty/NA formula gracefully", {
  df <- xlex("", print = FALSE)
  expect_equal(nrow(df), 0L)
})

test_that("xlex prints to console without error", {
  expect_output(xlex("=SUM(A1:A10)"), regexp = "root")
  expect_output(xlex("=SUM(A1:A10)"), regexp = "function")
  expect_output(xlex("=SUM(A1:A10)"), regexp = "SUM")
})

test_that("xlex handles LAMBDA with parameters", {
  df <- xlex("=LAMBDA(x, x*2)(5)", print = FALSE)
  fns <- df$val[df$label == "function"]
  expect_true("LAMBDA" %in% fns)
})

test_that("xlex handles anchor # operator", {
  df <- xlex("=SUM(A1#)", print = FALSE)
  expect_true(any(df$label == "operator" & df$val == "#"))
})

test_that("xlex handles implicit intersection @ operator", {
  df <- xlex("=@A1:A10", print = FALSE)
  expect_true(any(df$label == "operator" & df$val == "@"))
})

test_that("xlex labels intersection (space between two refs) distinctly", {
  df <- xlex("=SUM(A1:B10 A5:D5)", print = FALSE)
  expect_true(any(df$label == "intersection"))
})

test_that("xlex labels error literals as 'error'", {
  df <- xlex("=IFERROR(VLOOKUP(A1,B:C,2,0),#N/A)", print = FALSE)
  expect_true(any(df$label == "error" & df$val == "#N/A"))
})

test_that("xlex labels array literals as 'array'", {
  df <- xlex("=SUM({1,2;3,4})", print = FALSE)
  expect_true(any(df$label == "array"))
})
