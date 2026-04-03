# R/aaa.R
# This file is intentionally named aaa.R so it is sourced first
# (R loads R/ files in alphabetical order).
# It creates the shared package environment that registry.R and
# locales.R both write into at parse time.

#' @keywords internal
.spaghetti_env <- new.env(parent = emptyenv())
