# R/aaa.R
# This file is intentionally named aaa.R so it is sourced first
# (R loads R/ files in alphabetical order).
# It creates the shared package environment that registry.R and
# locales.R both write into at parse time.

#' @keywords internal
.spaghetti_env <- new.env(parent = emptyenv())

# Locale codes whose formula syntax uses ";" as the argument separator
# (because "," is the decimal separator). Lower-cased; matched against
# tolower(substring(locale, 1, 2)) so multi-segment codes like "de-DE"
# resolve correctly.
.spaghetti_env$SEMICOLON_LOCALES <- c(
  "af", "sq", "am", "ar", "hy", "as", "az", "be", "bs", "bg", "ca", "hr",
  "cs", "da", "nl", "et", "fi", "fr", "gl", "ka", "de", "el", "hu", "is",
  "it", "lv", "lt", "lb", "mk", "no", "pl", "pt", "ro", "ru", "sr", "sk",
  "sl", "es", "sv", "tr", "uk", "vi"
)
