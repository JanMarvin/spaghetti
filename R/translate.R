# R/translate.R
# Core translation functions.
#
# Key OOXML namespace rules:
#   1. _xlfn.       : future functions (post-Excel 2010)
#   2. _xlfn._xlws. : worksheet-scope dynamic arrays (FILTER, SORT)
#   3. _xlpm.       : LAMBDA / LET parameter and variable names
#   4. _xlfn.ANCHORARRAY(ref) : spilled range operator  ref#
#   5. _xlfn.SINGLE(ref)      : implicit intersection   @ref
#
# Separator handling:
#   OOXML storage always uses "," as the argument separator.
#   Localised Excel (e.g. German) uses ";" because "," is the decimal separator.
#   to_xml() accepts the local separator and normalises to "," in output.
#   from_xml() reads "," and emits the local separator when locale is set.
#
# Case rules:
#   - FUNC tokens (function names) are always uppercased — they live in the
#     registry and OOXML requires uppercase names.
#   - IDENT tokens (LAMBDA params, LET variable names, named ranges) preserve
#     their original case. Excel is case-preserving for _xlpm. names.
#     e.g. =LAMBDA(number, number+1)  ->  _xlfn.LAMBDA(_xlpm.number, _xlpm.number+1)
#
# LAMBDA / LET scoping:
#   Both functions bind named identifiers that must receive _xlpm. prefixes.
#   We track open LAMBDA/LET scopes as a stack. Each entry records:
#     depth  : paren depth relative to the function call opening paren
#     params : character vector of bound names (original case) seen so far
#   Any IDENT seen at depth == 1 inside a LAMBDA or LET is registered as a
#   bound name. Once registered, the name is recognised anywhere in the same
#   scope including inside nested calls in the body.


#' Convert a user-facing Excel formula to OOXML storage format
#'
#' @param formula Character scalar. The Excel formula, with or without `=`.
#' @param locale  Two-letter locale code (`"de"`, `"fr"`, …) or NULL.
#'   When set, localised function names are translated to English and the
#'   locale argument separator (`;` for many European locales) is accepted.
#' @param warn_unknown Logical; warn for unknown function names (default TRUE).
#'
#' @return Character scalar: OOXML formula starting with `=`.
#' @export
#' @examples
#' to_xml("=SEQUENCE(10)")
#' to_xml("=LAMBDA(temp, (5/9) * (temp-32))(100)")
#' to_xml("=FILTER(A1:A10, B1:B10 > 5)")
#' to_xml("=SUM(A1#)")
#' to_xml("=LET(tc,(B2-32)*5/9,rh,0.6,tc*ATAN(0.151977*(rh*100+8.313659)^0.5))")
#' to_xml("=SUMMEWENN(A1:A10;\"x\";B1:B10)", locale = "de")
to_xml <- function(formula, locale = NULL, warn_unknown = TRUE) {
  stopifnot(is.character(formula), length(formula) == 1)
  if (is.null(formula) || is.na(formula) || formula == "") return(formula)

  local_sep <- .get_sep(locale)
  tokens    <- .tokenise(formula, sep = local_sep)
  out       <- .transform_to_xml(tokens, locale = locale,
                                 warn_unknown = warn_unknown)
  # OOXML always stores "," regardless of locale
  .detokenise(out, prefix_eq = TRUE, sep = ",")
}


#' Convert an OOXML storage formula to user-facing Excel format
#'
#' @param formula Character scalar. The OOXML formula, with or without `=`.
#' @param locale  Two-letter locale code or NULL. When set, function names are
#'   translated to the target locale and the locale separator is used in output.
#'
#' @return Character scalar: user-facing formula starting with `=`.
#' @export
#' @examples
#' from_xml("=_xlfn.SEQUENCE(10)")
#' from_xml("=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(100)")
#' from_xml("=_xlfn._xlws.FILTER(A1:A10,B1:B10>5)")
#' from_xml("=SUM(_xlfn.ANCHORARRAY(A1))")
#' from_xml("=_xlfn.SEQUENCE(10)", locale = "de")
from_xml <- function(formula, locale = NULL) {
  stopifnot(is.character(formula), length(formula) == 1)
  if (is.null(formula) || is.na(formula) || formula == "") return(formula)

  # OOXML storage is always comma-separated
  tokens    <- .tokenise(formula, sep = ",")
  out       <- .transform_from_xml(tokens, locale = locale)

  # Emit using the locale separator (e.g. ";" for German)
  local_sep <- .get_sep(locale)
  .detokenise(out, prefix_eq = TRUE, sep = local_sep)
}


# ── Internal transformation: Excel → OOXML ──────────────────────────────────

#' @keywords internal
.transform_to_xml <- function(tokens, locale, warn_unknown) {
  n   <- length(tokens)
  out <- list()
  i   <- 1L

  # Scope stack for LAMBDA and LET.
  # Each entry: list(depth = int, params = char_vec)
  # depth  : paren nesting relative to the opening '(' of this call
  # params : original-case names of all bound identifiers seen so far
  lambda_scope <- list()

  while (i <= n) {
    tok <- tokens[[i]]

    # ── String literals pass through unchanged ──────────────────────────
    if (tok$type == TOKEN_TYPES$STRING) {
      out <- c(out, list(tok))
      i   <- i + 1L
      next
    }

    # ── Implicit intersection @ → _xlfn.SINGLE( ... ) ──────────────────
    if (tok$type == TOKEN_TYPES$IMPLICIT) {
      ref_tokens <- list()
      j <- i + 1L
      while (j <= n) {
        t <- tokens[[j]]
        if (t$type %in% c(TOKEN_TYPES$IDENT, TOKEN_TYPES$REF)) {
          ref_tokens <- c(ref_tokens, list(t))
          j <- j + 1L
          if (j <= n && tokens[[j]]$val == ":") {
            ref_tokens <- c(ref_tokens, list(tokens[[j]]))
            j <- j + 1L
          } else {
            break
          }
        } else {
          break
        }
      }
      ref_str <- paste(vapply(ref_tokens, `[[`, character(1), "val"),
                       collapse = "")
      out <- c(out, list(
        list(type = TOKEN_TYPES$OTHER, val = "_xlfn.SINGLE("),
        list(type = TOKEN_TYPES$OTHER, val = ref_str),
        list(type = TOKEN_TYPES$OTHER, val = ")")
      ))
      i <- j
      next
    }

    # ── FUNC token ───────────────────────────────────────────────────────
    if (tok$type == TOKEN_TYPES$FUNC) {
      fn_raw <- .strip_prefix(tok$val)
      fn_en  <- toupper(.locale_to_english(fn_raw, locale))
      tier   <- .prefix_for(fn_en)

      prefixed <- switch(tier,
                         xlws   = paste0("_xlfn._xlws.", fn_en),
                         xlfn   = paste0("_xlfn.",       fn_en),
                         legacy = fn_en
      )

      out <- c(out, list(list(type = TOKEN_TYPES$FUNC, val = prefixed)))

      # Both LAMBDA and LET bind identifier names that need _xlpm. prefixes.
      # Push a new scope entry; depth starts at 0 and becomes 1 when '(' fires.
      if (fn_en %in% c("LAMBDA", "LET")) {
        lambda_scope <- c(lambda_scope,
                          list(list(depth = 0L, params = character(0))))
      }
      i <- i + 1L
      next
    }

    # ── IDENT: LAMBDA/LET parameter or variable name ─────────────────────
    if (tok$type == TOKEN_TYPES$IDENT) {
      raw_name <- .strip_prefix(tok$val)   # preserve original case

      # Check if already registered as a bound name from any enclosing scope
      is_known_param <- any(vapply(
        lambda_scope,
        function(s) raw_name %in% s$params,
        logical(1)
      ))

      # If at depth == 1 of the innermost LAMBDA/LET, register this as a
      # new bound name (original case preserved)
      if (length(lambda_scope) > 0) {
        top_idx <- length(lambda_scope)
        if (lambda_scope[[top_idx]]$depth == 1L) {
          lambda_scope[[top_idx]]$params <-
            unique(c(lambda_scope[[top_idx]]$params, raw_name))
          is_known_param <- TRUE
        }
      }

      if (is_known_param) {
        # _xlpm. prefix + preserve original case
        out <- c(out, list(list(type = TOKEN_TYPES$IDENT,
                                val  = paste0("_xlpm.", raw_name))))
      } else {
        # Regular identifier (named range etc.) — pass through as-is
        out <- c(out, list(list(type = TOKEN_TYPES$IDENT, val = raw_name)))
      }
      i <- i + 1L
      next
    }

    # ── Anchor # → _xlfn.ANCHORARRAY(prev_ref) ──────────────────────────
    if (tok$type == TOKEN_TYPES$ANCHOR) {
      if (length(out) > 0) {
        last_tok <- out[[length(out)]]
        last_val <- last_tok$val
        m <- regmatches(last_val,
                        regexpr("[$A-Za-z]+[$0-9]+$", last_val))
        if (length(m) == 1 && nchar(m) > 0) {
          stripped <- substring(last_val, 1, nchar(last_val) - nchar(m))
          if (nchar(stripped) > 0) {
            out[[length(out)]]$val <- stripped
          } else {
            out <- out[-length(out)]
          }
          out <- c(out, list(list(
            type = TOKEN_TYPES$OTHER,
            val  = paste0("_xlfn.ANCHORARRAY(", m, ")")
          )))
        } else {
          out <- c(out, list(tok))
        }
      } else {
        out <- c(out, list(tok))
      }
      i <- i + 1L
      next
    }

    # ── OTHER: parens, operators, separators ────────────────────────────
    if (tok$type == TOKEN_TYPES$OTHER) {
      val <- tok$val
      if (val == "(") {
        lambda_scope <- .lambda_scope_open(lambda_scope)
      } else if (val == ")") {
        lambda_scope <- .lambda_scope_close(lambda_scope)
      }
      # Normalise any separator token to "," for OOXML storage
      if (val == "," || val == ";") {
        out <- c(out, list(list(type = TOKEN_TYPES$OTHER, val = ",")))
        i   <- i + 1L
        next
      }
    }

    out <- c(out, list(tok))
    i   <- i + 1L
  }

  out
}


# ── Internal transformation: OOXML → Excel ───────────────────────────────────

#' @keywords internal
.transform_from_xml <- function(tokens, locale) {
  n         <- length(tokens)
  out       <- list()
  i         <- 1L
  local_sep <- .get_sep(locale)

  while (i <= n) {
    tok <- tokens[[i]]

    # String literals pass through
    if (tok$type == TOKEN_TYPES$STRING) {
      out <- c(out, list(tok))
      i   <- i + 1L
      next
    }

    # FUNC: strip prefix, handle ANCHORARRAY/SINGLE unwrapping, then localise
    if (tok$type == TOKEN_TYPES$FUNC) {
      fn_clean <- .strip_prefix(tok$val)

      if (toupper(fn_clean) == "ANCHORARRAY") {
        j <- i + 1L
        if (j <= n && tokens[[j]]$val == "(") {
          j <- j + 1L
          ref_parts <- list()
          while (j <= n && tokens[[j]]$val != ")") {
            ref_parts <- c(ref_parts, list(tokens[[j]]))
            j <- j + 1L
          }
          if (j <= n) j <- j + 1L
          ref_str <- paste(vapply(ref_parts, `[[`, character(1), "val"),
                           collapse = "")
          out <- c(out, list(list(type = TOKEN_TYPES$OTHER,
                                  val  = paste0(ref_str, "#"))))
          i <- j
          next
        }
      }

      if (toupper(fn_clean) == "SINGLE") {
        j <- i + 1L
        if (j <= n && tokens[[j]]$val == "(") {
          j <- j + 1L
          ref_parts <- list()
          while (j <= n && tokens[[j]]$val != ")") {
            ref_parts <- c(ref_parts, list(tokens[[j]]))
            j <- j + 1L
          }
          if (j <= n) j <- j + 1L
          ref_str <- paste(vapply(ref_parts, `[[`, character(1), "val"),
                           collapse = "")
          out <- c(out, list(list(type = TOKEN_TYPES$OTHER,
                                  val  = paste0("@", ref_str))))
          i <- j
          next
        }
      }

      fn_out <- if (!is.null(locale)) .english_to_locale(fn_clean, locale) else fn_clean
      out <- c(out, list(list(type = TOKEN_TYPES$FUNC, val = fn_out)))
      i   <- i + 1L
      next
    }

    # IDENT: strip _xlpm. prefix, optionally localise
    if (tok$type == TOKEN_TYPES$IDENT) {
      fn_clean <- .strip_prefix(tok$val)
      fn_out   <- if (!is.null(locale)) .english_to_locale(fn_clean, locale) else fn_clean
      out <- c(out, list(list(type = TOKEN_TYPES$IDENT, val = fn_out)))
      i   <- i + 1L
      next
    }

    # Separator: swap stored "," to locale separator in output
    if (tok$type == TOKEN_TYPES$OTHER && tok$val == ",") {
      out <- c(out, list(list(type = TOKEN_TYPES$OTHER, val = local_sep)))
      i   <- i + 1L
      next
    }

    out <- c(out, list(tok))
    i   <- i + 1L
  }

  out
}


# ── Scope helpers ────────────────────────────────────────────────────────────

#' Increase paren depth of the innermost LAMBDA/LET scope.
#' @keywords internal
.lambda_scope_open <- function(scope) {
  if (length(scope) == 0) return(scope)
  idx <- length(scope)
  scope[[idx]]$depth <- scope[[idx]]$depth + 1L
  scope
}

#' Decrease paren depth; pop the scope when depth reaches 0.
#' @keywords internal
.lambda_scope_close <- function(scope) {
  if (length(scope) == 0) return(scope)
  idx <- length(scope)
  scope[[idx]]$depth <- scope[[idx]]$depth - 1L
  if (scope[[idx]]$depth == 0L) scope <- scope[-idx]
  scope
}


# ── Separator helper ─────────────────────────────────────────────────────────

#' Return the formula argument separator for a given locale.
#'
#' Locales that use "," as their decimal separator use ";" in formulas.
#'
#' @param locale Two-letter locale code or NULL.
#' @return ";" or ","
#' @keywords internal
.get_sep <- function(locale) {
  if (is.null(locale)) return(",")
  semicolon_locales <- c("de", "fr", "it", "es", "pt", "nl",
                         "ru", "da", "fi", "sv", "pl", "cs",
                         "sk", "hu", "ro", "hr", "bg", "el")
  lang <- tolower(substring(locale, 1, 2))
  if (lang %in% semicolon_locales) ";" else ","
}
