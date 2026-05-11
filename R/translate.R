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
#' @param formula Character scalar or vector. Excel formula(s), with or without `=`.
#' @param locale  Two-letter locale code (`"de"`, `"fr"`, …) or NULL.
#'   When set, localised function names are translated to English and the
#'   locale argument separator (`;` for many European locales) is accepted.
#' @param warn_unknown Logical; warn for unknown function names (default TRUE).
#'
#' @return Character scalar or vector: OOXML formula(s) starting with `=`.
#' @export
#' @examples
#' to_xml("=SEQUENCE(10)")
#' to_xml("=LAMBDA(temp, (5/9) * (temp-32))(100)")
#' to_xml("=FILTER(A1:A10, B1:B10 > 5)")
#' to_xml("=SUM(A1#)")
#' to_xml("=LET(tc,(B2-32)*5/9,rh,0.6,tc*ATAN(0.151977*(rh*100+8.313659)^0.5))")
#' to_xml("=SUMMEWENN(A1:A10;\"x\";B1:B10)", locale = "de")
#' to_xml(c("=SUM(A1:A10)", "=SEQUENCE(5)", "=FILTER(A1:A10, B1:B10 > 0)"))
to_xml <- function(formula, locale = NULL, warn_unknown = TRUE) {
  stopifnot(is.character(formula))

  local_sep <- .get_sep(locale)

  vapply(formula, function(f) {
    if (is.na(f) || !nzchar(f)) return(f)

    tokens <- .tokenise(f, sep = local_sep)
    out    <- .transform_to_xml(tokens, locale = locale,
                                warn_unknown = warn_unknown)
    # OOXML always stores "," regardless of locale
    .detokenise(out, prefix_eq = TRUE)
  }, character(1), USE.NAMES = FALSE)
}


#' Convert an OOXML storage formula to user-facing Excel format
#'
#' @param formula Character scalar or vector. OOXML formula(s), with or without `=`.
#' @param locale  Two-letter locale code or NULL. When set, function names are
#'   translated to the target locale and the locale separator is used in output.
#'
#' @return Character scalar or vector: user-facing formula(s) starting with `=`.
#' @export
#' @examples
#' from_xml("=_xlfn.SEQUENCE(10)")
#' from_xml("=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(100)")
#' from_xml("=_xlfn._xlws.FILTER(A1:A10,B1:B10>5)")
#' from_xml("=SUM(_xlfn.ANCHORARRAY(A1))")
#' from_xml("=_xlfn.SEQUENCE(10)", locale = "de")
#' from_xml(c("=_xlfn.SEQUENCE(5)", "=SUM(_xlfn.ANCHORARRAY(A1))"))
from_xml <- function(formula, locale = NULL) {
  stopifnot(is.character(formula))

  local_sep <- .get_sep(locale)

  vapply(formula, function(f) {
    if (is.na(f) || !nzchar(f)) return(f)

    # OOXML storage is always comma-separated
    tokens <- .tokenise(f, sep = ",")
    out    <- .transform_from_xml(tokens, locale = locale,
                                  local_sep = local_sep)
    .detokenise(out, prefix_eq = TRUE)
  }, character(1), USE.NAMES = FALSE)
}


# ── Internal transformation: Excel → OOXML ──────────────────────────────────

#' @keywords internal
.transform_to_xml <- function(tokens, locale, warn_unknown) {
  n   <- length(tokens)
  out <- vector("list", n)
  out_n <- 0L
  i   <- 1L

  push <- function(tok) {
    out_n <<- out_n + 1L
    if (out_n > length(out)) length(out) <<- length(out) * 2L
    out[[out_n]] <<- tok
  }

  # Absolute paren depth of the *input* token stream we've consumed so far.
  paren_depth <- 0L

  # Pending closures for `@FUNC(...)` -> `_xlfn.SINGLE(FUNC(...))`.
  # Each entry is the paren_depth at which the matching extra ')' should fire
  # (i.e. the depth the input was at when we saw `@`).
  pending_singles <- integer(0)

  # Scope stack for LAMBDA and LET.
  # Each entry: list(depth = int, params = char_vec)
  # depth  : paren nesting relative to the opening '(' of this call
  # params : original-case names of all bound identifiers seen so far
  lambda_scope <- list()

  # Peek the value of the next token, or "" if none.
  next_val <- function() if (i + 1L <= n) tokens[[i + 1L]]$val else ""
  next_type <- function() if (i + 1L <= n) tokens[[i + 1L]]$type else ""

  while (i <= n) {
    tok <- tokens[[i]]

    # ── String literals pass through unchanged ──────────────────────────
    if (tok$type == TOKEN_TYPES$STRING) {
      push(tok)
      i <- i + 1L
      next
    }

    # ── Implicit intersection @ ─────────────────────────────────────────
    if (tok$type == TOKEN_TYPES$IMPLICIT) {
      # @FUNC(...) -> _xlfn.SINGLE(FUNC(...))
      if (next_type() == TOKEN_TYPES$FUNC) {
        push(list(type = TOKEN_TYPES$OTHER, val = "_xlfn.SINGLE("))
        pending_singles <- c(pending_singles, paren_depth)
        i <- i + 1L
        next
      }
      # @ref / @ref:ref -> _xlfn.SINGLE(ref)
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
      push(list(type = TOKEN_TYPES$OTHER, val = "_xlfn.SINGLE("))
      push(list(type = TOKEN_TYPES$OTHER, val = ref_str))
      push(list(type = TOKEN_TYPES$OTHER, val = ")"))
      i <- j
      next
    }

    # ── FUNC token ───────────────────────────────────────────────────────
    if (tok$type == TOKEN_TYPES$FUNC) {
      fn_raw <- .strip_prefix(tok$val)
      fn_en  <- toupper(.locale_to_english(fn_raw, locale))
      tier   <- .prefix_for(fn_en)

      # Warn if the function name is not in any known registry tier.
      # .prefix_for() returns "xlfn" as a safe default for unknowns, so we
      # detect them by checking all three sets explicitly.
      if (warn_unknown &&
          !fn_en %in% .spaghetti_env$LEGACY &&
          !fn_en %in% .spaghetti_env$XLFN  &&
          !fn_en %in% .spaghetti_env$XLWS) {
        .warn_unknown_fn(fn_en)
      }

      prefixed <- switch(tier,
                         xlws   = paste0("_xlfn._xlws.", fn_en),
                         xlfn   = paste0("_xlfn.",       fn_en),
                         legacy = fn_en
      )

      push(list(type = TOKEN_TYPES$FUNC, val = prefixed))

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
      raw_name <- .strip_prefix(tok$val)

      # Already registered as a bound name in some enclosing scope?
      is_known_param <- FALSE
      for (s in lambda_scope) {
        if (raw_name %in% s$params) { is_known_param <- TRUE; break }
      }

      # Register as a new bound name only if at depth==1 of the innermost
      # LAMBDA/LET *and* the next token is "," — that's the invariant
      # distinguishing param/name slots from body expressions:
      #   LAMBDA(x, y, body)  -> x,y followed by "," ; body never is
      #   LET(a, 1, b, 2, expr) -> names followed by "," ; expr never is
      if (length(lambda_scope) > 0L) {
        top_idx <- length(lambda_scope)
        if (lambda_scope[[top_idx]]$depth == 1L && next_val() == ",") {
          if (!(raw_name %in% lambda_scope[[top_idx]]$params)) {
            lambda_scope[[top_idx]]$params <-
              c(lambda_scope[[top_idx]]$params, raw_name)
          }
          is_known_param <- TRUE
        }
      }

      if (is_known_param) {
        push(list(type = TOKEN_TYPES$IDENT,
                  val  = paste0("_xlpm.", raw_name)))
      } else {
        push(list(type = TOKEN_TYPES$IDENT, val = raw_name))
      }
      i <- i + 1L
      next
    }

    # ── Anchor # → _xlfn.ANCHORARRAY(prev_ref) ──────────────────────────
    if (tok$type == TOKEN_TYPES$ANCHOR) {
      if (out_n > 0L) {
        last_val <- out[[out_n]]$val
        m <- regmatches(last_val, regexpr("[$A-Za-z]+[$0-9]+$", last_val))
        if (length(m) == 1L && nchar(m) > 0L) {
          stripped <- substring(last_val, 1L, nchar(last_val) - nchar(m))
          if (nchar(stripped) > 0L) {
            out[[out_n]]$val <- stripped
            push(list(type = TOKEN_TYPES$OTHER,
                      val  = paste0("_xlfn.ANCHORARRAY(", m, ")")))
          } else {
            out[[out_n]] <- list(type = TOKEN_TYPES$OTHER,
                                 val  = paste0("_xlfn.ANCHORARRAY(", m, ")"))
          }
        } else {
          push(tok)
        }
      } else {
        push(tok)
      }
      i <- i + 1L
      next
    }

    # ── OTHER: parens, operators, separators ────────────────────────────
    if (tok$type == TOKEN_TYPES$OTHER) {
      val <- tok$val
      if (val == "(") {
        lambda_scope <- .lambda_scope_open(lambda_scope)
        paren_depth  <- paren_depth + 1L
      } else if (val == ")") {
        lambda_scope <- .lambda_scope_close(lambda_scope)
        paren_depth  <- paren_depth - 1L
      }
      # Normalise any separator token to "," for OOXML storage
      if (val == "," || val == ";") {
        push(list(type = TOKEN_TYPES$OTHER, val = ","))
        i <- i + 1L
        next
      }
      push(tok)
      # If this ')' just closed a pending @FUNC SINGLE wrapper, emit an
      # extra ')' to close _xlfn.SINGLE.
      if (val == ")" && length(pending_singles) > 0L &&
          pending_singles[length(pending_singles)] == paren_depth) {
        pending_singles <- pending_singles[-length(pending_singles)]
        push(list(type = TOKEN_TYPES$OTHER, val = ")"))
      }
      i <- i + 1L
      next
    }

    push(tok)
    i <- i + 1L
  }

  if (out_n < length(out)) length(out) <- out_n
  out
}


# ── Internal transformation: OOXML → Excel ───────────────────────────────────

#' @keywords internal
.transform_from_xml <- function(tokens, locale, local_sep = .get_sep(locale)) {
  n         <- length(tokens)
  out       <- vector("list", n)
  out_n     <- 0L
  i         <- 1L

  push <- function(tok) {
    out_n <<- out_n + 1L
    if (out_n > length(out)) length(out) <<- length(out) * 2L
    out[[out_n]] <<- tok
  }

  while (i <= n) {
    tok <- tokens[[i]]

    # String literals pass through
    if (tok$type == TOKEN_TYPES$STRING) {
      push(tok)
      i <- i + 1L
      next
    }

    # FUNC: strip prefix, handle ANCHORARRAY/SINGLE unwrapping, then localise
    if (tok$type == TOKEN_TYPES$FUNC) {
      fn_clean <- .strip_prefix(tok$val)
      fn_up    <- toupper(fn_clean)

      if (fn_up == "ANCHORARRAY" || fn_up == "SINGLE") {
        j <- i + 1L
        if (j <= n && tokens[[j]]$val == "(") {
          j <- j + 1L
          ref_parts <- character(0)
          while (j <= n && tokens[[j]]$val != ")") {
            ref_parts <- c(ref_parts, tokens[[j]]$val)
            j <- j + 1L
          }
          if (j <= n) j <- j + 1L
          ref_str <- paste(ref_parts, collapse = "")
          new_val <- if (fn_up == "ANCHORARRAY") paste0(ref_str, "#") else paste0("@", ref_str)
          push(list(type = TOKEN_TYPES$OTHER, val = new_val))
          i <- j
          next
        }
      }

      fn_out <- if (!is.null(locale)) .english_to_locale(fn_clean, locale) else fn_clean
      push(list(type = TOKEN_TYPES$FUNC, val = fn_out))
      i <- i + 1L
      next
    }

    # IDENT: strip _xlpm. prefix, optionally localise
    if (tok$type == TOKEN_TYPES$IDENT) {
      fn_clean <- .strip_prefix(tok$val)
      fn_out   <- if (!is.null(locale)) .english_to_locale(fn_clean, locale) else fn_clean
      push(list(type = TOKEN_TYPES$IDENT, val = fn_out))
      i <- i + 1L
      next
    }

    # Separator: swap stored "," to locale separator in output
    if (tok$type == TOKEN_TYPES$OTHER && tok$val == ",") {
      push(list(type = TOKEN_TYPES$OTHER, val = local_sep))
      i <- i + 1L
      next
    }

    push(tok)
    i <- i + 1L
  }

  if (out_n < length(out)) length(out) <- out_n
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

  # This list should ideally include most non-English/non-Asian locales
  # that follow the European decimal comma convention.
  semicolon_locales <- c(
    "af", "sq", "am", "ar", "hy", "as", "az", "be", "bs", "bg", "ca", "hr",
    "cs", "da", "nl", "et", "fi", "fr", "gl", "ka", "de", "el", "hu", "is",
    "it", "lv", "lt", "lb", "mk", "no", "pl", "pt", "ro", "ru", "sr", "sk",
    "sl", "es", "sv", "tr", "uk", "vi"
  )

  lang <- tolower(substring(locale, 1, 2))
  if (lang %in% semicolon_locales) ";" else ","
}
