# R/translate.R
# Core translation functions.
#
# Key OOXML namespace rules:
#   1. _xlfn.       : future functions (post-2010)
#   2. _xlfn._xlws. : worksheet-scope dynamic arrays (FILTER, SORT)
#   3. _xlpm.       : LAMBDA / LET parameter and variable names
#   4. _xlfn.ANCHORARRAY(ref) : spilled range operator  ref#
#   5. _xlfn.SINGLE(ref)      : implicit intersection   @ref
#
# Separator handling:
#   OOXML storage always uses "," as the argument separator.
#   Localised front-ends (e.g. German) use ";" because "," is the decimal
#   separator. to_xml() accepts the local separator and normalises to ","
#   in output. from_xml() reads "," and emits the local separator when
#   locale is set.
#
# Case rules:
#   - FUNC tokens (function names) are always uppercased — they live in the
#     registry and OOXML requires uppercase names.
#   - IDENT tokens (LAMBDA params, LET variable names, named ranges) preserve
#     their original case. _xlpm. names are case-preserving.
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


#' Convert a user-facing formula to OOXML storage format
#'
#' @param formula Character scalar or vector. Formula(s), with or without `=`.
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
#' \dontrun{to_xml("=SUMMEWENN(A1:A10;\"x\";B1:B10)", locale = "de")}
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


#' Convert an OOXML storage formula to user-facing format
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
#' \dontrun{from_xml("=_xlfn.SEQUENCE(10)", locale = "de")}
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


# ── Internal transformation: user-facing → OOXML ────────────────────────────

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

  # Helpers for skipping whitespace OTHER tokens around operators that
  # bind to a ref. Excel/spreadsheet UIs are tolerant of spaces like
  # `A1 #` and `@ A1` — these should bind as if there was no space.
  is_ws <- function(tok) {
    tok$type == TOKEN_TYPES$OTHER && grepl("^\\s+$", tok$val, perl = TRUE)
  }
  # Index of the next non-whitespace token in `tokens` starting from k,
  # or NA if none.
  skip_ws_forward <- function(k) {
    while (k <= n && is_ws(tokens[[k]])) k <- k + 1L
    if (k > n) NA_integer_ else k
  }
  # Index of the last non-whitespace token in `out` ending at out_n,
  # or NA if none.
  skip_ws_backward_out <- function() {
    k <- out_n
    while (k >= 1L && is_ws(out[[k]])) k <- k - 1L
    if (k < 1L) NA_integer_ else k
  }

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
      # Find the next significant token, skipping any whitespace between
      # @ and its target.
      k <- skip_ws_forward(i + 1L)
      tgt_type <- if (is.na(k)) "" else tokens[[k]]$type

      # @FUNC(...) -> _xlfn.SINGLE(FUNC(...))
      if (tgt_type == TOKEN_TYPES$FUNC) {
        push(list(type = TOKEN_TYPES$OTHER, val = "_xlfn.SINGLE("))
        pending_singles <- c(pending_singles, paren_depth)
        i <- k
        next
      }
      # @ref -> _xlfn.SINGLE(ref). REF tokens already cover Sheet!A1,
      # 'My Sheet'!A1, A1:B10, $A$1, etc. as a single token.
      if (tgt_type == TOKEN_TYPES$REF) {
        ref_val <- tokens[[k]]$val
        push(list(type = TOKEN_TYPES$OTHER,
                  val  = paste0("_xlfn.SINGLE(", ref_val, ")")))
        i <- k + 1L
        next
      }
      # @ident -> treat as ref-like (named range etc.)
      if (tgt_type == TOKEN_TYPES$IDENT) {
        ref_val <- tokens[[k]]$val
        push(list(type = TOKEN_TYPES$OTHER,
                  val  = paste0("_xlfn.SINGLE(", ref_val, ")")))
        i <- k + 1L
        next
      }
      # Standalone @ that we don't know how to wrap — emit verbatim
      push(tok)
      i <- i + 1L
      next
    }

    # ── FUNC token ───────────────────────────────────────────────────────
    if (tok$type == TOKEN_TYPES$FUNC) {
      fn_raw <- .strip_prefix(tok$val)
      fn_en  <- toupper(.locale_to_english(fn_raw, locale))
      tier   <- .prefix_for(fn_en)

      # Warn if the function name is not in any known registry tier.
      # .prefix_for() returns "xlfn" as a safe default for unknowns, so we
      # detect them via the cached union of all registry tiers.
      if (warn_unknown && !fn_en %in% .spaghetti_env$ALL_KNOWN) {
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
      # Look back past any whitespace OTHER tokens (e.g. user typed
      # "A1 #" with a space).
      ref_idx <- skip_ws_backward_out()
      if (!is.na(ref_idx)) {
        last <- out[[ref_idx]]
        if (last$type == TOKEN_TYPES$REF) {
          # Drop any whitespace tokens that sat between the REF and #.
          if (ref_idx < out_n) {
            # Keep only out[1..ref_idx]; the slot will be overwritten.
            out_n <- ref_idx
          }
          out[[out_n]] <- list(
            type = TOKEN_TYPES$OTHER,
            val  = paste0("_xlfn.ANCHORARRAY(", last$val, ")")
          )
        } else {
          # Fallback: pull a trailing cell-ref shape from the last token's
          # value (handles e.g. an IDENT that contained a ref-like suffix).
          m <- regmatches(last$val, regexpr("[$A-Za-z]+[$0-9]+$", last$val))
          if (length(m) == 1L && nchar(m) > 0L) {
            stripped <- substring(last$val, 1L, nchar(last$val) - nchar(m))
            if (ref_idx < out_n) out_n <- ref_idx
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


# ── Internal transformation: OOXML → user-facing ────────────────────────────

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
        # Find the inner token range between matching parens.
        if (i + 1L <= n && tokens[[i + 1L]]$val == "(") {
          inner_start <- i + 2L
          j           <- inner_start
          depth       <- 1L
          while (j <= n && depth > 0L) {
            v <- tokens[[j]]$val
            if (v == "(") depth <- depth + 1L
            else if (v == ")") depth <- depth - 1L
            if (depth > 0L) j <- j + 1L
          }
          inner_end <- j - 1L

          # Recursively transform the inner tokens so nested prefixes get
          # stripped/localised correctly.
          inner_toks <- if (inner_start <= inner_end) {
            .transform_from_xml(tokens[inner_start:inner_end],
                                locale = locale, local_sep = local_sep)
          } else {
            list()
          }
          inner_val <- paste(vapply(inner_toks, `[[`, character(1), "val"),
                             collapse = "")
          new_val <- if (fn_up == "ANCHORARRAY")
            paste0(inner_val, "#")
          else
            paste0("@", inner_val)
          push(list(type = TOKEN_TYPES$OTHER, val = new_val))
          i <- j + 1L
          next
        }
      }

      fn_out <- if (!is.null(locale)) .english_to_locale(fn_clean, locale) else fn_clean
      push(list(type = TOKEN_TYPES$FUNC, val = fn_out))
      i <- i + 1L
      next
    }

    # IDENT: strip _xlpm. prefix only. Do NOT localise — IDENT tokens are
    # user-bound names (LAMBDA params, LET variables, named ranges) that
    # happen to be valid identifiers; passing them through .english_to_locale
    # would mistranslate any name that collides with a function name.
    if (tok$type == TOKEN_TYPES$IDENT) {
      push(list(type = TOKEN_TYPES$IDENT, val = .strip_prefix(tok$val)))
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
#' Locales using "," as their decimal separator use ";" in formulas. The
#' list lives in `.spaghetti_env$SEMICOLON_LOCALES` (see R/aaa.R).
#'
#' @param locale Locale code or NULL.
#' @return ";" or ","
#' @keywords internal
.get_sep <- function(locale) {
  if (is.null(locale)) return(",")
  lang <- tolower(substring(locale, 1, 2))
  if (lang %in% .spaghetti_env$SEMICOLON_LOCALES) ";" else ","
}
